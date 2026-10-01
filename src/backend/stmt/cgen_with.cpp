#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_with.cpp: With 块语句生成 ---

void CCodeGen::visit(WithStmt& node) {
    // P17.1: With块 — 根据对象类型创建适当类型的临时变量
    WithObjInfo withInfo;
    std::string tempVar = "_vb6_with_" + std::to_string(tempCounter_++);
    std::string tempType = "void*";

    // 检测With对象类型: 遍历表达式判断
    if (node.object) {
        // --- Fix 086: 方法调用链对象推断 (With .Sql(...) / With obj.Method(...)) ---
        // With 对象是 IndexOrCallExpr/MemberAccessExpr/WithMemberExpr (方法调用链) 时,
        // 用 inferClassTypeOfExpr 推断返回类. 此前 `With .Sql(...)` 解析为
        // IndexOrCallExpr (非裸 WithMemberExpr), 嵌套With分支不命中 → tempType void*
        // → 成员解析回退全局 lookupModule 捡错符号 (如 .Fetch 命中模块级 Sub Fetch,
        // 生成 vb6_Demo_Fetch(_vb6_with_N) 这类带多余参数的非法调用).
        // 仅接受项目 Class (排除 COM 类/接口/UDT), 其余情况走原有分支.
        if (node.object->kind == ASTNodeKind::IndexOrCallExpr
            || node.object->kind == ASTNodeKind::MemberAccessExpr
            || node.object->kind == ASTNodeKind::WithMemberExpr) {
            std::string inferred = inferClassTypeOfExpr(*node.object);
            if (!inferred.empty()) {
                const Symbol* clsSym = lookupModuleDotted(inferred);
                if (clsSym && clsSym->kind == SymbolKind::Class && !clsSym->isInterface) {
                    withInfo.kind = WithObjKind::ClassInstance;
                    withInfo.className = cIdent(inferred);
                    tempType = "vb6_cls_" + withInfo.className + "*";
                    withInfo.ctrlOrigName = withInfo.className;
                }
            }
        }
        // --- Fix 010n: New表达式检测 (With New ClassName) ---
        if (node.object->kind == ASTNodeKind::NewExpr) {
            auto& newExpr = static_cast<NewExpr&>(*node.object);
            std::string clsLower = newExpr.className;
            std::transform(clsLower.begin(), clsLower.end(), clsLower.begin(), ::tolower);
            // Fix 086: COM类检测 — With New ADODB.Stream 等 COM 类型没有项目
            // Class 符号 (ComClass/ComInterface 或未注册), 不能按项目类生成
            // vb6_cls_<X>* 临时 (类型未定义 → C2065). 转 COMObject 走后期绑定.
            const Symbol* newClsSym = lookupModuleDotted(newExpr.className);
            if (!newClsSym || newClsSym->kind == SymbolKind::ComClass
                || newClsSym->kind == SymbolKind::ComInterface) {
                withInfo.kind = WithObjKind::COMObject;
                tempType = "void*";
                withInfo.ctrlOrigName = cIdent(newExpr.className);
            } else {
                // "With New X" 总是类实例
                withInfo.kind = WithObjKind::ClassInstance;
                withInfo.className = cIdent(newExpr.className);  // Fix 011r-1
                tempType = "vb6_cls_" + withInfo.className + "*";
                withInfo.ctrlOrigName = withInfo.className;
            }
        }
        // --- Fix 010n: WithMemberExpr (嵌套With: With .Method()) ---
        else if (node.object->kind == ASTNodeKind::WithMemberExpr) {
            // 嵌套With: .Member() — 继承外层With的对象类型
            if (!withObjectInfoStack_.empty()) {
                const auto& outerInfo = withObjectInfoStack_.back();
                if (outerInfo.kind == WithObjKind::ClassInstance ||
                    outerInfo.kind == WithObjKind::COMObject) {
                    // 方法返回值通常是同类或另一个类实例
                    withInfo.kind = WithObjKind::ClassInstance;
                    // Fix <vbeclipse>: 不能无条件继承外层类名 —— 先问该成员的**声明返回类**。
                    // ucPerspective.ctl `With l_Folder` 内的 `With .Position`
                    // (Position As Rectangle) 若继承外层 "Folder", 内层 .Left/.Right/
                    // .Top/.Bottom 会被解析到 Folder 上, 生成 vb6_Folder_prop_let_Left
                    // → LNK2019; 且临时变量被强转成 vb6_cls_Folder* (实际是 Rectangle*).
                    // 取不到返回类时再退回继承外层类名 (Fix 011r-1 原行为).
                    if (!outerInfo.className.empty()
                        && node.object->kind == ASTNodeKind::WithMemberExpr) {
                        auto& wmRet = static_cast<WithMemberExpr&>(*node.object);
                        if (!wmRet.memberName.empty()) {
                            const std::string retClsWM =
                                getClassMethodReturnType(outerInfo.className, wmRet.memberName);
                            if (!retClsWM.empty()) withInfo.className = cIdent(retClsWM);
                        }
                    }
                    if (withInfo.className.empty()) withInfo.className = outerInfo.className;
                    if (!withInfo.className.empty()) {
                        tempType = "vb6_cls_" + withInfo.className + "*";
                    }
                } else if (outerInfo.kind == WithObjKind::FormControl ||
                           outerInfo.kind == WithObjKind::WithEventsCtrl) {
                    // 控件属性返回的对象 → 类实例
                    withInfo.kind = WithObjKind::ClassInstance;
                    // 控件方法返回的对象类型未知, 不设置className
                }
                // UDT/Unknown/BuiltinObject: 保持Unknown (struct.field访问)
            }
        }

        // --- IdentifierExpr: 标识符With对象 ---
        std::string objNameLower;
        if (node.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& idExpr = static_cast<IdentifierExpr&>(*node.object);
            objNameLower = idExpr.name;
            std::transform(objNameLower.begin(), objNameLower.end(), objNameLower.begin(), ::tolower);
        }

        if (!objNameLower.empty()) {
            // Fix 161f-extlist: `With lv` —— lv 是**形参/局部持有的 ListView**
            // (`Sub RefillList(lv As ListView)`)。ListView 是真窗口, C 侧槽变量本身
            // 就是 HWND (调用点传的就是 vb6_hwnd_ListView1), 所以:
            //   HWND _vb6_with_N = lv;      ← 不是 (void*)(*lv)
            // 再把 _vb6_with_N 登记进 listViewSlotVars_, 让 .ListItems/.ColumnHeaders
            // 走 listViewHwndExprOf → vb6_ListView_ListItems((void*)_vb6_with_N)。
            // 否则落 COMObject 路径: `void* _vb6_with_N = (void*)(*lv)` +
            // vb6_ComGetObjectProp(_vb6_with_N, L"ListItems") —— 拿 HWND 当 IDispatch
            // 用, 编得过、运行期 0 行 0 列 (extlist RefillList 实测)。
            bool isLvSlot161f = listViewSlotVars_.count(objNameLower) > 0;
            auto itWE = knownWithEventsCtrlVars_.find(objNameLower);
            if (isLvSlot161f) {
                withInfo.kind = WithObjKind::COMObject;  // 让 .ListItems 走 COM 标记链
                withInfo.ctrlType = FrmControlType::ListView;
                withInfo.ctrlOrigName = objNameLower;
                tempType = "HWND";
            } else if (itWE != knownWithEventsCtrlVars_.end()) {  // P16: WithEvents控件 (优先于普通控件)
                withInfo.kind = WithObjKind::WithEventsCtrl;
                withInfo.ctrlType = itWE->second;
                auto itOrig = knownWithEventsCtrlOrigNames_.find(objNameLower);
                withInfo.ctrlOrigName = (itOrig != knownWithEventsCtrlOrigNames_.end())
                    ? "vb6_hwnd_" + cIdent(itOrig->second) : "vb6_hwnd_" + cIdent(objNameLower);
                tempType = "HWND";
            } else {
                // 窗体控件
                auto itCtrl = knownFormControls_.find(objNameLower);
                if (itCtrl != knownFormControls_.end()) {
                    withInfo.kind = WithObjKind::FormControl;
                    withInfo.ctrlType = itCtrl->second;
                    if (itCtrl->second == FrmControlType::Menu) {  // P20-36: Menu stores lowercase name for makeCtrlHwndArg
                        withInfo.ctrlOrigName = objNameLower;
                    } else {
                        withInfo.ctrlOrigName = "vb6_hwnd_" + cIdent(objNameLower);
                    }
                    tempType = "HWND";
                    // Fix <vbeclipse>: 工程内 UserControl 实例作 With 目标 (`With ucPerspective1`)
                    // —— 该名字是 .ctl 类的**实例**, 体内成员是类成员 (Sub/Property), 不是
                    // HWND 控件的 COM 属性. 此前一律 tempType=HWND → 体内 `WithEvents` 式
                    // 调用走 `vb6_ComCall(vb6_hwnd_ucPerspective1, L"AddView", ...)`, 拿 HWND
                    // 当 IDispatch 用 → RTL 报 "method AddView not found" (frmMain.Form_Load
                    // 建 perspective 全失败 → ucPerspective.Perspectives 恒为空 →
                    // InitializeAll 里 List.Item(0) 空表越界 → 0xC0000005 读 0xc).
                    // 与 cgen_expr_member_form_builtin.inc Fix 112 同口径: 成员交给
                    // resolveClassMemberCall 解析, this 取 vb6_UC_InstanceOf(hwnd).
                    auto itUCw = knownUserControlCtrlVars_.find(objNameLower);
                    if (itUCw != knownUserControlCtrlVars_.end()) {
                        withInfo.kind = WithObjKind::ClassInstance;
                        withInfo.className = cIdent(itUCw->second);
                        withInfo.ctrlOrigName = withInfo.className;
                        tempType = "vb6_cls_" + withInfo.className + "*";
                    }
                }
            }

            // COM对象变量检测
            if (withInfo.kind == WithObjKind::Unknown) {
                // Fix 157f: knownObjectVars_ 跨过程不清空 (模块级/Object局部变量需持久),
                // 但先前过程的 Object 类型参数注册 (如 ByRef This As Object) 会残留同名字
                // 段. 若当前已知该名称为 UDT (knownUdtVars_), 优先按 struct 处理,
                // 避免 `With This` (UDT参数) 被误判 COMObject → .field 走 vb6_ComGet*Prop
                // 生成 C2172/C2440.
                if (knownObjectVars_.count(objNameLower) && !knownUdtVars_.count(objNameLower)) {
                    withInfo.kind = WithObjKind::COMObject;
                }
            }

            // Fix 043b: Early-bound COM variable detection (Dim x As Dictionary, etc.)
            // knownObjectVars_ only contains void* (late-binding) variables.
            // Early-bound COM variables (vb6_ComIface_IDictionary*, etc.) are in
            // knownTypedComVars_. Without this check, With blocks on early-bound
            // COM locals fall through as Unknown, causing .Add to be resolved as
            // the class's own method instead of COM dispatch (C2198).
            if (withInfo.kind == WithObjKind::Unknown) {
                if (knownTypedComVars_.count(objNameLower) && !knownUdtVars_.count(objNameLower)) {
                    withInfo.kind = WithObjKind::COMObject;
                }
            }

            // Fix <vbeclipse>: `With <本函数返回值变量>` —— VB 允许直接 With 当前函数的
            // 返回值变量 (名字 = 函数名)。实证 Perspective.cls:
            //     Public Function AddFolder(...) As Folder
            //        Set AddFolder = New Folder
            //        With AddFolder
            //           .FolderId = FolderId : .Ratio = ... : .RefId = ... : .Relationship = ...
            //        End With
            // 该变量在 C 侧叫 `vb6_ret_AddFolder`, 而"按名登记类型"的集合 (Fix 088c 等)
            // 记的是**C 名**; 这里按**裸 VB 名**查 knownClassVars_ 必然落空 → 接收者被
            // 降级成 COMObject → `.FolderId = …` 生成
            // `vb6_ComSetProp((void*)vb6_cls_Folder*, L"FolderId", …)`, RTL 侧
            // vb6_ComIsDispatchable 拒绝裸结构体 → "property not found" **静默丢弃**
            // → perspective 里一个 folder 都没有 → 后面 Err.Raise 1
            // ("No perspective found!") → 未处理 → ExitProcess(1)。
            // 用 currentReturnCType_ (Fix 054 已为 With-UDT 保存的返回类型 C 名) 判类:
            // 它有 "vb6_cls_" 前缀就说明返回值是**原生工程类实例**。天然只在本过程内有效
            // —— 返回值变量本就是过程局部的, 不该去污染模块级集合 (那会让同名的
            // 模块变量/别的过程被误判)。
            if (withInfo.kind == WithObjKind::Unknown && currentProc_
                && !currentProc_->name.empty()
                && currentReturnCType_.rfind("vb6_cls_", 0) == 0
                && Symbol::toLower(objNameLower) == Symbol::toLower(currentProc_->name)) {
                std::string retClsVbe = currentReturnCType_.substr(8);
                if (!retClsVbe.empty() && retClsVbe.back() == '*') retClsVbe.pop_back();
                if (!retClsVbe.empty()) {
                    withInfo.kind = WithObjKind::ClassInstance;
                    withInfo.className = cIdent(retClsVbe);
                    tempType = "vb6_cls_" + withInfo.className + "*";
                }
            }

            // 类实例变量检测
            if (withInfo.kind == WithObjKind::Unknown) {
                auto itClassVar = knownClassVars_.find(objNameLower);
                if (itClassVar != knownClassVars_.end()) {
                    withInfo.kind = WithObjKind::ClassInstance;
                    // Fix 011r-1: 设置className, 让WithMemberExpr能精确解析该类方法
                    withInfo.className = cIdent(itClassVar->second);
                    tempType = "vb6_cls_" + withInfo.className + "*";
                }
            }

            // Fix 043b: Class field detection — knownObjectVars_ and knownClassVars_
            // only contain local variables/parameters, NOT class fields. When a With
            // block targets a class field (e.g., `With Dic` where Dic is `Dim Dic As
            // Dictionary`), we need to check classVoidFieldMap_ (for void*/COM fields)
            // and classTypedFieldMap_ (for typed class fields) to determine the correct
            // WithObjKind. Without this, COM fields like Dictionary fall through as
            // Unknown, and .Add inside the With block gets resolved to the class's own
            // Add method instead of COM dispatch (C2198).
            if (withInfo.kind == WithObjKind::Unknown && isClassModule_) {
                // Check void* (COM) fields first
                if (classVoidFieldMap_) {
                    auto itV = classVoidFieldMap_->find(moduleName_);
                    if (itV != classVoidFieldMap_->end() && itV->second.count(objNameLower)) {
                        withInfo.kind = WithObjKind::COMObject;
                    }
                }
                // Check typed class fields (project class or COM interface)
                if (withInfo.kind == WithObjKind::Unknown && classTypedFieldMap_) {
                    auto itT = classTypedFieldMap_->find(moduleName_);
                    if (itT != classTypedFieldMap_->end()) {
                        auto itField = itT->second.find(objNameLower);
                        if (itField != itT->second.end()) {
                            const std::string& typeName = itField->second;
                            if (typeName.rfind("COM:", 0) == 0) {
                                // COM interface field (e.g., "COM:Dictionary")
                                withInfo.kind = WithObjKind::COMObject;
                            } else {
                                // Project class field (e.g., "cAsyncSocket")
                                withInfo.kind = WithObjKind::ClassInstance;
                                withInfo.className = cIdent(typeName);
                                tempType = "vb6_cls_" + withInfo.className + "*";
                            }
                        }
                    }
                }
            }

            // Fix 010n: UDT变量检测 → 设置正确的struct类型 (而非void*)
            // 这样 struct.field 访问才能通过编译 (C2224修复)
            if (withInfo.kind == WithObjKind::Unknown) {
                auto itUdt = knownUdtVars_.find(objNameLower);
                if (itUdt != knownUdtVars_.end()) {
                    tempType = itUdt->second;  // e.g., "vb6_type_OPENFILENAME"
                    // Keep Unknown kind — struct.field 访问对UDT是正确的
                    // Fix 037: 注册 With 临时变量到 knownUdtVars_, 让嵌套 UDT 字段
                    // 访问 (如 _vb6_with_N.DecrBuffer.Data(0)) 能推断出 UDT 类型,
                    // 正确生成 VB6_SA_AT 而非误当函数调用 (C2064).
                    knownUdtVars_[tempVar] = itUdt->second;
                    knownLocalVars_.insert(tempVar);
                }
            }

            // Fix 054: With目标为当前函数UDT返回值 (如 With QRCodegenMakeBytes → vb6_ret_QRCodegenMakeBytes)
            // 函数返回变量不在 knownUdtVars_ 中, 但返回类型可能是UDT
            if (withInfo.kind == WithObjKind::Unknown && !currentReturnVar_.empty()) {
                std::string retLower = currentReturnVar_;
                std::transform(retLower.begin(), retLower.end(), retLower.begin(), ::tolower);
                if (retLower.find(objNameLower) != std::string::npos) {
                    // objName matches the return variable → check return type
                    if (currentReturnCType_.rfind("vb6_type_", 0) == 0) {
                        tempType = currentReturnCType_;
                        knownUdtVars_[tempVar] = currentReturnCType_;
                        knownLocalVars_.insert(tempVar);
                    }
                }
            }

            // Fix 010l: 内置全局对象检测 (Err/App/Screen/Printer/Clipboard/Debug)
            if (withInfo.kind == WithObjKind::Unknown) {
                if (objNameLower == "err" || objNameLower == "app" ||
                    objNameLower == "screen" || objNameLower == "printer" ||
                    objNameLower == "clipboard" || objNameLower == "debug" ||
                    objNameLower == "console") {  // Fix 161: Console (twinBASIC 口径)
                    withInfo.kind = WithObjKind::BuiltinObject;
                    withInfo.ctrlOrigName = objNameLower;  // store lowercase name
                }
            }
        }

        // --- Fix 010n: MemberAccessExpr (With obj.member / With me.member) ---
        if (node.object->kind == ASTNodeKind::MemberAccessExpr && withInfo.kind == WithObjKind::Unknown) {
            auto& memExpr = static_cast<MemberAccessExpr&>(*node.object);
            std::string memberLower = memExpr.memberName;
            std::transform(memberLower.begin(), memberLower.end(), memberLower.begin(), ::tolower);

            // Fix 090s: MAE 目标是「宿主类的字段」时按宿主类查字段表 — With
            // Http.RequestDataQuery (Http As cHttpClient, RequestDataQuery As New
            // Dictionary = COM void* 字段): 此前成员不在 knownClassVars_ 后走
            // memSym 全局查找, 捡到 Dictionary 符号按 ClassInstance +
            // vb6_cls_Dictionary* 处理 → .Item("k")=v 生成结构体字段调用 C2039.
            // 正确按宿主类 classVoidFieldMap_/classTypedFieldMap_ 分类: void*/
            // COM: → WithObjKind::COMObject (COM dispatch), 项目类 → ClassInstance.
            std::string hostCls90s = memExpr.object ? inferClassTypeOfExpr(*memExpr.object) : "";
            if (!hostCls90s.empty()) {
                if (classVoidFieldMap_) {
                    auto itV90s = classVoidFieldMap_->find(hostCls90s);
                    if (itV90s != classVoidFieldMap_->end() && itV90s->second.count(memberLower)) {
                        withInfo.kind = WithObjKind::COMObject;
                        withInfo.ctrlOrigName = hostCls90s;
                    }
                }
                if (withInfo.kind == WithObjKind::Unknown && classTypedFieldMap_) {
                    auto itT90s = classTypedFieldMap_->find(hostCls90s);
                    if (itT90s != classTypedFieldMap_->end()) {
                        auto itF90s = itT90s->second.find(memberLower);
                        if (itF90s != itT90s->second.end()) {
                            if (itF90s->second.rfind("COM:", 0) == 0) {
                                withInfo.kind = WithObjKind::COMObject;
                                withInfo.ctrlOrigName = hostCls90s;
                            } else {
                                withInfo.kind = WithObjKind::ClassInstance;
                                withInfo.className = cIdent(itF90s->second);
                                tempType = "vb6_cls_" + withInfo.className + "*";
                                withInfo.ctrlOrigName = withInfo.className;
                            }
                        }
                    }
                }
            }

            // Fix <vbeclipse> rev8: 宿主类已知时, 成员归属**由宿主类决定**, 不得再让
            // 全局同名变量 (knownClassVars_ 等) 认领。三个条件缺一不可:
            //   ① hostCls90s 推得出宿主类 ② resolveClassMemberCall 命中该成员
            //   —— 两者合起来才说明 "这个 .X 是宿主类的 X", 而不是别处的同名变量。
            //
            // 为什么必须加: ucFolder.AddView 的形参就叫 `View` (ByRef View As View),
            // 工程类兜底把 knownClassVars_["view"] 登记成 View 类; 而方法体里
            // `With View.View` 的**成员名**也是 "view" → 下面 line 350 的全局名字
            // 匹配撞车 → kind=ClassInstance(View) → tempType=vb6_cls_View*,
            // 而 vb6_View_prop_get_View() 实际返回 void* (View.cls:66 `As Object`)
            // → 块内 .hWnd/.Caption/.Icon 按类字段发 → C2039 "hWnd 不是
            // vb6_cls_View 的成员" ×12 (ucFolder.c 240/241/262)。
            //
            // 处置: 清空判定让下面 needHostCls113 段接手, 由 getClassMethodReturnType
            // 按属性**真实返回类型**分流 (工程类 → ClassInstance; As Object → 空,
            // 由本段末尾的 rev8 补成 COMObject 晚绑定)。
            const bool hostOwnsMemberV8 =
                !hostCls90s.empty()
                && !resolveClassMemberCall(hostCls90s, memExpr.memberName).empty();
            if (hostOwnsMemberV8) {
                withInfo.kind = WithObjKind::Unknown;
                withInfo.className.clear();
                withInfo.ctrlOrigName.clear();
                tempType = "void*";
            }

            // 检查成员是否为类实例变量 (me.member As SomeClass)
            if (!hostOwnsMemberV8
                && knownClassVars_.find(memberLower) != knownClassVars_.end()) {
                withInfo.kind = WithObjKind::ClassInstance;
                // Fix 011r-1: 获取成员的类名, 设置tempType
                auto itClassVar = knownClassVars_.find(memberLower);
                if (itClassVar != knownClassVars_.end()) {
                    withInfo.className = cIdent(itClassVar->second);
                    tempType = "vb6_cls_" + withInfo.className + "*";
                }
            } else if (!hostOwnsMemberV8 && knownObjectVars_.count(memberLower)) {
                withInfo.kind = WithObjKind::COMObject;
            } else if (!hostOwnsMemberV8 && knownUdtVars_.count(memberLower)) {
                // UDT成员: 使用struct类型 (如 With ofn → vb6_type_OPENFILENAME)
                auto itUdt = knownUdtVars_.find(memberLower);
                if (itUdt != knownUdtVars_.end()) {
                    tempType = itUdt->second;
                    // Fix 037: 注册 With 临时变量到 knownUdtVars_
                    knownUdtVars_[tempVar] = itUdt->second;
                    knownLocalVars_.insert(tempVar);
                }
            } else if (withInfo.kind == WithObjKind::Unknown) {
                // Fix 090s: kind 已被上面 hostCls90s 字段表解析 (COMObject/ClassInstance)
                // 时不再走 memSym 兜底 — 否则 With Http.RequestDataQuery (RequestDataQuery
                // As New Dictionary = COM void* 字段, A1 已置 COMObject) 被
                // lookup("RequestDataQuery") 捡到 Dictionary 符号 → 覆盖成 ClassInstance
                // + vb6_cls_Dictionary* → .Item(k)=v 生成结构体字段调用 C2039.
                // 尝试从符号表推断类型
                Symbol* memSym = symTab_.lookupModule(memExpr.memberName);
                if (!memSym) memSym = symTab_.lookup(memExpr.memberName);
                if (memSym) {
                    if (memSym->type == Vb6Type::UserDefinedType) {
                        // 查找UDT类型的C标识符
                        // TODO: Symbol没有存储typeRefName, 需要其他方式
                    } else if (memSym->type == Vb6Type::Object) {
                        withInfo.kind = WithObjKind::ClassInstance;
                        // Fix 011r-1: 若Symbol有variableTypeName, 用之; 否则className未知
                        // <vbeclipse>: 但 variableTypeName 里记的**可能是 COM/RTL 建模的类型**
                        // (stdole.StdFont 最典型) —— 那种没有类模块去发 prop_let_/成员定义,
                        // 按类实例发码就是"调一个不存在的函数" ⇒ LNK2019 (实测 Charts 2020
                        // `Property Set Font`: vb6_StdFont_prop_let_* 8 个全无定义, 每个
                        // .ctl 各 9 个未解析符号)。只有**工程类**才配当这个类名; 否则
                        // className 留空、tempType 保持 void*, 下游按"类名未知"走
                        // vb6_ComSetProp 晚绑定 (与 dev 同一形态)。
                        if (!memSym->variableTypeName.empty()
                            && isProjectClassName(memSym->variableTypeName)) {
                            withInfo.className = cIdent(memSym->variableTypeName);
                            tempType = "vb6_cls_" + withInfo.className + "*";
                        }
                    }
                }
                // czUI fix: 宿主伪对象的属性 (UserControl.Parent / Extender.X 等)
                // 是容器窗体对象 — 运行时只有 HWND, 没有项目类。落到下面的
                // ClassInstance 兜底会让 .Left 等成员撞上 VB 内置函数符号
                // (Left$) 被解析成 vb6_<类>_Left(withObj) → LNK2019
                // (czUI.ctl ToggleFullScreen 实测)。转 COMObject 晚绑定,
                // .Left/.Top/.Width/.Height 走 vb6_ComGetIntProp。
                if (withInfo.kind == WithObjKind::Unknown && memExpr.object &&
                    memExpr.object->kind == ASTNodeKind::IdentifierExpr) {
                    auto& objId133w = static_cast<IdentifierExpr&>(*memExpr.object);
                    std::string objLower133w = Symbol::toLower(objId133w.name);
                    if (objLower133w == "usercontrol" || objLower133w == "ambient" ||
                        objLower133w == "extender" || objLower133w == "propertypage") {
                        withInfo.kind = WithObjKind::COMObject;
                    }
                }
                // 无法确定类型时默认为类实例 (void*不支持.member访问)
                if (withInfo.kind == WithObjKind::Unknown && tempType == "void*") {
                    withInfo.kind = WithObjKind::ClassInstance;
                }
            }
            // Fix <vbeclipse>: MAE 目标是「宿主类的属性 Get」而非数据字段时, 上面
            // classVoidFieldMap_/classTypedFieldMap_ 都查不到 → kind 落到**空
            // className** 的 ClassInstance。ucPerspective.ctl `With Folder.Views`
            // (Folder 是 ByRef Folder 类形参, Views 是 Folder 类的 Property Get
            // As List) 即此形: 块内 .IsEmpty 被 cgen_expr_with.cpp 的 className
            // 为空兜底 (symTab_.lookupModule 全局捡同名成员) 套上**当前模块**前缀
            // → vb6_ucPerspective_IsEmpty (LNK2019); .Count/.Item 只是碰巧唯一命中
            // List 类才显得正常。
            // 与 inferClassTypeOfExpr 的 MemberAccessExpr 分支对齐, 用
            // getClassMethodReturnType 取属性返回类。该函数仅对真实项目 Class
            // 返回类名, String/Long 等标量属性返回 "" (Fix 085b 语义不变)。
            //
            // 触发条件是 **className 为空** 而非 kind==Unknown: 上面的
            // symTab_.lookupModule(memberName) 兜底在 memSym->type==Object 而
            // variableTypeName 为空时, 会把 kind 置成 ClassInstance 却留下空
            // className — 只判 Unknown 抓不到。
            // (l_Folder.Views 等能工作是因 knownClassVars_ 里恰好有别的过程登记过
            //  名为 "views" 的 List 变量, 属巧合, 不可依赖。)
            const bool needHostCls113 =
                (withInfo.kind == WithObjKind::ClassInstance && withInfo.className.empty())
                || (withInfo.kind == WithObjKind::Unknown);
            if (needHostCls113) {
                // Fix <vbeclipse>: `<host>.<PropertyGet>` 且推断不出宿主类时, 用该
                // PropertyGet 的返回类型当 With 块类 (见 getClassMethodReturnType 中
                // 同名 Fix 的说明: 跨模块 PropertyGet 的 type 常记 Variant, 靠
                // variableTypeName 才能取到返回类). 治 `With Folder.Views` 推不出
                // List → 块内 .IsEmpty 落到模块级兜底 → vb6_ucPerspective_IsEmpty.
                std::string retClsProp = hostCls90s.empty()
                    ? std::string()
                    : getClassMethodReturnType(hostCls90s, memberLower);
                // <vbeclipse>: 但**只有工程类**才配当 ClassInstance —— 它的 prop_let_/
                // prop_set_ 由类模块自己发定义。COM/RTL 建模的类型 (stdole.StdFont 最典型:
                // RTL 里只有 `typedef vb6_ComIface_Font vb6_cls_StdFont;`) 没有那份定义,
                // 按类实例发码就是"调一个不存在的函数" ⇒ LNK2019 (实测 Charts 2020 的
                // ucChartArea/ucChartBar/ucPieChart/ucTreeMaps 的 `Property Set Font`:
                // 8 个 vb6_StdFont_prop_let_* 全无定义, 每个 .ctl 各 9 个未解析符号)。
                // 不满足就维持原来的 kind (COMObject 那条路发 vb6_ComSetProp, 与 dev 相同)。
                if (!retClsProp.empty() && isProjectClassName(retClsProp)) {
                    withInfo.kind = WithObjKind::ClassInstance;
                    withInfo.className = cIdent(retClsProp);
                    tempType = "vb6_cls_" + withInfo.className + "*";
                    withInfo.ctrlOrigName = withInfo.className;
                } else if (hostOwnsMemberV8) {
                    // Fix <vbeclipse> rev8: 成员确属宿主类, 但返回类型**不是工程类**
                    // —— 最典型就是 `Property Get View() As Object` (View.cls:66)。
                    // getClassMethodReturnType 对 As Object 返回空 (rev7 收紧后的
                    // 正确行为), 若放任 kind 留 Unknown, 下面的 `tempType=="void*"`
                    // 兜底会把它按 ClassInstance + 空 className 处理, 块内 .hWnd
                    // 走 `tempVar->hWnd` (void* 取成员) → C2223。
                    // As Object 的语义就是 IDispatch → 必须转 COMObject 晚绑定,
                    // 块内 .hWnd/.Caption 由 vb6_ComGetProp 运行时问窗体。
                    withInfo.kind = WithObjKind::COMObject;
                    withInfo.className.clear();
                    tempType = "void*";
                }
            }
        }

        // --- Fix 010n: IndexOrCallExpr (With arr(idx) / With func()) ---
        if (node.object->kind == ASTNodeKind::IndexOrCallExpr && withInfo.kind == WithObjKind::Unknown) {
            // 数组元素或函数返回值 — 通常是类实例或VARIANT
            // 数组元素访问如 m_uWindowState(0) → VARIANT UDT
            auto& callExpr = static_cast<IndexOrCallExpr&>(*node.object);
            if (callExpr.callee && callExpr.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& idExpr = static_cast<IdentifierExpr&>(*callExpr.callee);
                std::string arrLower = idExpr.name;
                std::transform(arrLower.begin(), arrLower.end(), arrLower.begin(), ::tolower);
                // Fix 055: 优先检查UDT数组元素类型
                auto itUdtArr = arrayUdtElemTypes_.find(arrLower);
                if (itUdtArr != arrayUdtElemTypes_.end()) {
                    tempType = itUdtArr->second;  // e.g. "vb6_type_RECT"
                    // 注册到knownUdtVars_, 让后续字段访问正确
                    knownUdtVars_[tempVar] = itUdtArr->second;
                    knownLocalVars_.insert(tempVar);
                } else {
                    // 检查是否为已知数组 → 元素类型
                    auto itArr = arrayElemTypes_.find(arrLower);
                    if (itArr != arrayElemTypes_.end()) {
                        if (itArr->second == Vb6Type::UserDefinedType) {
                            tempType = "vb6_VARIANT";  // UDT数组元素存储为VARIANT (fallback, should be caught above)
                        } else if (itArr->second == Vb6Type::Variant || itArr->second == Vb6Type::Object) {
                            tempType = "vb6_VARIANT";
                        }
                    }
                }
            }
            // Fix 160w: Function 返回 UDT 的 With 目标 (Common.bas `With
            // GetAppVersionInfo()`, Function As VS_FIXEDFILEINFO). callee 符号是
            // Function 且返回 UserDefinedType → tempType = vb6_type_<UDT名>, kind 保持
            // Unknown 走 struct 字段访问. 否则 void* fallback 发射
            // `(void*)vb6_Common_GetAppVersionInfo()` → C2440 (struct→void*).
            if (withInfo.kind == WithObjKind::Unknown && tempType == "void*"
                && callExpr.callee && callExpr.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& fnId160w = static_cast<IdentifierExpr&>(*callExpr.callee);
                Symbol* fnSym160w = symTab_.lookupModule(fnId160w.name);
                if (!fnSym160w) fnSym160w = symTab_.lookup(fnId160w.name);
                if (fnSym160w && fnSym160w->kind == SymbolKind::Function
                    && fnSym160w->type == Vb6Type::UserDefinedType
                    && !fnSym160w->variableTypeName.empty()) {
                    tempType = "vb6_type_" + cIdent(fnSym160w->variableTypeName);
                    knownUdtVars_[tempVar] = tempType;
                    knownLocalVars_.insert(tempVar);
                }
            }
            // 函数返回值且仍为void* → 默认按 COM 后期绑定分发
            if (withInfo.kind == WithObjKind::Unknown && tempType == "void*") {
                // Fix 090y: void* With 目标 (COM 方法返回对象, 如 cIni.Section As
                // Dictionary) → COMObject (后期绑定 dispatch). 此前 ClassInstance
                // (className 空) → WithMemberExpr 成员解析落入全局符号表撞名
                // (cTimers.Item) → 左值错误 C2106. UDT 目标 tempType=vb6_type_*
                // 不受影响; Variant-对象目标运行时 dispatch 也更贴合 COM 语义.
                withInfo.kind = WithObjKind::COMObject;
            }
        }

        // --- 最终回退: void* 不支持 .member 访问 → COM 后期绑定分发 ---
        // Fix 090y: (同上方 Fix, 独立于 isClassModule_ 字段检测的最终兜底)
        // UDT struct (vb6_type_*) 目标 tempType 非 void* → 不受影响.
        if (withInfo.kind == WithObjKind::Unknown && tempType == "void*") {
            withInfo.kind = WithObjKind::COMObject;
        }
        // Fix 054: tempType 已解析为 UDT struct (vb6_type_*) → 保持 Unknown, 用 struct.field 访问
        if (withInfo.kind == WithObjKind::ClassInstance && tempType.rfind("vb6_type_", 0) == 0) {
            withInfo.kind = WithObjKind::Unknown;
        }
    }

    // Fix 092o: 本帧 withInfo 延后到目标表达式生成之后再入栈 — 若提前入栈, 嵌套
    // With 的目标表达式 (.ReturnJson()) 会以内层 className (cJson) 解析外层成员
    // (实为 cHttpClient 的 ReturnJson) → 解析失败退化为数据字段访问 (cAliyunCaptcha
    // 222: `_vb6_with_3->ReturnJson()` C2039 + 实参丢失). 见下方 emitExpr 之后的入栈.

    // Fix 010l: BuiltinObject 不需要临时变量 — 属性读写直接映射为RTL函数调用
    if (withInfo.kind == WithObjKind::BuiltinObject) {
        withObjectInfoStack_.push_back(withInfo);  // Fix 092o: BuiltinObject 无目标表达式, 直接入栈
        // 推入占位符以保持 withObjectVars_ 与 withObjectInfoStack_ 同步
        withObjectVars_.push_back(withInfo.ctrlOrigName);
        c_.emitLine("{");
        c_.indent();
        emitStmtList(node.body);
        c_.dedent();
        c_.emitLine("}");
        withObjectVars_.pop_back();
        withObjectInfoStack_.pop_back();
        return;
    }

    // P17.1: 抑制With对象表达式的默认属性解析
    bool prevSuppress = suppressDefaultProp_;
    if (withInfo.kind == WithObjKind::FormControl || withInfo.kind == WithObjKind::WithEventsCtrl) {
        suppressDefaultProp_ = true;
    }

    emitExpr(*node.object);

    suppressDefaultProp_ = prevSuppress;

    // Fix <vbeclipse>: 工程内 UserControl 实例作 With 目标时, 目标表达式要取**宿主窗口
    // 反查到的实例指针** (vb6_UC_InstanceOf(hwnd)), 而不是裸 HWND 槽 —— 否则 tempType
    // vb6_cls_<UC>* 被灌入一个 HWND, 体内成员调用等于拿 HWND 当结构体解引用.
    // (与 cgen_expr_member_form_builtin.inc:264 的 thisArg 同构.)
    if (withInfo.kind == WithObjKind::ClassInstance
        && node.object->kind == ASTNodeKind::IdentifierExpr) {
        auto& idW = static_cast<IdentifierExpr&>(*node.object);
        auto itUCw2 = knownUserControlCtrlVars_.find(Symbol::toLower(idW.name));
        if (itUCw2 != knownUserControlCtrlVars_.end()) {
            // emitExpr 对 FormControl 标识符不发 HWND 槽 (那由 FormControl 专用发射分支
            // 用 ctrlOrigName 拼), 这里必须显式取 hwnd 实参再反查实例.
            lastExpr_ = "(vb6_cls_" + cIdent(itUCw2->second) + "*)vb6_UC_InstanceOf("
                      + makeCtrlHwndArg(Symbol::toLower(idW.name), withInfo.ctrlType) + ")";
        }
    }

    // Fix 110c: With <UDT 数组元素 / UDT 嵌套字段> — 形如 m_Serie(i).Rects(j)、
    // m_Item(i).LegendRect 等. 这类目标在 C 里是**结构体值**, 若 tempType 仍是
    // void*, 下方会发射 `(void*)<结构体>` → C2440 ("无法从 vb6_type_RectL 转换
    // 为 void *"), 或对 void* 目标做 ->members 访问 → C2224/C2039.
    // 用已有的 UDT 类型推断 (走 arrayUdtElemTypes_ + udtMembers 成员表) 得到
    // 精确的 vb6_type_X, 让 With 体按 struct 字段访问 (_vb6_with_N->Left).
    if (tempType == "void*") {
        std::string withUdt = inferUdtTypeOfExpr(*node.object);
        if (!withUdt.empty() && withUdt.rfind("vb6_type_", 0) == 0) {
            tempType = withUdt;
            withInfo.kind = WithObjKind::Unknown;  // Fix 054: UDT struct → 字段访问
            knownUdtVars_[tempVar] = withUdt;
            knownLocalVars_.insert(tempVar);
        }
    }

    // Fix 092o: 目标表达式已生成完毕 (期间保持外层栈顶), 现在把本帧 withInfo 入栈 —
    // body 内的 .成员 解析与下方的 Menu 判定都依赖它.
    withObjectInfoStack_.push_back(withInfo);

    if (!withObjectInfoStack_.empty() && withObjectInfoStack_.back().kind == WithObjKind::FormControl && withObjectInfoStack_.back().ctrlType == FrmControlType::Menu) {  // P20-36
        c_.emitLine("int " + tempVar + " = 0;  /* Menu: no HWND, props use (hmenu,menuId) */");
    } else {
        // Fix 160w: 宿主伪对象 `With UserControl` / `With PropertyPage` (UserControl
        // 类模块内) — emitExpr(<IdentifierExpr "UserControl">) 生成裸 `UserControl`,
        // C 无该声明 → C2065 (VBFlexGrid.c:1552). UserControl 伪对象在此作用域即当前
        // 控件的宿主窗口 (vb6_UserControl_hWnd, extern HWND), 直接替换表达式.
        // PropertyPage 同理用 vb6_PropertyPage_hwnd.
        if (tempType == "void*" && node.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& hid160w = static_cast<IdentifierExpr&>(*node.object);
            if (hid160w.name == "UserControl") {
                lastExpr_ = "vb6_UserControl_hWnd";
            } else if (hid160w.name.compare(0, 12, "PropertyPage") == 0) {
                lastExpr_ = "vb6_PropertyPage_hwnd";
            }
        }
        // Fix 038: C2440 修复 — UDT 同类型转换和 UDT/VARIANT → void* 转换
        bool isUdtTempType = (tempType.rfind("vb6_type_", 0) == 0);
        if (isUdtTempType) {
            // Fix 081j: UDT With块使用指针引用，而非值拷贝
            // VB6中 With uPoints(lIdx) 内 .X = ... 直接修改数组元素
            // C中需要用指针: vb6_type_RECT* _vb6_with = &VB6_SA_AT(...)
            // Fix 160w: 目标是 Function 返回 UDT (With GetAppVersionInfo()) —
            // 返回值是右值, &(fn()) → C2102 非法取址, 且 (void*) 强转 → C2440.
            // 用复合字面量承载拷贝再取址 (VB6 在 With <udtFunc> 的临时副本上读写).
            bool isFnRet160w = false;
            if (node.object->kind == ASTNodeKind::IndexOrCallExpr) {
                auto& ioc160w = static_cast<IndexOrCallExpr&>(*node.object);
                if (ioc160w.callee && ioc160w.callee->kind == ASTNodeKind::IdentifierExpr) {
                    auto& fid160w = static_cast<IdentifierExpr&>(*ioc160w.callee);
                    Symbol* fn160w = symTab_.lookupModule(fid160w.name);
                    if (!fn160w) fn160w = symTab_.lookup(fid160w.name);
                    isFnRet160w = (fn160w && fn160w->kind == SymbolKind::Function);
                }
            }
            if (isFnRet160w) {
                // Fix 160w Final: `(T){fn()}` 是**位置初始化** (fn() 结果赋给首成员
                // T.dwSignature), 不是整体拷贝 — MSVC 报 C2440 "无法从 vb6_type_X
                // 转换到 int32_t" (Common.c 617/630/643). 改用 pending 临时承载拷贝
                // 再取址 (与 arg_emit Fix 090q/160-H 同款: 声明先于引用落地).
                std::string fnRetTmp160w = "_vb6_withret" + std::to_string(tempCounter_++);
                c_.addPending(tempType + " " + fnRetTmp160w + " = " + lastExpr_ + ";");
                c_.emitLine(tempType + "* " + tempVar + " = &" + fnRetTmp160w + "  /* With obj ref (fn ret copy) */;");
            } else {
                c_.emitLine(tempType + "* " + tempVar + " = &(" + lastExpr_ + ")  /* With object ref (ptr) */;");
            }
        } else if (tempType == "void*") {
            // Fix 160w: 宿主伪结构体全局 (UserControl.Extender / UserControl.Ambient)
            // 在 vb6rtl_userctl.h 是 struct 值; (void*)(struct) → C2440 (VBFlexGrid.c
            // 1521 "无法从 vb6_UserControl_Extender_Type 转换到 void *"). 取地址即可,
            // With 体内 .成员 由 WithMemberExpr 按 COM dispatch 解析.
            if (lastExpr_ == "vb6_UserControl_Extender" || lastExpr_ == "vb6_UserControl_Ambient") {
                c_.emitLine(tempType + " " + tempVar + " = &(" + lastExpr_ + ")  /* With object ref (host struct) */;");
            } else {
            // 检查表达式是否为 UDT 或 VARIANT — 这些类型不能直接 cast 到 void*
            std::string udtCType = inferUdtTypeOfExpr(*node.object);
            if (!udtCType.empty()) {
                // UDT → void*: 取地址获取指针
                c_.emitLine(tempType + " " + tempVar + " = &(" + lastExpr_ + ")  /* With object ref */;");
            } else if (isDefinitelyVariantExpr(*node.object) && inferClassTypeOfExpr(*node.object).empty()) {
                // VARIANT → void*: 用 VariantToObjectVal 提取对象指针
                // Fix 084e: 若表达式实际是类对象 (如 Cookies("name") 返回
                // vb6_cls_cHttpServerCookieAttr* 但被误判为 Variant), 则
                // VariantToObjectVal(对象指针) 触发 C2440, 走下方 (void*) 直转.
                // Fix 090e: 符号表把项目类默认成员属性 (Property Get Cookie()
                // As cHttpServerCookieAttr) 的返回类型误注册为 Variant 时,
                // 生成的 C 表达式是 vb6_cHttpServerCookies_prop_get_Cookie(...)
                // (返回 vb6_cls_cHttpServerCookieAttr*), 并非 vb6_VARIANT 值;
                // 对类指针调 VariantToObjectVal → C2440 (cHttpServerCookies
                // ExpireCookie: With Cookie(Key)). 仅当 C 级确认实参是
                // vb6_VARIANT (cExprIsVariant / 已知 Variant 变量/字段) 时
                // 走 VariantToObjectVal, 否则按对象指针 (void*) 直转.
                bool withIsVariantVal090e = cExprIsVariant(lastExpr_);
                if (!withIsVariantVal090e) {
                    std::string lower090e = lastExpr_;
                    std::transform(lower090e.begin(), lower090e.end(),
                                   lower090e.begin(), ::tolower);
                    if (knownVariantVars_.count(lower090e)) {
                        withIsVariantVal090e = true;
                    } else if (lower090e.compare(0, 4, "me->") == 0) {
                        std::string mem090e = lower090e.substr(4);
                        if (classVariantMembers_.count(mem090e)) {
                            withIsVariantVal090e = true;
                        }
                    }
                }
                if (withIsVariantVal090e) {
                    c_.emitLine(tempType + " " + tempVar + " = vb6_VariantToObjectVal(" + lastExpr_ + ")  /* With object ref */;");
                } else {
                    c_.emitLine(tempType + " " + tempVar + " = (" + tempType + ")" + lastExpr_ + "  /* With object ref */;");
                }
            } else {
                // Fix 092j: inferClassTypeOfExpr 非空 (按 VB 声明推断出项目类) 但
                // C 级表达式实际是 Variant 值 (COM 链结果) → 不能 (void*) 硬转:
                //   cLang 144: With me->LangInfo.Item("LangList").Item(Idx+1)
                //     → (void*)vb6_VariantFromComResult(vb6_ComCall(...)) C2440
                //       "无法从 vb6_VARIANT 转换为 void *";
                // 仅当 C 级确认是 vb6_VARIANT 时改用 VariantToObjectVal 提取.
                // vbeclipse: 本分支原先只认 cExprIsVariant 的**字符串前缀**, 裸名
                // Variant 局部 (如 ucPerspective.ctl Refresh 的
                // `Dim l_ucFolder As Variant` + `With l_ucFolder`) 一律落 573 直转
                // → C2440 ×8 (ucPerspective.c 1953/1983/2036/2054/2078/2096/2126/2146
                // `void* _vb6_with_N = (void*)l_ucFolder`). isDefinitelyVariantExpr
                // 入口处也可能被 lookupModule 的同名符号误导 (Fix 049b 同款陷阱),
                // 故与上方 090e 同口径: knownVariantVars_ 裸名 + me-> classVariantMembers_。
                bool withIsVariantVal092j = cExprIsVariant(lastExpr_);
                if (!withIsVariantVal092j) {
                    std::string lower092j = lastExpr_;
                    std::transform(lower092j.begin(), lower092j.end(),
                                   lower092j.begin(), ::tolower);
                    if (knownVariantVars_.count(lower092j)) {
                        withIsVariantVal092j = true;
                    } else if (lower092j.compare(0, 4, "me->") == 0) {
                        std::string mem092j = lower092j.substr(4);
                        if (classVariantMembers_.count(mem092j)) {
                            withIsVariantVal092j = true;
                        }
                    }
                }
                if (withIsVariantVal092j) {
                    c_.emitLine(tempType + " " + tempVar + " = vb6_VariantToObjectVal(" + lastExpr_ + ")  /* With object ref */;");
                } else {
                    c_.emitLine(tempType + " " + tempVar + " = (" + tempType + ")" + lastExpr_ + "  /* With object ref */;");
                }
            }
            }  /* Fix 160w: 宿主结构体 else 闭合 (tempType == "void*" 分支) */
        } else {
            c_.emitLine(tempType + " " + tempVar + " = (" + tempType + ")" + lastExpr_ + "  /* With object ref */;");
        }
    }

    withObjectVars_.push_back(tempVar);

    // Fix 161f-extlist: With 目标是 ListView 槽 (形参/局部) 时, 把 _vb6_with_N 本身
    // 也登记为 ListView 槽 —— 体 `.ListItems` / `.ColumnHeaders` 生成的 comObjExpr_
    // 就是这个 tempVar, 登记后 listViewHwndExprOf 才能把它解析成 HWND。
    // 弹出时 (下方 pop_back) 必须撤销登记, 否则同名 tempVar 在后续函数复用 → 假命中。
    bool lvSlotRegistered161f = false;
    if (withInfo.ctrlType == FrmControlType::ListView
        && withInfo.kind == WithObjKind::COMObject) {
        std::string tvLower161f = Symbol::toLower(tempVar);
        lvSlotRegistered161f = listViewSlotVars_.insert(tvLower161f).second;
    }

    c_.emitLine("{");
    c_.indent();
    emitStmtList(node.body);
    c_.dedent();
    c_.emitLine("}");

    withObjectVars_.pop_back();
    withObjectInfoStack_.pop_back();
    if (lvSlotRegistered161f) listViewSlotVars_.erase(Symbol::toLower(tempVar));
}


} // namespace vb6c3
