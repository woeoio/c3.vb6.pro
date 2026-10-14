#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <cstdio>
#include <cstdlib>
#include <functional>
#include <map>

namespace vb6c3 {

// --- cgen_util_classtype.cpp: 类/UDT 类型推断与字段 C 类型 ---


std::string CCodeGen::inferClassTypeOfExpr(const ASTNode& expr) const {
    switch (expr.kind) {
        case ASTNodeKind::IdentifierExpr: {
            // base case: 变量 → knownClassVars_
            auto& id = static_cast<const IdentifierExpr&>(expr);
            std::string lower = Symbol::toLower(id.name);
            auto it = knownClassVars_.find(lower);
            if (it != knownClassVars_.end()) return it->second;
            if (std::getenv("C3_DBG110") && lower == "usercontrol") {
                const Symbol* s2 = symTab_.lookup(id.name);
                std::fprintf(stderr, "[DBG110] inferClassTypeOfExpr(UserControl) mod=%s sym=%p kind=%d vtn=%s\n",
                             moduleName_.c_str(), (const void*)s2, s2 ? (int)s2->kind : -1,
                             s2 ? s2->variableTypeName.c_str() : "<null>");
            }
            // Fix 084g: 局部变量/参数声明为 As ClassName (如 Dim Response As cHttpServerResponse)
            // 不在 knownClassVars_ (跨模块类变量表) 中, 从符号表 variableTypeName 推断类名
            const Symbol* sym = symTab_.lookup(id.name);
            if (sym && (sym->kind == SymbolKind::Variable || sym->kind == SymbolKind::Parameter)
                && !sym->variableTypeName.empty()) {
                const Symbol* clsSym = symTab_.lookup(sym->variableTypeName);
                if (clsSym && clsSym->kind == SymbolKind::Class) {
                    return sym->variableTypeName;
                }
            }
            // Fix 084z-2: 当前模块的属性 (Property Get) 返回类 — 属性符号在模块
            // 作用域符号表中 (如 cTlsRemaster.pvSocket() As cTlsSocket). 无参属性
            // 可推断类实例, 用于 pvSocket.SyncReceiveArray(...) 的方法/形参签名解析
            // (否则 Fix 033 把属性名当模块名, 回退 lookupModule 命中 storageKey
            // 同名冲突的错误类 cWinsock.SyncReceiveArray → C2197 参数过多).
            if (sym && sym->kind == SymbolKind::PropertyGet && !sym->variableTypeName.empty()) {
                const Symbol* clsSym = symTab_.lookup(sym->variableTypeName);
                if (clsSym && clsSym->kind == SymbolKind::Class) {
                    return sym->variableTypeName;
                }
            }
            // Fix 218: 当前模块的**裸函数名 (返回项目类)** 也是一个类实例表达式 ——
            // 与 084g 的局部变量/参数、084z-2 的 PropertyGet 同档。VB6 里 `Fn.Member`
            // 中 Fn 是无参 Function 时, 语义就是 "(调用 Fn).Member", 接收者是这个
            // 函数**返回的那个对象**, 不是"模块 Fn"。
            //
            // 实测 ComCtlsDemo (两处 pattern 完全一致, 都是"影子访问器"惯用法):
            //   Builds/CoolBar/CbrBand.cls:44   Private Function ShadowCoolBar() As CoolBar
            //   Builds/CoolBar/CbrBand.cls:114  Caption = ShadowCoolBar.FBandCaption(PropID)
            //   Builds/ListView/LvwListItem.cls:52  Private Function ShadowListView() As ListView
            //   Builds/ListView/LvwListItem.cls:80  If ShadowListView.FListItemVerify(PropPtr, PropIndex) …
            // 修复前: Fix 083c (cgen_expr_member_class_module.inc) 只兜了 PropertyGet,
            // 于是这两处 inferClassTypeOfExpr 返回空 → 第 18-22 行的
            // `itClassVar.emplace` 不成立 → 落优先级3"模块名.方法名" → 接收者丢失,
            // 生成的 callee 里根本没有这个对象 (实测 CbrBand.c:110 发成
            // `vb6_CoolBar_prop_get_FBandCaption(me->PropID)`、LvwListItem.c:69 发成
            // `vb6_ListView_FListItemVerify(me->PropPtr, &(me->PropIndex))`) →
            // 下游 opt_pad 只补 `(void*)me` / com_bind 只前置部分实参 ⇒ C2198 参数太少
            // (an5.py 归因: "Friend 方法调用(可能缺 receiver)" 408 条, 是 MSVC 侧最大一族)。
            //
            // 修好之后 callee 是 `vb6_<Mod>_ShadowCoolBar((void*)me)` (值上下文发射,
            // 见本文件 path 上方 asCallCallee_ 的存取), 外层 com_bind 的 P6.5 拆分
            // (cgen_expr_call_com_bind.inc:859) 会自动把接收者前置进实参列表。
            //
            // 注意排除**当前过程自身**的函数名: 那是"引用函数名 = 引用返回值槽"
            // 的另一条语义 (Fix 084z-4 / cgen_expr_ident_dispatch.inc:118 已处理),
            // 在这里抢先命中会把 `vb6_ret_X` 槽再包一层, 属于另一条通路, 不动。
            if (sym && sym->kind == SymbolKind::Function && !sym->variableTypeName.empty()
                && !(currentProc_ && Symbol::toLower(currentProc_->name) == lower)) {
                const Symbol* clsSym = symTab_.lookup(sym->variableTypeName);
                if (clsSym && clsSym->kind == SymbolKind::Class) {
                    return sym->variableTypeName;
                }
            }
            // Fix <vbeclipse>: 上面符号表路径全落空时, 该标识符的名字**本身就是一个
            // 工程类** → 按该工程类推断 (VB6 语义: 工程内定义优先于宿主同名符号)。
            // 用 projectClassNameOf (driver 注入的工程类名表) 判, 而不是顺着符号表
            // 逐级查 Class/ComClass: 名字被同名符号占满时符号表根本查不到类。
            // ucPerspective.ctl `CreateFolder(ByRef Folder As Folder, ...)` 实测
            // symTab_.lookup("Folder") 两次都返回宿主同名 PropertyGet(kind=4) 而非
            // coclass (消费作用域里 "folder" 键是 ComClass, 名字被占满), 于是
            // Folder.Views / Folder.FolderId 整条跨模块属性链推不出类, With 块
            // className 留空 → 块内 .IsEmpty 套当前模块前缀 → vb6_ucPerspective_IsEmpty。
            // 放在最后: 局部变量/参数/属性各自的判定优先, 不会把与类同名的**局部
            // 变量**误当类 (已知局部走 knownClassVars_ 与 Variable/Parameter 分支)。
            // 表里没有的名字行为完全不变。
            //
            // ⚠ Fix <vbeclipse> rev7: 但**必须先排除"同名成员访问"**。
            // View.cls 里 `Property Get View() As Object` (返回窗体对象) 与类名 `View`
            // 同名 ⇒ 消费点 `l_View.View` 的 `.View` 段名字撞工程类表 ⇒ 整条链被推成
            // "vb6_cls_View 实例" ⇒ `With View.View` 把 tempType 定成 vb6_cls_View*,
            // 于是 `_vb6_with_2->hWnd` 按类字段发, 而 vb6_View_prop_get_View() 实际
            // 返回 void* (As Object) ⇒ 编译期 C2039 "hWnd 不是 vb6_cls_View 的成员"
            // ×216 (ucFolder.c/SubClass.c/modSubClass.bas 全线)。
            // 判据: 该标识符在**本模块**是个成员 (PropertyGet/PropertyLet/PropertySet/
            // 变量/参数) 时, 它的类型由**声明**决定, 与同名工程类无关 —— 名字撞车
            // 不是类型依据。已知局部在上面几条分支已处理, 到这里的是"跨模块成员 +
            // 名字撞类"这一种。
            {
                const Symbol* s2 = symTab_.lookup(id.name);
                if (s2 && (s2->kind == SymbolKind::PropertyGet
                           || s2->kind == SymbolKind::PropertyLet
                           || s2->kind == SymbolKind::PropertySet
                           || s2->kind == SymbolKind::Variable
                           || s2->kind == SymbolKind::Parameter)) {
                    // 有声明就看声明: 声明类型确实解析为项目类才返回类名, 否则返回空
                    // (交下游走 COM 晚绑定 —— As Object 的成员就该那样走)。
                    if (!s2->variableTypeName.empty()) {
                        const std::string projHit =
                            projectClassNameOf(s2->variableTypeName);
                        if (!projHit.empty()) return projHit;
                    }
                    return "";
                }
            }
            const std::string projClsVbe = projectClassNameOf(id.name);
            if (!projClsVbe.empty()) return projClsVbe;
            return "";
        }
        case ASTNodeKind::IndexOrCallExpr: {
            // recursive case: 类方法调用 obj.Method(args) → 返回类
            auto& call = static_cast<const IndexOrCallExpr&>(expr);
            if (!call.callee) return "";
            if (call.callee->kind == ASTNodeKind::IdentifierExpr) {
                // Fix 085c: 模块内裸函数调用返回类实例 (如 cAsyncSocket 内
                // pvToSocket(idx) As cAsyncSocket), 后续 .frNotifyGetHostByName(...)
                // 链式调用需要知道返回类以拆成 vb6_cAsyncSocket_frNotify...(this,...).
                // 此前仅支持 MemberAccessExpr/WithMemberExpr callee, 裸函数推断断链 →
                // 生成 (ret).Method(...) 非法字段访问 (C2039: 不是 vb6_cls_X 的成员).
                auto& id = static_cast<const IdentifierExpr&>(*call.callee);
                // Fix 192: **类数组元素** arr(i) — 元素类型是项目类时, arr(i) 本身就是
                // 一个类实例 (void* 槽里放的是 vb6_cls_X*), 后续 .Method/.Field 必须按
                // 该类解析。此前这里只查"裸函数调用返回类", 类数组元素漏掉 →
                // 接收者推断不出类 → 通用成员访问发 VB6_SA_AT(void*, arr, i).Move(...)
                // → MSVC C2224 "左侧必须具有结构/联合类型" (void* 取成员)。
                // 静态 `Dim s(1) As C` 与 `ReDim a(1) As C` 都踩同一处。
                // 先查数组表: 名字在 knownArrays_ 里就是下标访问, 不是函数调用。
                {
                    std::string arrLower = Symbol::toLower(id.name);
                    if (knownArrays_.count(arrLower)) {
                        auto itArr = arrayClassElemTypes_.find(arrLower);
                        if (itArr != arrayClassElemTypes_.end()) return itArr->second;
                    }
                }
                const Symbol* fn = symTab_.lookupModule(id.name);
                if (fn && fn->kind == SymbolKind::Function
                    && !fn->variableTypeName.empty()) {
                    const Symbol* clsSym = symTab_.lookupModule(fn->variableTypeName);
                    if (clsSym && clsSym->kind == SymbolKind::Class) return fn->variableTypeName;
                }
                return "";
            }
            if (call.callee->kind != ASTNodeKind::MemberAccessExpr
                && call.callee->kind != ASTNodeKind::WithMemberExpr) return "";
            // Fix 085b: With 块内方法链 .Data(...).CalculateCRC16(...) — callee 是
            // WithMemberExpr, 基类是 With 栈顶对象类, 不能按 MemberAccessExpr 解析
            // (否则 With 链方法在推断中断链, 生成 (ret).Method(...) 非法字段调用).
            if (call.callee->kind == ASTNodeKind::WithMemberExpr) {
                auto& wm = static_cast<const WithMemberExpr&>(*call.callee);
                if (withObjectInfoStack_.empty()) return "";
                const auto& winfo = withObjectInfoStack_.back();
                if (winfo.kind != WithObjKind::ClassInstance || winfo.className.empty()) return "";
                return getClassMethodReturnType(winfo.className, wm.memberName);
            }
            auto& ma = static_cast<const MemberAccessExpr&>(*call.callee);
            if (!ma.object) return "";
            // 递归推断对象表达式的类名
            std::string baseClassName = inferClassTypeOfExpr(*ma.object);
            if (baseClassName.empty()) return "";
            // 查找该方法的返回类型
            const std::string retCls242 = getClassMethodReturnType(baseClassName, ma.memberName);
            if (retCls242.empty()) return "";
            // Fix 242: `<集合>(i)` —— callee 是**无参属性**且返回项目集合类时, `(i)` 不是
            // 该属性的实参, 而是返回的那个集合的**默认成员** (VB_UserMemId = 0, VB6 集合
            // 惯例名 `Item`) 的索引; 表达式类型是**元素类**, 不是集合类本身.
            //
            // 实测 ComCtlsDemo Builds/ToolBar/ToolBar.ctl:4199
            //     With Button.ButtonMenus(i)
            //         If .Visible = True Then … .Separator … .Text … .Picture …
            //   TbrButton.cls:308  Property Get ButtonMenus() As TbrButtonMenus   (无参)
            //   TbrButtonMenus.cls:96 Property Get Item(ByVal Index As Variant) As TbrButtonMenu
            //                        (Attribute Item.VB_UserMemId = 0)
            // 修复前把 With 目标推成 vb6_cls_TbrButtonMenus*, 块内 .Visible/.Separator/
            // .Text/.Picture/.Enabled/.Checked 全按**集合类**找成员 ⇒ C2039×18.
            //
            // 两道闸, 都不命中就维持原返回值 (集合类), 既不劣化既有场景也不乱动
            // `obj.Method(args)` 这类"member 自己带参"的调用:
            //   (a) 确实带索引 (`call.positional` 非空);
            //   (b) member 在 base 类里是**无参**成员 (读方向形参表为空) — 带参属性/
            //       方法走不到这里, `(i)` 就该是它自己的实参, 返回类型仍是 member 的返回类.
            // 元素类取不出 (Item 返回非项目类) 时同样回退.
            if (!call.positional.empty()) {
                std::vector<ParameterInfo> collParams242;
                if (findClassMemberGetParams(baseClassName, ma.memberName, collParams242)
                    && collParams242.empty()) {
                    const std::string elemCls242 = getClassMethodReturnType(retCls242, "Item");
                    if (!elemCls242.empty()) return elemCls242;
                }
            }
            return retCls242;
        }
        // Fix 037: MeExpr → 类模块内 me 即当前类
        case ASTNodeKind::MeExpr: {
            if (isClassModule_) return moduleName_;
            return "";
        }
        // Fix 037b: MemberAccessExpr → 递归推断 object 的类类型, 再从
        // classTypedFieldMap_ 查找字段的类类型 (仅项目类字段, 非 COM).
        // 用于链式访问 ctx.Request.QueryString(idx) 中 Request 的类型推断:
        //   ctx (cHttpServerContext) → Request (cHttpServerRequest) → QueryString (COM:Dictionary)
        case ASTNodeKind::MemberAccessExpr: {
            auto& ma = static_cast<const MemberAccessExpr&>(expr);
            if (!ma.object) return "";
            std::string baseClassName = inferClassTypeOfExpr(*ma.object);
            if (baseClassName.empty()) return "";
            if (classTypedFieldMap_) {
                auto it = classTypedFieldMap_->find(baseClassName);
                if (it != classTypedFieldMap_->end()) {
                    std::string memLower = Symbol::toLower(ma.memberName);
                    auto itF = it->second.find(memLower);
                    if (itF != it->second.end()) {
                        // 仅返回项目类字段类型 (COM: 前缀的不是项目类)
                        if (itF->second.compare(0, 4, "COM:") == 0) return "";
                        return itF->second;
                    }
                }
            }
            // Fix 084z: 属性 Get 回退 — 成员是属性(返回类实例)而非数据字段时
            // (如 pvSocket 是 cTlsReMaster 的 Property Get, 不在字段表中),
            // 用 getClassMethodReturnType 推断返回类. 否则调用方类型推断失败
            // → fallback 到错误类解析参数, 生成 C2197/C2198 (参数过多/过少:
            // SyncReceiveArray 声明6参却按 cWinsock 的12参展开, Connect 声明
            // 11参却按 cWinsock 的5参展开).
            return getClassMethodReturnType(baseClassName, ma.memberName);
        }
        // Fix 037: WithMemberExpr → 当前 With 块 tempVar 的类类型 (仅 ClassInstance kind)
        case ASTNodeKind::WithMemberExpr: {
            if (withObjectInfoStack_.empty() || withObjectVars_.empty()) return "";
            const auto& info = withObjectInfoStack_.back();
            if (info.kind == WithObjKind::ClassInstance && !info.className.empty()) {
                // Fix 090s: .X 若是 With 目标类的 typed 项目类字段 (如
                // With HttpSvr: .Router.Reg → .Router 字段 As cHttpServerRouter),
                // 返回字段的类, 供外层 .Reg/.Encode 等成员/方法按字段类解析
                // (findClassMemberCallParams/resolveClassMemberCall 用对类, 否则
                // 形参表空 → .Router.Reg "Test" 实参全丢 C2198). 与 MemberAccessExpr
                // 分支 (classTypedFieldMap_ 查询) 对齐.
                auto& wmRef = static_cast<const WithMemberExpr&>(expr);
                if (classTypedFieldMap_) {
                    auto it = classTypedFieldMap_->find(info.className);
                    if (it != classTypedFieldMap_->end()) {
                        std::string memLower = Symbol::toLower(wmRef.memberName);
                        auto itF = it->second.find(memLower);
                        if (itF != it->second.end()) {
                            // COM:/void* 字段 (COM: 前缀) 不是项目类 → 返回空让外层
                            // 走 COM dispatch; 项目类字段返回类名
                            if (itF->second.compare(0, 4, "COM:") == 0) return "";
                            return itF->second;
                        }
                    }
                }
                // Fix <vbeclipse>: .X 是 With 目标类的属性(返回项目类)而非字段/
                // 标量 (ucPerspective.ctl `With l_ucFolder` 内 `.Views.IsEmpty`,
                // Views 是 Property Get 返回 List 类) → 用返回类推断, 否则外层
                // .IsEmpty/.Count/.Item 落到类兜底 `->` 字段 → C2039 / Item 乱绑
                // (vb6_ucFolder_prop_get_Views->Item(...)). getClassMethodReturnType
                // 仅对真实项目 Class 返回类名 (String/Long 属性返回 ""), 对纯字段
                // /标量属性维持 info.className 原行为 (Fix 085b 语义不变).
                {
                    std::string retClsFromProp =
                        getClassMethodReturnType(info.className, wmRef.memberName);
                    if (!retClsFromProp.empty()) return retClsFromProp;
                }
                // Fix 246: `.X` 是**返回 As Object/Variant 的成员**时, 它本身是后期绑定
                // 对象, 成员链必须走 COM —— 不能继续沿用 With 目标类。
                // 实测 ComCtlsDemo ToolBar.ctl:4199 `With Button.ButtonMenus(i)` 块内
                // `.Picture.Handle` / `.Picture.Type`: TbrButtonMenu.cls:137
                // `Property Get Picture() As IPictureDisp` (C 侧 void*)。修复前把
                // `.Picture` 推成 vb6_cls_TbrButtonMenu* ⇒ `.Handle` 发成
                // `vb6_TbrButtonMenu_prop_get_Picture(x)->Handle` → C2223×6 (ToolBar.c
                // 7147/7151/7153/7198/7202/7204)。返回空让下游按后期绑定发
                // vb6_ComGetProp(<prop_get_Picture(x)>, L"Handle")。
                if (classMemberReturnsAsObject(info.className, wmRef.memberName)) return "";
                // .X 非数据字段 (方法/属性等) → 维持原行为: With 目标类自身
                // (Fix 085b 在调用链推断处对 callee=WithMemberExpr 已按方法返回类
                // 特判, 此处不做方法返回类型推断以免误伤 String/Long 属性场景)
                return info.className;
            }
            return "";
        }
        default:
            return "";
    }
}


// Fix 219: Variant 载体判定 —— 见 cgen_helpers.inc 的声明注释。
bool CCodeGen::isVariantCarrierExpr(const std::string& cExpr) const {
    return cExprIsVariant(cExpr) || cExpr.rfind("vb6_List_", 0) == 0;
}

// Fix 219: Variant 载体上的成员访问 = 后期绑定的属性读。
std::string CCodeGen::lateBoundReadOnVariantCarrier(const std::string& carrierExpr,
                                                    const std::string& memberName) const {
    return std::string("vb6_VariantFromComResult(vb6_ComGetProp(vb6_VariantToObjectVal(")
         + carrierExpr + "), L\"" + memberName + "\"))";
}



// Fix 037: 递归推断表达式的 UDT C 类型标识符 (如 "vb6_type_UcsBuffer").
// 支持 IdentifierExpr (knownUdtVars_ 直查) 和 MemberAccessExpr (嵌套 UDT 字段递归).
// 返回空串表示非 UDT 表达式.
std::string CCodeGen::inferUdtTypeOfExpr(const ASTNode& expr) const {
    switch (expr.kind) {
        case ASTNodeKind::IdentifierExpr: {
            auto& id = static_cast<const IdentifierExpr&>(expr);
            std::string lower = Symbol::toLower(id.name);
            auto it = knownUdtVars_.find(lower);
            if (it != knownUdtVars_.end()) return it->second;
            // Fix 090j: 函数体内函数名标识符 = 本函数返回对象 (VB6: Function
            // pvVfsOpen As ZipVfsType 内写 pvVfsOpen.BufferArray, 即返回 UDT 的
            // 字段). 与发射层 Fix 084z-4/088d (函数名→类返回对象) 及注册层
            // Fix 090i (vb6_ret_X → knownUdtVars_) 对称: 推断层必须把 "函数名"
            // 映射到返回 UDT 类型, 否则字段类型推断落空 → 字段整体按 Unknown:
            // Variant 字段被当函数 (SourceFileInfo(3) → C2064)、LongPtr 字段被
            // Variant 化 (BufferPtr = BufferBase → VariantToLong → C2440),
            // As Any 实参把 Variant 字段强转指针 (C2440) — cZipArchive VFS 簇.
            if (currentProc_ && !currentReturnCType_.empty()
                && currentReturnCType_.rfind("vb6_type_", 0) == 0
                && lower == Symbol::toLower(currentProc_->name)) {
                return currentReturnCType_;
            }
            return "";
        }
        // Fix 081i: IndexOrCallExpr — UDT数组元素访问 arr(idx).field
        // 查 arrayUdtElemTypes_ 获取数组元素UDT类型
        case ASTNodeKind::IndexOrCallExpr: {
            auto& call = static_cast<const IndexOrCallExpr&>(expr);
            if (call.callee && call.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& id = static_cast<const IdentifierExpr&>(*call.callee);
                std::string lower = Symbol::toLower(id.name);
                auto it = arrayUdtElemTypes_.find(lower);
                if (it != arrayUdtElemTypes_.end()) return it->second;
            }
            // Fix 110c: UDT 数组的**成员**数组 — m_Serie(i).Rects(j).
            // callee 是 MemberAccessExpr(父对象, 成员名), 成员本身是 UDT 数组
            // (如 tSerie.Rects() As RectL) → 元素类型 = 成员的 UDT 类型.
            // 递归推断 callee 的 UDT 类型即可 (tSerie → Rects → RectL).
            if (call.callee && call.callee->kind == ASTNodeKind::MemberAccessExpr) {
                std::string elemUdt = inferUdtTypeOfExpr(*call.callee);
                if (!elemUdt.empty()) return elemUdt;
            }
            // Fix 110s: With 块内的成员 UDT 数组 — With .Rects(j).
            // callee 是 WithMemberExpr (当前 With 对象的字段), 字段本身是 UDT 数组
            // (如 tSerie.Rects() As RectF) → 元素类型 = 字段的 UDT 类型.
            // 此前只处理 MemberAccessExpr, WithMemberExpr 落空 → With 临时变量退化为
            // void* → `(void*)VB6_SA_AT(vb6_type_RectF, ...)` C2440 (ucTreeMaps.c 1878).
            if (call.callee && call.callee->kind == ASTNodeKind::WithMemberExpr) {
                std::string elemUdt = inferUdtTypeOfExpr(*call.callee);
                if (!elemUdt.empty()) return elemUdt;
            }
            return "";
        }
        case ASTNodeKind::MemberAccessExpr: {
            auto& ma = static_cast<const MemberAccessExpr&>(expr);
            if (!ma.object) return "";
            // 递归推断父对象的 UDT 类型
            std::string parentUdtCType = inferUdtTypeOfExpr(*ma.object);
            if (parentUdtCType.empty()) return "";
            // parentUdtCType 形如 "vb6_type_UcsBuffer", 剥前缀得到 UDT 名
            const std::string prefix = "vb6_type_";
            if (parentUdtCType.size() <= prefix.size()
                || parentUdtCType.compare(0, prefix.size(), prefix) != 0) return "";
            std::string udtName = parentUdtCType.substr(prefix.size());
            Symbol* udtSym = symTab_.lookupModule(udtName);
            if (!udtSym || udtSym->kind != SymbolKind::UserDefinedType) return "";
            std::string memLower = Symbol::toLower(ma.memberName);
            for (auto& mi : udtSym->udtMembers) {
                if (Symbol::toLower(mi.name) == memLower) {
                    // 若该成员本身是 UDT (typeRefName 非空且能查到 UserDefinedType 符号)
                    if (!mi.typeRefName.empty()) {
                        Symbol* refSym = symTab_.lookupModule(mi.typeRefName);
                        if (refSym && refSym->kind == SymbolKind::UserDefinedType) {
                            return "vb6_type_" + cIdent(mi.typeRefName);
                        }
                    }
                    return "";  // 成员是标量/数组, 不是嵌套 UDT
                }
            }
            return "";
        }
        // Fix 037: WithMemberExpr → 当前 With 块 tempVar 的 UDT 类型递归.
        // With 块临时变量 (_vb6_with_N) 已被 cgen_stmt.cpp 注册到 knownUdtVars_
        // (仅 WithObjKind::Unknown — UDT — 才注册). 若 tempVar 不是 UDT (ClassInstance/
        // COMObject 等其他 kind), 此处返回空串.
        // 递归: .member 即 With 块 UDT 的某字段; 若该字段本身是嵌套 UDT (typeRefName
        // 在符号表中查到 UserDefinedType), 返回 "vb6_type_<memberUdtName>".
        // 例: With uCtx (UcsTlsContext) 内的 .DecrBuffer (UcsBuffer) → 返回
        // "vb6_type_UcsBuffer"; 让外层 .DecrBuffer.Data(0) 的 IndexOrCallExpr 能
        // 在 UcsBuffer 的 udtMembers 中找到 Data (动态数组成员) 并生成 VB6_SA_AT.
        case ASTNodeKind::WithMemberExpr: {
            if (withObjectInfoStack_.empty() || withObjectVars_.empty()) return "";
            const auto& info = withObjectInfoStack_.back();
            if (info.kind != WithObjKind::Unknown) return "";  // 仅 UDT
            const std::string& tempVar = withObjectVars_.back();
            std::string tempLower = Symbol::toLower(tempVar);
            auto it = knownUdtVars_.find(tempLower);
            if (it == knownUdtVars_.end()) return "";
            const std::string parentUdtCType = it->second;
            const std::string prefix = "vb6_type_";
            if (parentUdtCType.size() <= prefix.size()
                || parentUdtCType.compare(0, prefix.size(), prefix) != 0) return "";
            std::string udtName = parentUdtCType.substr(prefix.size());
            Symbol* udtSym = symTab_.lookupModule(udtName);
            if (!udtSym || udtSym->kind != SymbolKind::UserDefinedType) return "";
            auto& wm = static_cast<const WithMemberExpr&>(expr);
            std::string memLower = Symbol::toLower(wm.memberName);
            for (auto& mi : udtSym->udtMembers) {
                if (Symbol::toLower(mi.name) == memLower) {
                    if (!mi.typeRefName.empty()) {
                        Symbol* refSym = symTab_.lookupModule(mi.typeRefName);
                        if (refSym && refSym->kind == SymbolKind::UserDefinedType) {
                            return "vb6_type_" + cIdent(mi.typeRefName);
                        }
                    }
                    return "";
                }
            }
            return "";
        }
        default:
            return "";
    }
}


// Fix 084n: 推断 target 是否为 UDT 字段链, 是则返回字段 Vb6Type, 否则 Unknown.
// 供赋值语句 (cgen_stmt) 将 Variant RHS 转换为目标字段类型.
// ⚠ 账 #261 订正头注释: 这里**从来没带过 Array 标志** —— 语义层把数组成员记在
// `mi.isArrayDynamic` / `mi.arraySize` 上 (semantic_analyzer_decl_type.cpp:130-136),
// `mi.type` 存的是**元素**档。所以"`Data() As Byte` 这个字段"答出来就是 Byte。
// 想知道这一槽是不是数组本体 (拿标量检查套它会 error 6: 实测 VbQRCodegen 的
// `.Data = baData` 被包成 `vb6_ChkByte(指针)` ⇒ 启动期 Unhandled VB6 Error #6)，
// 由 outIsArray 在**同一趟** udtMembers 走查里带回 —— 不开第二张表、不再抄一遍查表。
Vb6Type CCodeGen::inferUdtFieldVb6Type(const ASTNode* target, bool* outIsArray) const {
    if (outIsArray) *outIsArray = false;
    if (!target) return Vb6Type::Unknown;
    std::string memName;
    if (target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<const MemberAccessExpr&>(*target);
        if (inferUdtTypeOfExpr(*ma.object).empty()) return Vb6Type::Unknown;
        memName = ma.memberName;
    } else if (target->kind == ASTNodeKind::WithMemberExpr) {
        if (withObjectInfoStack_.empty() || withObjectVars_.empty()) return Vb6Type::Unknown;
        const auto& info = withObjectInfoStack_.back();
        if (info.kind != WithObjKind::Unknown) return Vb6Type::Unknown;  // 仅 UDT
        auto& wm = static_cast<const WithMemberExpr&>(*target);
        memName = wm.memberName;
    } else {
        return Vb6Type::Unknown;
    }
    if (memName.empty()) return Vb6Type::Unknown;

    // 解析对象 UDT C 类型
    std::string udtCType;
    if (target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<const MemberAccessExpr&>(*target);
        udtCType = inferUdtTypeOfExpr(*ma.object);
    } else {
        const std::string& tempVar = withObjectVars_.back();
        auto it = knownUdtVars_.find(Symbol::toLower(tempVar));
        if (it == knownUdtVars_.end()) return Vb6Type::Unknown;
        udtCType = it->second;
    }
    const std::string prefix = "vb6_type_";
    if (udtCType.size() <= prefix.size() || udtCType.compare(0, prefix.size(), prefix) != 0)
        return Vb6Type::Unknown;
    std::string udtName = udtCType.substr(prefix.size());
    Symbol* udtSym = symTab_.lookupModule(udtName);
    if (!udtSym || udtSym->kind != SymbolKind::UserDefinedType) return Vb6Type::Unknown;
    std::string memLower = Symbol::toLower(memName);
    for (auto& mi : udtSym->udtMembers) {
        if (Symbol::toLower(mi.name) == memLower) {
            if (outIsArray) *outIsArray = (mi.isArrayDynamic || mi.arraySize > 0);
            return mi.type;
        }
    }
    return Vb6Type::Unknown;
}

// 账 #262: 定长成员数组的下标折算 —— 秩在这里，层数就只在这里定。
// 发码形态是 C 的多维数组 (`float M[5][5]`)，所以每层一个括号、步长归 C 算，
// 不需要另开一张"每行几格"的表 (那张表迟早和 cgen_decl.cpp 的折尺寸分家)。
std::string CCodeGen::udtFixedMemberIndex(const std::string& fieldExpr,
                                          const std::vector<ExprPtr>& args, int32_t rank) {
    int32_t layers = rank < 1 ? 1 : rank;
    std::string out = fieldExpr;
    for (int32_t k = 0; k < layers; k++) {
        std::string idx;
        if (k < (int32_t)args.size() && args[k]) {
            emitExpr(*args[k]);   // ⚠ emitExpr 覆写 lastExpr_，每层当场取走
            idx = std::move(lastExpr_);
            if (idx.empty()) idx = "0";   // 空表达式不发 `field[]`
        } else {
            idx = "0";            // 缺的那层 (含省略实参): 给一个确定的在范围内的格子 (VB6 此处报下标越界)
        }
        out += "[" + idx + "]";
    }
    return out;
}

// <vbeclipse> <VBFlexGridDemo>: 返回左值 UDT 字段的**声明类型名** (`mi.typeRefName`)。
// 空串 = 无声明类型 (真 Variant 槽), 非空 = 声明了具体类型 (Object/接口/类/UDT)。
// 与 Set-RHS 装箱 guard 配套 —— VTableHandle.bas `Set VTableIPAOData.OriginalIOleIPAO = This`
// 一条, 字段声明 `As OLEGuids.IOleInPlaceActiveObject`, TLB 命中的 caller 侧
// resolveTypeRef 把 mi.type 兜底成 Variant (同 arg_emit Fix 210 那条同源问题), 若
// 装箱就发出 `field = vb6_VariantFromValue(This)` → 字段是 typed ptr 收 VARIANT →
// C2440。用**声明名**做判据: 只要非空, 这条槽就不是 Variant 容器。
std::string CCodeGen::udtFieldTypeRefNameOfTarget(const ASTNode* target) const {
    if (!target) return std::string();
    std::string udtCType, memName;
    if (target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<const MemberAccessExpr&>(*target);
        if (!ma.object) return std::string();
        udtCType = inferUdtTypeOfExpr(*ma.object);
        memName = ma.memberName;
    } else if (target->kind == ASTNodeKind::WithMemberExpr) {
        if (withObjectInfoStack_.empty() || withObjectVars_.empty()) return std::string();
        const auto& info = withObjectInfoStack_.back();
        if (info.kind != WithObjKind::Unknown) return std::string();
        memName = static_cast<const WithMemberExpr&>(*target).memberName;
        auto it = knownUdtVars_.find(Symbol::toLower(withObjectVars_.back()));
        if (it == knownUdtVars_.end()) return std::string();
        udtCType = it->second;
    } else {
        return std::string();
    }
    const std::string prefix = "vb6_type_";
    if (memName.empty() || udtCType.size() <= prefix.size()
        || udtCType.compare(0, prefix.size(), prefix) != 0)
        return std::string();
    std::string udtName = udtCType.substr(prefix.size());
    Symbol* udtSym = symTab_.lookupModule(udtName);
    if (!udtSym || udtSym->kind != SymbolKind::UserDefinedType) return std::string();
    std::string memLower = Symbol::toLower(memName);
    for (auto& mi : udtSym->udtMembers) {
        if (Symbol::toLower(mi.name) == memLower) return mi.typeRefName;
    }
    return std::string();
}


// ai/022 B08f-1 (D37): 与上一条走同一条 UDT 解析路, 但返回字段在 **C 里的对象类型**
// (`vb6_cls_X*` / `void*` / "" = 不是对象字段或推不出)。为什么不直接改 inferUdtFieldVb6Type
// 的返回值: `As <项目类>` 的 UDT 成员在语义层就是 Variant，而 Vb6Type::Variant 另有消费者
// （实参打包、Let 赋值语义），改它等于同时改那几处行为。这里要回答的问题只有一个 ——
// "这个左值到底是不是 Variant 容器" —— 那只有字段的真实 C 类型说了算。
std::string CCodeGen::udtFieldCTypeOfTarget(const ASTNode* target) const {
    if (!target) return std::string();
    std::string udtCType, memName;
    if (target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<const MemberAccessExpr&>(*target);
        if (!ma.object) return std::string();
        udtCType = inferUdtTypeOfExpr(*ma.object);
        memName = ma.memberName;
    } else if (target->kind == ASTNodeKind::WithMemberExpr) {
        if (withObjectInfoStack_.empty() || withObjectVars_.empty()) return std::string();
        const auto& info = withObjectInfoStack_.back();
        if (info.kind != WithObjKind::Unknown) return std::string();  // 仅 UDT
        memName = static_cast<const WithMemberExpr&>(*target).memberName;
        auto it = knownUdtVars_.find(Symbol::toLower(withObjectVars_.back()));
        if (it == knownUdtVars_.end()) return std::string();
        udtCType = it->second;
    } else {
        return std::string();
    }
    const std::string prefix = "vb6_type_";
    if (memName.empty() || udtCType.size() <= prefix.size()
        || udtCType.compare(0, prefix.size(), prefix) != 0)
        return std::string();
    return udtFieldObjCType(udtCType, Symbol::toLower(memName));
}


// ============================================================
// Fix 085: UDT 对象字段类型推断
// ============================================================
std::string CCodeGen::udtFieldObjCType(const std::string& udtCType,
                                       const std::string& memberLower) const {
    const std::string prefix = "vb6_type_";
    if (udtCType.size() <= prefix.size() || udtCType.compare(0, prefix.size(), prefix) != 0)
        return "";
    std::string udtName = udtCType.substr(prefix.size());
    Symbol* udtSym = symTab_.lookupModule(udtName);
    if (!udtSym || udtSym->kind != SymbolKind::UserDefinedType) return "";
    std::string memLower = Symbol::toLower(memberLower);
    for (const auto& mi : udtSym->udtMembers) {
        if (Symbol::toLower(mi.name) != memLower) continue;
        if (mi.type == Vb6Type::UserDefinedType) {
            // 嵌套 UDT 字段 (非对象) — 返回其 UDT C 类型供链式推断
            return mi.typeRefName.empty() ? "" : "vb6_type_" + cIdent(mi.typeRefName);
        }
        if (mi.type == Vb6Type::Object) {
            if (!mi.typeRefName.empty()) {
                // VBA. 前缀剥离 (如 VBA.Collection → Collection)
                std::string tn = mi.typeRefName;
                if (tn.size() > 4 && tn.compare(0, 4, "VBA.") == 0) tn = tn.substr(4);
                Symbol* refSym = symTab_.lookupModule(tn);
                if (refSym && refSym->kind == SymbolKind::Class) {
                    // 项目类对象字段 → 类方法/属性调度 (early bound).
                    // 注意: C 层类类型名须用类的规范模块名 (sourceModule), 而非 UDT
                    // 字段中的引用名 (如 tZipFileItem.SourceArchive As "ZipArchive" 实际
                    // 对应 vb6_cls_cZipArchive* — 引用名可能与模块名不同, 用了引用名
                    // 会让 resolveClassMemberCall 查不到成员 (类符号按模块名登记)).
                    // sourceModule 仅在跨模块注入时填写; 类自身符号(当前类模块
                    // 编译中)为空, 此时规范名即当前模块名 moduleName_.
                    std::string clsCanon = !refSym->sourceModule.empty()
                                               ? refSym->sourceModule
                                               : moduleName_;
                    return "vb6_cls_" + cIdent(clsCanon) + "*";
                }
            }
            // Collection/COM/接口 等对象字段 → COM dispatch
            return "void*";
        }
        // ai/022 B08f-1 (D37): 跨模块的 `As <项目类>` 在语义层落到 Variant 兜底 (类名靠
        // semantic_analyzer_decl_type.cpp 那条新增分支存进 typeRefName)。这里按**当前**
        // (stage 3.5 之后，Class 符号已注入)的符号表回判：认得出工程类才当对象字段，
        // 认不出照旧返回 "" —— 现在标量/真 Variant 字段也带名字了，不能顺手误判成对象。
        if (mi.type == Vb6Type::Variant && !mi.typeRefName.empty()) {
            std::string tn = mi.typeRefName;
            if (tn.size() > 4 && tn.compare(0, 4, "VBA.") == 0) tn = tn.substr(4);
            Symbol* refSym = symTab_.lookupModule(tn);
            if (refSym && refSym->kind == SymbolKind::Class) {
                std::string clsCanon = !refSym->sourceModule.empty()
                                           ? refSym->sourceModule
                                           : moduleName_;
                return "vb6_cls_" + cIdent(clsCanon) + "*";
            }
            // Fix <VBFlexGridDemo>: 与 struct **字段声明侧同源**的一条判据 —— 拿
            // mapTypeRef (定义侧权威, 见 cgen_base_type.cpp:340 ComInterface 分支,
            // 未识别类型 line 478 兜底 return "void*") 会**发**出的那份 C 类型决定
            // marker 归属. 只要字段 C 类型是**指针**, 语义 = 对象槽, 返回 "void*"
            // 让 appendUdtObjFieldMarker 追加 `/* udt objfield void* */` 标记 →
            // 外层 MemberAccessExpr 消费走 vb6_ComCall 晚绑定
            // (cgen_expr_member_generic_access.inc:222 那一支)。
            // CI 现场 (diag212b 的 VTableHandle.h + c3-error.log 双 dump 实证):
            // GA windows-latest runner **没有** SysWOW64\OLEGuids.tlb, 且注册表
            // {5A2B9220-...} 也没登记 (VB4001 warning ×2 直证), 于是 driver
            // 层 typelibRefs 双 miss (loadByClsid + loadByPath) → TLB 里的
            // `IOleInPlaceActiveObject` 从未进符号表 → semantic_analyzer_typeref.cpp:99/116
            // 两处 lookupModule 也 miss → resolveTypeRef 落 line 134 Variant 兜底,
            // typeRefName 存 "OLEGuids.IOleInPlaceActiveObject"。同份字段 mapTypeRef
            // 也 miss → line 478 兜底 → 结构里的字段 C 类型就是 **void***。旧
            // udtFieldObjCType 这条 Variant 分支返回 "" (无 marker) → 外层 MAE
            // fallback (class_fallback.inc:297/300) 发裸 `(*This).OriginalIOleIPAO
            // .QueryInterface(...)` → C2224 ×13 + 级联 C2197/C2198 (VTableHandle.c
            // 459/462/470/476/482/488/510/512/516/524/530/536/542, GA t2 2/2).
            // local 同一份代码走的是**不同**的 TLB 载入结果 (SysWOW64 文件在、
            // 注册表 GUID 在): TLB load → ComInterface 进模块符号表 → pass 1
            // 认得 → mi.type = Object → line 494 分支返回 "void*" → marker 生效
            // → 49ed1b22 那条 lookupTypeSymbol(短名) 兜底本地也不 fire (它已经
            // 从 Object 分支出); 而 CI 上 lookupTypeSymbol 也 miss (TLB 没 load
            // 就没这个符号) —— 所以判据**不能靠符号命中**, 得回到"限定名 (含 '.')
            // 在 VB6 语境里就是 COM 引用库的类型"这条**语法**事实: 无论符号表有
            // 没有, 只要 mapTypeRef 会兜底 "void*", 字段 C 类型就是 void*。
            if (tn.find('.') != std::string::npos) return "void*";
            return "";
        }
        // Fix 177: String 字段 → "BSTR"。调用方 appendUdtObjFieldMarker 只对
        // "void*"/"vb6_cls_*" 追加对象标记, 故新增此返回不影响既有分派;
        // 供 udtFieldIsBstrInCTarget 判定"该字段赋值必须走 vb6_BSTR_Assign 深拷贝"。
        // Fix <vbeclipse> 2026-10-06: 数组成员 (`parts() As String`) 的 mi.type 是
        // String|Array 位组合, 上面的 == 判等天然排除, 其 C 类型是
        // vb6_SafeArray1D* 不是 BSTR, 另走动态数组链路。
        if (mi.type == Vb6Type::String) return "BSTR";
        // 标量/数组等非对象字段
        return "";
    }
    return "";
}


// Fix 177: 从**已生成的 C 目标串**判断"UDT 字段是否为 String"。
// 场景: `VBFlexGridCells.Rows(iRow).Cols(iCol).Text = TextIn` 发射成
//   `VB6_SA_AT(vb6_type_TCELL, VB6_SA_AT(vb6_type_TCOLS, me->..Rows, iRow).Cols, iCol).Text`
// 这类两层 UDT 数组链在 AST 侧解析不出来 (类字段不在 knownUdtVars_, 且
// `Cols() As TCELL` 的成员类型带 Array 标志), 于是 String 字段退化成
// **裸指针赋值** → 存进去的是调用方临时 BSTR 的地址, 返回即悬垂
// (demo 表现: 除前两行外整表读到同一个/垃圾值)。
// 这里改为直接读宏的第一个实参 (元素 UDT 名) + 宏右部的成员名。
bool CCodeGen::udtFieldIsBstrInCTarget(const std::string& target) const {
    // 目标形如 `VB6_SA_AT(vb6_type_TCELL, VB6_SA_AT(vb6_type_TCOLS, me->..Rows,
    // iRow).Cols, iCol).Text`。成员名 = 末尾的 `.Name`; 元素 UDT = **外层**宏的
    // 第一个实参 —— 不能用"第一个逗号", 因为外层实参②自身含嵌套宏的逗号,
    // 必须按括号深度配平后再取逗号。
    size_t dot = target.rfind('.');
    if (dot == std::string::npos || dot + 1 >= target.size()) return false;
    size_t nameEnd = dot + 1;
    while (nameEnd < target.size() &&
           (isalnum((unsigned char)target[nameEnd]) || target[nameEnd] == '_')) ++nameEnd;
    if (nameEnd != target.size() || nameEnd == dot + 1) return false;
    if (target[dot - 1] != ')') return false;
    std::string member = target.substr(dot + 1, nameEnd - dot - 1);

    // 成员所属的宏 = 闭合括号**紧贴** `.` 的那个候选 (即最内层)。逐候选按宏起始
    // 位置推进 (find(macro, cand+1) 会跳过宏名本身从而漏掉嵌套内层); 也不能取
    // "第一个能配平的候选" —— 外层的深度 0 收尾同样合法, 会把元素类型误取成外层
    // 的 TCOLS 而非内层的 TCELL。
    static const std::string macro = "VB6_SA_AT(";
    size_t open = std::string::npos;
    size_t close = std::string::npos;
    for (size_t cand = target.find(macro); cand != std::string::npos && cand < dot;
         cand = target.find(macro, cand + macro.size())) {
        int depth = 1;                       // 宏名自带的 '(' 已被跳过, 未计入
        for (size_t k = cand + macro.size(); k < dot; ++k) {
            if (target[k] == '(') ++depth;
            else if (target[k] == ')' && --depth == 0) {
                if (k + 1 == dot) { open = cand; close = k; }
                break;
            }
        }
    }
    if (open == std::string::npos) return false;
    size_t args = open + macro.size();
    size_t comma = target.find(',', args);
    if (comma == std::string::npos || comma >= close) return false;
    std::string elemCType = target.substr(args, comma - args);
    while (!elemCType.empty() && isspace((unsigned char)elemCType.back())) elemCType.pop_back();
    if (elemCType.rfind("vb6_type_", 0) != 0) return false;
    return udtFieldObjCType(elemCType, Symbol::toLower(member)) == "BSTR";
}


// ---- 账 #118: 定长串 (String * N) 的长度查询 ----
// 见 cgen_helpers.inc 的声明注释。返回非空 = 是定长串, 值即长度表达式。
std::string CCodeGen::fixedStrLenOfExpr(const ASTNode& expr,
                                        const std::string& altCName) const {
    // ① 裸标识符: 模块级变量 / 局部 / 形参 / 返回值槽
    if (expr.kind == ASTNodeKind::IdentifierExpr) {
        const std::string lower = Symbol::toLower(
            static_cast<const IdentifierExpr&>(expr).name);
        auto it = knownFixedStringLen_.find(lower);
        if (it != knownFixedStringLen_.end()) return it->second;
        if (!altCName.empty()) {
            auto itAlt = knownFixedStringLen_.find(Symbol::toLower(altCName));
            if (itAlt != knownFixedStringLen_.end()) return itAlt->second;
        }
        return "";
    }
    // ② UDT 字段: `r.f` —— 先求对象是哪个 UDT (inferUdtTypeOfExpr 认得裸 UDT 变量、
    //    嵌套 UDT 字段、SA_AT 数组元素等既有形状), 再查该 UDT 的定长字段表。
    //    刻意**不问**字段名跨 UDT 是否唯一: 类型来自对象, 同名不同 UDT 各查各的。
    if (expr.kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<const MemberAccessExpr&>(expr);
        if (!ma.object) return "";
        const std::string objUdt = inferUdtTypeOfExpr(*ma.object);
        if (objUdt.rfind("vb6_type_", 0) != 0) return "";
        auto uit = udtFixedStrFieldLen_.find(objUdt);
        if (uit == udtFixedStrFieldLen_.end()) return "";
        auto mit = uit->second.find(Symbol::toLower(ma.memberName));
        if (mit == uit->second.end()) return "";
        return mit->second;
    }
    return "";
}


std::string CCodeGen::appendUdtObjFieldMarker(const std::string& objExpr,
                                              const std::string& udtCType,
                                              const std::string& member,
                                              const std::string& accessOp) const {
    std::string fieldCType = udtFieldObjCType(udtCType, Symbol::toLower(member));
    std::string fieldAccess = objExpr + accessOp + cIdent(member);
    // 仅对象字段 (项目类 vb6_cls_* / Collection·COM void*) 才追加标记;
    // 嵌套 UDT (vb6_type_*) 与标量/字符串等原样返回 — 嵌套 UDT 继续由
    // inferUdtTypeOfExpr / 普通字段拼接处理, 标记残留会干扰函数参数等上下文.
    if (fieldCType != "void*" && fieldCType.rfind("vb6_cls_", 0) != 0) return fieldAccess;
    return fieldAccess + "  /* udt objfield " + fieldCType + " */";
}



// Fix 084o: 需要 int32_t 上下文中的 Variant 表达式 → vb6_VariantToLong 包装
std::string CCodeGen::toLongIfVariant(const std::string& cExpr, const Expr* astExpr) {
    if (cExprIsVariant(cExpr)) return "vb6_VariantToLong(" + cExpr + ")";
    if (astExpr && astExpr->kind == ASTNodeKind::IdentifierExpr) {
        auto& ident = static_cast<IdentifierExpr&>(const_cast<Expr&>(*astExpr));
        std::string lower = ident.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownVariantVars_.count(lower)) return "vb6_VariantToLong(" + cExpr + ")";
    }
    return cExpr;
}

// Fix 156: 需要 double 上下文中的 Variant 表达式 → vb6_VariantToDouble 包装.
// VB6 浮点除法 `/` 生成 `((double)(L) / (double)(R))`; 当 L/R 是 vb6_VARIANT
// 结构体时 C 强制转换非法 → C2440 "函数/类型强制转换表达式: 无法从
// “vb6_VARIANT”转换为“double”".
//   VBFlexGrid.ctl Property Let FloatFromVariant 系: Select Case VarType(Value)
//     Case vbDouble: Int64 = Value / 10000   (Value 是 ByVal Variant)
//     → 原生成 Int64 = ((double)(Value) / (double)(10000.0)); C2440.
// 与 toLongIfVariant 同构 (整除 \ 已在 Fix 084o 用同款处理), 仅提取函数换成
// vb6_VariantToDouble. 未判为 Variant 时原样返回, 故对今天能编译的表达式零影响.
std::string CCodeGen::toDoubleIfVariant(const std::string& cExpr, const Expr* astExpr) {
    if (cExprIsVariant(cExpr)) return "vb6_VariantToDouble(" + cExpr + ")";
    if (astExpr && astExpr->kind == ASTNodeKind::IdentifierExpr) {
        auto& ident = static_cast<IdentifierExpr&>(const_cast<Expr&>(*astExpr));
        std::string lower = ident.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownVariantVars_.count(lower)) return "vb6_VariantToDouble(" + cExpr + ")";
    }
    return cExpr;
}
// ============================================================
// Fix 178: 含所有权成员 (String / 动态数组 / Variant / 嵌套 UDT) 的 UDT 深拷贝
// ============================================================
// VB6 运行时按类型描述符复制 UDT: `b = a` 与 `LSet b = a` 之后 b 拥有**自己的一份**
// String/子数组/Variant。C 端 `b = a` (以及 LSet 落的 memcpy) 是浅拷贝, 成员指针被
// 两个变量共享 → 一处写、多处变。VBFlexGridDemo 的 PropRows Let 用
//   LSet VBFlexGridCells.Rows(i) = VBFlexGridDefaultCols
// 填新行, 于是行 2..149 的 Cols 载体全部别名到行 0 (探针: 写 (9,15)="NINE" 后
// 读 0/2/9/11 行四格同值)。这里按符号表的成员元数据为每个此类 UDT 生成一个 static
// 拷贝函数: 标量直赋、String 复制、动态数组走 vb6_ArrayAssign1D_Cb、嵌套 UDT 递归。
// 对象/接口成员维持按位 (与既有代码同口径, 不引入 AddRef/Release 失衡)。

namespace {

// "vb6_type_TCELL" → "TCELL"; 非 UDT C 类型串返回空
std::string udtNameOfCType(const std::string& udtCType) {
    const std::string prefix = "vb6_type_";
    if (udtCType.size() <= prefix.size() || udtCType.compare(0, prefix.size(), prefix) != 0)
        return "";
    return udtCType.substr(prefix.size());
}

// 剥掉 Array/ByRef 位标志, 取基础类型
Vb6Type udtBaseType(Vb6Type t) {
    uint16_t v = static_cast<uint16_t>(t);
    v &= ~static_cast<uint16_t>(Vb6Type::Array);
    v &= ~static_cast<uint16_t>(Vb6Type::ByRef);
    return static_cast<Vb6Type>(v);
}

} // namespace

// 符号表里的 UDT 符号 (vb6_type_X ← X); 查不到返回 nullptr
static const Symbol* udtSymbolOf(const SymbolTable& symTab, const std::string& udtCType) {
    std::string name = udtNameOfCType(udtCType);
    if (name.empty()) return nullptr;
    const Symbol* s = symTab.lookupModule(name);
    if (s && s->kind == SymbolKind::UserDefinedType) return s;
    return nullptr;
}

bool CCodeGen::udtHasOwnedMembersR(const std::string& udtCType,
                                   std::vector<std::string>& stack) const {
    const Symbol* sym = udtSymbolOf(symTab_, udtCType);
    if (!sym) return false;                     // 无元数据 → 不敢当所有权处理
    for (const auto& s : stack) if (s == udtCType) return false;  // 自引用防御
    stack.push_back(udtCType);
    bool owned = false;
    for (const auto& mi : sym->udtMembers) {
        Vb6Type bt = udtBaseType(mi.type);
        // 定长数组成员与动态数组成员: 载体本身/元素可能持所有权
        if (mi.isArrayDynamic) { owned = true; break; }
        if (bt == Vb6Type::String || bt == Vb6Type::Variant) { owned = true; break; }
        if (bt == Vb6Type::UserDefinedType && !mi.typeRefName.empty()) {
            if (udtHasOwnedMembersR("vb6_type_" + cIdent(mi.typeRefName), stack)) {
                owned = true; break;
            }
        }
    }
    stack.pop_back();
    return owned;
}

bool CCodeGen::udtHasOwnedMembers(const std::string& udtCType) const {
    std::vector<std::string> stack;
    return udtHasOwnedMembersR(udtCType, stack);
}

void CCodeGen::requestUdtCopy(const std::string& udtCType, bool needVoidPtrWrapper) const {
    if (udtNameOfCType(udtCType).empty()) return;
    auto it = udtCopyRequested_.find(udtCType);
    if (it == udtCopyRequested_.end()) {
        udtCopyRequested_[udtCType] = needVoidPtrWrapper;
    } else if (needVoidPtrWrapper) {
        it->second = true;   // 只能升级: 有一处当数组元素用就需要 void* 版
    }
}

std::string CCodeGen::udtDeepCopyAssign(const Expr& targetNode, const std::string& target,
                                        const Expr* valueNode,
                                        const std::string& value) const {
    if (value.empty() || target.empty()) return "";
    // 整体数组引用 (`A()` / `.Cols()`) 不是 UDT 结构体而是载体指针: 那里
    // inferUdtTypeOfExpr 给出的是**元素** UDT 类型, 按结构体深拷贝会把指针当结构体拷。
    // 该形态由 Fix 170/178 的 wrapWholeArrayAssign (vb6_ArrayAssign1D_Cb) 负责。
    if (isWholeArrayRef(&targetNode) || isWholeArrayRef(valueNode)) return "";
    // 目标 UDT: 优先从已生成的 C 串解析 (下标宏最具体), 退回 AST 推断
    std::string tgtUdt;
    if (target.compare(0, 7, "VB6_SA_") == 0) {
        size_t vp = target.find("vb6_type_");
        if (vp != std::string::npos) {
            size_t end = target.find_first_of(",)", vp);
            if (end != std::string::npos) tgtUdt = target.substr(vp, end - vp);
        }
    }
    if (tgtUdt.empty()) tgtUdt = inferUdtTypeOfExpr(targetNode);
    if (udtNameOfCType(tgtUdt).empty()) return "";
    if (!valueNode) return "";
    std::string valUdt = inferUdtTypeOfExpr(*valueNode);
    if (valUdt != tgtUdt) return "";             // 异型 UDT / 非 UDT → 维持原路径
    if (!udtHasOwnedMembers(tgtUdt)) return "";
    requestUdtCopy(tgtUdt, false);
    return "vb6_udtcpy_" + udtNameOfCType(tgtUdt) + "(&" + target + ", &" + value
           + ");  /* Fix 178: UDT 深拷贝 */";
}

// 本模块请求到的全部 UDT 拷贝函数 (含依赖闭包)。先给前向声明再给定义, 因此集合内
// 的相互引用与输出顺序无关。
std::string CCodeGen::emitUdtCopyBlock() const {
    if (udtCopyRequested_.empty()) return "";

    // udtCType → 是否需要 void* 薄封装 (作为 vb6_udt_elem_copy 传入)
    std::map<std::string, bool> need;
    std::vector<std::string> work;
    std::set<std::string> queued;
    for (const auto& kv : udtCopyRequested_) {
        need[kv.first] = kv.second;
        work.push_back(kv.first);
        queued.insert(kv.first);
    }
    while (!work.empty()) {
        std::string udtCType = work.back();
        work.pop_back();
        const Symbol* sym = udtSymbolOf(symTab_, udtCType);
        if (!sym) continue;
        for (const auto& mi : sym->udtMembers) {
            if (udtBaseType(mi.type) != Vb6Type::UserDefinedType || mi.typeRefName.empty())
                continue;
            std::string sub = "vb6_type_" + cIdent(mi.typeRefName);
            if (udtNameOfCType(sub).empty() || !udtHasOwnedMembers(sub)) continue;
            bool asElemCb = mi.isArrayDynamic;   // 动态数组元素 → 回调要 void* 版
            auto it = need.find(sub);
            if (it == need.end()) need[sub] = asElemCb;
            else if (asElemCb) it->second = true;
            if (queued.insert(sub).second) work.push_back(sub);
        }
    }

    std::string out;
    out += "\n/* === Fix 178: 含所有权成员的 UDT 深拷贝 (自动生成) === */\n";
    for (const auto& kv : need) {
        std::string n = udtNameOfCType(kv.first);
        out += "static void vb6_udtcpy_" + n + "(vb6_type_" + n + "* d, const vb6_type_"
             + n + "* s);\n";
        if (kv.second)
            out += "static void vb6_udtcpy_" + n + "_v(void* d, const void* s);\n";
    }
    out += "\n";

    for (const auto& kv : need) {
        const std::string& udtCType = kv.first;
        std::string n = udtNameOfCType(udtCType);
        const Symbol* sym = udtSymbolOf(symTab_, udtCType);
        out += "static void vb6_udtcpy_" + n + "(vb6_type_" + n + "* d, const vb6_type_"
             + n + "* s) {\n";
        out += "    if (d == s) return;\n";
        if (!sym) {                       // 元数据缺失: 退化成按位 (不会误释放)
            out += "    *d = *s;\n}\n\n";
            if (kv.second)
                out += "static void vb6_udtcpy_" + n + "_v(void* d, const void* s) {\n"
                       "    vb6_udtcpy_" + n + "((vb6_type_" + n + "*)d, (const vb6_type_"
                     + n + "*)s);\n}\n\n";
            continue;
        }
        for (const auto& mi : sym->udtMembers) {
            std::string fn = cIdent(mi.name);
            if (fn.empty()) continue;
            std::string dl = "d->" + fn, sl = "s->" + fn;
            Vb6Type bt = udtBaseType(mi.type);
            std::string subUdt = (bt == Vb6Type::UserDefinedType && !mi.typeRefName.empty())
                                     ? ("vb6_type_" + cIdent(mi.typeRefName)) : "";
            bool subOwned = !subUdt.empty() && udtHasOwnedMembers(subUdt);
            bool scalarOwned = (bt == Vb6Type::String || bt == Vb6Type::Variant);

            if (mi.arraySize > 0) {
                // 定长数组成员: 发射形态是 `T Name[...]` (账 #262 起多维也是 —— `T Name[r][c]`)，
                // 不能整体赋值。
                if (!scalarOwned && !subOwned) {
                    out += "    memcpy(" + dl + ", " + sl + ", sizeof(" + dl + "));\n";
                    continue;
                }
                // 账 #262/#265: 逐格走**扁平元素指针**，秩有几维都只数一遍格子。
                // 旧写法是 `sizeof(d->S) / sizeof(d->S[_i])` —— 两个毛病: ① `_i` 写在
                // 它自己那条声明**之前** ⇒ C2065 (实物: 任何带定长 String/Variant 数组成员
                // 且会被赋值的 UDT 都编不过，语料里 0 处所以从没响过)；② 多维成员那样数出来
                // 的是**行数**，会把 4 格当 2 格拷。`sizeof(*_dp)` 由 C 按元素类型算，与秩无关。
                std::string elemC = (bt == Vb6Type::String) ? "BSTR"
                              : (bt == Vb6Type::Variant) ? "vb6_VARIANT" : subUdt;
                if (elemC.empty()) continue;   // 认不出元素档就不动这一格 (按位拷已过)
                out += "    { " + elemC + "* _dp = (" + elemC + "*)&d->" + fn
                     + "; const " + elemC + "* _sp = (const " + elemC + "*)&s->" + fn + ";\n";
                out += "      int32_t _i; int32_t _n = (int32_t)(sizeof(d->" + fn
                     + ") / sizeof(*_dp));\n";
                out += "      for (_i = 0; _i < _n; _i++) {\n";
                if (bt == Vb6Type::String) {
                    out += "        BSTR _o = _dp[_i]; _dp[_i] = _sp[_i] ? SysAllocString(_sp[_i]) : NULL;\n"
                           "        if (_o && _o != _sp[_i]) vb6_BSTR_Free(_o);\n";
                } else if (bt == Vb6Type::Variant) {
                    out += "        vb6_VariantClear(&_dp[_i]); vb6_VariantCopy(&_dp[_i], &_sp[_i]);\n";
                } else {
                    out += "        vb6_udtcpy_" + udtNameOfCType(subUdt) + "(&_dp[_i], &_sp[_i]);\n";
                }
                out += "      }\n    }\n";
                continue;
            }
            if (mi.isArrayDynamic) {
                // 动态数组成员: 载体必须各持一份; 元素含所有权时交给回调逐个深拷贝
                std::string cb = "NULL";
                if (subOwned) cb = "vb6_udtcpy_" + udtNameOfCType(subUdt) + "_v";
                out += "    { vb6_SafeArray1D* _o = d->" + fn + ";\n"
                       "      d->" + fn + " = vb6_ArrayAssign1D_Cb(NULL, s->" + fn + ", "
                     + cb + ");\n"
                       "      if (_o && _o != s->" + fn + ") vb6_SafeArrayDestroy1D(_o);\n"
                       "    }\n";
                continue;
            }
            if (bt == Vb6Type::String) {
                out += "    { BSTR _o = d->" + fn + "; d->" + fn
                     + " = s->" + fn + " ? SysAllocString(s->" + fn + ") : NULL;\n"
                       "      if (_o && _o != s->" + fn + ") vb6_BSTR_Free(_o);\n    }\n";
                continue;
            }
            if (bt == Vb6Type::Variant) {
                out += "    vb6_VariantClear(&" + dl + "); vb6_VariantCopy(&" + dl + ", &"
                     + sl + ");\n";
                continue;
            }
            if (subOwned) {
                out += "    vb6_udtcpy_" + udtNameOfCType(subUdt) + "(&" + dl + ", &" + sl
                     + ");\n";
                continue;
            }
            out += "    " + dl + " = " + sl + ";\n";   // 标量/对象指针: 按位
        }
        out += "}\n";
        if (kv.second)
            out += "static void vb6_udtcpy_" + n + "_v(void* d, const void* s) {\n"
                   "    vb6_udtcpy_" + n + "((vb6_type_" + n + "*)d, (const vb6_type_"
                 + n + "*)s);\n}\n";
        out += "\n";
    }
    return out;
}

} // namespace vb6c3
