#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_decl_proc.cpp: 过程声明生成（SubDecl / FunctionDecl） ---


// ============================================================
// 声明 visit 方法
// ============================================================

void CCodeGen::visit(SubDecl& node) {
    // 泛型模板 (tB, G2/G3): 模板本体不发码 (泛型器注入特化副本)
    if (!node.typeParams.empty()) return;
    // ai/vb-asm-extension-spec: Asm 块过程 → x64 独立 MASM 过程 / x86 内联 __asm 块
    if (tryEmitAsmProc(node.name, node.access, node.params, nullptr, node.body, node.loc,
                       node.isNaked)) return;

    // Fix <vbeclipse> rev17: WithEvents 事件处理器 (<控件>_<事件名>) 的形参一律按**事件
    // ABI = ByVal** 发。事件是"跨对象边界"的回调, emitEventSink 早就立了这条规矩
    // (`pi.isByVal = true` + typedef 用 mapTypeRef 的单指针), RaiseEvent 发送侧也照它发
    // (`onXxx(handler, (*Editor))`) —— 唯独处理器的**定义**是从它自己的 VB 签名来的:
    // `Sub ucPerspective1_OpenEditor(Editor As Object)` 里没写 ByVal ⇒ VB6 默认 ByRef
    // ⇒ C 侧 `void** Editor`, 体里 `(*Editor)`。于是处理器把"对象指针"再解一层引用:
    //   实证 play78: Editor.Caption 取到 NULL (av read 0x0)
    //   @ _vb6_frmMain_ucPerspective1_OpenEditor+0x19。
    // 这里直接把 AST 形参的 isByVal 置真, 签名与函数体**一起**跟着变 (两处都读它),
    // 不改符号表也不动事件侧, 与 prelude 只发一次前向声明互相自洽。
    // 账 #222: 这段从 visit 里抽成 applyEventHandlerAbi 一处实现 —— 模块级前置声明那一趟与
    // 定义这一趟翻的必须是同一份 AST, 否则声明按 ByRef、定义按 ByVal, 中间还冒出一份零形参
    // 声明 (实测 ucProgressCircular: 一枚函数三种形参表 => 1 条 C2084 + 24 条 C2198)。
    if (ucEventHandlers_.count(Symbol::toLower(node.name)) > 0) applyEventHandlerAbi(node);

    std::string sig = makeProcSignature(node);

    // Fix 055: Form事件处理函数不能为static, 因为wndproc用extern引用它们
    bool isFormEventProc = isFormModule_ && node.name.find("Form_") == 0;
    // Fix 089e: 仅 Private 成员编译为 static — Friend 成员跨模块调用
    // (VB6 Friend = 工程内可见, 如 cHttpServer.OnDataArrival 被
    // cClientCallback 调用 / cSerialConfig.BuildTimeouts 被 cSerialPort
    // 调用), 若 static 则调用方 TU 中"声明但未定义" → C2129.
    if (node.access == AccessLevel::Private && !isFormEventProc) {
        c_.emitLine("static " + sig + " {");
    } else {
        c_.emitLine(sig + " {");
    }

    c_.indent();

    // 查找符号获取参数信息 (重载组内按声明位置取本变体, 无重载时等价旧 lookupModule)
    auto* sym = symTab_.lookupModuleOverloadByLoc(node.name, node.loc);
    currentProc_ = sym;
    currentReturnVar_ = "";
    currentReturnCType_ = "";  // Fix 054

    // Fix 056b: 清理局部数组注册 (模块级/类成员数组跨过程保留)
    clearProcArrayTracking();
    ansiTempsToFree_.clear();
    ansiOutParams_.clear();
    ivrefLocalsToRelease_.clear();  // tB Interface B05
    ansiCounter_ = 0;
    asmMixedBlockCounter_ = 0;   // ai/vb-asm-extension-spec 项2: 混排片段序号按过程重置
    knownBstrVars_.clear();
    knownDoubleVars_.clear();
    knownSingleVars_.clear();
    knownDateVars_.clear();   // Fix 175
    knownBoolVars_.clear();     // ai/022 W1
    knownByteVars_.clear();     // 账 #123
    knownIntVars_.clear();       // ai/009 5.10
    knownLongVars_.clear();
    knownLongPtrVars_.clear();  // Bug #2 fix: 也清空LongPtr集合
    knownVariantVars_.clear();
    // M22-fix: 只清空UDT变量map(旧条目会冲突), set类型不清空(WithEvents等模块级条目需跨过程保留)
    knownUdtVars_.clear();
    knownFixedStringLen_.clear();
    // Fix 010o: 清空局部变量集合
    knownLocalVars_.clear();
    knownNewVars_.clear();
    knownNewVars_.insert(moduleNewVars_.begin(), moduleNewVars_.end());  // Fix 090v
    // Fix 091m: 回灌模块级 Variant 变量 (knownVariantVars_ 已被 clear)
    knownVariantVars_.insert(moduleVariantVars_.begin(), moduleVariantVars_.end());
    knownByRefParams_.clear();  // Fix 081g
    // Fix 160w: 过程入口重置 COM 属性派发标记与 With 栈 — 上一个过程的 WithMemberExpr
    // 值 (如 `.Result`, `PropCellFont.Size`) 只把对象表达式留在 lastExpr_, 成员名挂
    // 在 comObjExpr_/comMemberName_/isComMarker_… 若不清, 会泄漏到下一个过程被其
    // BinaryExpr/赋值误当作 COM 属性读取 → C2065 (_vb6_with_14/_vb6_with_15,
    // Command5_Click -> Command10_Click) / C2440 (vb6_ComIface_Font*→float).
    isComMarker_ = false;
    comObjExpr_.clear();
    comMemberName_.clear();
    isEarlyBoundCom_ = false;
    earlyBoundSym_ = nullptr;
    withObjectVars_.clear();
    withObjectInfoStack_.clear();
    // Fix 010r/010r-10: 类模块中注册me到knownClassVars_ (使Me.Method()正确分发)
    // 改为map赋值: me → 当前模块名(类名)
    if (isClassModule_) knownClassVars_["me"] = moduleName_;
    // P6.11: 恢复类模块成员变量类型 (clear后从持久化集合恢复)
    knownBstrVars_.insert(classBstrMembers_.begin(), classBstrMembers_.end());
    knownDoubleVars_.insert(classDoubleMembers_.begin(), classDoubleMembers_.end());
    knownLongVars_.insert(classLongMembers_.begin(), classLongMembers_.end());
    // Fix 010n: 恢复类模块UDT成员变量 (knownUdtVars_被clear后需要从classUdtMembers_恢复)
    knownUdtVars_.insert(classUdtMembers_.begin(), classUdtMembers_.end());
    // Fix 010n (扩展): 恢复普通模块模块级UDT变量 (同 classUdtMembers_ 机制)
    knownUdtVars_.insert(moduleUdtMembers_.begin(), moduleUdtMembers_.end());

    for (auto& p : node.params) {
        // Fix 081g: Register ByRef params for For-loop dereference fix
        if (!p->isByVal) {
            std::string brKey = p->name;
            std::transform(brKey.begin(), brKey.end(), brKey.begin(), ::tolower);
            knownByRefParams_.insert(brKey);
        }
        // Fix 084m: 无类型子句的 Optional 参数 (如 Optional RecordsAffected) 默认是
        // Variant, C 类型 vb6_VARIANT*; 必须注册到 knownVariantVars_,
        // 否则 `RecordsAffected = 123` 生成裸赋值 → C2440 (cDataBase Exec).
        if (!p->asType) {
            std::string pLower = p->name;
            std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);
            knownVariantVars_.insert(pLower);
        }
        if (p->asType && p->asType->kind == ASTNodeKind::SimpleTypeRef) {
            auto& simpleP = static_cast<SimpleTypeRef&>(*p->asType);
            std::string pLower = p->name;
            std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);
            // P20-44: `As DataObject` 形参 → 记入 dataObjectParams_ (与 func.cpp 同款;
            // 窗体事件处理器走的是本文件不是 func.cpp, 漏这里就收不到)。
            if (Symbol::toLower(simpleP.name) == "dataobject")
                dataObjectParams_.insert(pLower);
            // Fix 161f: 同 cgen_decl_func.cpp — 限定名 (ComctlLib.ColumnHeader) 用
            // lookupTypeSymbol 按全名→末段查, 与 mapTypeRef 同口径。
            auto* pSym = lookupTypeSymbol(simpleP.name);
            if (!pSym) {
                size_t pDot161f = simpleP.name.find('.');
                if (pDot161f != std::string::npos)
                    pSym = lookupTypeSymbol(simpleP.name.substr(pDot161f + 1));
            }
            // Fix <vbeclipse> rev7: **工程类名表兜底** —— 与 mapTypeRef 的上移同根因。
            // `ByRef View As View` (ucFolder.AddView) / `ByRef Folder As Folder`
            // 这类**跨模块** .cls 形参在过程作用域查不到 (lookupTypeSymbol 只按本
            // 模块+类型库), 下面的 if 链全落空 ⇒ 形参不进 knownClassVars_ ⇒
            // 过程体内 `View.View` / `l_View.hWnd` 被当"模块限定符"或裸标识符
            // ⇒ 生成 `vb6_View_prop_get_View.hWnd` (丢实参, C2224)。
            // 三分支自带 pSym 守卫, 这里置的 knownClassVars_ 不会被后面覆盖。
            if (!pSym || pSym->kind != SymbolKind::Class) {
                const std::string projClsP092r = projectClassNameOf(simpleP.name);
                if (!projClsP092r.empty()) knownClassVars_[pLower] = projClsP092r;
            }
            if (pSym && pSym->kind == SymbolKind::UserDefinedType) {
                knownUdtVars_[pLower] = "vb6_type_" + cIdent(simpleP.name);
            } else if (pSym && pSym->kind == SymbolKind::Class) {
                knownClassVars_[pLower] = pSym->name;
            } else if (pSym && (pSym->kind == SymbolKind::ComClass || pSym->kind == SymbolKind::ComInterface)) {
                // Fix <vbeclipse>-2: 同 cgen_decl_func.cpp —— 本工程有同名类模块时走原生
                // (判据 projectClassNameOf, 不是 isExternal; 类型库自动加载注进来的
                // 内建 coclass isExternal=false).
                const std::string projCls2 = projectClassNameOf(simpleP.name);
                if (!projCls2.empty()) {
                    knownClassVars_[pLower] = projCls2;
                } else {
                    knownTypedComVars_[pLower] = pSym;
                }
            }
            // 接口类型的参数
            auto* pSym2 = symTab_.lookup(simpleP.name);
            if (pSym2 && pSym2->kind == SymbolKind::Class && pSym2->isInterface) {
                knownIfaceVars_[pLower] = pSym2->name;
            }
            // Fix 161f: `lv As ListView` 形参槽就是 HWND —— 同 cgen_decl_func.cpp。
            if (Symbol::toLower(simpleP.name).find("listview") != std::string::npos)
                listViewSlotVars_.insert(pLower);
            // 注册BSTR/Double/Long类型参数到类型跟踪集合
            Vb6Type paramType = typeSys_.resolveTypeName(simpleP.name);
            if (paramType == Vb6Type::String) knownBstrVars_.insert(pLower);
            // Fix 123b: Currency/Single/Date 形参同属 C double 组 — 此前漏注册,
            // inferExprType(形参) 落符号表查找 (129 模块工程里同名符号撞车) 或
            // Variant → RaiseEvent 实参打包误走 BSTR 分支 (cZipArchive
            // frFireProgress 的 Current/Total As Currency → vb6_BSTR_FromStr(double)
            // C2440, cZipArchive.c 2341/2343). 与 visit(FunctionDecl) 同步.
            else if (paramType == Vb6Type::Double || paramType == Vb6Type::Currency
                     || paramType == Vb6Type::Single || paramType == Vb6Type::Date) {
                knownDoubleVars_.insert(pLower);
                // Fix 175: Date 形参额外登记 —— C 层 Date/Double 同型, 只按
                // C 类型串分派会让 inferExprType 看不见 Date (打出序列号)。
                if (paramType == Vb6Type::Date) knownDateVars_.insert(pLower);
            }
            else if (paramType == Vb6Type::Long || paramType == Vb6Type::Integer || paramType == Vb6Type::Boolean
                     || paramType == Vb6Type::Byte) {
                knownLongVars_.insert(pLower);
                // ai/022 W1: 布尔形参另登记一份 (口径同 Fix 175 的 Date 形参)
                if (paramType == Vb6Type::Boolean) knownBoolVars_.insert(pLower);
                // 账 #123: Byte 形参一并进这一支 (口径同 ai/022 W1 的 Boolean —— C 型不同串
                // 就永远看不见 ⇒ 比较被当 Variant 取地址)。另登记一份到 knownByteVars_,
                // inferExprType 先判 Byte 那张表, 所以这里进 knownLongVars_ 不会把它读成 Long。
                if (paramType == Vb6Type::Byte) knownByteVars_.insert(pLower);
            }
            // Bug #2 fix: LongPtr 参数注册到独立集合
            else if (paramType == Vb6Type::LongPtr || paramType == Vb6Type::LongLong) knownLongPtrVars_.insert(pLower);   // Fix 084m
            // Fix 035: Variant 参数也要注册, 否则 `(*X) = concrete` 赋值不会触发
            // wrapVariantValue 包装, 导致 C2440 (ByRef Variant 参数写穿透场景).
            else if (paramType == Vb6Type::Variant) knownVariantVars_.insert(pLower);

            // Fix 023c: 注册 void* 参数 (As Object / As Collection / 外部 COM 类型如
            // ADODB.Recordset / Scripting.Dictionary 等) 到 knownObjectVars_ —
            // 让成员访问走 COM dispatch (vb6_ComCall / vb6_ComGet*Prop),
            // 而非直接 obj.member 字段访问, 避免 C2224 (void* 上 .member).
            // 仅当参数未被前面分支精确注册为 Class / ComClass / Interface / UDT 时
            // 才查 C 类型, 避免对 vb6_cls_* / vb6_ComIface_* 等 C 类型参数的错误
            // 注册. 与 visit(VariableDecl) line 651-656 行为一致 (局部 void* 同样注册).
            // Fix <VBFlexGridDemo>: UDT 形参**优先**登记 (口径同 cgen_decl_func.cpp 那处)。
            // knownClassVars_ 全程不清空, 别的模块 `Dim This As <工程类>` 会把同名条目
            // 泄漏过来 ⇒ UDT 形参被当类实例 ⇒ 成员访问回落 COM 后期绑定
            // vb6_ComGetObjectProp((*This), …) C2172。UDT 与类互斥, 故擦掉残留类条目。
            {
                std::string udtCTypeF = mapTypeRef(p->asType.get());
                while (!udtCTypeF.empty() && (udtCTypeF.back() == '*' || udtCTypeF.back() == ' '))
                    udtCTypeF.pop_back();
                if (udtCTypeF.compare(0, 9, "vb6_type_") == 0) {
                    knownUdtVars_[pLower] = udtCTypeF;
                    // 同 cgen_decl_func.cpp: 擦掉互斥表里的同名残留 (它们都全程不清空,
                    // 且判定分支排在 obj_dispatch 的 UDT 字段分支之前)。
                    knownClassVars_.erase(pLower);
                    knownTypedComVars_.erase(pLower);
                    knownIfaceVars_.erase(pLower);
                    knownIvrefVars_.erase(pLower);
                    knownObjectVars_.erase(pLower);
                    knownVariantVars_.erase(pLower);
                }
            }
            if (!knownClassVars_.count(pLower) && !knownTypedComVars_.count(pLower)
                && !knownIfaceVars_.count(pLower) && !knownUdtVars_.count(pLower)) {
                std::string paramCType = mapTypeRef(p->asType.get());
                if (paramCType == "void*") {
                    knownObjectVars_.insert(pLower);
                }
                // Bug #2 fix: Enum等未知类型参数, C类型为int32_t时注册为Long
                else if (paramCType == "int32_t" || paramCType == "int16_t" || paramCType == "VBABOOL") {
                    if (!knownLongVars_.count(pLower)) knownLongVars_.insert(pLower);
                }
                // Bug #2 fix: C类型为intptr_t时注册为LongPtr
                else if (paramCType == "intptr_t") {
                    if (!knownLongPtrVars_.count(pLower)) knownLongPtrVars_.insert(pLower);
                }
                // Fix 082: COM interface pointer types (vb6_ComIface_*) are pointer-sized on x64
                else if (paramCType.find("vb6_ComIface_") != std::string::npos) {
                    if (!knownLongPtrVars_.count(pLower)) knownLongPtrVars_.insert(pLower);
                }
            }
        } else if (p->asType && p->asType->kind == ASTNodeKind::ArrayTypeRef) {
            // Fix 010r-6: Register array parameters so arr(idx) generates VB6_SA_AT instead of (*arr)(idx)
            Vb6Type elemType = resolveArrayElemType(p->asType.get());
            std::string pLower = p->name;
            std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);
            knownArrays_.insert(pLower);
            arrayElemTypes_[pLower] = elemType;
            arrayDimCounts_[pLower] = 1;
            // Fix 055: 注册UDT数组元素C类型
            std::string udtCType = resolveArrayUdtElemCType(p->asType.get());
            if (!udtCType.empty()) {
                arrayUdtElemTypes_[pLower] = udtCType;
            }
            knownLocalVars_.insert(pLower);
        }
    }

    // VB6 Static Sub: 过程内所有局部变量都是static
    inStaticProc_ = node.isStatic;

    // 检测GoSub并声明返回地址栈
    hasGoSub_ = hasGoSubInStmts(node.body);
    gosubReturnCounter_ = 0;
    if (hasGoSub_) {
        c_.emitLine("int vb6_gosub_stack[32];");
        c_.emitLine("int vb6_gosub_sp = 0;");
    }

    // Bug #1 fix (082h): 预扫描UBound/LBound(arr,N>1)收集ND数组名
    scanNDArraysInStmts(node.body);

    // Fix 081: Apply default values for Optional parameters when not passed
    // VB6: Optional ByVal Ecl As Long = 1  →  if (!_has_Ecl) Ecl = 1;
    // This ensures the parameter variable has the correct default value
    // when the caller omits it, instead of the type's zero value.
    if (sym) {
        for (auto& pi : sym->params) {
            if (pi.isOptional && !pi.isParamArray && pi.hasDefaultValue && !pi.defaultValueExpr.empty()) {
                std::string pName = cIdent(pi.name);
                std::string hasFlag = "_has_" + pName;
                if (pi.isByVal) {
                    c_.emitLine("if (!" + hasFlag + ") " + pName + " = " + pi.defaultValueExpr + ";");
                } else {
                    std::string cType = mapType(pi.type);
                    if (pi.type == Vb6Type::String) {
                        c_.emitLine("if (!" + hasFlag + ") vb6_BSTR_Assign(" + pName + ", " + pi.defaultValueExpr + ");");
                    } else if (pi.type == Vb6Type::Variant) {
                        c_.emitLine("if (!" + hasFlag + ") (*" + pName + ") = " + pi.defaultValueExpr + ";");
                    } else {
                        c_.emitLine("if (!" + hasFlag + ") (*" + pName + ") = " + pi.defaultValueExpr + ";");
                    }
                }
            }
        }
    }

    // P12.3: 检测On Error并声明局部错误处理
    hasOnError_ = hasOnErrorInStmts(node.body);
    // P14.1.2: 检测Resume/Resume Next
    hasResume_ = hasResumeInStmts(node.body);
    inProtectedBlock_ = false;
    resumePointCounter_ = 0;
    dispatchPoints_.clear();
    currentErrorHandlerLabel_.clear();
    procExitLabelUsed_ = false;
    if (hasOnError_) {
        c_.emitLine("jmp_buf vb6_local_err_jmp;");
        if (hasResume_) {
            c_.emitLine("int32_t vb6_err_resume_point = 0;");
            c_.emitLine("int32_t vb6_err_resume_next_point = 0;");
        }
        c_.emitLine("vb6_SaveErrState();");
    }

    // 生成过程体 (P14.1.2: 传入hasResume_以启用resume点生成)
    // Fix 086: 先将块内 Dim/Const 提升到过程顶部 (VB6 局部声明是过程级作用域)
    hoistLocalDecls(node.body);
    // Fix <vbeclipse>: VB6 隐式变量 (无 Option Explicit 时未声明即使用) ——
    // 语义层登记的名字在此预声明为 Variant C 局部并注册 knownVariantVars_。
    implicitMacroNames_.clear();
    if (currentProc_) {
        auto* impl = symTab_.implicitVarsFor(Symbol::toLower(moduleName_),
                                             Symbol::toLower(currentProc_->name));
        if (impl) {
            for (const auto& n : *impl) {
                std::string nLower = Symbol::toLower(n);
                if (!knownLocalVars_.count(nLower)) {
                    knownLocalVars_.insert(nLower);
                    knownVariantVars_.insert(nLower);
                    // Fix <vbeclipse>: 隐式变量名可能撞 Win32 宏 (PopupMenu.cls 的
                    // MF_BYPOSITION)。过程内 push_macro/undef 隔离, 出过程 pop 复原。
                    c_.emitLine("#pragma push_macro(\"" + n + "\")");
                    c_.emitLine("#undef " + n);
                    implicitMacroNames_.push_back(n);
                    c_.emitLine("vb6_VARIANT " + cIdent(n) + " = vb6_VariantEmpty();  /* 隐式变量 */");
                }
            }
        }
    }
    // 账 #179: 控件代码运行在自己的宿主上下文里 (与下面的 PopInstance 成对)。
    if (ucCtxScoped()) c_.emitLine("vb6_UC_PushInstance((void*)me);");
    emitStmtList(node.body, hasResume_);

    // Fix <vbeclipse>: 过程统一出口。Exit Sub 发的 `goto vb6_proc_exit;` 落在这里,
    // 保证 vb6_RestoreErrState() 与 ivref Release 在所有退出路径上都会执行
    // (VB6: 过程退出自动恢复错误处理状态)。见 cgen_state.inc procExitLabelUsed_ 注释。
    if (procExitLabelUsed_) {
        c_.emitLine("vb6_proc_exit:;");
    }

        // P12.3: 恢复调用者的错误处理状态
    if (hasOnError_) {
        c_.emitLine("vb6_RestoreErrState();");
    }

    // M22: 释放当前过程中残留的ANSI临时变量 (正常退出路径)
    for (auto& ansiVar : ansiTempsToFree_) {
        c_.emitLine("vb6_FreeANSI(" + ansiVar + ");");
    }
    ansiTempsToFree_.clear();
    ansiOutParams_.clear();

    // tB Interface B05: 接口变量持有引用, 正常出口处经槽 Release (Exit Sub 例外, 同 ANSI 临时变量)
    emitIvrefScopeRelease();

    // 正常退出守卫 - 防止落入dispatch switch
    if (ucCtxScoped()) c_.emitLine("vb6_UC_PopInstance();");   // 账 #179: 与体首成对
    c_.emitLine("return;");

    // P14.1.2: Resume dispatch switch - 仅通过goto可达
    if (hasResume_ && !dispatchPoints_.empty()) {
        c_.emitLine("vb6_err_dispatch_switch:;");
        c_.emitLine("switch(vb6_err_dispatch) {");
        c_.indent();
        for (int pt : dispatchPoints_) {
            c_.emitLine("case " + std::to_string(pt) + ": goto vb6_resume_" + std::to_string(pt) + ";");
        }
        c_.dedent();
        c_.emitLine("}");
    }

    // Fix <vbeclipse>: 复原过程内被 #undef 的 Win32 宏 (隐式变量名隔离)
    for (const auto& n : implicitMacroNames_) {
        c_.emitLine("#pragma pop_macro(\"" + n + "\")");
    }
    implicitMacroNames_.clear();
    currentProc_ = nullptr;
    currentProc_ = nullptr;
    inStaticProc_ = false;
    hasGoSub_ = false;
    hasOnError_ = false;
    hasResume_ = false;
    inProtectedBlock_ = false;
    dispatchPoints_.clear();
    currentErrorHandlerLabel_.clear();
    procExitLabelUsed_ = false;
    c_.dedent();
    c_.emitLine("}");
    c_.emitBlank();
}


// 事件处理器的 ABI: 形参一律 ByVal (跨对象边界的回调)。AST 与 proc 符号两处都要翻 ——
// 只翻 AST 会得到"签名 void* Editor + 体里 (*Editor)" (cgen_expr_ident_dispatch.inc 读的是
// 符号表里的 isByVal), 实测 frmMain.c 三条 C2100。声明侧与定义侧共用这一处 (账 #222)。
void CCodeGen::applyEventHandlerAbi(SubDecl& node) {
    for (auto& prm : node.params) {
        if (prm && !prm->isParamArray) prm->isByVal = true;
    }
    if (auto* evtSym = symTab_.lookupModuleOverloadByLoc(node.name, node.loc)) {
        for (auto& pi : evtSym->params) {
            if (!pi.isParamArray) pi.isByVal = true;
        }
    }
}

// 按名字在本模块里找那枚 <控件>_<事件> 处理器: 翻 ABI, 并回答"它的第一形参是不是元素号 Index"
// (VB6 的控件数组处理器就是这么写的, 与非数组的差别只在这一枚形参)。找不到 = false。
bool CCodeGen::prepareEventHandlerProc(Module& mod, const std::string& loweredName) {
    for (auto& d : mod.declarations) {
        if (!d || d->kind != ASTNodeKind::SubDecl) continue;
        auto& sub = static_cast<SubDecl&>(*d);
        if (!sub.typeParams.empty()) continue;          // 泛型模板本体不发码
        if (Symbol::toLower(sub.name) != loweredName) continue;
        applyEventHandlerAbi(sub);
        return !sub.params.empty() && sub.params[0] &&
               Symbol::toLower(sub.params[0]->name) == "index";
    }
    return false;
}

} // namespace vb6c3
