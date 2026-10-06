// cgen_expr_ident_symbol.inc - visit(IdentifierExpr&) 函数体片段：符号解析（参数/局部变量/窗体属性/枚举/COM全局/变量常量）
// 由 src/backend/expr/cgen_expr_ident.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 234~420 行，逐行未改。

    // Fix 056: Form 内置属性隐式访问 (WindowState/ScaleWidth/ScaleHeight)
    // VB6: 在 Form 模块内, 裸属性名等同于 Me.属性名
    if (isFormModule_ && !knownFormName_.empty()) {
        std::string formHwnd;
        auto itOrig = knownFormControlOriginalNames_.find(knownFormName_);
        if (itOrig != knownFormControlOriginalNames_.end()) {
            formHwnd = "vb6_hwnd_" + cIdent(itOrig->second);
        } else {
            formHwnd = "vb6_hwnd_" + moduleName_;
        }
        std::string readFn = getControlPropReadFn(FrmControlType::Form, node.name);
        // 账 #68: VB6 里裸 `Left(s, n)` 是**内置函数**（窗体的 Left 属性必须写 Me.Left）。
        // 这条分支对"与窗体属性同名的裸标识符"一律按属性读，而 Left 恰好两头都占 ⇒
        // `Left(x, 1)` 发成 vb6_GetControlLeft(hwnd) 再被调用方拼上实参：编得过、运行期
        // 读的是窗体坐标（C29-1b 那轮还踩成 0xC0000005）。**函数调用上下文里让内置函数赢**：
        // 符号表把 VB6 内置函数登记成 isBuiltin 的 Function（实测 Left ⇒ kind=Function
        // builtin=1），所以这是一条"问表"的规则，不是硬编码名单。值上下文（`x = Left`）
        // 照旧是窗体属性。
        bool builtinCallWins68 = false;
        if (asCallCallee_ && !readFn.empty()) {
            Symbol* bs68 = symTab_.lookup(node.name);
            if (!bs68) bs68 = symTab_.lookupModule(node.name);
            builtinCallWins68 = (bs68 && bs68->isBuiltin
                                 && bs68->kind == SymbolKind::Function);
        }
        if (!readFn.empty() && !builtinCallWins68) {
            lastExpr_ = readFn + "(" + formHwnd + ")";
            return;
        }
    }

    // Fix <vbeclipse> rev27: `Dim X As New <窗体名>` 的自动实例化守卫由
    // cgen_localdecl.cpp 在**声明处**发出 (VB6 语义 = 声明即构造)。这里**不要**再发 ——
    // 窗体变量在符号表里是普通 Variable, 引用多走 MemberAccess 路径, 在此发守卫
    // 既不可靠 (实测一条都没落到生成 C 里) 又会在纯标识符引用时重复发。
    // 兜底: 声明处那条要求 `storageClass.empty()` (静态局部不实例化), 若变量来自
    // Static/Const 等其他声明形态则此处不处理, 保持改动前行为。

    // Fix 086: 跨模块窗体默认实例引用 (cLogs 里 Unload FLogs / FLogs.Show).
    // 窗体名作为值 → 该窗体HWND访问器 vb6_form_hwnd_<Form>().
    // 成员访问 (.Visible/.Show) 由 MemberAccessExpr 的窗体模块分支处理.
    if (knownFormModuleNames_.count(lower) && lower != knownFormName_) {
        lastExpr_ = "vb6_form_hwnd_" + cIdent(node.name) + "()  /* form default instance */";
        return;
    }

    // 枚举成员引用
    if (foundSym && foundSym->kind == SymbolKind::EnumMember) {
        if (foundSym->hasConstValue) {
            lastExpr_ = std::to_string(foundSym->constIntValue);
        } else {
            lastExpr_ = cName;
        }
        return;
    }

    // P24-04: ComGlobalNs — VB_GlobalNameSpace promoted函数 (如VBMAN)
    // 当VBMAN作为独立标识符出现时 (不在MemberAccessExpr.object位置),
    // 生成sGlobal单例创建 + promoted方法调用
    if (foundSym && foundSym->kind == SymbolKind::ComGlobalNs) {
        std::string progIdWide = "L\"" + foundSym->comProgId + "\"";
        std::string methodName = foundSym->comGlobalNsMethodName;
        lastExpr_ = "vb6_ComCallObject(vb6_CreateObject(" + progIdWide + "), L\"" + methodName + "\", NULL, 0)";
        return;
    }

    // 用户变量/常量 (非参数、非函数)
    if (foundSym && (foundSym->kind == SymbolKind::Variable
                  || foundSym->kind == SymbolKind::Constant)) {
        // Fix 017: 数值常量 (内置 vbMethod=2/vbDirectory=16, 或用户 Public Const,
        // 或跨模块注入的 EnumMember-as-Constant) 直接输出数值, 避免发出裸标识符
        // (C 代码中无对应 #define → C2065).
        // Fix 081d: 仅对整型/浮点/布尔常量内联数值; 字符串常量(hasConstValue=true
        // 但 constType==String) 必须走正常标识符路径, 引用 #define 宏名, 否则
        // constIntValue==0 被错误输出为 "0".
        // Fix 161f: 未写 As Type 的 Const 在符号表里 constType=Variant (且实测
        // hasConstValue=0), 走不到下面的 constType 分支; 但它的值已在
        // visit(ConstDecl) 折叠成功 (C 侧 emit 成 `#define NAME (4127)`), 登记在
        // moduleIntConstValues_. 若发裸宏名, 下游按符号表判型会得到 Variant →
        // 实参被包成 vb6_VariantToLong(LVM_GETHEADER) → C2440 "无法从 int 转换为
        // vb6_VARIANT" (extlist MListView SendMessage 实测)。故**先查这张表**,
        // 按已折叠整数值内联, 与 `As Long` Const 同口径。字符串 Const 不会进这张
        // 表 (tryEvalConstInt 对字符串失败), 故无 081d 误伤。
        {
            auto itCV161f = moduleIntConstValues_.find(lower);
            if (itCV161f != moduleIntConstValues_.end()) {
                lastExpr_ = std::to_string(itCV161f->second);
                return;
            }
        }
        if (foundSym->kind == SymbolKind::Constant && foundSym->hasConstValue) {
            if (foundSym->constType == Vb6Type::Long || foundSym->constType == Vb6Type::Integer
                || foundSym->constType == Vb6Type::Boolean || foundSym->constType == Vb6Type::Byte
                || foundSym->constType == Vb6Type::Error) {
                lastExpr_ = std::to_string(foundSym->constIntValue);
                return;
            } else if (foundSym->constType == Vb6Type::Single || foundSym->constType == Vb6Type::Double) {
                lastExpr_ = std::to_string(foundSym->constFloatValue);
                return;
            }
            // 字符串常量和其他类型: 不内联, 让代码继续走标识符路径引用 #define 宏
        }
        // P11.7: 如果是内置Object类型变量(=窗体控件), 优先走默认属性读取
        if (foundSym->isBuiltin && foundSym->type == Vb6Type::Object) {
            // P17.1: With块内抑制默认属性解析, 返回HWND引用
            if (suppressDefaultProp_) {
                lastExpr_ = "vb6_hwnd_" + cIdent(node.name);
                return;
            }

            auto itCtrl = knownFormControls_.find(lower);
            if (itCtrl != knownFormControls_.end()) {
                const char* defaultProp = getDefaultPropertyName(itCtrl->second);
                if (defaultProp) {
                    std::string readFn = getControlPropReadFn(itCtrl->second, defaultProp);
                    if (!readFn.empty()) {
                        lastExpr_ = readFn + "(" + makeCtrlHwndArg(lower, itCtrl->second) + ")  /* default prop: ." + std::string(defaultProp) + " */";
                        return;
                    }
                }
            }
            // P20-31: WithEvents控件变量默认属性读取
            if (!suppressDefaultProp_) {
                auto itWECtrl = knownWithEventsCtrlVars_.find(lower);
                if (itWECtrl != knownWithEventsCtrlVars_.end()) {
                    const char* defaultProp = getDefaultPropertyName(itWECtrl->second);
                    if (defaultProp) {
                        std::string readFn = getControlPropReadFn(itWECtrl->second, defaultProp);
                        if (!readFn.empty()) {
                            auto itOrig = knownWithEventsCtrlOrigNames_.find(lower);
                            std::string weVarName = (itOrig != knownWithEventsCtrlOrigNames_.end()) ? itOrig->second : cName;
                            lastExpr_ = readFn + "(" + weVarName + ")  /* WithEvents ctrl default prop: ." + std::string(defaultProp) + " */";
                            return;
                        }
                    }
                }
            }
        }
        // 类模块变量通过me->访问
        // Fix 010o: 但如果该名称是局部变量(Dim/For循环变量)或外部模块变量, 不加me->前缀
        if (isClassModule_ && currentProc_ && foundSym->kind == SymbolKind::Variable
            && !knownLocalVars_.count(lower)
            && !foundSym->isExternal) {
            Symbol* paramSym = symTab_.lookupLocal(node.name);
            if (!paramSym || paramSym->kind != SymbolKind::Parameter) {
                lastExpr_ = "me->" + cName;
                // P14.3.1: Dim As New自动实例化 (类模块成员)
                auto itNewM = knownNewVars_.find(lower);
                if (itNewM != knownNewVars_.end()) {
                    {
                std::string newExpr;
                auto itCom = knownTypedComVars_.find(lower);
                if (itCom != knownTypedComVars_.end()) {
                    newExpr = comNewExprFor(itCom->second);  // Fix 177b: 遮蔽→工程类工厂+IDispatch包装
                    if (newExpr.empty()) {
                        const std::string& _progId = itCom->second->comProgId.empty() ? itCom->second->name : itCom->second->comProgId;
                        newExpr = "(void*)vb6_NewObject(L\"" + _progId + "\")";
                    }
                } else {
                    newExpr = "vb6_cls_" + itNewM->second + "_New()";
                }
                c_.emitLine("if (!me->" + cName + ") me->" + cName + " = " + newExpr + ";  /* Dim As New auto-instantiate */");
            }
                }
                return;
            }
        }
        // P14.3.1: Dim As New自动实例化守卫
        // (注: `Dim X As New <窗体>` 的守卫在本文件更前面统一发出 —— 那些局部变量
        //  是普通 Variable, 会先被下面的 Variable 分支 return 掉。)
        auto itNew = knownNewVars_.find(lower);
        if (itNew != knownNewVars_.end()) {
            {
            std::string newExpr;
            auto itCom = knownTypedComVars_.find(lower);
            if (itCom != knownTypedComVars_.end()) {
                newExpr = comNewExprFor(itCom->second);  // Fix 177b: 遮蔽→工程类工厂+IDispatch包装; 未遮蔽→空串
                if (newExpr.empty()) {
                    const std::string& _progId = itCom->second->comProgId.empty() ? itCom->second->name : itCom->second->comProgId;
                    newExpr = "(void*)vb6_NewObject(L\"" + _progId + "\")";
                }
                // Fix 090v-com: 类型化COM接口指针 (vb6_ComIface_X*) 转型:
                // vb6_NewObject 返回 void*, 直接赋给类型化指针在 C 下是约束违例 → 编译失败.
                if (knownTypedComVarCType_.count(lower)) {
                    std::string _t = knownTypedComVarCType_[lower];
                    if (!_t.empty() && _t.back() == '*') _t.pop_back();
                    if (_t.find("vb6_ComIface_") != std::string::npos) newExpr = "(" + _t + ")" + newExpr;
                }
            } else {
                newExpr = "vb6_cls_" + itNew->second + "_New()";
            }
            // Fix 086: Variant局部变量持有对象时, 守卫用 VariantToObjectVal 判空,
            // 赋值用 _Generic vb6_VariantFromValue 包装 (避免 !VARIANT 与 C2440)
            if (knownVariantVars_.count(lower)) {
                c_.emitLine("if (vb6_VariantToObjectVal(" + cName + ") == NULL) " + cName
                            + " = vb6_VariantFromValue(" + newExpr + ");  /* Dim As New auto-instantiate (Variant) */");
            } else {
                c_.emitLine("if (!" + cName + ") " + cName + " = " + newExpr + ";  /* Dim As New auto-instantiate */");
            }
        }
        }
        lastExpr_ = cName;
        return;
    }

    // M22-Issue5: foundSym为null但可能是cgen过程级跟踪的Dim As New变量
    // 符号表在语义分析后current_指向模块scope, 过程级局部变量查找失败
    // 此时knownNewVars_/knownClassVars_仍然有效(cgen过程中维护)
    if (!foundSym) {
        auto itNew = knownNewVars_.find(lower);
        if (itNew != knownNewVars_.end()) {
            {
            std::string newExpr;
            auto itCom = knownTypedComVars_.find(lower);
            if (itCom != knownTypedComVars_.end()) {
                newExpr = comNewExprFor(itCom->second);  // Fix 177b: 遮蔽→工程类工厂+IDispatch包装; 未遮蔽→空串
                if (newExpr.empty()) {
                    const std::string& _progId = itCom->second->comProgId.empty() ? itCom->second->name : itCom->second->comProgId;
                    newExpr = "(void*)vb6_NewObject(L\"" + _progId + "\")";
                }
                // Fix 090v-com: 类型化COM接口指针 (vb6_ComIface_X*) 转型:
                // vb6_NewObject 返回 void*, 直接赋给类型化指针在 C 下是约束违例 → 编译失败.
                if (knownTypedComVarCType_.count(lower)) {
                    std::string _t = knownTypedComVarCType_[lower];
                    if (!_t.empty() && _t.back() == '*') _t.pop_back();
                    if (_t.find("vb6_ComIface_") != std::string::npos) newExpr = "(" + _t + ")" + newExpr;
                }
            } else {
                newExpr = "vb6_cls_" + itNew->second + "_New()";
            }
            // Fix 086: Variant局部变量持有对象时, 守卫用 VariantToObjectVal 判空,
            // 赋值用 _Generic vb6_VariantFromValue 包装 (避免 !VARIANT 与 C2440)
            if (knownVariantVars_.count(lower)) {
                c_.emitLine("if (vb6_VariantToObjectVal(" + cName + ") == NULL) " + cName
                            + " = vb6_VariantFromValue(" + newExpr + ");  /* Dim As New auto-instantiate (Variant) */");
            } else {
                c_.emitLine("if (!" + cName + ") " + cName + " = " + newExpr + ";  /* Dim As New auto-instantiate */");
            }
        }
            lastExpr_ = cName;
            return;
        }
        if (knownClassVars_.find(lower) != knownClassVars_.end()) {
            lastExpr_ = cName;
            return;
        }
    }
