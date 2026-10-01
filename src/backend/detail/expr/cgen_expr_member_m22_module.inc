// cgen_expr_member_m22_module.inc - 函数体片段：M22 跨模块变量访问与函数调用结果上的方法调用
// 由 src/backend/expr/cgen_expr_member.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 775~943 行，逐行未改。

    // M22: 跨模块变量访问 Module1.myName → vb6_Module1_myName
    // 优先级4: object不是已知变量, 成员也不是函数 → 模块名.变量名
    if (node.object && node.object->kind == ASTNodeKind::IdentifierExpr) {
        auto& objIdent2 = static_cast<IdentifierExpr&>(*node.object);
        std::string objLower2 = objIdent2.name;
        std::transform(objLower2.begin(), objLower2.end(), objLower2.begin(), ::tolower);
        bool isVarName2 = false;
        auto* objSym2 = symTab_.lookup(objIdent2.name);
        if (!objSym2) objSym2 = symTab_.lookupModule(objIdent2.name);
        if (objSym2 && (objSym2->kind == SymbolKind::Variable || objSym2->kind == SymbolKind::Parameter)) {
            isVarName2 = true;
        }
        // M22-fix: symTab_在模块作用域, 找不到过程级局部变量
        // 补充检查cgen层的跟踪集合: UDT变量、类实例变量、Dim As New变量、COM对象变量
        if (!isVarName2 && knownUdtVars_.count(objLower2)) isVarName2 = true;
        if (!isVarName2 && knownClassVars_.find(objLower2) != knownClassVars_.end()) isVarName2 = true;
        if (!isVarName2 && knownNewVars_.count(objLower2)) isVarName2 = true;
        if (!isVarName2 && knownObjectVars_.count(objLower2)) isVarName2 = true;
        if (!isVarName2 && knownTypedComVars_.count(objLower2)) isVarName2 = true;
        if (!isVarName2 && knownIfaceVars_.count(objLower2)) isVarName2 = true;
        // Fix 056b: 属性符号也是有效对象表达式 (PropertyGet/Let/Set),
        // 不能当作"模块名限定符"处理, 否则 LastError.Number → vb6_LastError_Number (C2065)
        if (!isVarName2 && objSym2 &&
            (objSym2->kind == SymbolKind::PropertyGet ||
             objSym2->kind == SymbolKind::PropertyLet ||
             objSym2->kind == SymbolKind::PropertySet)) {
            isVarName2 = true;
        }
        // Fix 084y-8: 枚举类型名.成员 (EnumLogLevel.LvCustom) — 枚举类型不是
        // 变量也不是模块, 原逻辑落入模块名限定符 → vb6_EnumLogLevel_LvCustom
        // (C2065, 声明名是 vb6_enum_EnumLogLevel_LvCustom 带 enum_ 层). 成员
        // 有常量值则内联数值 (LvCustom=5), 否则生成声明同构名.
        if (!isVarName2 && objSym2 && objSym2->kind == SymbolKind::EnumType) {
            Symbol* enumMemSym = symTab_.lookup(node.memberName);
            if (enumMemSym && enumMemSym->kind == SymbolKind::EnumMember
                && enumMemSym->hasConstValue) {
                lastExpr_ = std::to_string(enumMemSym->constIntValue);
                return;
            }
            lastExpr_ = "vb6_enum_" + cIdent(objIdent2.name) + "_" + cIdent(node.memberName);
            return;
        }

        if (!isVarName2) {
            // object不是已知变量 → 假设是模块名限定符
            // 成员是变量: Module1.myName
            // When module is #included, use the unprefixed name directly
            std::string modName2 = objIdent2.name;
            std::string varName2 = cIdent(node.memberName);
            std::string modLower2 = modName2;
            std::transform(modLower2.begin(), modLower2.end(), modLower2.begin(), ::tolower);
            // Fix 086: 在指定模块中解析成员实际符号 (大小写无关).
            // 1) 常量: 内联数值 (跨模块常量引用 modSerialPortAPI.SetDTR —
            //    模块头文件只发 #define SETDTR, 无 vb6_<mod>_<name> 标识符,
            //    且使用处大小写可能与 #define 不同 → C2065).
            // 2) 变量: 用符号的规范大小写 (extern 声明名与使用处不同).
            Symbol* qSym = nullptr;
            if (symTab_.moduleScope()) {
                const std::string wantLower = Symbol::toLower(node.memberName);
                for (const auto& [qkey, qsym] : symTab_.moduleScope()->symbols()) {
                    if (!qsym) continue;
                    if (qsym->lowerName != wantLower && Symbol::toLower(qsym->name) != wantLower) continue;
                    if (!qsym->isExternal || Symbol::toLower(qsym->sourceModule) != modLower2) continue;
                    qSym = qsym.get();
                    break;
                }
            }
            if (qSym && qSym->kind == SymbolKind::Constant && qSym->hasConstValue) {
                if (qSym->constType == Vb6Type::Long || qSym->constType == Vb6Type::Integer
                    || qSym->constType == Vb6Type::Boolean || qSym->constType == Vb6Type::Byte
                    || qSym->constType == Vb6Type::Error) {
                    lastExpr_ = std::to_string(qSym->constIntValue);
                    return;
                }
                if (qSym->constType == Vb6Type::Single || qSym->constType == Vb6Type::Double) {
                    lastExpr_ = std::to_string(qSym->constFloatValue);
                    return;
                }
                varName2 = cIdent(qSym->name);
            } else if (qSym && qSym->kind == SymbolKind::Variable) {
                varName2 = cIdent(qSym->name);
            } else if (!qSym && modulePublicConsts_) {
                // Fix 086: 注入被本地同名符号阻止时, 查 driver 预扫描常量表
                auto itMod = modulePublicConsts_->find(modLower2);
                if (itMod != modulePublicConsts_->end()) {
                    auto itVal = itMod->second.find(Symbol::toLower(node.memberName));
                    if (itVal != itMod->second.end()) {
                        lastExpr_ = std::to_string(itVal->second);
                        return;
                    }
                }
            }
            bool isIncluded2 = false;
            for (const auto& extMod : externalModules_) {
                std::string extLower = extMod;
                std::transform(extLower.begin(), extLower.end(), extLower.begin(), ::tolower);
                if (extLower == modLower2) { isIncluded2 = true; break; }
            }
            // Fix <vbeclipse>: 宿主伪对象 (UserControl/Ambient/...) 的成员名按 RTL 约定
            // 规范化 (源码 `UserControl.hDc` → RTL 只有 vb6_UserControl_hDC), 否则
            // 生成未声明的 vb6_UserControl_hDc (C2065 ×3: ucButton/ucCaption/ucTab).
            varName2 = cIdent(canonicalHostPseudoMember(modName2, varName2));
            lastExpr_ = isIncluded2 ? varName2 : ("vb6_" + cIdent(modName2) + "_" + varName2);
            return;
        }
    }

    // Fix 083d: 函数调用结果上的方法调用 — pvToSocket(.SocketPtr(lIdx)).frNotifyGetHostByName(...)
    // object 是非标识符表达式(方法调用返回类实例) → 用类型推断得到类名,
    // 生成全局方法调用 vb6_<Class>_<Method>((void*)<objExpr>, args...),
    // 避免生成 objExpr.member(args) 非法 C (C2039/结构体无此成员)
    // Fix 084y-4: 调用上下文 (asCallCallee_=true, 如 db.Sql(s).Exec(...)) 时跳过本
    // 分支 — 它生成 func((void*)objExpr) 只含 this 指针, 绕过外层 IndexOrCallExpr
    // 的 Optional 默认参数填充 → C2198 参数太少 (cDataBase.Exec 声明5参却传1参).
    // 调用上下文应交给 Fix 015 (IndexOrCallExpr 链 → pendingChainObj_ 裸函数名,
    // 由外层补齐参数). 值上下文 (bX = obj.Fn().Prop) 才走本分支.
    if (node.object && node.object->kind != ASTNodeKind::IdentifierExpr && !asCallCallee_) {
        std::string className = inferClassTypeOfExpr(*node.object);
        if (!className.empty()) {
            std::string resolvedFn = resolveClassMemberCall(className, node.memberName);
            if (!resolvedFn.empty()) {
                emitExpr(*node.object);
                std::string objExpr = std::move(lastExpr_);
                // 类虚表 (tB, ai/022 B08e 第⑪站): 这一支的 object 是"非标识符表达式", 但
                // `Me.<字段>` 这种字段链发出来仍是纯读 (`me->m_up`), 所以 `Me.m_up.<属性>` 的读
                // 实测落在这里 —— 上一版地图把它一概判成"按构造不纯 → 只能判死", 本轮按证据订正。
                // 口径同 ①⑤: 纯的换成按 __cvtbl 索引, 不纯的由 mustDispatch 出 VB3027;
                // 派发串与 this 实参共用下面这一个 thisArg083d (拼两遍 = 双副作用)。
                // 两个方向坑: 属性写 (`_prop_let_`/`_prop_set_`) 不改写 (3.4b 无该方向的槽),
                // 左值求值 (`suppressVirtDispatch_`) 也不改写 —— Pattern C/D2 会字符串级换后缀。
                std::string thisArg083d = "(void*)" + objExpr;
                if (!suppressVirtDispatch_
                    && resolvedFn.find("_prop_let_") == std::string::npos
                    && resolvedFn.find("_prop_set_") == std::string::npos) {
                    const std::string dispatchFn083d = virtDispatchCallee(
                        className, node.memberName, thisArg083d,
                        /*mustDispatch=*/true, node.loc);
                    if (!dispatchFn083d.empty()) resolvedFn = dispatchFn083d;
                }
                // Fix 090am: 值上下文无括号方法引用 (If Db.Sql(s).Param(p).QueryParam Then)
                // — object 是非标识符链式调用 → 本 Fix 083d 分支只发 this ((void*)obj)
                // → C2198 参数太少 (QueryParam 声明5参传1参). 与 Fix 089h/090ak 同模板:
                // 按 findClassMemberCallParams 形参表补默认值. 属性 (prop_get_/prop_let_)
                // 不 pad (无括号属性读只发 this 即正确, pad 反致 C2197).
                std::vector<ParameterInfo> params090am;
                bool isB090am = false;
                if (resolvedFn.find("_prop_") == std::string::npos
                    && findClassMemberCallParams(className, node.memberName,
                                                 params090am, isB090am)
                    && !params090am.empty() && !isB090am) {
                    std::string argList090am = thisArg083d;
                    for (size_t i = 0; i < params090am.size(); i++) {
                        const auto& param = params090am[i];
                        argList090am += ", ";
                        std::string defVal090am;
                        if (param.hasDefaultValue && !param.defaultValueExpr.empty()) {
                            defVal090am = param.defaultValueExpr;
                        } else {
                            defVal090am = defaultValue(param.type);
                        }
                        if (param.isByVal) {
                            argList090am += defVal090am;
                        } else {
                            std::string cType090am = mapType(param.type);
                            if (param.type == Vb6Type::Variant
                                || param.type == Vb6Type::Empty
                                || param.type == Vb6Type::Null
                                || param.type == Vb6Type::Object) {
                                argList090am += "&(" + cType090am + "){0}";
                            } else {
                                argList090am += "&(" + cType090am + "){" + defVal090am + "}";
                            }
                        }
                    }
                    for (size_t i = 0; i < params090am.size(); i++) {
                        const auto& param = params090am[i];
                        if (param.isOptional && !param.isParamArray) {
                            argList090am += ", 0";
                        }
                    }
                    lastExpr_ = resolvedFn + "(" + argList090am + ")";
                } else {
                    lastExpr_ = resolvedFn + "(" + thisArg083d + ")";
                }
                return;
            }
        }
    }
