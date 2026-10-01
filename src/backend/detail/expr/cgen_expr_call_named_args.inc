// cgen_expr_call_named_args.inc - visit(IndexOrCallExpr&) 函数体片段：named 实参、argList 组装与 ParamArray 展开
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 2454~2860 行，逐行未改。

    std::vector<bool> actuallyPassedParams;

    // P14.3.3: 命名参数位置展开 - 按参数名映射到正确位置
    if (!node.named.empty() && !calleeParams.empty()) {
        // Build name->index map from callee params (case-insensitive)
        std::unordered_map<std::string, size_t> paramMap;
        for (size_t i = 0; i < calleeParams.size(); i++) {
            std::string lower = calleeParams[i].name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (!lower.empty()) paramMap[lower] = i;
        }

        // Create position-mapped args vector
        size_t totalParams = calleeParams.size();
        std::vector<std::string> orderedArgs(totalParams);
        std::vector<bool> filled(totalParams, false);

        // Place positional args (already in args[])
        for (size_t i = 0; i < args.size() && i < totalParams; i++) {
            orderedArgs[i] = std::move(args[i]);
            filled[i] = true;
        }

        // ByRef handling helper lambda
        auto applyByRef = [&](std::string& argVal, size_t pi) {
            // Fix 078 rev2: ByRef array params are now vb6_SafeArray1D**,
            // so they need & at call site — same as other ByRef params.
            bool piHasByValOverride = node.byvalOverrides.count(pi) > 0;
            bool isByRef = (pi < calleeParams.size() && !calleeParams[pi].isByVal && !calleeParams[pi].isParamArray)
                           && !piHasByValOverride;
            if (!isByRef) return;
            bool piCalleeParamIsArray = (pi < calleeParams.size()
                && (static_cast<uint16_t>(calleeParams[pi].type) & static_cast<uint16_t>(Vb6Type::Array)));
            if (argVal.size() > 3 && argVal.substr(0, 2) == "(*" && argVal.back() == ')') {
                // (*x) → depends on callee param type (see Fix 078 rev2 details in position-arg path)
                std::string innerName = argVal.substr(2, argVal.size() - 3);
                if (piCalleeParamIsArray) {
                    argVal = innerName;
                } else {
                    // Fix 079: check if innerName is a ByRef param of the current function
                    bool piIsCurrentByRefParam = false;
                    Vb6Type piCurPType189 = Vb6Type::Unknown;
                    if (currentProc_) {
                        for (auto& p : currentProc_->params) {
                            if (Symbol::toLower(p.name) == Symbol::toLower(innerName) && !p.isByVal && !p.isParamArray) {
                                bool paramIsArray = (static_cast<uint16_t>(p.type) & static_cast<uint16_t>(Vb6Type::Array)) != 0;
                                if (!paramIsArray) {
                                    piIsCurrentByRefParam = true;
                                    piCurPType189 = p.type;
                                }
                                break;
                            }
                        }
                    }
                    if (piIsCurrentByRefParam) {
                        // Fix 189: 被调方形参是 ByRef Variant 时不可直传 T*
                        // (同位置实参路径 — vt 取到 T 的字节 → 0xC0000005)
                        std::string piRep189;
                        if (pi < calleeParams.size()
                            && (calleeParams[pi].type == Vb6Type::Variant
                                || calleeParams[pi].type == Vb6Type::Empty)) {
                            piRep189 = wrapByRefVariantParamArg(innerName, piCurPType189);
                        }
                        argVal = piRep189.empty() ? innerName : piRep189;
                    } else {
                        argVal = "&" + innerName;
                    }
                }
            } else if (argVal.size() > 2 && argVal[0] == '&') {
                // already has &, keep as-is
            } else if (argVal.size() > 2 && argVal.substr(0, 2) == "me" && argVal[2] == '-') {
                argVal = "&(" + argVal + ")";
            } else {
                bool isSimpleIdent = !argVal.empty() && (std::isalpha(static_cast<unsigned char>(argVal[0])) || argVal[0] == '_');
                if (isSimpleIdent) {
                    for (char c : argVal) {
                        if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_') {
                            isSimpleIdent = false; break;
                        }
                    }
                }
                if (isSimpleIdent && !argVal.empty()) {
                    // Fix 079: If the simple identifier is a ByRef param of the current function,
                    // it's already a pointer (T*) — pass directly, don't add &.
                    bool piArgIsCurrentByRefParam = false;
                    Vb6Type piArgCurPType189 = Vb6Type::Unknown;
                    if (currentProc_) {
                        for (auto& p : currentProc_->params) {
                            if (Symbol::toLower(p.name) == Symbol::toLower(argVal) && !p.isByVal && !p.isParamArray) {
                                piArgIsCurrentByRefParam = true;
                                piArgCurPType189 = p.type;
                                break;
                            }
                        }
                    }
                    if (piArgIsCurrentByRefParam) {
                        // ByRef param of current function → already T*, pass directly
                        // Fix 189: ByRef Variant 形参时需重打包 (见上)
                        std::string piRep189b;
                        if (pi < calleeParams.size()
                            && (calleeParams[pi].type == Vb6Type::Variant
                                || calleeParams[pi].type == Vb6Type::Empty)) {
                            piRep189b = wrapByRefVariantParamArg(argVal, piArgCurPType189);
                        }
                        if (!piRep189b.empty()) argVal = piRep189b;
                    } else if (argVal == "NULL" || argVal == "0") {
                        // Fix 086: NULL/0 无可取址, 直接传空指针 (ByRef 形参收到 NULL)
                    } else if (isConstIdent(argVal)) {
                        // Fix 084aa: 常量宏不可取址 → 按形参类型复合字面量包装
                        Vb6Type pType = Vb6Type::Variant;
                        if (pi < calleeParams.size()) pType = calleeParams[pi].type;
                        argVal = wrapConstArgForByRef(argVal, pType);
                    } else {
                        argVal = "&" + argVal;
                    }
                } else {
                    std::string cType = "int32_t";
                    if (pi < calleeParams.size()) cType = mapType(calleeParams[pi].type);
                    if (cType == "vb6_VARIANT") {
                        // Fix 086: 非左值表达式传 ByRef Variant — 不能用 {argVal}
                        // 首字段初始化 (C2440: 指针初始化VARTYPE). 空对象/零值
                        // (如内置App对象 (void*)0) → 零初始化VARIANT (VT_EMPTY,
                        // 与VB6传Nothing语义一致); 其余退回零初始化以保编译通过.
                        if (!(argVal == "(void*)0" || argVal == "NULL" || argVal == "0")) {
                            diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                                "P7.5: ByRef Variant arg is not addressable, passing empty Variant (value dropped): " + argVal);
                        }
                        argVal = "&(vb6_VARIANT){0}";
                    } else {
                        argVal = "&(" + cType + "){" + argVal + "}";
                    }
                }
            }
        };

        // Place named args at their parameter positions
        for (auto& named : node.named) {
            std::string lower = named.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            auto it = paramMap.find(lower);
            if (it != paramMap.end()) {
                size_t pi = it->second;
                emitExpr(*named.value);
                std::string argVal = std::move(lastExpr_);
                applyByRef(argVal, pi);
                // Fix 091o: 命名实参 (Name:=expr) 此前只做 ByRef 处理, 漏了 ByVal
                // 形参类型适配 — ByVal Variant 形参收具体类型实参需
                // vb6_VariantFromValue 打包 (_Generic 宏按实类型选 ctor).
                // 例: oListenSocket.Accept(m_oSocket, UseTls:=bUseTls)
                //     (Accept 第 4 参 Optional ByVal UseTls As Variant) →
                //     C2440 int16_t→vb6_VARIANT (cWinsock.c 1604).
                // 类对象实参保持指针直传 (同 4824 Fix 084d).
                bool argIsClassObj091o = !inferClassTypeOfExpr(*named.value).empty();
                // <vbeclipse>: 实参是类对象 ≠ 槽收类对象。工程类形参 (p.type 是 Variant、
                // C 侧却是 vb6_cls_X*) 在本路径同样要按**槽**交付, 与 arg_emit/arg_variant
                // 同一判据 (cParamClassPtrType)。
                if (pi < calleeParams.size()) {
                    const std::string cpNo = cParamClassPtrType(calleeParams[pi]);
                    if (!cpNo.empty()) { argVal = deliverToClassSlot(cpNo, argVal); }
                }
                if (pi < calleeParams.size() && calleeParams[pi].isByVal
                    && calleeParams[pi].type == Vb6Type::Variant && !argIsClassObj091o
                    && cParamClassPtrType(calleeParams[pi]).empty()) {
                    // 同 arg_emit 那一处: 装箱只走 boxToVariant 这一权威 (见该处注释)
                    argVal = boxToVariant(named.value.get(), argVal);
                }
                if (pi < orderedArgs.size()) {
                    orderedArgs[pi] = std::move(argVal);
                    filled[pi] = true;
                }
            }
        }

        // Fill gaps with optional/missing param defaults
        for (size_t i = 0; i < totalParams; i++) {
            if (!filled[i]) {
                if (calleeParams[i].isParamArray) {
                    orderedArgs[i] = "NULL";
                    // filled[i] stays false for actuallyPassed tracking
                } else {
                    // Fix 190: 缺省值须按形参传值方式包装 (ByRef → &(T){缺省值}).
                    // 原先直接给 defaultValue(type) 的值, Optional ByRef String 会
                    // 把 BSTR 值交给声明为 BSTR* 的形参 → 被调方解引用後 AV.
                    orderedArgs[i] = defaultArgForParam(calleeParams[i].type,
                                                        calleeParams[i].isByVal,
                                                        calleeParams[i].defaultValueExpr);
                    // filled[i] stays false for actuallyPassed tracking
                }
            }
        }

        // Fix 081c: Save which params were actually passed (before gap-filling set all to true)
        actuallyPassedParams = filled;
        // Fix 142: 省略的位置实参不算真正传入 (_has_ 标志须为 0), 但其缺省值
        // 已由上方实参发射阶段处理并写入 orderedArgs/args, 这里只改标志.
        for (size_t oi142 : node.omittedArgs) {
            if (oi142 < actuallyPassedParams.size()) actuallyPassedParams[oi142] = false;
        }

        args = std::move(orderedArgs);
    } else {
        // No named args or no param info: append named args as-is (fallback)
        for (auto& named : node.named) {
            emitExpr(*named.value);
            args.push_back(std::move(lastExpr_));
        }
    }

    // P14.1.5: ParamArray packing - if callee has a ParamArray parameter,

    // pack extra arguments into a SAFEARRAY* and adjust argList

    std::string argList;

    int paIndex = -1;  // index of ParamArray parameter in calleeParams

    for (size_t i = 0; i < calleeParams.size(); i++) {

        if (calleeParams[i].isParamArray) { paIndex = (int)i; break; }

    }



    if (paIndex >= 0) {

        // Split args: normal args [0..paIndex-1] + ParamArray args [paIndex..end]

        int normalCount = paIndex;  // number of non-ParamArray params

        int paArgCount = (int)args.size() - normalCount;

        if (paArgCount < 0) paArgCount = 0;



        // Build normal argList

        for (int i = 0; i < normalCount && i < (int)args.size(); i++) {

            if (i > 0) argList += ", ";

            argList += args[i];

        }



        // Pack ParamArray args into SAFEARRAY* using a temp variable

        std::string paVar = "_pa_" + std::to_string(tempCounter_++);

        if (paArgCount > 0) {

            c_.emitLine("SAFEARRAY* " + paVar + " = vb6_PA_Create(" + std::to_string(paArgCount) + ");");

            for (int i = 0; i < paArgCount; i++) {

                int argIdx = normalCount + i;

                if (argIdx < (int)args.size()) {

                    // Fix 122: ParamArray 省略实参位 (DeArray x, A, , C 的空位) 在上方
                    // 缺省值填充阶段被填成 defaultValue(Variant)=vb6_VariantEmpty(),
                    // 此处按 PA_SetLong 发射 → PA_SetLong(pa, i, vb6_VARIANT) C2440
                    // (cToolsArray.c 123). VB6 语义: ParamArray 省略位 = Missing
                    // (VT_ERROR/DISP_E_PARAMNOTFOUND), IsMissing(OutVars(i)) 才为真.
                    // 复用 vb6_ComPackMissing() (返回堆上 VARIANT*, SafeArrayPutElement
                    // 会复制内容), 发射后立即释放.
                    if (node.omittedArgs.count((size_t)argIdx)) {
                        c_.emitLine("{ void* _pa_missing_" + std::to_string(tempCounter_) + " = vb6_ComPackMissing(); vb6_PA_SetVariant(" + paVar + ", " + std::to_string(i) + ", (VARIANT*)_pa_missing_" + std::to_string(tempCounter_) + "); vb6_ComVarFree(_pa_missing_" + std::to_string(tempCounter_) + "); }");
                        tempCounter_++;
                        continue;
                    }

                    std::string paArgExpr = args[argIdx];

                    bool isLongArg = false;

                    bool isDoubleArg = false;

                    if (argIdx < (int)node.positional.size()) {

                        auto& paArg = node.positional[argIdx];

                        if (paArg->kind == ASTNodeKind::LiteralExpr) {

                            auto& lit = static_cast<LiteralExpr&>(*paArg);
                            if (lit.literalKind == LiteralKind::Integer) isLongArg = true;
                            else if (lit.literalKind == LiteralKind::Double) isDoubleArg = true;

                        }
                    }

                    /* Fix 082: Check if arg is a COM interface pointer or VarPtr result.
                       These are pointer-sized and must use PA_SetLongPtr on x64. */
                    bool isComPtrArg = false;
                    if (argIdx < (int)node.positional.size()) {
                        auto& paArg = node.positional[argIdx];
                        // Check for VarPtr/ObjPtr/StrPtr call
                        if (paArg->kind == ASTNodeKind::IndexOrCallExpr) {
                            auto& callNode = static_cast<IndexOrCallExpr&>(*paArg);
                            if (callNode.callee && callNode.callee->kind == ASTNodeKind::IdentifierExpr) {
                                auto& ident = static_cast<IdentifierExpr&>(*callNode.callee);
                                std::string vpLower = ident.name;
                                std::transform(vpLower.begin(), vpLower.end(), vpLower.begin(), ::tolower);
                                if (vpLower == "varptr" || vpLower == "objptr" || vpLower == "strptr") {
                                    isComPtrArg = true;
                                }
                            }
                        }
                        // Check for LongPtr/Object variable (COM interface pointer)
                        if (paArg->kind == ASTNodeKind::IdentifierExpr) {
                            auto& ident = static_cast<IdentifierExpr&>(*paArg);
                            std::string identLower = ident.name;
                            std::transform(identLower.begin(), identLower.end(), identLower.begin(), ::tolower);
                            if (knownLongPtrVars_.count(identLower)) {
                                isComPtrArg = true;
                            }
                        }
                    }
                    /* Also detect from C expression: VarPtr result, address-of, COM iface */
                    if (!isComPtrArg) {
                        isComPtrArg = (paArgExpr.find("&(") != std::string::npos ||
                                      paArgExpr.find("(intptr_t)") != std::string::npos ||
                                      paArgExpr.find("vb6_ComIface_") != std::string::npos);
                    }

                    if (isLongArg) {

                        c_.emitLine("vb6_PA_SetLong(" + paVar + ", " + std::to_string(i) + ", " + paArgExpr + ");");

                    } else if (isDoubleArg) {

                        c_.emitLine("vb6_PA_SetDouble(" + paVar + ", " + std::to_string(i) + ", " + paArgExpr + ");");

                    } else {

                        bool looksLikeBSTR = (paArgExpr.find("vb6_BSTR") != std::string::npos ||

                                             paArgExpr.find("L\"") != std::string::npos);

                        if (looksLikeBSTR) {

                            c_.emitLine("vb6_PA_SetBSTR(" + paVar + ", " + std::to_string(i) + ", " + paArgExpr + ");");

                        } else if (isComPtrArg) {

                            /* Fix 082: COM interface pointers and VarPtr() results need PA_SetLongPtr
                               (VT_I8 / intptr_t) to avoid pointer truncation on x64. */
                            c_.emitLine("vb6_PA_SetLongPtr(" + paVar + ", " + std::to_string(i) + ", (intptr_t)" + paArgExpr + ");");

                        } else {

                            c_.emitLine("vb6_PA_SetLong(" + paVar + ", " + std::to_string(i) + ", " + paArgExpr + ");");

                        }

                    }

                }

            }

        } else {

            c_.emitLine("SAFEARRAY* " + paVar + " = NULL;");

        }



        // Append SAFEARRAY* to argList

        if (!argList.empty()) argList += ", ";

        argList += paVar;



        // P14.1.4 Optional padding for params BEFORE the ParamArray

        if (normalCount > (int)args.size()) {

            for (int i = (int)args.size(); i < normalCount; i++) {

                if (!argList.empty()) argList += ", ";

                const auto& param = calleeParams[i];

                std::string defVal;

                if (param.hasDefaultValue && !param.defaultValueExpr.empty()) {

                    defVal = param.defaultValueExpr;

                } else {

                    defVal = defaultValue(param.type);

                }

                if (param.isByVal) {

                    argList += defVal;

                } else {

                    std::string cType = mapType(param.type);

                    // P20-36: Variant/struct types can't use {funcCall()} compound literal
                    if (param.type == Vb6Type::Variant || param.type == Vb6Type::Empty ||
                        param.type == Vb6Type::Null || param.type == Vb6Type::Object) {
                        argList += "&(" + cType + "){0}";
                    } else {
                        argList += "&(" + cType + "){" + defVal + "}";
                    }

                }

            }

        }

        // P20-36: IsMissing _has_ flags for Optional params before ParamArray
        for (int i = 0; i < normalCount; i++) {
            const auto& param = calleeParams[i];
            if (param.isOptional && !param.isParamArray) {
                if (!argList.empty()) argList += ", ";
                // Fix 142: 省略的位置实参视为未传 → _has_=0
                argList += (i < (int)args.size() && !node.omittedArgs.count((size_t)i)) ? "1" : "0";
            }
        }

    } else {

        // No ParamArray - normal argList construction
        // Fix 041c: Truncate extra args when more args than calleeParams.
        // This handles Declare functions called with more args than their C signature
        // (e.g., CallWindowProcW with 8 VB6 args but 5 C params) and Property Get
        // called with index args that should have been default-member calls.
        size_t maxArgs = args.size();
        if (calleeParamsFound && !calleeIsBuiltin && args.size() > calleeParams.size()) {
            maxArgs = calleeParams.size();
        }

        for (size_t i = 0; i < maxArgs; i++) {

            if (i > 0) argList += ", ";

            argList += args[i];

        }

    }

    // P8.1: UBound/LBound - 1D鐢╲b6_UBound, ND鐢╲b6_UBoundND/vb6_LBoundND

    // P14.1.5: Also handle UBound/LBound on ParamArray parameters
