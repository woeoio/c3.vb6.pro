// cgen_expr_call_opt_pad.inc - visit(IndexOrCallExpr&) 函数体片段：Optional 参数补齐 + IsMissing _has_ 标志 + 类方法隐式 me 实参插入
// 由 src/backend/detail/expr/cgen_expr_call_pad_conv.inc 拆出（2026-09-19），
// 在函数体内被 #include，非独立编译单元。内容 = 拆分前该文件第 5~81 行，逐行未改。

    if (calleeParams.size() > 0 && args.size() < calleeParams.size() && paIndex < 0 && !calleeIsBuiltin) {
        for (size_t i = args.size(); i < calleeParams.size(); i++) {
            if (i > 0 || !args.empty()) argList += ", ";
            const auto& param = calleeParams[i];
            // Determine the default value (explicit or type-zero)
            std::string defVal;
            if (param.hasDefaultValue && !param.defaultValueExpr.empty()) {
                defVal = param.defaultValueExpr;
            } else {
                defVal = defaultValueForParam(param);
            }
            // ByRef params need pointer, ByVal need value
            if (param.isByVal) {
                argList += defVal;
            } else {
                // ByRef: pass address of compound literal: &(type){defVal}
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

                    // Fix 056: ND版本需要vb6_SafeArrayND*参数, 动态数组声明为1D*
                    if (!args.empty()) {
                        args[0] = "(vb6_SafeArrayND*)" + args[0];
                    }

                }
    // P20-36: IsMissing support - append _has_ flags for Optional params
    // For each Optional param in calleeParams: 1 if actually passed, 0 if padded
    // Fix 030b: builtin 跳过 (RTL C 函数无 _has_ 尾叜)
    // Fix 042a: Declare 函数也跳过 (Declare C 签名无 _has_ 尾叜, 但 Optional padding 仍需要)
    if (calleeParams.size() > 0 && paIndex < 0 && !calleeIsBuiltin && !calleeIsDeclare) {
        bool hasOptional = false;
        for (size_t i = 0; i < calleeParams.size(); i++) if (calleeParams[i].isOptional && !calleeParams[i].isParamArray) { hasOptional = true; break; }
        if (hasOptional) {
        for (size_t i = 0; i < calleeParams.size(); i++) {
            const auto& param = calleeParams[i];
            if (param.isOptional && !param.isParamArray) {
                if (!argList.empty()) argList += ", ";
                // Fix 081c: For named-arg path, use actuallyPassedParams to determine _has_ flag.
                // For positional-arg path, args.size() already equals actual arg count (padding
                // doesn't increase args.size()), so (i < args.size()) is correct.
                if (!actuallyPassedParams.empty() && i < actuallyPassedParams.size()) {
                    argList += actuallyPassedParams[i] ? "1" : "0";
                } else {
                    // Fix 142: 省略的位置实参 (node.omittedArgs) 视为未传 → _has_=0
                    argList += (i < args.size() && !node.omittedArgs.count(i)) ? "1" : "0";
                }
            }
        }
        }  // end if (hasOptional)
    }  // end if (calleeParams.size() > 0 && ...)

    // P6.6: 类模块中调用同类方法(包括递归), 需要自动添加me作为第一个参数
    // 如 Factorial(N-1) -> vb6_MathLib_Factorial(me, (N-1))
    if (classMethodObjArg.empty() && isClassModule_ && currentProc_) {
        std::string modPrefix = "vb6_" + cIdent(moduleName_) + "_";
        if (callee.find(modPrefix) == 0) {
            classMethodObjArg = "(void*)me";
        }
    }

    // P6.5: 如果classMethodObjArg非空, 需要将其作为第一个参数插入
    if (!classMethodObjArg.empty()) {
        if (argList.empty()) {
            argList = classMethodObjArg;
        } else {
            argList = classMethodObjArg + ", " + argList;
        }
    }

