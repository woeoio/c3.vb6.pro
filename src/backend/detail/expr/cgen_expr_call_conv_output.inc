// cgen_expr_call_conv_output.inc - visit(IndexOrCallExpr&) 函数体片段：输出侧适配：MsgBox 首参包装 + Format 首参包装 + CStr 兜底包装
// 由 src/backend/detail/expr/cgen_expr_call_pad_conv.inc 拆出（2026-09-19），
// 在函数体内被 #include，非独立编译单元。内容 = 拆分前该文件第 249~353 行，逐行未改。

    // MsgBox自动BSTR转换: MsgBox期望BSTR, 非BSTR的prompt参数需包装
    // 注意: 只转换第一个参数(prompt), 不能把整个argList包进去(MsgBox v, , title时argList含3个参数)
    if (callee == "vb6_MsgBox" || callee == "vb6_MsgBox1") {
        if (!args.empty()) {
            std::string& firstArg = args[0];
            if (firstArg.find("vb6_ComGetIntProp") == 0 ||
                firstArg.find("vb6_ComVtableGetInt") == 0) {
                firstArg = "vb6_CStrLong(" + firstArg + ")";
            } else if (firstArg.find("vb6_ComGetDoubleProp") == 0 ||
                       firstArg.find("vb6_ComVtableGetDouble") == 0) {
                firstArg = "vb6_CStrDbl(" + firstArg + ")";
            } else if (firstArg.find("vb6_VariantFromComResult") == 0) {
                // COM调用结果(VARIANT*)→vb6_VARIANT, 需转BSTR
                firstArg = "vb6_VariantToString(" + firstArg + ")";
            } else if (firstArg.find("vb6_ComCall(") == 0) {
                // P24-01: 后期绑定COM调用返回VARIANT*, 需解包转BSTR
                firstArg = "vb6_VariantToString(vb6_VariantFromComResult(" + firstArg + "))";
            } else if (firstArg.find("vb6_VariantToString(") == 0) {
                // Fix 090u: 实参已是 VariantToString(...) (参数打包段对 Variant 实参
                // 按 BSTR 形参已转换一次, 如 MsgBox .FindFirst(...) As Variant /
                // MsgBox .Value(1,6) As Variant) — 再包一层 → VariantToString(BSTR)
                // C2440 "BSTR→vb6_VARIANT". 此处保持一层.
            } else if (!node.positional.empty() && inferExprType(*node.positional[0]) == Vb6Type::Variant) {
                // Variant类型变量/表达式: MsgBox v → vb6_VariantToString(v)
                firstArg = "vb6_VariantToString(" + firstArg + ")";
            } else if (!node.positional.empty()) {
                // Fix 161e: 上面那串只认领了 **COM 读 / Variant** 两类形态,
                // **纯数值**一条都没包 ⇒ `MsgBox Len("你好hello")` / `MsgBox 7` /
                // `MsgBox (n + 1)` / `MsgBox 3.5` / `MsgBox True` 全都把 int32_t
                // (double / int16_t) **直接当 BSTR 指针**传给 vb6_MsgBox1 ——
                // 弹窗是空的: 地址 7、123 之类根本不可读, 真读得到时显示的也是
                // 垃圾而不是数字。VB6 会隐式把数值转成字符串再显示。
                // 这里对上面没认领的形态统一过 wrapToBSTR
                // (cgen_expr_binary_util.cpp:81) —— 它按 inferExprType 分流
                // (Long/Integer/Boolean/Byte/Double/Single/Date/LongLong…),
                // 且 String 分支原样返回, 所以不会把已正确的字符串实参再包一层。
                firstArg = wrapToBSTR(firstArg, *node.positional[0]);
            }
        }
    }
    // P25: MsgBox BSTR转换可能修改了args[0], 需重建argList
    if (callee == "vb6_MsgBox" || callee == "vb6_MsgBox1") {
        argList.clear();
        for (size_t i = 0; i < args.size(); i++) {
            if (i > 0) argList += ", ";
            argList += args[i];
        }
    }
    // MsgBox(prompt) -> vb6_MsgBox1(prompt)
    // MsgBox(prompt, buttons) -> vb6_MsgBox(prompt, buttons, NULL)
    if (callee == "vb6_MsgBox") {
        if (node.positional.size() == 1) {
            callee = "vb6_MsgBox1";
        } else if (node.positional.size() == 2) {
            argList += ", NULL";
        }
    }
    // P26: Format 第一个参数需要包装为 vb6_VARIANT
    if (callee == "vb6_Format" && !args.empty()) {
        Vb6Type argType = inferExprType(*node.positional[0]);
        if (argType == Vb6Type::Long || argType == Vb6Type::Integer) {
            args[0] = "vb6_VariantLong(" + args[0] + ")";
        } else if (argType == Vb6Type::Double || argType == Vb6Type::Single) {
            args[0] = "vb6_VariantDouble(" + args[0] + ")";
        } else if (argType == Vb6Type::String) {
            args[0] = "vb6_VariantString(" + args[0] + ")";
        } else if (argType == Vb6Type::Boolean) {
            // Fix 198: 原为 vb6_VariantInt((int16_t)x) —— 那是把 VB 的 Boolean 当
            // C 的 short 装箱 (VT_I2), 于是 Format(True,"G") 打出 -1. 换真 VT_BOOL.
            args[0] = boxToVariant(node.positional[0].get(), args[0]);
        } else if (argType == Vb6Type::Byte) {
            args[0] = "vb6_VariantInt((int16_t)(" + args[0] + "))";
        } else if (argType == Vb6Type::Date) {
            args[0] = "vb6_VariantDouble((double)(" + args[0] + "))";
        }
        // Fix 036: Format fallback — 未匹配类型 (Currency/Unknown 等, 非 Variant) 用
        // vb6_VariantFromValue 包装. Variant 类型无需包装 (已是 vb6_VARIANT).
        else if (argType != Vb6Type::Variant) {
            args[0] = "vb6_VariantFromValue(" + args[0] + ")";
        }
        argList.clear();
        for (size_t i = 0; i < args.size(); i++) {
            if (i > 0) argList += ", ";
            argList += args[i];
        }
    }
    // Fix 036: CStr fallback — 当特殊分支未匹配 (Boolean/Byte/Date/Currency/常量/
    // 未注册函数返回值/Const BSTR 等), callee 仍为 vb6_CStr (接 vb6_VARIANT). 用
    // vb6_VariantFromValue 包装 args[0], _Generic 按实参 C 类型自动选择 Variant ctor,
    // 避免 concrete → VARIANT C2440. 对已是 vb6_VARIANT 的实参为 identity (no-op), 安全.
    // 但需排除确定 Variant 的表达式 (如 Me.Segments(i) 返回 Variant): _Generic 宏
    // 对某些 Variant 表达式展开可能产生逗号问题 → C2197, 故用 isDefinitelyVariantExpr
    // 跳过, 保留 vb6_CStr(variantExpr) 原样 (vb6_CStr 接 VARIANT, 直接可用).
    if (callee == "vb6_CStr" && !args.empty()) {
        bool argIsDefVariant = false;
        if (!node.positional.empty()) {
            argIsDefVariant = isDefinitelyVariantExpr(*node.positional[0]);
        }
        // Fix 038b-4: 字符串级 Variant 检测 — 如果是 C 级 Variant 表达式,
        // 也跳过 VariantFromValue 包装 (vb6_CStr 直接接 VARIANT)
        if (!argIsDefVariant) {
            argIsDefVariant = cExprIsVariant(args[0]);
        }
        if (!argIsDefVariant) {
            // Fix 121b: 同 Fix 121 — 裸 COM 结果必须解引用 (VariantFromComResult),
            // 否则 _Generic 按 void* 包装得到空 Variant。
            bool comRes121b = (args[0].compare(0, 12, "vb6_ComCall(") == 0)
                || (args[0].compare(0, 15, "vb6_ComGetProp(") == 0)
                || (args[0].compare(0, 21, "vb6_ComGetObjectProp(") == 0);
            args[0] = comRes121b
                ? "vb6_VariantFromComResult(" + args[0] + ")"
                : boxToVariant(node.positional.empty() ? nullptr
                                 : node.positional[0].get(), args[0]);
            argList.clear();
            for (size_t i = 0; i < args.size(); i++) {
                if (i > 0) argList += ", ";
                argList += args[i];
            }
        }
    }
