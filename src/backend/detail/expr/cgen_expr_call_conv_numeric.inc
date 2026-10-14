// cgen_expr_call_conv_numeric.inc - visit(IndexOrCallExpr&) 函数体片段：数值转换函数适配：CInt/CLng/CDbl/CByte/CSng/CBool 的 V 后缀与 BSTR 实参数值化
// 由 src/backend/detail/expr/cgen_expr_call_pad_conv.inc 拆出（2026-09-19），
// 在函数体内被 #include，非独立编译单元。内容 = 拆分前该文件第 165~248 行，逐行未改。

    // P8.4: Variant参数适配 — 如果目标函数不接受Variant但参数是Variant类型, 使用V后缀函数
    // CInt(Variant)→vb6_CIntV, CDbl(Variant)→vb6_CDblV, CLng(Variant)→vb6_CLngV
    // Fix 090d: 扩展覆盖 CByte/CSng/CBool (CByte("&H"&hex) 字符串实参同样需解析适配)
    if (callee == "vb6_CInt" || callee == "vb6_CLng" || callee == "vb6_CDbl"
        || callee == "vb6_CByte" || callee == "vb6_CSng" || callee == "vb6_CBool"
        // ai/032: 四档无符号/窄整型转换函数 —— 同口径 (Variant → *V, BSTR → vb6_NumVal)
        || callee == "vb6_CSByte" || callee == "vb6_CUInt"
        || callee == "vb6_CULng" || callee == "vb6_CULngLng"
        // twinBASIC 形 CLngPtr 同路 (RTL vb6_CLngPtr* 已在头里备齐三档)
        || callee == "vb6_CLngPtr") {
        // 检查第一个参数是否为Variant变量
        bool firstArgIsVariant = false;
        bool firstArgIsBstr = false;
        if (!node.positional.empty()) {
            auto& firstArg = node.positional[0];
            if (firstArg->kind == ASTNodeKind::IdentifierExpr) {
                auto& idArg = static_cast<IdentifierExpr&>(*firstArg);
                std::string argLower = idArg.name;
                std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);
                if (knownVariantVars_.count(argLower)) firstArgIsVariant = true;
                else if (knownBstrVars_.count(argLower)) firstArgIsBstr = true;
            }
            // Fix 036: 非标识符 Variant 表达式 (函数返回 Variant 的调用/类方法等) 也需 V 后缀
            if (!firstArgIsVariant) {
                firstArgIsVariant = isDefinitelyVariantExpr(*firstArg);
            }
            // Fix 038b-4: 字符串级 Variant 检测回退 — 捕获 vb6_VariantArrayGet 等表达式
            if (!firstArgIsVariant && !args.empty()) {
                firstArgIsVariant = cExprIsVariant(args[0]);
            }
        }
        if (firstArgIsVariant) {
            if (callee == "vb6_CInt") callee = "vb6_CIntV";
            else if (callee == "vb6_CLng") callee = "vb6_CLngV";
            else if (callee == "vb6_CDbl") callee = "vb6_CDblV";
            else if (callee == "vb6_CByte") callee = "vb6_CByteV";
            else if (callee == "vb6_CSng") callee = "vb6_CSngV";
            else if (callee == "vb6_CBool") callee = "vb6_CBoolV";
            // ai/032: 四档无符号/窄整型同上
            else if (callee == "vb6_CSByte")   callee = "vb6_CSByteV";
            else if (callee == "vb6_CUInt")    callee = "vb6_CUIntV";
            else if (callee == "vb6_CULng")    callee = "vb6_CULngV";
            else if (callee == "vb6_CULngLng") callee = "vb6_CULngLngV";
            else if (callee == "vb6_CLngPtr")  callee = "vb6_CLngPtrV";
            // Fix 084l: 撤销 Fix 038b-2 的 Variant→double 提前提取 (getRuntimeParamCType
            // 把 CLng/CDbl/CInt 形参当作 double, 已将 lRet 替换为 vb6_VariantToDouble(lRet)).
            // V 后缀函数直接接受 vb6_VARIANT — 保留提取会产生 vb6_CLngV(double) → C2440.
            if (!args.empty()) {
                static const char* extractPrefixes[] = {
                    "vb6_VariantToDouble(", "vb6_VariantToLong("};
                for (auto* pre : extractPrefixes) {
                    size_t pl = strlen(pre);
                    if (args[0].compare(0, pl, pre) == 0 && args[0].size() > pl + 1) {
                        args[0] = args[0].substr(pl, args[0].size() - pl - 1);
                        argList.clear();
                        for (size_t ai = 0; ai < args.size(); ai++) {
                            if (ai > 0) argList += ", ";
                            argList += args[ai];
                        }
                        break;
                    }
                }
            }
        }
        // Fix 036: BSTR 参数 → vb6_NumVal 转换为 double (CInt/CLng/CDbl/CByte/CSng/CBool
        // 接 double). Fix 090d: vb6_Val 不识别 "&H" 前缀 (Val("&HFF")=0), 而 VB6 类型
        // 转换函数解析 &H/&O → CByte("&H" & hex) 需 vb6_NumVal; 十进制语义不变.
        // Fix 084c: 扩展检测到字符串级 BSTR 表达式 (vb6_Trim/vb6_BSTR_Concat/vb6_CStr
        // /vb6_VariantToString 等), 不仅限于已知 BSTR 标识符 — 否则
        // vb6_CLng(vb6_Trim(vb6_VariantToString(...))) 触发 C2440 "BSTR → double".
        bool argIsBstrExpr = firstArgIsBstr;
        if (!argIsBstrExpr && !args.empty()) {
            std::string& a0 = args[0];
            argIsBstrExpr =
                a0.find("vb6_Trim(") == 0 || a0.find("vb6_LTrim(") == 0 ||
                a0.find("vb6_RTrim(") == 0 || a0.find("vb6_StrConv(") == 0 ||
                a0.find("vb6_Mid(") == 0 || a0.find("vb6_Left(") == 0 ||
                a0.find("vb6_Right(") == 0 || a0.find("vb6_Replace(") == 0 ||
                a0.find("vb6_String(") == 0 || a0.find("vb6_Format(") == 0 ||
                a0.find("vb6_UCase(") == 0 || a0.find("vb6_LCase(") == 0 ||
                a0.find("vb6_Space(") == 0 || a0.find("vb6_IIfBSTR(") == 0 ||
                a0.find("vb6_Chr(") == 0 || a0.find("vb6_ChrW(") == 0 ||
                a0.find("vb6_BSTR_Concat(") == 0 || a0.find("vb6_BSTR_FromStr(") == 0 ||
                a0.find("vb6_CStr") == 0 || a0.find("vb6_VariantToString(") == 0 ||
                a0.find("VB6_SA_AT(BSTR,") != std::string::npos ||
                // Fix 152g: 控件字符串属性读取器返回 BSTR/wchar_t*, 之前未列入
                // 白名单 → vb6_CLng(vb6_GetControlText(...)) 等把 BSTR 直接喂给
                // 接 double 的转换函数 → C2440 "wchar_t* → double"。VB6 里
                // CLng(Text1)/CDbl(Text1) 合法 (解析字符串), 需用 vb6_NumVal 桥接。
                // 生成码在右括号后带 "  /* ctrl prop read */" 注释, 前缀匹配不受影响。
                a0.find("vb6_GetControlText(") == 0 ||
                a0.find("vb6_GetSelText(") == 0 ||
                a0.find("vb6_GetControlTag(") == 0 ||
                a0.find("vb6_GetToolTipText(") == 0 ||
                a0.find("vb6_GetListItem(") == 0 ||
                a0.find("vb6_GetMenuCaption(") == 0;
        }
        if (argIsBstrExpr && !args.empty()) {
            args[0] = "vb6_NumVal(" + args[0] + ")";
            argList.clear();
            for (size_t i = 0; i < args.size(); i++) {
                if (i > 0) argList += ", ";
                argList += args[i];
            }
        }
    }
