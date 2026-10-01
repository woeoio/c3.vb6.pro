// cgen_expr_call_builtin_fixup.inc - visit(IndexOrCallExpr&) 函数体片段：内置函数调用点修正（UBound/LBound、字符串/日期/财务函数、实参个数）
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 2861~3250 行，逐行未改。

    if (callee == "vb6_UBound" || callee == "vb6_LBound") {

        // P14.1.5: Check if first arg is a ParamArray parameter

        bool firstArgIsPA = false;

        if (node.positional.size() >= 1 && currentProc_) {

            auto& firstArg = node.positional[0];

            if (firstArg->kind == ASTNodeKind::IdentifierExpr) {

                std::string argLower = static_cast<IdentifierExpr&>(*firstArg).name;

                std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);

                for (auto& p : currentProc_->params) {

                    if (p.isParamArray) {

                        std::string pLower = p.name;

                        std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);

                        if (pLower == argLower) {

                            firstArgIsPA = true;

                            break;

                        }

                    }

                }

            }

        }



        if (firstArgIsPA) {

            // ParamArray: UBound(args) -> vb6_PA_UBound(args), LBound(args) -> vb6_PA_LBound(args)

            if (callee == "vb6_UBound") {

                callee = "vb6_PA_UBound";

            } else {

                callee = "vb6_PA_LBound";

            }

            // argList should be just the PA variable name (no dimension param)

            if (args.size() == 1) {

                argList = args[0];

            }

        } else {

            // Fix 090n: UBound/LBound 首参若是 Variant 数组表达式 (Variant 变量 /
            // UDT 的 As Variant 字段如 vb6_ret_X.BufferArray), 需提取 SafeArray1D*:
            // RTL 签名 vb6_UBound(vb6_SafeArray1D*, ...), 直接传 VARIANT → C2440
            // (cZipArchive pvVfsCreate: UBound(pvVfsCreate.BufferArray, 1)).
            // <vbeclipse>: 同一条边界 —— 首参若是"已经返回载体"的内置调用
            // (`UBound(Split(s, ","))`) 再包一层就是 C2440, 用 carrier 判定跳过。
            if (!args.empty() && node.positional.size() >= 1
                && !cExprIsSafeArrayCarrier(args[0])) {
                bool arrFlag090n = false;
                bool vFlag090n = isDefinitelyVariantExpr(*node.positional[0], &arrFlag090n);
                if (vFlag090n) {
                    args[0] = "vb6_VariantToSafeArray1D(&" + args[0] + ")";
                }
            }

            if (args.size() == 1) {

                // 缺省维度参数, 补1

                argList += ", 1";

            }

            // 检查是否为ND数组, 需要用ND版本

            // Bug #1 fix: 两种检测方式:
            // 1) arrayDimCounts_中有记录且维度>1 (适用于Dim声明的ND数组)
            // 2) dimension参数值>1 (适用于ByRef参数, 声明时维度未知但调用时指定)

            bool useND = false;

            if (node.positional.size() >= 1) {

                auto& firstArg = node.positional[0];

                std::string arrLower;

                if (firstArg->kind == ASTNodeKind::IdentifierExpr) {

                    arrLower = static_cast<IdentifierExpr&>(*firstArg).name;

                    std::transform(arrLower.begin(), arrLower.end(), arrLower.begin(), ::tolower);

                }

                auto itDc = arrayDimCounts_.find(arrLower);

                if (itDc != arrayDimCounts_.end() && itDc->second > 1) {

                    useND = true;

                }

                // Bug #1 fix: 如果dimension参数值>1, 也使用ND版本

                // dimension参数是第二个参数 (index 1 in args, after the array arg)

                if (!useND && args.size() >= 2) {

                    // args[1] 是 dimension 参数的C表达式, 检查是否是常量>1

                    const std::string& dimExpr = args[1];

                    // 尝试解析为整数常量

                    try {

                        int dimVal = std::stoi(dimExpr);

                        if (dimVal > 1) {

                            useND = true;

                        }

                    } catch (...) {

                        // 非常量表达式, 无法确定; 不启用ND

                    }

                }

                // Bug #1 fix (082h): 第三种检测 - 同一过程中已有UBound(arr,N>1)使用ND版本

                // 例如: UBound(uVectors, 2) 用了ND版本, 则 UBound(uVectors, 1) 也应使用ND版本

                if (!useND && knownNDArraysInProc_.count(arrLower)) {

                    useND = true;

                }

                if (useND) {

                    // 注册到过程内ND数组集合, 后续 UBound(arr,1) 也会用ND版本

                    if (!arrLower.empty()) {

                        knownNDArraysInProc_.insert(arrLower);

                    }

                    // ND数组 -> 使用vb6_UBoundND/vb6_LBoundND

                    if (callee == "vb6_UBound") {

                        callee = "vb6_UBoundND";

                    } else {

                        callee = "vb6_LBoundND";

                    }

                    // ND版本需要(vb6_SafeArrayND*)转换第一个参数

                    // argList格式是 "arrExpr, dimExpr", 需要改为 "(vb6_SafeArrayND*)(arrExpr), dimExpr"

                    if (args.size() >= 2) {

                        argList = "(vb6_SafeArrayND*)(" + args[0] + "), " + args[1];

                    } else if (args.size() == 1) {

                        argList = "(vb6_SafeArrayND*)(" + args[0] + "), 1";

                    }

                }

            }

        }

    }    // InStr: VB6允许2参数形式 InStr(string1, string2)
    // RTL: vb6_InStr(start, haystack, needle) → 2参数时补start=1
    if (callee == "vb6_InStr" || callee == "vb6_InStrB") {
        if (args.size() == 2) {
            argList = "1, " + argList;
        }
    }

    // P14.2.2: CurDir - VB6 allows 0-arg CurDir() → vb6_CurDir(NULL)
    if (callee == "vb6_CurDir") {
        if (args.empty()) {
            argList = "NULL";
        }
    }

    // Fix <vbeclipse>: LoadPicture() — VB6 允许**无参**调用 (取回"空图片", 常用于
    // `Screen.MouseIcon = LoadPicture()` 复位鼠标指针; ucPerspective.ctl:596/600)。
    // RTL 签名是 vb6_LoadPictureEx(BSTR pathname) (1 参), 0 参调用 → C2198 参数太少
    // (ucPerspective.c:467/472)。补 NULL: vb6_LoadPictureEx 遇 NULL 直接返回 NULL,
    // 正是"空图片"的语义。
    if (callee == "vb6_LoadPictureEx") {
        if (args.empty()) {
            argList = "NULL";
        }
    }

    // P14.2.2: Dir - VB6 allows 1-arg Dir(pattern) → vb6_Dir(pattern, 0)
    if (callee == "vb6_Dir") {
        if (args.size() == 1) {
            argList += ", 0";
        }
        // Fix 086: 无参 Dir() (继续上次搜索) → 传 NULL, 0
        if (args.empty()) {
            argList = "NULL, 0";
        }
    }

    // <vbeclipse>: VB6 里 InStr/InStrRev/StrComp/Replace/Split/Filter 的 compare 形
    // **缺省值不是常量 0, 而是本模块的 Option Compare**。下面各 padding 段统一用这个
    // 串填 compare 槽 (旧写法一律硬编码 0 ⇒ Option Compare Text 模块里省略 compare 的
    // 内置函数仍按二进制比, 与同模块的 `=` 判定互相矛盾)。
    const std::string cmpDefault = optionCompareText_ ? "1" : "0";

    // P14.2.3: Split - VB6 Split(expr[, delim[, limit[, compare]]])
    // RTL: vb6_Split(expr, delim, limit, compare)
    if (callee == "vb6_Split") {
        if (args.size() == 1) {
            argList += ", NULL";        // default delimiter = space
        }
        if (args.size() <= 2) {
            argList += ", -1";          // limit = -1 (unlimited)
        }
        if (args.size() <= 3) {
            argList += ", " + cmpDefault;
        }
    }

    // P14.2.3: Join - VB6 Join(arr[, delimiter])
    // RTL: vb6_Join(arr, delimiter)
    if (callee == "vb6_Join") {
        if (args.size() == 1) {
            argList += ", NULL";        // default delimiter = space
        }
    }

    // <vbeclipse>: Format(expression[, format]) 的缺省 format 补在 conv_output 段
    // (那里 args[0] 装箱后会重建 argList, 在这补会被抹掉), 见 cgen_expr_call_conv_output.inc。

    // P14.2.4: DateDiff - VB6 DateDiff(interval, date1, date2[, firstDayOfWeek[, firstWeekOfYear]])
    // RTL: vb6_DateDiff(interval, date1, date2, firstDayOfWeek, firstWeekOfYear)
    if (callee == "vb6_DateDiff") {
        if (args.size() == 3) {
            argList += ", 1, 1";  // vbSunday, vbFirstJan1
        } else if (args.size() == 4) {
            argList += ", 1";     // vbFirstJan1
        }
    }

    // P14.2.4: DatePart - VB6 DatePart(interval, date[, firstDayOfWeek[, firstWeekOfYear]])
    // RTL: vb6_DatePart(interval, date, firstDayOfWeek, firstWeekOfYear)
    if (callee == "vb6_DatePart") {
        if (args.size() == 2) {
            argList += ", 1, 1";  // vbSunday, vbFirstJan1
        } else if (args.size() == 3) {
            argList += ", 1";     // vbFirstJan1
        }
    }

    // Replace: VB6允许3参数形式 Replace(string, find, replacement)
    // RTL: Replace(string, find, replacement, start, count, compare)
    if (callee == "vb6_Replace") {
        if (args.size() == 3) {
            argList += ", 1, -1, " + cmpDefault;
        } else if (args.size() == 4) {
            argList += ", -1, " + cmpDefault;
        } else if (args.size() == 5) {
            argList += ", " + cmpDefault;
        }
    }

    // P21-B: Weekday(date[, firstDayOfWeek]) - default firstDayOfWeek=1 (vbSunday)
    if (callee == "vb6_Weekday") {
        if (args.size() == 1) {
            argList += ", 1";
        }
    }

    // Fix 110l: InputBox(prompt[, title[, default[, xpos[, ypos[, helpfile[, context]]]]]]])
    // RTL: vb6_InputBox(prompt, title, defaultstr, xpos, ypos, helpfile, context) — 7 参.
    // VB6 允许只写前 1~3 个实参 (Charts 2020 ppProgressCircular.pag 351:
    //   InputBox("...", , 100) → vb6_InputBox(prompt, 0L, 100) → C2198 参数太少).
    if (callee == "vb6_InputBox") {
        static const char* kInputBoxDefaults[7] = {
            nullptr,                 // prompt (必填)
            "vb6_BSTR_Empty()",      // title
            "vb6_BSTR_Empty()",      // defaultstr
            "0",                     // xpos
            "0",                     // ypos
            "vb6_BSTR_Empty()",      // helpfile
            "0",                     // context
        };
        for (size_t i = args.size(); i < 7; i++) {
            argList += ", ";
            argList += kInputBoxDefaults[i];
        }
    }

    // P21-10: FormatDateTime(date[, namedFormat]) - default namedFormat=0 (vbGeneralDate)
    if (callee == "vb6_FormatDateTime") {
        if (args.size() == 1) {
            argList += ", 0";
        }
    }

    // P21-16: NPer(rate, pmt, pv[, fv][, type]) - defaults: fv=0, type=0
    if (callee == "vb6_NPer") {
        if (args.size() == 3) argList += ", 0, 0";
        else if (args.size() == 4) argList += ", 0";
    }

    // P21-19: IRR(values[, guess]) - default guess=0.1
    if (callee == "vb6_IRR") {
        if (args.size() == 1) argList += ", 0.1";
    }

    // P20-37: GetSetting(app, section, key[, default]) - default is empty string
    if (callee == "vb6_GetSetting") {
        if (args.size() == 3) argList += ", vb6_BSTR_Empty()";
    }

    // Fix 010e: Err.Raise with Source/Description arguments
    // Err.Raise always calls vb6_ErrRaise(number, source, description) — 3 args
    // Pad with NULL BSTRs when fewer args provided
    if (callee == "vb6_ErrRaise" && args.size() < 3) {
        // Pad to 3 args
        while (args.size() < 3) {
            args.push_back("(BSTR)0");
            if (!argList.empty()) argList += ", ";
            argList += "(BSTR)0";
        }
    }
    if (callee == "vb6_ErrRaise" && args.size() > 3) {
        // Truncate to 3 (Err.Raise can have up to 5 VB6 args, C RTL handles 3)
        argList = args[0] + ", " + args[1] + ", " + args[2];
    }

    // Fix 010r: Builtin RTL functions with optional parameters — pad to required arg count
    // VB6 allows omitting optional args; C RTL functions require all args
    if (callee == "vb6_Mid" && args.size() < 3) {
        // czUI fix: 省略的 length 参数用 -1 哨兵 (RTL 按"取到结尾"处理),
        // 不能补 0 — 那是字面长度 0 (VB6 返回空串), 曾让 HexToColor("#243447")
        // 解析成 0 (纯黑), 整个 czUI demo 颜色全错。
        while (args.size() < 3) {
            args.push_back("-1");
            if (!argList.empty()) argList += ", ";
            argList += "-1";
        }
    }
    if (callee == "vb6_StrConv" && args.size() < 3) {
        while (args.size() < 3) {
            args.push_back("0");
            if (!argList.empty()) argList += ", ";
            argList += "0";
        }
    }
    if (callee == "vb6_FormatNumber" && args.size() < 5) {
        while (args.size() < 5) {
            args.push_back("-1");  // -1 = vbUseDefault for optional args
            if (!argList.empty()) argList += ", ";
            argList += "-1";
        }
    }
    // Fix <vbeclipse> <VBFlexGridDemo>: Clipboard.SetData 二形 data / data, format
    //   VB 语义 `[format]` 可省, 但 RTL 已扩成 2 参 (format=vbCF* 常量, 0=Bitmap)。
    //   1 参调用直接补 format=0，与 `vb6_Clipboard_SetData(void*, int32_t)` 对齐。
    if (callee == "vb6_Clipboard_SetData" && args.size() < 2) {
        while (args.size() < 2) {
            args.push_back("0");
            if (!argList.empty()) argList += ", ";
            argList += "0";
        }
    }
    // Fix 092w: StrConv(Byte数组, vbUnicode) → vb6_StrConvFromByteArray
    // 反向转换: Byte() -> String. 第一实参为已知 Byte 数组变量 (vb6_SafeArray1D*)
    // 时, RTL vb6_StrConv 期望 BSTR 会类型不匹配; 改用返回 BSTR 的 helper.
    if (callee == "vb6_StrConv" && !args.empty()) {
        std::string arg0name = args[0];
        if (arg0name.compare(0, 4, "me->") == 0) arg0name = arg0name.substr(4);
        if (arg0name.size() > 4 && arg0name[0] == '(' && arg0name[1] == '*'
            && arg0name.back() == ')') {
            arg0name = arg0name.substr(2, arg0name.size() - 3);
        }
        std::string lower = arg0name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownByteArrayVars_.count(lower)) {
            callee = "vb6_StrConvFromByteArray";
        }
    }
    // Fix 140: StrPtr(Byte数组) → 数组数据指针 (而非 SafeArray1D* 结构体指针).
    // 场景: LabelPlus.ctl GdipMeasureString/GdipAddPathString(StrPtr(m_Caption), -1, ...)
    // — m_Caption 是 Byte() 数组, 但 GDI+ 需要 UTF-16 字符指针, 必须取 arr->data.
    if (callee == "vb6_StrPtr" && !args.empty()) {
        std::string a140 = args[0];
        if (a140.compare(0, 4, "me->") == 0) a140 = a140.substr(4);
        if (a140.size() > 4 && a140[0] == '(' && a140[1] == '*'
            && a140.back() == ')') {
            a140 = a140.substr(2, a140.size() - 3);
        }
        std::string l140 = a140;
        std::transform(l140.begin(), l140.end(), l140.begin(), ::tolower);
        if (a140.find_first_of("( .") == std::string::npos
            && (knownByteArrayVars_.count(l140) || classByteArrayMembers_.count(l140))) {
            callee = "vb6_SafeArrayDataPtr";
        }
    }
    // Fix 034: Builtin RTL functions with Optional params — pad to required C RTL arg count.
    // calleeParams 已对 builtin 标 isOptional, 但 general padding (line 3291+) 被 !calleeIsBuiltin
    // 跳过 (RTL C 函数不接受 _has_ 尾叜). 各函数硬编码补默认值, 默认值取自 VB6 语义.
    // InStrRev(string1, string2[, start[, compare]]) — start=-1 (从末尾), compare=本模块默认
    if (callee == "vb6_InStrRev") {
        if (args.size() == 2) {
            argList += ", -1, " + cmpDefault;
        } else if (args.size() == 3) {
            argList += ", " + cmpDefault;
        }
    }
    // Round(x[, decimals]) — decimals default = 0
    if (callee == "vb6_Round" && args.size() == 1) {
        argList += ", 0";
    }
    // StrComp(s1, s2[, compare]) — compare 缺省 = 本模块 Option Compare
    if (callee == "vb6_StrComp" && args.size() == 2) {
        argList += ", " + cmpDefault;
    }
    // Shell(pathname[, windowstyle]) — windowstyle default = 2 (vbMinimizedFocus)
    if (callee == "vb6_Shell" && args.size() == 1) {
        argList += ", 2";
    }
    // Fix 158j: AppActivate(title[, wait]) — RTL 签名 vb6_AppActivate(BSTR, int32_t wait),
    // VB6 允许省略 wait (Startup.bas: AppActivate "VBFlexGrid Demo" → 只传 1 参
    // → C2198 参数太少). 缺省值 0 (不等待).
    if (callee == "vb6_AppActivate" && args.size() == 1) {
        argList += ", 0";
    }
    // Fix 135b: 无参 Rnd 语义 = Rnd(1) (取下一个随机数); 0 才是"重复上一值"
    // Randomize([seed]) — handled as Sub bare-call; see cgen_stmt.cpp Randomize dispatch
    if (callee == "vb6_Rnd" && args.empty()) {
        argList = "0";
    }
    // Randomize As Function-call form `Randomize()` — pad seed = 0.0
    if (callee == "vb6_Randomize" && args.empty()) {
        argList = "0.0";
    }
    // InStr 的 compare 形 (<vbeclipse>, 替换 Fix 041 的"截掉第 4 参"):
    //   InStr(s1, s2[, compare])            首参是字符串 ⇒ (s1, s2, compare)
    //   InStr(start, s1, s2[, compare])     首参是数值   ⇒ (start, s1, s2[, compare])
    // 旧写法把 4 参截成 3 参 ⇒ vbTextCompare 被静默丢掉; 而"字符串优先的三参形"三个实参
    // 错接到 (start, haystack, needle) 三个槽 ⇒ needle 收到整数 1 → 解引用地址 1 崩溃
    // (实测 InStr(s, "world", vbTextCompare) 段错误)。带 compare 的一律改发 vb6_InStrC。
    if (callee == "vb6_InStr") {
        bool firstIsString = false;
        if (!node.positional.empty()) {
            firstIsString = (inferExprType(*node.positional[0]) == Vb6Type::String);
        }
        if (args.size() == 4) {
            callee = "vb6_InStrC";
            argList = args[0] + ", " + args[1] + ", " + args[2] + ", " + args[3];
        } else if (args.size() == 3 && firstIsString) {
            callee = "vb6_InStrC";
            argList = "1, " + args[0] + ", " + args[1] + ", " + args[2];
        } else if (optionCompareText_ && (args.size() == 2 || args.size() == 3)) {
            // 省略 compare ⇒ 缺省取本模块的 Option Compare (上面那段 2 参补 "1, "
            // 只动 argList 不动 args, 所以这里按 args 重建)。
            callee = "vb6_InStrC";
            argList = (args.size() == 2 ? "1, " : "") + args[0] + ", " + args[1]
                      + (args.size() == 2 ? "" : ", " + args[2]) + ", 1";
        }
    }
    // <vbeclipse>: Filter(array, string[, include[, compare]]) — RTL 签名 4 参, 此前完全没有
    // padding ⇒ `Filter(a, "x")` 这一 VB6 合法形根本编不过 (C2198; 首参登记成 Variant|Array
    // 之前 Variant 实参还 C2440)。include 缺省 True (VB6 的 True = -1, RTL 侧按 !=0 归一),
    // compare 缺省 0 = vbBinaryCompare。
    if (callee == "vb6_Filter") {
        if (args.size() == 2) {
            argList += ", -1, " + cmpDefault;
        } else if (args.size() == 3) {
            argList += ", " + cmpDefault;
        }
    }

    // P14.1.4: General Optional parameter padding for user-defined functions
    // calleeParams is empty for builtin RTL functions (registered without params), so they're auto-skipped
    // Fix 030b: builtin 即使现在有 calleeParams (用于触发包装), 也不走 padding/IsMissing 路径
    // (RTL C 签名不接受尾叜 _has_ flag, 默认值由 builtin 的特殊 codegen 处理如 UBound 补 dimension=