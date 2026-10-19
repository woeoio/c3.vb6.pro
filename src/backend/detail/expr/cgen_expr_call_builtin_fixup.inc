// cgen_expr_call_builtin_fixup.inc - visit(IndexOrCallExpr&) 函数体片段：内置函数调用点修正（UBound/LBound、字符串/日期/财务函数、实参个数）
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 2861~3250 行，逐行未改。

    // Fix <vbeclipse> (2026-10-06): LoadRes* 实参统一按 VB6 Variant 形包装。
    // RTL 签名是 (vb6_VARIANT, ...) (VB6 里三函数的实参本就是 Variant: 数字 id 与
    // 字符串资源名都合法 —— 实测语料两边都有, cAesCBC 的 LoadResData("AES.CBC",
    // "JSCRIPT") 与 modResources 的 LoadResPicture(102, 0))。不包装的话字符串形
    // 会把 BSTR 截断进 int32 形参 (Fix 090bz 时代起就是坏的)。包一层
    // vb6_VariantFromValue: 整数字面量 → VT_I4, BSTR 表达式 → VT_BSTR, Long 变量
    // → VT_I4, RTL 按 vt 分派按名/按 id 查资源。
    if (callee == "vb6_LoadResString" || callee == "vb6_LoadResData"
        || callee == "vb6_LoadResPicture") {
        std::string joinedRes;
        for (size_t iRes = 0; iRes < args.size(); iRes++) {
            if (iRes) joinedRes += ", ";
            joinedRes += "vb6_VariantFromValue(" + args[iRes] + ")";
        }
        argList = joinedRes;
        args.clear();
    }


    // Fix <vbeclipse>: `Load <控件数组>(i)` / `Unload <控件数组>(i)` 发码接入。
    // 控件数组元素的表达式路发的是 `vb6_CtrlArr_GetAt(&vb6_arr_<名>, <idx>)`
    // (cgen_util_ctrl.cpp / cgen_expr_member_precheck.inc 的同一出口)，于是
    // `Load cmdBtn(3)` 落到 {"load","vb6_LoadForm"} 那一行后发成
    // `vb6_LoadForm(默认属性读数)` —— 对 NULL 句柄读 Caption 再拿去 Load 窗体,
    // 控件根本没被创建 (vb6_CtrlArr_Load 全仓零调用者, 2026-10-05 实测发码)。
    // 这里按 args[0] 里的 GetAt 形状拦下, 改发 vb6_CtrlArr_Load/Unload。
    // 仅认 `vb6_CtrlArr_GetAt(&vb6_arr_<名>, <idx>)` 这一个形状 (菜单数组没有
    // 这个出口, Load/Unload 对菜单数组本来也不适用); 普通窗体 `Load FrmX`
    // 不含该形状, 原样落回 vb6_LoadForm。父窗口: 窗体模块用本窗体句柄;
    // UserControl 模块用容器句柄; 都不是才给 NULL。
    if ((callee == "vb6_LoadForm" || callee == "vb6_UnloadForm") && args.size() == 1) {
        static const std::string kGetAtArr224 = "vb6_CtrlArr_GetAt(&vb6_arr_";
        size_t getAtPos224 = args[0].find(kGetAtArr224);
        if (getAtPos224 != std::string::npos) {
            size_t nameB224 = getAtPos224 + kGetAtArr224.size();
            size_t comma224 = args[0].find(',', nameB224);
            size_t open224 = args[0].find("GetAt(", getAtPos224);
            if (comma224 != std::string::npos && open224 != std::string::npos) {
                std::string arrName224 = args[0].substr(nameB224, comma224 - nameB224);
                // 从 GetAt 的 '(' 起配对括号, 取出完整第二实参 (下标可能本身是表达式)
                int depth224 = 0;
                size_t close224 = std::string::npos;
                for (size_t i224 = open224 + 5; i224 < args[0].size(); ++i224) {
                    char ch224 = args[0][i224];
                    if (ch224 == '(') depth224++;
                    else if (ch224 == ')' && --depth224 == 0) { close224 = i224; break; }
                }
                if (close224 != std::string::npos) {
                    std::string getAtArgs224 = args[0].substr(open224 + 6, close224 - (open224 + 6));
                    size_t idxComma224 = getAtArgs224.find(',');
                    std::string idxArg224 = (idxComma224 != std::string::npos)
                        ? getAtArgs224.substr(idxComma224 + 1) : "0";
                    while (!idxArg224.empty() && (idxArg224.front() == ' ' || idxArg224.front() == '\t')) idxArg224.erase(idxArg224.begin());
                    while (!idxArg224.empty() && (idxArg224.back() == ' ' || idxArg224.back() == '\t')) idxArg224.pop_back();
                    if (callee == "vb6_LoadForm") {
                        std::string parent224 = "NULL";
                        if (isFormModule_ && !knownFormName_.empty()) {
                            auto itF224 = knownFormControlOriginalNames_.find(knownFormName_);
                            parent224 = "(void*)vb6_hwnd_"
                                      + cIdent(itF224 != knownFormControlOriginalNames_.end()
                                                 ? itF224->second : knownFormName_);
                        } else if (isDesignerModule_) {
                            parent224 = "(void*)vb6_UserControl_ContainerHwnd";
                        }
                        callee = "vb6_CtrlArr_Load";
                        argList = "&vb6_arr_" + arrName224 + ", " + idxArg224
                                + ", " + parent224 + ", (void*)GetModuleHandleW(NULL)";
                    } else {
                        callee = "vb6_CtrlArr_Unload";
                        argList = "&vb6_arr_" + arrName224 + ", " + idxArg224;
                    }
                    args.clear();
                }
            }
        }
    }

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

    // <vbeclipse>: VB6 里 InStr/InStrRev/StrComp/Replace/Split/Filter 的 compare 形
    // **缺省值不是常量 0, 而是本模块的 Option Compare**。下面那张表里写 "%cmp%" 的档
    // 统一由这个串填 (旧写法一律硬编码 0 ⇒ Option Compare Text 模块里省略 compare 的
    // 内置函数仍按二进制比, 与同模块的 `=` 判定互相矛盾)。
    const std::string cmpDefault = optionCompareText_ ? "1" : "0";


    // ==== 账 #305 §B153 = 第四十一刀: 「内置出口省略实参时补哪一档」收成这一张表 ====
    // 旧世界里这件事有 28 位主人 (一枚内置函数一段 `if (callee == "vb6_X") { ... argList += ", 1"; }`),
    // 而它的**答案住在 VB6 的形参表 + RTL 的 C 声明里** —— 两份从不汇合。一行 = 一枚 RTL 出口:
    //   slots[i]  = 第 i 档的 VB6 缺省实参, **按 RTL 声明的绝对槽位编号** (0 起)
    //   nullptr   = 这一档 VB6 里必填, 没有缺省答案 —— 缺省链在此**停住**:
    //               少写的实参原样交给 cl, 由 C2198 响亮报错。旧写法里有两枚 nullptr 是
    //               要拼进 std::string 的 (kInputBoxDefaults[0] + `for (i = args.size(); ...)`),
    //               0 参 InputBox 那形当场是 UB —— 停住比补一个假答案诚实。
    //   "%cmp%"   = 唯一的非字面档: VB6 里 InStr/InStrRev/StrComp/Replace/Split/Filter 的 compare
    //               缺省不是常量 0, 而是**本模块的 Option Compare** (旧写法硬编码 0 ⇒ Text 模块里
    //               省略 compare 的内置函数与同模块的 `=` 判定互相矛盾, 这条口径原样搬过来)。
    //   intoArgs  = 缺档同时写进 args —— 这六枚下游有「按 args 重建 argList」的消费点
    //               (conv_numeric / conv_output; Format 那一格的注释就是同一课, 见下), 只补
    //               argList 会被抹掉。
    //   Mid/MidB 的省略 length 补 **-1 哨兵** (RTL 按"取到结尾"处理) 而不是 0 —— 0 是字面
    //               长度 0, 曾让 czUI 的 HexToColor("#243447") 解析成纯黑、整个 demo 颜色全错。
    // 语义侧 calleeParams 对 builtin 标的 isOptional 今天仍被通用 padding 的 !calleeIsBuiltin
    // 挡在外面 (RTL 的 C 过程不收 _has_ 尾巴), 所以补档只发生在这一个地方; 那张表与本表的
    // 逐名对账 (槽数 = RTL 声明的形参数、可选尾巴 = 本表非 nullptr 的档) 由第 55 道哨兵
    // scripts/check_builtin_opt_defaults.ps1 钉住, 新增一形硬编码补档会红。
    struct BuiltinOptDefault {
        const char* callee;
        const char* const* slots;
        size_t nslots;
        bool intoArgs;
    };
    static const char* const kOD_Split[]       = { nullptr, "NULL", "-1", "%cmp%" };
    static const char* const kOD_Join[]        = { nullptr, "NULL" };
    static const char* const kOD_Filter[]      = { nullptr, nullptr, "-1", "%cmp%" };
    static const char* const kOD_Replace[]     = { nullptr, nullptr, nullptr, "1", "-1", "%cmp%" };
    static const char* const kOD_InStrRev[]    = { nullptr, nullptr, "-1", "%cmp%" };
    static const char* const kOD_StrComp[]     = { nullptr, nullptr, "%cmp%" };
    static const char* const kOD_DateDiff[]    = { nullptr, nullptr, nullptr, "1", "1" };
    static const char* const kOD_DatePart[]    = { nullptr, nullptr, "1", "1" };
    static const char* const kOD_Weekday[]     = { nullptr, "1" };
    static const char* const kOD_FormatDT[]    = { nullptr, "0" };
    static const char* const kOD_Round[]       = { nullptr, "0" };
    static const char* const kOD_Shell[]       = { nullptr, "2" };
    static const char* const kOD_AppActivate[] = { nullptr, "0" };
    static const char* const kOD_GetSetting[]  = { nullptr, nullptr, nullptr, "vb6_BSTR_Empty()" };
    static const char* const kOD_InputBox[]    = { nullptr, "vb6_BSTR_Empty()", "vb6_BSTR_Empty()",
                                                   "0", "0", "vb6_BSTR_Empty()", "0" };
    static const char* const kOD_NPer[]        = { nullptr, nullptr, nullptr, "0", "0" };
    static const char* const kOD_IRR[]         = { nullptr, "0.1" };
    static const char* const kOD_Dir[]         = { "NULL", "0" };
    static const char* const kOD_CurDir[]      = { "NULL" };
    static const char* const kOD_LoadPicEx[]   = { "NULL" };
static const char* const kOD_Rnd[]         = { "0" };        // VB6 无参 Rnd = 取下一个数
static const char* const kOD_Randomize[]   = { "0.0" };
// 账 #439 (§B212): 位 3 = **UseParentheses**，默认是 **False(0)**，**不是 -1**。
//   ⚠ 这一格原先和另外两位一样填 "-1"，于是「**省略**」与「**显式写 True**」在 C 侧**同码**
//   （VB6 的 True 恒为 -1）⇒ RTL 收到 -1 分不清是「省略、走默认」还是「要括号」。
//     按 -1 走默认 ⇒ 显式 True 被吃掉（实测 FormatNumber(-1234.5678,2,True,True) 交 -1,234.57，
//     运行期交 (1,234.57)）；按 -1 当真 ⇒ **省略档**全部变括号 = 把绝大多数正常调用一起改坏。
//   位 1/2/4 的默认（NumDigits=-1、IncludeLeadingDigits=True、GroupDigits=True）都还是 -1。
    static const char* const kOD_Mid[]         = { nullptr, nullptr, "-1" };
    static const char* const kOD_MidB[]        = { nullptr, nullptr, "-1" };
    static const char* const kOD_StrConv[]     = { nullptr, nullptr, "0" };
    static const char* const kOD_FormatNum[]   = { nullptr, "-1", "-1", "0", "-1" };
    static const char* const kOD_ClipSet[]     = { nullptr, "0" };
    static const char* const kOD_ErrRaise[]    = { nullptr, "(BSTR)0", "(BSTR)0" };
    // §B149 格③: VB6 的 DeleteSetting 有 1/2/3 参三形, RTL 出口是 3 枚必填 BSTR。
    // 两参 = 删整节 (RTL 在 key 为空时走 RegDeleteKeyW, vb6rtl_registry.c:107), 三参 = 删一个键。
    // 一参那档 (删整个应用) 刻意**不补**: RTL 在 section 为空时直接 return, 补一个空串等于
    // 把「删应用」静默变成「什么都不做」; 语料里那一形 0 处 (见 §B153) —— 让它响亮编不过。
    static const char* const kOD_DelSetting[]  = { nullptr, nullptr, "vb6_BSTR_Empty()" };
    static const BuiltinOptDefault kBuiltinOptDefaults[] = {
        { "vb6_Split",           kOD_Split,       sizeof(kOD_Split) / sizeof(kOD_Split[0]),       false },
        { "vb6_Join",            kOD_Join,        sizeof(kOD_Join) / sizeof(kOD_Join[0]),         false },
        { "vb6_Filter",          kOD_Filter,      sizeof(kOD_Filter) / sizeof(kOD_Filter[0]),     false },
        { "vb6_Replace",         kOD_Replace,     sizeof(kOD_Replace) / sizeof(kOD_Replace[0]),   false },
        { "vb6_InStrRev",        kOD_InStrRev,    sizeof(kOD_InStrRev) / sizeof(kOD_InStrRev[0]), false },
        { "vb6_StrComp",         kOD_StrComp,     sizeof(kOD_StrComp) / sizeof(kOD_StrComp[0]),   false },
        { "vb6_DateDiff",        kOD_DateDiff,    sizeof(kOD_DateDiff) / sizeof(kOD_DateDiff[0]), false },
        { "vb6_DatePart",        kOD_DatePart,    sizeof(kOD_DatePart) / sizeof(kOD_DatePart[0]), false },
        { "vb6_Weekday",         kOD_Weekday,     sizeof(kOD_Weekday) / sizeof(kOD_Weekday[0]),   false },
        { "vb6_FormatDateTime",  kOD_FormatDT,    sizeof(kOD_FormatDT) / sizeof(kOD_FormatDT[0]), false },
        { "vb6_Round",           kOD_Round,       sizeof(kOD_Round) / sizeof(kOD_Round[0]),       false },
        { "vb6_Shell",           kOD_Shell,       sizeof(kOD_Shell) / sizeof(kOD_Shell[0]),       false },
        { "vb6_AppActivate",     kOD_AppActivate, sizeof(kOD_AppActivate) / sizeof(kOD_AppActivate[0]), false },
        { "vb6_GetSetting",      kOD_GetSetting,  sizeof(kOD_GetSetting) / sizeof(kOD_GetSetting[0]),   false },
        { "vb6_InputBox",        kOD_InputBox,    sizeof(kOD_InputBox) / sizeof(kOD_InputBox[0]),       false },
        { "vb6_NPer",            kOD_NPer,        sizeof(kOD_NPer) / sizeof(kOD_NPer[0]),               false },
        { "vb6_IRR",             kOD_IRR,         sizeof(kOD_IRR) / sizeof(kOD_IRR[0]),                 false },
        { "vb6_Dir",             kOD_Dir,         sizeof(kOD_Dir) / sizeof(kOD_Dir[0]),                 false },
        { "vb6_CurDir",          kOD_CurDir,      sizeof(kOD_CurDir) / sizeof(kOD_CurDir[0]),           false },
        { "vb6_LoadPictureEx",   kOD_LoadPicEx,   sizeof(kOD_LoadPicEx) / sizeof(kOD_LoadPicEx[0]),     false },
        { "vb6_Rnd",             kOD_Rnd,         sizeof(kOD_Rnd) / sizeof(kOD_Rnd[0]),                 false },
        { "vb6_Randomize",       kOD_Randomize,   sizeof(kOD_Randomize) / sizeof(kOD_Randomize[0]),     false },
        { "vb6_Mid",             kOD_Mid,         sizeof(kOD_Mid) / sizeof(kOD_Mid[0]),                 true  },
        { "vb6_MidB",            kOD_MidB,        sizeof(kOD_MidB) / sizeof(kOD_MidB[0]),               true  },
        { "vb6_StrConv",         kOD_StrConv,     sizeof(kOD_StrConv) / sizeof(kOD_StrConv[0]),         true  },
        { "vb6_FormatNumber",    kOD_FormatNum,   sizeof(kOD_FormatNum) / sizeof(kOD_FormatNum[0]),     true  },
    // 账 #439 (§B212): 这两家与 FormatNumber **同形**（同一条 5 位 C 签名、同样的可选实参），
    //   改前**只有 FormatNumber 登记了**默认档 ⇒ `FormatCurrency(1234.5678)` / `FormatPercent(...)`
    //   发成 1 个实参进 5 参的 RTL 出口 = **cl C2198 实参太少**，整份文件编不过。
    //   三家共用 kOD_FormatNum（位置 0 是 value，其余 1..4 省略档 = "-1" = 各用自己的默认）。
    { "vb6_FormatCurrency", kOD_FormatNum,   sizeof(kOD_FormatNum) / sizeof(kOD_FormatNum[0]),     true  },
    { "vb6_FormatPercent",  kOD_FormatNum,   sizeof(kOD_FormatNum) / sizeof(kOD_FormatNum[0]),     true  },
        { "vb6_Clipboard_SetData", kOD_ClipSet,   sizeof(kOD_ClipSet) / sizeof(kOD_ClipSet[0]),         true  },
        { "vb6_ErrRaise",        kOD_ErrRaise,    sizeof(kOD_ErrRaise) / sizeof(kOD_ErrRaise[0]),       true  },
        { "vb6_DeleteSetting",   kOD_DelSetting,  sizeof(kOD_DelSetting) / sizeof(kOD_DelSetting[0]),   false },
    };
    for (const BuiltinOptDefault& bod : kBuiltinOptDefaults) {
        if (callee != bod.callee) continue;
        for (size_t i = args.size(); i < bod.nslots; i++) {
            const char* lit = bod.slots[i];
            if (!lit) break;                      // 必填档: 缺省链停住, 让 cl 的 C2198 说话
            const std::string v = std::string(lit) == "%cmp%" ? cmpDefault : lit;
            if (!argList.empty()) argList += ", ";
            argList += v;
            if (bod.intoArgs) args.push_back(v);
        }
    }
    // <vbeclipse>: Format(expression[, format]) 的缺省 format 补在 conv_output 段
    // (那里 args[0] 装箱后会重建 argList, 在这补会被抹掉), 见 cgen_expr_call_conv_output.inc。

    if (callee == "vb6_ErrRaise" && args.size() > 3) {
        // Truncate to 3 (Err.Raise can have up to 5 VB6 args, C RTL handles 3)
        argList = args[0] + ", " + args[1] + ", " + args[2];
    }

    // Fix 092w: StrConv(Byte数组, vbUnicode) → vb6_StrConvFromByteArray
    // 反向转换: Byte() -> String. 第一实参为已知 Byte 数组变量 (vb6_SafeArray1D*)
    // 时, RTL vb6_StrConv 期望 BSTR 会类型不匹配; 改用返回 BSTR 的 helper.
    if (callee == "vb6_StrConv" && !args.empty()) {
        std::string arg0name = args[0];
        if (arg0name.compare(0, 4, "me->") == 0) arg0name = arg0name.substr(4);
        // 账 #305 §B181 (第七十一刀): 长度闸原来是 `> 4`，于是枚长正好 4 的 `(*b)`
        // (单字母名) 剥不出 `b` ⇒ 这一族的规则在 backend 里有 24 位主人、各家闸不同
        // (`> 2` / `> 3` / `> 4`)，同一个问题在不同站点答案不一样。这里收成「至少 `(*x)`」。
        if (arg0name.size() >= 4 && arg0name[0] == '(' && arg0name[1] == '*'
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
        // 账 #305 §B181 (第七十一刀): 原来写 `> 4` ⇒ 枚长正好 4 的 `(*b)` (单字母名) 剥不出
        // `b`，而这一站剥不出来就等于认不出是字节数组 ⇒ 静默把描述符指针交给要 BSTR 的形参。
        // 闸改成「至少 `(*x)`」。同一条剥壳规则在 backend 里有 24 位主人、闸各有三种
        // (`> 2` / `> 3` / `> 4`) ⇒ 收成一处是下一格 (§B181 先手)，本刀先把这一格修对。
        if (a140.size() >= 4 && a140[0] == '(' && a140[1] == '*'
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
    // 账 #305 §B180 (第六十九刀): StrPtr 收**整枚槽**那一形改道到收槽的出口。
    // 上面 Fix 140 已经立了「按实参换 callee 名」这一形状的先例 (Byte() → SafeArrayDataPtr)，
    // 本格只是把同一形状用到 Variant 上 —— 而**问句只有一处主人**
    // (`builtinWantsRawSlotFromVariant`，另一枚读者 = arg_variant 撤回解包；两站判据同源，
    // 免得再长一次 §B176 形② 那种「一位两主、各判各的」)。改道后发码 =
    // `vb6_StrPtrVar(v)`，RTL 那一枚在 `vt` 已是串时**借**槽里的 BSTR (空串仍答 0，
    // 与账 #257 第十三层同一条口径)，不再每调用现造一枚没人回收的临时串。
    if (callee == "vb6_StrPtr" && !node.positional.empty() && node.positional[0]
        && builtinWantsRawSlotFromVariant(callee, node.positional.size(), 0,
                                          *node.positional[0])) {
        callee = "vb6_StrPtrVar";
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
    // P14.1.4: General Optional parameter padding for user-defined functions
    // calleeParams is empty for builtin RTL functions (registered without params), so they're auto-skipped
    // Fix 030b: builtin 即使现在有 calleeParams (用于触发包装), 也不走 padding/IsMissing 路径
    // (RTL C 签名不接受尾叜 _has_ flag, 默认值由 builtin 的特殊 codegen 处理如 UBound 补 dimension=