// builtin_funcs.inc - registerBuiltins 的内置对象与内置函数注册段（函数体片段）
// 由 src/semantics/semantic_analyzer_builtin.cpp 拆出（2026-09-17），在函数体内被
// #include，非独立文件。内容 = 原文件第 504~793 行，逐行未改（含局部 lambda addObj /
// addBuiltinFunc / addBuiltinFuncWithParams 与局部常量 kVariantArray）。


    // 内置对象: Debug, Err, Screen, App, Printer
    auto addObj = [&](const char* name) {
        auto sym = std::make_unique<Symbol>(SymbolKind::Variable, name, Vb6Type::Object,
            SourceLocation{}, AccessLevel::Public);
        sym->isBuiltin = true;
        symTab_.define(std::move(sym));
    };
    addObj("Debug");
    addObj("Err");
    addObj("Screen");
    addObj("App");
    addObj("Printer");
    addObj("Forms");
    addObj("Clipboard");
    addObj("Console");  // Fix 161: Console 对象 (对齐 twinBASIC) —— 控制台 I/O

    // 内置函数 (声明为 Function 符号, 避免未声明错误)
    auto addBuiltinFunc = [&](const char* name, Vb6Type retType) {
        auto sym = std::make_unique<Symbol>(SymbolKind::Function, name, retType,
            SourceLocation{}, AccessLevel::Public);
        sym->isBuiltin = true;
        symTab_.define(std::move(sym));
    };
    // Fix 030b: 带参数签名的内置函数注册 — 填充 sym->params, 让 IndexOrCallExpr 的
    // calleeParams 查找命中, 从而触发 Fix 024 P2 (ByVal Variant 正向包装) 和
    // Fix 029 (ByVal 具体类型反向提取). params: vector of (name, type, isByVal, isOptional).
    auto addBuiltinFuncWithParams = [&](const char* name, Vb6Type retType,
                                        std::initializer_list<std::tuple<const char*, Vb6Type, bool, bool>> params) {
        auto sym = std::make_unique<Symbol>(SymbolKind::Function, name, retType,
            SourceLocation{}, AccessLevel::Public);
        sym->isBuiltin = true;
        for (auto& p : params) {
            ParameterInfo pi;
            pi.name = std::get<0>(p);
            pi.type = std::get<1>(p);
            pi.isByVal = std::get<2>(p);
            pi.isOptional = std::get<3>(p);
            sym->params.push_back(std::move(pi));
        }
        symTab_.define(std::move(sym));
    };
    // Fix 034: Variant|Array 复合类型 (Variant 含数组), 供多个内置函数的数组参数
    // 声明使用 — 反向提取时 paramIsArray 分支会触发 vb6_VariantToSafeArray1D.
    const Vb6Type kVariantArray = static_cast<Vb6Type>(
        static_cast<uint16_t>(Vb6Type::Variant) | static_cast<uint16_t>(Vb6Type::Array));
    // Fix 034: 字符串/转换/数值/类型检查/数学/文件/日期等内置函数补全参数签名 —
    // 让 IndexOrCallExpr 的 calleeParams 查找命中, 触发 Fix 024 P2 (ByVal Variant 正向包装) 和
    // Fix 029 (ByVal 具体类型反向提取), 消除大量 C2440 (函数实参 → VARIANT/BSTR/double 等).
    // RTL C 签名参考 src/rtl/core/vb6rtl/vb6rtl.h. VB6 Optional 参数标记 isOptional=true (语义保留),
    // 但 RTL C 函数不接受 Optional padding 和 IsMissing _has_ 尾叜 — cgen 用硬编码补默认值
    // (见 cgen_expr.cpp line 3140+ 和 line 3263+).
    // 字符串函数 (RTL 签名: vb6_Len/Left/Right/Mid/InStr/InStrRev 等均接受 BSTR)
    addBuiltinFuncWithParams("Len", Vb6Type::Long,
        {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("Left", Vb6Type::String,
        {{"s", Vb6Type::String, true, false}, {"n", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("Right", Vb6Type::String,
        {{"s", Vb6Type::String, true, false}, {"n", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("Mid", Vb6Type::String,
        {{"s", Vb6Type::String, true, false}, {"start", Vb6Type::Long, true, false},
         {"len", Vb6Type::Long, true, true}});
    // InStr: VB6 双形态 InStr(s1, s2) 和 InStr(start, s1, s2). cgen 用 argList="1, "+argList
    // 把 2-arg 形状调整为 3-arg (start=1). 为避免 2-arg 形态下 arg[0]=s1 被 calleeParams[0]
    // (start:Long) 错误包装 (Variant→vb6_VariantToLong 而非 vb6_VariantToString),
    // 不声明 calleeParams — 让 cgen 硬编码重排实参位置后直接调用.
    addBuiltinFunc("InStr", Vb6Type::Long);
    addBuiltinFuncWithParams("InStrRev", Vb6Type::Long,
        {{"string1", Vb6Type::String, true, false}, {"string2", Vb6Type::String, true, false},
         {"start", Vb6Type::Long, true, true}, {"compare", Vb6Type::Long, true, true}});
    addBuiltinFuncWithParams("LTrim", Vb6Type::String, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("RTrim", Vb6Type::String, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("Trim", Vb6Type::String, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("LCase", Vb6Type::String, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("UCase", Vb6Type::String, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("Replace", Vb6Type::String,
        {{"expr", Vb6Type::String, true, false}, {"find", Vb6Type::String, true, false},
         {"rep", Vb6Type::String, true, false}, {"start", Vb6Type::Long, true, true},
         {"count", Vb6Type::Long, true, true}, {"compare", Vb6Type::Long, true, true}});
    addBuiltinFuncWithParams("Split", Vb6Type::Variant,
        {{"expr", Vb6Type::String, true, false}, {"delimiter", Vb6Type::String, true, true},
         {"limit", Vb6Type::Long, true, true}, {"compare", Vb6Type::Long, true, true}});
    addBuiltinFuncWithParams("Join", Vb6Type::String,
        {{"arr", kVariantArray, true, false}, {"delimiter", Vb6Type::String, true, true}});
    addBuiltinFuncWithParams("StrComp", Vb6Type::Long,
        {{"s1", Vb6Type::String, true, false}, {"s2", Vb6Type::String, true, false},
         {"compare", Vb6Type::Long, true, true}});
    addBuiltinFuncWithParams("StrReverse", Vb6Type::String, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("Space", Vb6Type::String, {{"n", Vb6Type::Long, true, false}});
    // String(n, charCode) — charCode 在 VB6 可为 String 或 Integer; RTL 取 int32_t.
    addBuiltinFuncWithParams("String", Vb6Type::String,
        {{"n", Vb6Type::Long, true, false}, {"charCode", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("Asc", Vb6Type::Long, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("Chr", Vb6Type::String, {{"code", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("Val", Vb6Type::Double, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("Str", Vb6Type::String, {{"n", Vb6Type::Long, true, false}});
    // Format: 保留 addBuiltinFunc (无 calleeParams) — cgen_expr.cpp:3487 会按 inferExprType
    // 把 args[0] 包装成 vb6_VariantLong/Double/String/Int. 若此处加 ByVal Variant calleeParams,
    // Fix 024 P2 会先用 vb6_VariantFromValue 包装(返回 VARIANT), 然后 Format 特殊包装
    // vb6_VariantLong(VARIANT) 传 int32_t 参数 → C2440. 双重包装冲突, 故不注册参数.
    addBuiltinFunc("Format", Vb6Type::String);
    // 类型转换函数 — cgen 特殊处理决定参数策略.
    // CStr/CInt/CLng/CDbl: cgen_expr.cpp:3383/3428 有 callee-rewrite 特殊分支
    //   - CStr: 根据 arg 类型改写 callee 为 vb6_CStrLong/Dbl/Bool/Date (具体类型参数)
    //   - CInt/CLng/CDbl: 已知 Variant 变量改写 callee 为 vb6_CIntV/CLngV/CDblV (VARIANT 参数)
    // 若此处加 calleeParams, Fix 024 P2/Fix 029 会先 wrapping, 与 callee-rewrite 冲突 → C2440.
    // 故保持 addBuiltinFunc (无 params), 让特殊分支独立工作.
    addBuiltinFunc("CStr", Vb6Type::String);
    addBuiltinFunc("CInt", Vb6Type::Integer);
    addBuiltinFunc("CLng", Vb6Type::Long);
    addBuiltinFuncWithParams("CSng", Vb6Type::Single, {{"v", Vb6Type::Double, true, false}});
    addBuiltinFunc("CDbl", Vb6Type::Double);
    addBuiltinFuncWithParams("CBool", Vb6Type::Boolean, {{"v", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("CDate", Vb6Type::Date, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFuncWithParams("CByte", Vb6Type::Byte, {{"v", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("CCur", Vb6Type::Currency, {{"v", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("CDec", Vb6Type::Variant, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFunc("CVar", Vb6Type::Variant);
    addBuiltinFuncWithParams("Hex", Vb6Type::String, {{"n", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("Oct", Vb6Type::String, {{"n", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("CVErr", Vb6Type::Variant, {{"errorNumber", Vb6Type::Long, true, false}});
    // 数值函数 — RTL C 签名: 所有数值函数接受 double.
    addBuiltinFuncWithParams("Abs", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Int", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Fix", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    // Fix 029: Sgn 返回 Long 但接受 double. Variant 实参 → vb6_VariantToDouble 提取.
    addBuiltinFuncWithParams("Sgn", Vb6Type::Long, {{"x", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Sqr", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Round", Vb6Type::Double,
        {{"x", Vb6Type::Double, true, false}, {"decimals", Vb6Type::Long, true, true}});
    // Rnd/Randomize: Optional 参数. 不补 calleeParams — 现有 cgen 不补默认值会 C2198,
    // 但补了 calleeParams 又会被 !calleeIsBuiltin guard 跳过 padding. Fix 034 额外在
    // cgen_expr.cpp 中为这两个加硬编码 padding (无 calleeParams, 走老路径).
    addBuiltinFunc("Rnd", Vb6Type::Single);
    // 转换与类型检查 — RTL: IsXxx 接受 vb6_VARIANT by-value.
    addBuiltinFuncWithParams("IsNumeric", Vb6Type::Boolean, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFuncWithParams("IsDate", Vb6Type::Boolean, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFuncWithParams("IsEmpty", Vb6Type::Boolean, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFuncWithParams("IsNull", Vb6Type::Boolean, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFuncWithParams("IsObject", Vb6Type::Boolean, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFuncWithParams("IsArray", Vb6Type::Boolean, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFunc("IsNothing", Vb6Type::Boolean);
    addBuiltinFuncWithParams("IsError", Vb6Type::Boolean, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFuncWithParams("TypeName", Vb6Type::String, {{"v", Vb6Type::Variant, true, false}});
    addBuiltinFuncWithParams("VarType", Vb6Type::Long, {{"v", Vb6Type::Variant, true, false}});
    // 数组
    // Fix 030b: UBound/LBound 注册带参数签名 — 让 Fix 029 反向提取能在 Variant 实参上
    // 触发 vb6_VariantToSafeArray1D, 修正 C2440 (vb6_VARIANT → vb6_SafeArray1D*) 子类.
    // RTL: int32_t vb6_UBound(vb6_SafeArray1D* safeArray, int32_t dimension);
    // VB6: UBound(arrayname[, dimension]) — arrayname 接受 Variant 含数组或类型化数组.
    // 第 1 参数用 Variant|Array (不纯 Variant, 让反向提取 paramIsArray 分支命中).
    addBuiltinFuncWithParams("UBound", Vb6Type::Long,
        {{"array", kVariantArray, true, false}, {"dimension", Vb6Type::Long, true, true}});
    addBuiltinFuncWithParams("LBound", Vb6Type::Long,
        {{"array", kVariantArray, true, false}, {"dimension", Vb6Type::Long, true, true}});
    addBuiltinFunc("Array", Vb6Type::Variant);
    // 数学 — RTL: sin/cos/tan/atan/exp/log 均接受 double.
    addBuiltinFuncWithParams("Sin", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Cos", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Tan", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Atn", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Exp", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Log", Vb6Type::Double, {{"x", Vb6Type::Double, true, false}});
    // 文件 — RTL: LOF/EOF/Loc/Seek 接受 int32_t filenumber; GetAttr/FileLen/FileDateTime/Environ 接受 BSTR.
    addBuiltinFunc("FreeFile", Vb6Type::Long);
    addBuiltinFuncWithParams("LOF", Vb6Type::Long, {{"filenumber", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("EOF", Vb6Type::Boolean, {{"filenumber", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("Loc", Vb6Type::Long, {{"filenumber", Vb6Type::Long, true, false}});
    // GetAttr 在 vbcrtl.h 中声明为接受 BSTR pathname.
    addBuiltinFuncWithParams("GetAttr", Vb6Type::Long, {{"pathname", Vb6Type::String, true, false}});
    addBuiltinFunc("SetAttr", Vb6Type::Void);
    addBuiltinFuncWithParams("Seek", Vb6Type::Long, {{"filenumber", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("FileLen", Vb6Type::Long, {{"pathname", Vb6Type::String, true, false}});
    // FileAttr(filenumber, attribute) — 2 个 ByRefLong 参数
    addBuiltinFuncWithParams("FileAttr", Vb6Type::Long,
        {{"filenumber", Vb6Type::Long, true, false}, {"attribute", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("FileDateTime", Vb6Type::Date, {{"pathname", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("Dir", Vb6Type::String,
        {{"pathname", Vb6Type::String, true, false}, {"attributes", Vb6Type::Long, true, true}});
    addBuiltinFuncWithParams("CurDir", Vb6Type::String, {{"drive", Vb6Type::String, true, true}});
    addBuiltinFuncWithParams("Shell", Vb6Type::Long,
        {{"pathname", Vb6Type::String, true, false}, {"windowstyle", Vb6Type::Long, true, true}});
    addBuiltinFuncWithParams("Environ", Vb6Type::String, {{"envstring", Vb6Type::String, true, false}});
    addBuiltinFunc("Command", Vb6Type::String);
    // 日期时间 — RTL: Year/Month/Day/Hour/Minute/Second 接受 double (Date),
    // DateAdd(interval, number, date), DateDiff(interval, d1, d2, ...), DateSerial(y,m,d),
    // DateValue(BSTR), TimeSerial(h,m,s), TimeValue(BSTR), Weekday(date[, firstDayOfWeek]).
    addBuiltinFunc("Now", Vb6Type::Date);
    addBuiltinFunc("Date", Vb6Type::Date);
    addBuiltinFunc("Time", Vb6Type::Date);
    addBuiltinFuncWithParams("DateAdd", Vb6Type::Date,
        {{"interval", Vb6Type::String, true, false}, {"number", Vb6Type::Double, true, false},
         {"date", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("DateDiff", Vb6Type::Long,
        {{"interval", Vb6Type::String, true, false}, {"date1", Vb6Type::Double, true, false},
         {"date2", Vb6Type::Double, true, false}, {"firstDayOfWeek", Vb6Type::Long, true, true},
         {"firstWeekOfYear", Vb6Type::Long, true, true}});
    addBuiltinFuncWithParams("DatePart", Vb6Type::Long,
        {{"interval", Vb6Type::String, true, false}, {"date", Vb6Type::Double, true, false},
         {"firstDayOfWeek", Vb6Type::Long, true, true}, {"firstWeekOfYear", Vb6Type::Long, true, true}});
    addBuiltinFuncWithParams("DateSerial", Vb6Type::Double,
        {{"year", Vb6Type::Long, true, false}, {"month", Vb6Type::Long, true, false},
         {"day", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("DateValue", Vb6Type::Date, {{"dateStr", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("TimeSerial", Vb6Type::Date,
        {{"hour", Vb6Type::Long, true, false}, {"minute", Vb6Type::Long, true, false},
         {"second", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("TimeValue", Vb6Type::Date, {{"timeStr", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("Year", Vb6Type::Long, {{"date", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Month", Vb6Type::Long, {{"date", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Day", Vb6Type::Long, {{"date", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Hour", Vb6Type::Long, {{"time", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Minute", Vb6Type::Long, {{"time", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Second", Vb6Type::Long, {{"time", Vb6Type::Double, true, false}});
    addBuiltinFuncWithParams("Weekday", Vb6Type::Long,
        {{"date", Vb6Type::Double, true, false}, {"firstDayOfWeek", Vb6Type::Long, true, true}});

    // 交互
    addBuiltinFunc("MsgBox", Vb6Type::Long);
    addBuiltinFunc("InputBox", Vb6Type::String);
    // RTL: vb6_RGB(int32_t r, int32_t g, int32_t b); vb6_QBColor(int32_t n).
    addBuiltinFuncWithParams("RGB", Vb6Type::Long,
        {{"r", Vb6Type::Long, true, false}, {"g", Vb6Type::Long, true, false},
         {"b", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("QBColor", Vb6Type::Long, {{"n", Vb6Type::Long, true, false}});
    // 杂项
    addBuiltinFunc("DoEvents", Vb6Type::Long);
    addBuiltinFunc("Erl", Vb6Type::Long);
    // Tab/Spc 用于 Print 语句, RTL 接受 int32_t.
    addBuiltinFuncWithParams("Tab", Vb6Type::String, {{"column", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("Spc", Vb6Type::String, {{"count", Vb6Type::Long, true, false}});
    // RTL: vb6_CreateObject(const wchar_t* progId); vb6_GetObject(pathName, progId).
    addBuiltinFuncWithParams("CreateObject", Vb6Type::Object,
        {{"progId", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("GetObject", Vb6Type::Object,
        {{"pathName", Vb6Type::String, true, true}, {"progId", Vb6Type::String, true, true}});
    // LoadPicture(pathname) 返回 void* (Object). RTL: vb6_LoadPictureEx(BSTR).
    addBuiltinFuncWithParams("LoadPicture", Vb6Type::Object, {{"pathname", Vb6Type::String, true, false}});
    addBuiltinFunc("SavePicture", Vb6Type::Void);
    addBuiltinFunc("SaveSetting", Vb6Type::Void);
    addBuiltinFunc("GetSetting", Vb6Type::String);
    addBuiltinFunc("DeleteSetting", Vb6Type::Void);
    addBuiltinFunc("GetAllSettings", Vb6Type::Variant);
    addBuiltinFunc("Load", Vb6Type::Void);
    addBuiltinFunc("Unload", Vb6Type::Void);
    addBuiltinFunc("SendKeys", Vb6Type::Void);
    addBuiltinFunc("AppActivate", Vb6Type::Void);
    addBuiltinFunc("IsMissing", Vb6Type::Boolean);
    addBuiltinFunc("IIf", Vb6Type::Variant);
    addBuiltinFunc("Beep", Vb6Type::Void);
// P14.3.5: CallByName(obj, procName$, callType, [args...])
addBuiltinFunc("CallByName", Vb6Type::Variant);
    // P18-D: 兼容性填平新增内置函数 — RTL: AscW/AscB 接受 BSTR, ChrW/ChrB 接受 int32_t,
    // StrConv(text, conversion, localeID) 全 3 参.
    addBuiltinFuncWithParams("AscW", Vb6Type::Long, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("ChrW", Vb6Type::String, {{"code", Vb6Type::Long, true, false}});
    addBuiltinFuncWithParams("AscB", Vb6Type::Long, {{"s", Vb6Type::String, true, false}});
    addBuiltinFuncWithParams("ChrB", Vb6Type::String, {{"code", Vb6Type::Long, true, false}});
    addBuiltinFunc("Timer", Vb6Type::Single);
    addBuiltinFuncWithParams("StrConv", Vb6Type::String,
        {{"text", Vb6Type::String, true, false}, {"conversion", Vb6Type::Long, true, false},
         {"localeID", Vb6Type::Long, true, true}});
    // <vbeclipse>: Filter 此前用 addBuiltinFunc (零参数签名) 注册 ⇒ 数组槽拿不到
    // Fix 029 的反向提取, `Filter(v, "x")` 里 v 是 Variant 时直接发 vb6_VARIANT 给
    // SafeArray1D* 形参 → cl C2440 (实测), 而 Variant 装数组正是 Filter 的常态。
    // 与 Join/UBound/LBound 同形: 首参标 Variant|Array, 尾两参 VB6 Optional。
    addBuiltinFuncWithParams("Filter", Vb6Type::Variant,
        {{"sourcearray", kVariantArray, true, false}, {"match", Vb6Type::String, true, false},
         {"include", Vb6Type::Long, true, true}, {"compare", Vb6Type::Long, true, true}});
    addBuiltinFunc("VarPtr", Vb6Type::Long);
    addBuiltinFunc("StrPtr", Vb6Type::Long);
    addBuiltinFunc("ObjPtr", Vb6Type::Long);
    addBuiltinFunc("LSet", Vb6Type::String);
    addBuiltinFunc("RSet", Vb6Type::String);
    addBuiltinFunc("WeekdayName", Vb6Type::String);
    addBuiltinFunc("MonthName", Vb6Type::String);
    addBuiltinFunc("FormatCurrency", Vb6Type::String);
    addBuiltinFunc("FormatNumber", Vb6Type::String);
    addBuiltinFunc("FormatPercent", Vb6Type::String);
    addBuiltinFunc("FormatDateTime", Vb6Type::String);
    // P18-E: Financial / Choose / Switch / Lock / Reset / Partition
    addBuiltinFunc("Choose", Vb6Type::Variant);
    addBuiltinFunc("Switch", Vb6Type::Variant);
    addBuiltinFunc("SLN", Vb6Type::Double);
    addBuiltinFunc("SYD", Vb6Type::Double);
    addBuiltinFunc("DDB", Vb6Type::Double);
    addBuiltinFunc("FV", Vb6Type::Double);
    addBuiltinFunc("PV", Vb6Type::Double);
    addBuiltinFunc("Pmt", Vb6Type::Double);
    addBuiltinFunc("IPmt", Vb6Type::Double);
    addBuiltinFunc("PPmt", Vb6Type::Double);
    addBuiltinFunc("RATE", Vb6Type::Double);
    addBuiltinFunc("NPer", Vb6Type::Double);
    addBuiltinFunc("IRR", Vb6Type::Double);
    addBuiltinFunc("MIRR", Vb6Type::Double);
    addBuiltinFunc("NPV", Vb6Type::Double);
    addBuiltinFunc("Partition", Vb6Type::String);
