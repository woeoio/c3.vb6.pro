// cgen_expr_ident_builtin_table.inc - 内置函数表 (VB6 名 -> RTL 函数名) + 派生的键集合
//
// 原定义在 cgen_expr_ident_builtin.inc 的 visit(IdentifierExpr&) **函数体内** (static local),
// 别的 visit() 片段 (IndexOrCallExpr 的 Pattern G 等) 够不到。Fix <VBFlexGridDemo> 把它提到
// 命名空间作用域: 两个 visit() 所在 TU 各 include 一次 (内部链接, 各持一份副本),
// 键集合 kBuiltinFuncNames 供"内置函数优先于同名类成员属性"的判定复用,
// 免得再抄第二份名单走样。表内容逐行未改, 仅缩进 4 -> 0。

static const std::unordered_map<std::string, std::string> builtinFuncs = {
    {"len",      "vb6_Len"},
    {"lenb",     "vb6_LenB"},  // Fix 048: LenB built-in
    {"msgbox",   "vb6_MsgBox"},
    {"unload",   "vb6_UnloadForm"},
    {"inputbox", "vb6_InputBox"},
    {"input",    "vb6_InputString"},  // Task #40: parseIdentifierOrCall 剥 $ 后
                                      // Input$ 到达此处的键名是 input (双键兜底)
    {"input$",  "vb6_InputString"},  // P15.4: Input function
    {"format",   "vb6_Format"},
    {"cstr",     "vb6_CStr"},
    {"cint",     "vb6_CInt"},
    {"clng",     "vb6_CLng"},
    {"cdbl",     "vb6_CDbl"},
    {"csng",     "vb6_CSng"},
    {"cbool",    "vb6_CBool"},
    {"cdate",    "vb6_CDate"},
    {"cbyt",     "vb6_CByte"},
    {"instr",    "vb6_InStr"},
    {"instrb",   "vb6_InStrB"},
    {"left",     "vb6_Left"},
    {"right",    "vb6_Right"},
    {"mid",      "vb6_Mid"},
    {"trim",     "vb6_Trim"},
    {"ltrim",    "vb6_LTrim"},
    {"rtrim",    "vb6_RTrim"},
    {"ucase",    "vb6_UCase"},
    {"lcase",    "vb6_LCase"},
    {"chr",      "vb6_Chr"},
    {"asc",      "vb6_Asc"},
    {"abs",      "vb6_Abs"},
    {"int",      "vb6_Int"},
    {"fix",      "vb6_Fix"},
    {"sgn",      "vb6_Sgn"},
    {"sqr",      "vb6_Sqr"},
    {"ubound",   "vb6_UBound"},
    {"lbound",   "vb6_LBound"},
    {"isarray",  "vb6_IsArray"},
    {"isnumeric","vb6_IsNumeric"},
    {"isnothing","vb6_IsNothing"},
    {"typeof",   "vb6_TypeOf"},
    {"typename", "vb6_TypeName"},
    {"createobject","vb6_CreateObject"},
    {"getobject","vb6_GetObject"},
    {"loadpicture","vb6_LoadPictureEx"},
    // 字符串函数 (P4新增)
    {"replace",  "vb6_Replace"},
    {"space",    "vb6_Space"},
    {"string",   "vb6_String"},   // String$函数
    {"strcomp",  "vb6_StrComp"},
    {"strreverse","vb6_StrReverse"},
    {"instrrev", "vb6_InStrRev"},
    {"val",      "vb6_Val"},
    {"str",      "vb6_Str"},
    // 数学函数 (P4新增)
    {"sin",      "vb6_Sin"},
    {"cos",      "vb6_Cos"},
    {"tan",      "vb6_Tan"},
    {"atn",      "vb6_Atn"},
    {"log",      "vb6_Log"},
    {"exp",      "vb6_Exp"},
    {"round",    "vb6_Round"},
    {"rnd",      "vb6_Rnd"},
    {"randomize","vb6_Randomize"},
    // 日期时间 (P4新增)
    {"now",      "vb6_Now"},
    {"date",     "vb6_Date"},
    {"time",     "vb6_Time"},
    {"year",     "vb6_Year"},
    {"month",    "vb6_Month"},
    {"day",      "vb6_Day"},
    {"hour",     "vb6_Hour"},
    {"minute",   "vb6_Minute"},
    {"second",   "vb6_Second"},
    // 日期函数 P14.2.4
    {"dateadd",  "vb6_DateAdd"},
    {"datediff", "vb6_DateDiff"},
    {"datepart", "vb6_DatePart"},
    {"dateserial", "vb6_DateSerial"},
    // P21-B: 日期/数组函数
    {"weekday",  "vb6_Weekday"},
    {"datevalue","vb6_DateValue"},
    {"timeserial","vb6_TimeSerial"},
    {"timevalue","vb6_TimeValue"},
    // 类型转换 (P4新增)
    {"cbyte",    "vb6_CByte"},
    {"cvar",     "vb6_CVar"},
    {"hex",      "vb6_Hex"},
    {"oct",      "vb6_Oct"},
    // 类型检查 (P4新增)
    {"isnull",   "vb6_IsNull"},
    {"isempty",  "vb6_IsEmpty"},
    {"isobject", "vb6_IsObject"},
    {"isdate",   "vb6_IsDate"},
    {"iserror",  "vb6_IsError"},
    {"vartype",  "vb6_VarType"},
    // 文件I/O (P4新增)
    {"freefile", "vb6_FreeFile"},
    {"eof",      "vb6_EOF"},
    {"lof",      "vb6_LOF"},
    {"loc",      "vb6_Loc"},
    {"kill",     "vb6_Kill"},
    // 系统函数 (P14.2.2新增)
    
    {"shell",    "vb6_Shell"},
    {"environ",  "vb6_Environ"},
    {"dir",      "vb6_Dir"},
    {"curdir",   "vb6_CurDir"},
    {"command",  "vb6_Command"},
    {"command$", "vb6_Command"},  // P14.2.2: Command$ alias
    {"split",    "vb6_Split"},
    {"join",     "vb6_Join"},

    // ParamArray (P14.1.5)

    {"ismissing", "vb6_IsMissing"},
    // P14.3.5: CallByName
    {"callbyname", "vb6_CallByName"},
    // P18-A: 兼容性填平新增内置函数
    {"ccur",       "vb6_CCur"},
    {"cdec",       "vb6_CDec"},
    {"rgb",        "vb6_RGB"},
    {"qbcolor",    "vb6_QBColor"},
    {"filedatetime", "vb6_FileDateTime"},
    {"filelen",    "vb6_FileLen"},
    {"sendkeys",   "vb6_SendKeys"},
    {"appactivate", "vb6_AppActivate"},
    // P18-D: 兼容性填平新增内置函数
    {"ascw",       "vb6_AscW"},
    {"chrw",       "vb6_ChrW"},
    {"ascb",       "vb6_AscB"},
    {"chrb",       "vb6_ChrB"},
    {"timer",      "vb6_Timer"},
    {"strconv",    "vb6_StrConv"},
    {"filter",     "vb6_Filter"},
    {"strptr",     "vb6_StrPtr"},
    {"objptr",     "vb6_ObjPtr"},
    {"lset",       "vb6_LSet"},
    {"rset",       "vb6_RSet"},
    {"weekdayname", "vb6_WeekdayName"},
    {"monthname",  "vb6_MonthName"},
    {"formatcurrency", "vb6_FormatCurrency"},
    {"formatnumber",  "vb6_FormatNumber"},
    {"formatpercent",  "vb6_FormatPercent"},
    // P18-E: 
    {"sln",       "vb6_SLN"},
    {"syd",       "vb6_SYD"},
    {"ddb",       "vb6_DDB"},
    {"fv",        "vb6_FV"},
    {"pv",        "vb6_PV"},
    {"pmt",       "vb6_Pmt"},
    {"ipmt",      "vb6_IPmt"},
    {"ppmt",      "vb6_PPmt"},
    {"rate",      "vb6_RATE"},
    {"npv",       "vb6_NPV"},
    {"partition", "vb6_Partition"},
    // P21-C: new functions
    {"cverr",          "vb6_CVErr"},
    {"formatdatetime", "vb6_FormatDateTime"},
    {"getattr",        "vb6_GetAttr"},
    {"setattr",        "vb6_SetAttr"},
    {"seek",           "vb6_SeekFunc"},
    {"doevents",       "vb6_DoEvents"},
    {"savepicture",    "vb6_SavePicture"},
    {"load",           "vb6_LoadForm"},
    {"nper",           "vb6_NPer"},
    {"fileattr",       "vb6_FileAttr"},
    {"erl",            "vb6_Erl"},
    {"tab",            "vb6_Tab"},
    {"spc",            "vb6_Spc"},
    {"irr",            "vb6_IRR"},
    {"mirr",           "vb6_MIRR"},
    // Fix 048: LoadResData
    {"loadresdata", "vb6_LoadResData"},
    // Fix <vbeclipse>: LoadResPicture / LoadResString (VbEclipse modResources)
    {"loadrespicture", "vb6_LoadResPicture"},
    {"loadresstring", "vb6_LoadResString"},
    // P20-37: Registry functions
    {"savesetting",    "vb6_SaveSetting"},
    {"getsetting",     "vb6_GetSetting"},
    {"deletesetting",  "vb6_DeleteSetting"},
    {"getallsettings", "vb6_GetAllSettings"},

};

// 键集合: IndexOrCallExpr 的 Pattern G (类模块隐式 Me.field) 等处判断
// "这个名字是内置函数" 时复用, 与上面同一份表派生, 不会走样。
static const std::unordered_set<std::string> kBuiltinFuncNames = [] {
    std::unordered_set<std::string> s;
    for (const auto& kv : builtinFuncs) s.insert(kv.first);
    return s;
}();
