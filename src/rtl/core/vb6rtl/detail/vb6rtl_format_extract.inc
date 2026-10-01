// vb6rtl_format_extract.inc - vb6_Format 函数体片段：数值提取与无格式串时的默认转换
// 由 src/rtl/core/vb6rtl/vb6rtl_format.c 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 27~81 行，逐行未改。

    // Full VB6 Format implementation:
    // - Named formats: General Number, Currency, Fixed, Standard, Percent, Scientific,
    //   Yes/No, True/False, On/Off, Long Date, Short Date, Long Time, Short Time
    // - User-defined numeric formats: 0 # . , % E- E+ e- e+ \ "xxx" ;
    // - User-defined string formats: < > @ & !
    // - Multi-section: positive;negative;zero;null
    // - Date/time formats handled by FormatDateTime path

    // === Step 1: Extract numeric value ===
    double numVal = 0.0;
    int isNumeric = 0;
    int isString = 0;
    BSTR strVal = NULL;
    int isNull = 0;

    switch (expr.vt) {
        case vb6_vtEmpty:
            return vb6_BSTR_FromStr(L"");
        case vb6_vtNull:
            isNull = 1;
            break;
        case vb6_vtInteger:
            numVal = (double)expr.iVal; isNumeric = 1; break;
        case vb6_vtLong:
            numVal = (double)expr.lVal; isNumeric = 1; break;
        case vb6_vtDouble:
            numVal = expr.dblVal; isNumeric = 1; break;
        case vb6_vtSingle:
            numVal = (double)expr.fltVal; isNumeric = 1; break;
        // Task #44: VT_I8 = LongLong / x64 的 LongPtr(intptr_t)。此前缺此分支 →
        // 落 default 返回**空串** (Fix 084m 注释已记)。包装侧 (vb6rtl_variant.h 的
        // _Generic long long:) 现已统一产出 VT_I8, 这里必须接住, 否则 x64 下
        // CStr(LongPtr) 仍是空串 (实测 PtrProbe 打 "Q=" 而非 "Q=4294967296")。
        // 无格式串时走 vb6_CStrLongLong 的 %lld 精确十进制 —— 转 double 会在
        // >2^53 处丢精度; 带格式串才进通用数值分支 (此时用户已要求数值语义)。
        case (vb6_vartype)VT_I8:
            if (!fmt) return vb6_CStrLongLong(expr.llVal);
            numVal = (double)expr.llVal; isNumeric = 1; break;
        case vb6_vtByte:
            numVal = (double)expr.bVal; isNumeric = 1; break;
        case vb6_vtCurrency:
            numVal = expr.cyVal / 10000.0; isNumeric = 1; break;
        case vb6_vtDate: {
            // Fix 118: 带日期/时间格式串时走日期格式化 (此前落进数值分支 →
            // Format(DateSerial(2020,1,1),"mmm") 输出 "43831")
            // Fix 162b-extlist: 命名日期/时间格式 ("short date"/"medium date"/
            // "long time"/...) 先于 IsDateFormat 判别 —— IsDateFormat 对命名
            // 格式返回 0 (排除语义), 原先会漏到这里落数值分支输出 "45331"。
            if (fmt && vb6_BSTR_Len(fmt) > 0 && vb6_fmtNamedDateTimeKind(fmt))
                return vb6_fmtNamedDateTime(expr.dblVal, fmt);
            if (fmt && vb6_BSTR_Len(fmt) > 0 && vb6_fmtIsDateFormat(fmt))
                return vb6_fmtDateSerial(expr.dblVal, fmt);
            if (!fmt) {
                SYSTEMTIME st;
                if (VariantTimeToSystemTime(expr.dblVal, &st)) {
                    wchar_t dateBuf[64], timeBuf[64];
                    GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, NULL, dateBuf, 64);
                    GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, NULL, timeBuf, 64);
                    wchar_t fullBuf[128];
                    swprintf(fullBuf, 128, L"%s %s", dateBuf, timeBuf);
                    return vb6_BSTR_FromStr(fullBuf);
                }
            }
            numVal = expr.dblVal; isNumeric = 1; break;
        }
        case vb6_vtBSTR:
            strVal = expr.bstrVal; isString = 1; break;
        case vb6_vtBoolean:
            return vb6_BSTR_FromStr(expr.boolVal ? L"True" : L"False");
        default:
            return vb6_BSTR_FromStr(L"");
    }

