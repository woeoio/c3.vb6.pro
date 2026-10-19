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
        /* ai/032 §6.6: 无符号两档 —— 形状照上面 VT_I8 先例: 无格式串走精确的
         * 无符号十进制出口 (转 double 会在 >2^53 处丢精度), 带格式串进通用数值分支。 */
        case (vb6_vartype)VT_UI4:
            if (!fmt) return vb6_CStrULong(expr.ulVal);
            numVal = (double)expr.ulVal; isNumeric = 1; break;
        case (vb6_vartype)VT_UI8:
            if (!fmt) return vb6_CStrULongLong(expr.ullVal);
            numVal = (double)expr.ullVal; isNumeric = 1; break;
        /* ai/032 §6.9: SByte / UInteger 两档进这一族。
         * 以前这两档被 _Generic 的 `char:` / `wchar_t:` 档装成 VT_I4, 而**这一张表**
         * 里没有 VT_I4 之外的小整型分支 —— 落进 default ⇒ CStr 交**空串**(实测
         * tests/tu_holder `UF-ui-box=` 变空)。与 VT_I2 同一写法即可: 8/16 位扩到
         * double 不丢精度, 所以不必像 VT_I8/UI4/UI8 那样单开一条无格式串的快路。 */
        case (vb6_vartype)VT_I1:
            numVal = (double)expr.sbVal; isNumeric = 1; break;
        case (vb6_vartype)VT_UI2:
            numVal = (double)expr.usVal; isNumeric = 1; break;
        /* 账 #305 §B142: VT_DECIMAL 这一档以前根本不在这张表里(落 default ⇒ CStr(Decimal)
         * 交空串)。形状照上面 :40 那条 VT_I8 的先例: 无格式串走精确出口, 带格式串才
         * 转 double 进通用数值分支(用户要的就是数值语义)。两个出口都在 vb6rtl.c。 */
        case (vb6_vartype)VT_DECIMAL: {
            if (!fmt) return vb6_DecimalToBSTR(expr);
            { double dv = 0.0;
              if (!vb6_DecimalToR8(expr, &dv)) return vb6_BSTR_FromStr(L"");
              numVal = dv; isNumeric = 1; }
            break;
        }
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
            // 值真是日期 => 这条臂多认 w/ww/q/ttttt (账 #427 §B201 的 allowWQT=1)
            if (fmt && vb6_BSTR_Len(fmt) > 0 && vb6_fmtIsDateFormat(fmt, 1))
                return vb6_fmtDateSerial(expr.dblVal, fmt, firstDayOfWeek, firstWeekOfYear);
            if (!fmt) {
                // 没有格式串 = general date 那一档。原先这里无条件拼「短 + 空格 + 长」⇒ 整点日期多发
                // 一枚空格与 0:00:00 (第一百零一刀 §B203 的①: 这一问的答案住在 vb6_fmtLocaleGeneralDate 里)
                SYSTEMTIME st;
                if (VariantTimeToSystemTime(expr.dblVal, &st)) {
                    wchar_t genBuf[128];
                    vb6_fmtLocaleGeneralDate(&st, genBuf, 128);
                    return vb6_BSTR_FromStr(genBuf);
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

