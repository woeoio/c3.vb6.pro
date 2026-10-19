// vb6rtl_format_parse.inc - vb6_Format 函数体片段：命名格式识别与用户自定义格式串解析
// 由 src/rtl/core/vb6rtl/vb6rtl_format.c 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 82~197 行，逐行未改。

    // === Step 2: No format string → default conversion ===
    if (!fmt || vb6_BSTR_Len(fmt) == 0) {
        if (isNull) return vb6_BSTR_FromStr(L"");
        if (isNumeric) {
            wchar_t buf[64];
            /* Fix 117: VB6 默认数值格式 (Single 7 位有效数字, 最短往返) */
            return vb6_NumToBSTRDefault(numVal, expr.vt == vb6_vtSingle);
        }
        if (isString && strVal) return vb6_BSTR_FromBSTR(strVal);
        return vb6_BSTR_FromStr(L"");
    }

    // Convert fmt BSTR to narrow string for parsing
    int32_t fmtLen = vb6_BSTR_Len(fmt);
    char* fmtStr = (char*)malloc(fmtLen + 1);
    for (int32_t i = 0; i < fmtLen; i++) fmtStr[i] = (char)fmt[i];
    fmtStr[fmtLen] = '\0';

    // === Step 3: Check for named formats ===
    // Case-insensitive comparison
    char fmtLower[128];
    for (int i = 0; i < fmtLen && i < 127; i++) fmtLower[i] = (char)tolower((unsigned char)fmtStr[i]);
    fmtLower[fmtLen < 127 ? fmtLen : 127] = '\0';

    if (isNumeric) {
        // Fix 162b-extlist: 数值 + 命名日期/时间格式 → 按日期序列号走 locale/
        // VB6 规格输出 ("short date" 原先落空 → 数值分支输出 "45331")。
        if (vb6_fmtNamedDateTimeKind(fmt)) {
            free(fmtStr);
            return vb6_fmtNamedDateTime(numVal, fmt);
        }
        // Fix 118b: 数值 + 日期/时间格式串 → 按日期序列号格式化。
        // (VB6 把数值当日期序列号。Charts 2020 的 X 轴月份标签 Format(DateSerial(2020,i,1), "mmm")
        //  改前改后吃的都是这条臂 —— 第一百零二刀 (账 #431 §B205) 把 DateSerial 的返回档改回 Date 之后，
        //  Format 的**首参**仍由账 #430 那张手抄装箱表决定 (Date 档交 vb6_VariantDouble)，所以那一形
        //  没换站；换站的是「值先落进一枚 Variant 槽再交给 Format」那一形 (尺 .build/k638_new_run.txt:
        //  Format(槽) 交 [2026/1/5] 而改前交 [46027])。这条臂同时是 Format(43831, "mmm") 那族真数值的
        //  唯一出口 —— **臂不撤，理由换掉了**。)
        // allowWQT=0: 这条臂是「数值 + 日期字母」的兜底, 不许扩到 w/q/t ——
        // 语料里「#,##0 Twip」「T0」「#### Byte」那一族 15 处吃的就是这条臂 (账 #427 §B201)
        if (vb6_fmtIsDateFormat(fmt, 0)) {
            free(fmtStr);
            return vb6_fmtDateSerial(numVal, fmt, firstDayOfWeek, firstWeekOfYear);
        }
        // Named numeric formats
        if (strcmp(fmtLower, "general number") == 0) {
            free(fmtStr);
            wchar_t buf[64];
            /* Fix 117: VB6 默认数值格式 (Single 7 位有效数字, 最短往返) */
            return vb6_NumToBSTRDefault(numVal, expr.vt == vb6_vtSingle);
        }
        if (strcmp(fmtLower, "currency") == 0) {
            free(fmtStr);
            return vb6_FormatCurrency(numVal, -1, -1, 0, -1);
        }
        if (strcmp(fmtLower, "fixed") == 0) {
            free(fmtStr);
            wchar_t buf[64];
            swprintf(buf, 64, L"%.2f", numVal);
            return vb6_BSTR_FromStr(buf);
        }
        if (strcmp(fmtLower, "standard") == 0) {
            free(fmtStr);
            // Standard = thousand separator + 2 decimal places
            return vb6_FormatNumber(numVal, 2, -1, 0, -1);
        }
        if (strcmp(fmtLower, "percent") == 0) {
            free(fmtStr);
            return vb6_FormatPercent(numVal, 2, -1, 0, -1);
        }
        if (strcmp(fmtLower, "scientific") == 0) {
            free(fmtStr);
            wchar_t buf[64];
            swprintf(buf, 64, L"%.2E", numVal);
            return vb6_BSTR_FromStr(buf);
        }
        if (strcmp(fmtLower, "yes/no") == 0) {
            free(fmtStr);
            return vb6_BSTR_FromStr((numVal == 0.0) ? L"No" : L"Yes");
        }
        if (strcmp(fmtLower, "true/false") == 0) {
            free(fmtStr);
            return vb6_BSTR_FromStr((numVal == 0.0) ? L"False" : L"True");
        }
        if (strcmp(fmtLower, "on/off") == 0) {
            free(fmtStr);
            return vb6_BSTR_FromStr((numVal == 0.0) ? L"Off" : L"On");
        }
    }

    // Named date/time formats for date values
    // (Not yet implemented; fall through to user-defined)

    // === Step 4: User-defined format string ===
    // Parse multi-section format: positive;negative;zero[;null]
    // Find semicolons (not inside quoted strings)
    int sectionStart[4] = {0, -1, -1, -1};
    int sectionLen[4] = {0, 0, 0, 0};
    int numSections = 1;
    int inQuote = 0;
    for (int i = 0; i < fmtLen; i++) {
        if (fmtStr[i] == '"') { inQuote = !inQuote; continue; }
        if (inQuote) continue;
        if (fmtStr[i] == '\\') { i++; continue; }  // skip escaped char
        if (fmtStr[i] == ';') {
            sectionLen[numSections - 1] = i - sectionStart[numSections - 1];
            if (numSections < 4) {
                numSections++;
                sectionStart[numSections - 1] = i + 1;
            }
        }
    }
    sectionLen[numSections - 1] = fmtLen - sectionStart[numSections - 1];

    // Select section based on value
    int secIdx = 0;
    if (isNull) {
        if (numSections >= 4) secIdx = 3;
        else { free(fmtStr); return vb6_BSTR_FromStr(L""); }
    } else if (isString) {
        secIdx = 0;  // string always uses first section
    } else {
        // Numeric
        double absVal = fabs(numVal);
        if (numSections == 1) {
            secIdx = 0;
        } else if (numSections == 2) {
            secIdx = (numVal >= 0) ? 0 : 1;
        } else if (numSections >= 3) {
            secIdx = (numVal > 0) ? 0 : (numVal < 0) ? 1 : 2;
        }
    }
    char* sec = fmtStr + sectionStart[secIdx];
    int secLen = sectionLen[secIdx];

