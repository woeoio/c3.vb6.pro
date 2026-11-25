// vb6rtl_format_numeric_tail.inc - vb6_Format 函数体片段（位于 if (isNumericFormat || !isString) 块内）：字面量拼装与回退
// 由 src/rtl/core/vb6rtl/vb6rtl_format.c 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 533~632 行，逐行未改。

        // §11.167 (第一百八十四刀): 字面量按位置拼装。
        // 改前这一段自己在注释里写着 "Simplified … doesn't position them" —— 字面量全被丢掉:
        //   Format(1234.5678, "(0.00)")        -> "1234.57"   (真机 H1: "(1234.57)")
        //   Format(-d, "0.00;(0.00)")          -> "1234.57"   (真机 H3: "(1234.57)")
        //   Format(d, "0.00\"x\"") / "0.00-"    -> "1234.57"   (真机 H4/H5: 带尾巴)
        //   Format(d, "\\#0.00")               -> "1234.57"   (真机 H7: "#1234.57")
        //   Format(0, "0.00;0.0;zero")         -> "0"          (真机 H9: "zero")
        //   Format(0.5, "#.##")                -> "0.5"        (真机 H15: ".5" —— 整数段只有 '#'
        //                                                       且值为 0 时不打那个 0)
        // 本刀只做**前缀/后缀**两段 (占位符之间夹字面量那一形语料里没有), 科学计数那一路
        // 保持原样 —— 它的 numBuf 已经含尾数+指数。
        wchar_t result[256];
        int rLen = 0;

        if (ePos >= 0) {
            if (isNegative) result[rLen++] = L'-';
            int nl = (int)wcslen(numBuf);
            for (int i = 0; i < nl && rLen < 240; i++) result[rLen++] = numBuf[i];
            result[rLen] = L'\0';
            free(fmtStr);
            return vb6_BSTR_FromStr(result);
        }

        // 1) 最后一个数字占位符的位置 —— 用它把字面量分成前缀/后缀
        int lastPh = -1;
        {
            int iq = 0;
            for (int i = 0; i < secLen; i++) {
                wchar_t c = sec[i];
                if (c == '"') { iq = !iq; continue; }
                if (iq) continue;
                if (c == '\\') { i++; continue; }
                if (c == '0' || c == '#') lastPh = i;
            }
        }
        // 2) 逐字符收字面量: 引号内原样、反斜杠转义一位; 占位符与 , . % E e 跳过
        wchar_t prefix[64]; int preLen = 0;
        wchar_t suffix[64]; int sufLen = 0;
        {
            int iq = 0;
            for (int i = 0; i < secLen; i++) {
                wchar_t c = sec[i];
                if (c == '"') { iq = !iq; continue; }
                if (c == '\\') {
                    if (i + 1 >= secLen) break;
                    c = sec[++i];
                } else if (!iq) {
                    // ⚠ 只有 `E+` / `e-` 那两形是科学计数记号 (`zero` 里的 e 是字面量 ——
                    // 真机 H9 = "zero"，改前把它当记号丢掉后打 "zro")
                    int isSciMark = ((c == 'E' || c == 'e') && i + 1 < secLen
                                     && (sec[i + 1] == '+' || sec[i + 1] == '-'));
                    if (c == '0' || c == '#' || c == '.' || c == ',' || c == '%'
                        || isSciMark) continue;
                    if (c == 'E' || c == 'e') { /* 字面量 e，照常落下去 */ }
                }
                if (i < lastPh) { if (preLen < 60) prefix[preLen++] = c; }
                else            { if (sufLen < 60) suffix[sufLen++] = c; }
            }
        }
        prefix[preLen] = L'\0';
        suffix[sufLen] = L'\0';

        // 3) 拼装: 符号 + 前缀 + 整数 + 小数 + % + 后缀
        if (isNegative) result[rLen++] = L'-';
        for (int i = 0; i < preLen; i++) result[rLen++] = prefix[i];

        const int totalIntPh = beforeDot0 + beforeDotHash;
        const int totalFracPh = afterDot0 + afterDotHash;
        const int dropLeadZero = (beforeDot0 == 0 && totalIntPh > 0
                                  && wcscmp(intPart, L"0") == 0);
        if (totalIntPh > 0 || totalFracPh > 0) {
            int ipStart = dropLeadZero ? 1 : 0;
            int ipLen2 = (int)wcslen(intPart);
            for (int i = ipStart; i < ipLen2; i++) result[rLen++] = intPart[i];

            if (totalFracPh > 0 || dotPos >= 0) {
                int showFrac = (int)wcslen(fracPart);
                int minFrac = afterDot0;
                while (showFrac > minFrac && fracPart[showFrac - 1] == L'0') showFrac--;
                if (showFrac > 0 || afterDot0 > 0) {
                    result[rLen++] = L'.';
                    for (int i = 0; i < showFrac; i++) result[rLen++] = fracPart[i];
                }
            }
        }
        if (hasPercent) result[rLen++] = L'%';
        for (int i = 0; i < sufLen; i++) result[rLen++] = suffix[i];
        result[rLen] = L'\0';

        free(fmtStr);
        return vb6_BSTR_FromStr(result);
    }

    // === Step 8: Fallback ===
    free(fmtStr);
    if (isNumeric) {
        /* Fix 117: VB6 默认数值格式 (Single 7 位有效数字, 最短往返) */
        return vb6_NumToBSTRDefault(numVal, expr.vt == vb6_vtSingle);
    }
    if (isString && strVal) return vb6_BSTR_FromBSTR(strVal);
    return vb6_BSTR_FromStr(L"");
