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

        // §11.169 (第一百八十六刀): 整数区有 **≥2 组** 占位符时, 数字要按组分配。
        // 真机 (.temp/k186p/P204.bas):
        //   "0 0"@1234 -> "123 4"      "0 00"@1234 -> "12 34"    "0 000"@1234 -> "1 234"
        //   "#,##0 0" -> "1,23 4"      "0 #,##0" -> "0 1,234"    "00 00"@12 -> "00 12"
        // 口径 = ① 整串先按千分位分组 (所以 "1,23" 那个逗号是整数的, 不是组间的);
        //        ② 从右往左每组按容量切片, **最左组吸收余数**;
        //        ③ 切不到数的组按自己的 0/# 补位 ('0' 补 0, '#' 补空)。
        // 只在**没有小数点**时走这条路 (带点的多组真机不合这几条规律, 见在册)。
        if (dotPos < 0) {
            int cur = -1, ng = 0, iq0 = 0, commaAny = 0;
            int gCap[8], gZero[8];
            for (int i = 0; i < secLen; i++) {
                wchar_t c = sec[i];
                if (c == '"') { iq0 = !iq0; if (cur >= 0) cur = -1; continue; }
                if (iq0) continue;
                if (c == '\\') { i++; if (cur >= 0) cur = -1; continue; }
                if (c == '0' || c == '#') {
                    if (cur < 0) {
                        if (ng >= 8) { cur = -2; continue; }
                        cur = ng++; gCap[cur] = 0; gZero[cur] = 0;
                    }
                    if (cur >= 0) { gCap[cur] += 1; if (c == '0') gZero[cur] += 1; }
                    continue;
                }
                if (c == ',') { commaAny = 1; continue; }
                if (cur >= 0) cur = -1;
            }
            if (ng >= 2 && cur != -2) {
                // ① 数字源: 整只整数部分 (值为 0 且全是 '#' 时为空)
                int allHash = 1;
                for (int gi = 0; gi < ng; gi++) if (gZero[gi] > 0) allHash = 0;
                int dropAll = (allHash && wcscmp(intPart, L"0") == 0);
                wchar_t gsrc[160]; int gsLen = 0;
                if (!dropAll) {
                    // intPart 可能已被上游按千分位插过逗号 (§千分位那段) —— 这里只取数字位,
                    // 逗号下面按 commaAny 重算一次, 否则会出 "1,,23"。
                    wchar_t digs[128]; int dLen1 = 0;
                    int dLen0 = (int)wcslen(intPart);
                    for (int i = 0; i < dLen0 && dLen1 < 120; i++)
                        if (intPart[i] != L',') digs[dLen1++] = intPart[i];
                    if (!commaAny) {
                        for (int i = 0; i < dLen1 && gsLen < 150; i++) gsrc[gsLen++] = digs[i];
                    } else {
                        for (int i = 0; i < dLen1; i++) {
                            if (i > 0 && (dLen1 - i) % 3 == 0 && gsLen < 150) gsrc[gsLen++] = L',';
                            if (gsLen < 150) gsrc[gsLen++] = digs[i];
                        }
                    }
                }
                gsrc[gsLen] = L'\0';
                // ② 右往左切片
                int slStart[8], slEnd[8];
                int p = gsLen;
                for (int gi = ng - 1; gi >= 0; gi--) {
                    if (gi == 0) { slStart[0] = 0; slEnd[0] = p; }
                    else {
                        int taken = 0, q = p;
                        while (q > 0 && taken < gCap[gi]) {
                            q--;
                            if (gsrc[q] != L',') taken++;
                        }
                        slStart[gi] = q; slEnd[gi] = p; p = q;
                    }
                }
                // ③ 按出现顺序拼: 字面量照抄, 组按切片 (空则补位)
                if (isNegative) result[rLen++] = L'-';
                int gi2 = 0;
                int cur2 = -1, iq2 = 0;
                for (int i = 0; i < secLen; i++) {
                    wchar_t c = sec[i];
                    if (c == '"') { iq2 = !iq2; if (cur2 >= 0) cur2 = -1; continue; }
                    if (c == '\\') {
                        if (i + 1 >= secLen) break;
                        c = sec[++i];
                        if (cur2 >= 0) cur2 = -1;
                    } else if (!iq2) {
                        if (c == '0' || c == '#') {
                            if (cur2 < 0) {
                                cur2 = gi2++;
                                int k = slStart[cur2];
                                int n = 0;
                                for (; k < slEnd[cur2]; k++) { result[rLen++] = gsrc[k]; if (gsrc[k] != L',') n++; }
                                if (n == 0) {
                                    for (int z = 0; z < gZero[cur2]; z++) result[rLen++] = L'0';
                                }
                            }
                            continue;
                        }
                        if (c == ',') {
                            continue;   // 整数区里的逗号一律是千分位记号 (真机 "0 h, 0 m" -> "1,23 h 4 m"), 不当字面量
                        } else if (c == '%' || c == '.' || c == 'E' || c == 'e') {
                            if (!(c == 'E' || c == 'e')) continue;
                        }
                        if (cur2 >= 0) cur2 = -1;
                    }
                    if (rLen < 240) result[rLen++] = c;
                }
                if (hasPercent) result[rLen++] = L'%';
                result[rLen] = L'\0';
                free(fmtStr);
                return vb6_BSTR_FromStr(result);
            }
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
                    // §11.168: **段末那一枚 e/E 也是记号** (真机 P201: "0 Byte"->"1234 Byt"、
                    // "0e"->"1234"、"0 Be"->"1234 B"、"0 e"->"1234 "；而 "0 Bytes"/"0 ex"/
                    // "0 E5" 里的 e/E 后面还有字，照旧是字面量)
                    if ((c == 'E' || c == 'e') && i == secLen - 1) continue;
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
