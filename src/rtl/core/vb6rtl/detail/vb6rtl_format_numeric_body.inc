// vb6rtl_format_numeric_body.inc - vb6_Format 函数体片段（位于 if (isNumericFormat || !isString) 块内）：数字格式化与 0/# 占位符补位
// 由 src/rtl/core/vb6rtl/vb6rtl_format.c 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 399~532 行，逐行未改。

        // Format the number
        wchar_t numBuf[128];

        if (ePos >= 0) {
            // Scientific notation
            // Calculate exponent
            double absVal = fabs(scaleVal);
            int exponent = 0;
            if (absVal != 0.0) {
                exponent = (int)floor(log10(absVal));
                // Adjust: VB6 normalizes so mantissa is 1.0 <= m < 10.0 if leading #,
                // or 0.x if no integer placeholder
            }
            double mantissa = (absVal != 0.0) ? scaleVal / pow(10.0, (double)exponent) : 0.0;
            if (val < 0) mantissa = -fabs(mantissa);
            if (val < 0 && scaleVal < 0) { /* mantissa already negative */ }

            // Round mantissa to decimalPlaces
            if (decimalPlaces > 0) {
                double factor = pow(10.0, (double)decimalPlaces);
                mantissa = round(mantissa * factor) / factor;
            }

            wchar_t mantStr[64];
            swprintf(mantStr, 64, L"%.*f", decimalPlaces, mantissa);

            // Format exponent
            int expVal = exponent;
            wchar_t expStr[32];
            wchar_t eChar = (ePlusType <= 2) ? L'E' : L'e';
            if (ePlusType == 1 || ePlusType == 3) {
                // Always show sign
                swprintf(expStr, 32, L"%c+%0*d", eChar, eDigits > 0 ? eDigits : 2, expVal);
            } else {
                swprintf(expStr, 32, L"%c%0*d", eChar, eDigits > 0 ? eDigits : 2, abs(expVal));
                if (expVal < 0) {
                    // Insert minus sign
                    wchar_t tmp[32];
                    swprintf(tmp, 32, L"%c-%s", eChar, expStr + 2);
                    wcscpy(expStr, tmp);
                }
            }

            swprintf(numBuf, 128, L"%s%s", mantStr, expStr);
        } else {
            // Fixed-point notation
            // §11.171: 真机 `Format` 格式串的进位与 FormatNumber 同口径 —— **十进制**半值远离零。
            // 旧写法 `round(scaleVal * factor) / factor` 是在 double 上进位（二进制），于是
            //   Format(1.005, "0.00") = 1.00（真机 1.01）、Format(9.995, "0.00") = 9.99（真机 10.00）、
            //   Format(0.4999999999999999, "0") = 0（真机 1，15 位十进制文本已是 0.5）。
            // 反证：Format(1.0049999, "0.00") 必须仍是 1.00 —— 所以不能拿二进制完整展开进位。
            if (decimalPlaces >= 0) {
                vb6_fmtFixedText(scaleVal, decimalPlaces, numBuf, 128);
            } else {
                swprintf(numBuf, 128, L"%.*f", decimalPlaces, scaleVal);
            }
        }

        // Now reformat according to 0/# placeholders
        // Parse numBuf into integer and fractional parts
        wchar_t intPart[64] = L"0";
        wchar_t fracPart[64] = L"";
        int isNegative = 0;
        {
            int ni = 0;
            if (numBuf[0] == L'-') { isNegative = 1; ni = 1; }
            int di = 0;
            while (numBuf[ni] && numBuf[ni] != L'.') { intPart[di++] = numBuf[ni++]; }
            intPart[di] = L'\0';
            if (numBuf[ni] == L'.') {
                ni++;
                di = 0;
                while (numBuf[ni]) { fracPart[di++] = numBuf[ni++]; }
                fracPart[di] = L'\0';
            }
        }

        // Handle negative display for multi-section (section 1 = negative format)
        // In multi-section, the negative section's format is applied to the absolute value
        if (secIdx == 1 && isNegative) {
            isNegative = 0;  // don't double-negate
        }

        // Pad/truncate fractional part according to placeholders
        {
            int fLen = (int)wcslen(fracPart);
            // Extend if afterDot0 requires more digits
            while (fLen < afterDot0) { fracPart[fLen++] = L'0'; }
            fracPart[fLen] = L'\0';
            // Truncate if too many digits (already rounded above)
            if (fLen > totalAfterDot) {
                fracPart[totalAfterDot] = L'\0';
            }
        }

        // Pad integer part according to 0 placeholders
        {
            int iLen = (int)wcslen(intPart);
            while (iLen < beforeDot0) {
                // Prepend '0'
                for (int j = iLen; j >= 0; j--) intPart[j + 1] = intPart[j];
                intPart[0] = L'0';
                iLen++;
            }
        }

        // Apply thousand separator if format has commas
        {
            // Check if format has commas as thousand separators
            int hasThousandSep = 0;
            int iq = 0;
            for (int i = 0; i < secLen; i++) {
                if (sec[i] == '"') { iq = !iq; continue; }
                if (iq) continue;
                if (sec[i] == '\\') { i++; continue; }
                if (sec[i] == ',' && !hasComma1000) hasThousandSep = 1;
            }
            if (hasThousandSep) {
                // Insert thousand separators into intPart
                int iLen = (int)wcslen(intPart);
                wchar_t tmp[64];
                int ti = 0;
                for (int j = iLen - 1, cnt = 0; j >= 0; j--) {
                    tmp[ti++] = intPart[j];
                    cnt++;
                    if (cnt % 3 == 0 && j > 0) tmp[ti++] = L',';
                }
                tmp[ti] = L'\0';
                // Reverse
                for (int j = 0; j < ti / 2; j++) {
                    wchar_t t = tmp[j]; tmp[j] = tmp[ti - 1 - j]; tmp[ti - 1 - j] = t;
                }
                wcscpy(intPart, tmp);
            }
        }

