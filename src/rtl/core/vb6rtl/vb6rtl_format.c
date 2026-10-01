// vb6rtl_format.c - VB6 运行时库: 格式化家族：vb6_Format 用户自定义格式实现
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 284~893 行
// 2026-09-17 二次拆分：vb6_Format 函数体 606 行按既有 Step 分节切为 6 个片段，
// 本文件（伞文件）只保留 25 行前置说明 + 函数签名与片段 #include + 收尾两行。
// 片段与原 vb6rtl.c 行区间对照（本文件行号 = 原 vb6rtl.c 行号 - 258）：
//   vb6rtl_format_extract.inc       原 285~339    数值提取与无格式串默认转换（Step 1~2）
//   vb6rtl_format_parse.inc         原 340~455    命名格式识别与用户格式串解析（Step 3~4）
//   vb6rtl_format_string.inc        原 456~551    格式类型判定与字符串格式应用（Step 5~6）
//   vb6rtl_format_numeric_pre.inc   原 552~656    数值格式：校验、千分位缩放、占位符计数
//   vb6rtl_format_numeric_body.inc  原 657~790    数值格式：数字格式化与 0/# 补位
//   vb6rtl_format_numeric_tail.inc  原 791~890    数值格式：字面量拼装与回退（Step 7 尾 + Step 8）
// 六个 .inc 是「函数体片段」，在 vb6_Format() 函数体内被 #include（C 允许），故不用 .c/.h 后缀 ——
// 它们不是独立编译单元，单独编译会不过。片段逐行未改，局部变量与 static 表原样不动 → 零行为改动。

#include "vb6rtl.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <stdarg.h>
#include <wchar.h>
#include <wctype.h>
#include <time.h>
#ifdef _WIN32
#include <direct.h>
#include <io.h>
#include <oleauto.h>
#include <olectl.h>
#include <windows.h>
#endif

// P24-08: MessageBoxW (user32) + GetConsoleWindow (kernel32)
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "kernel32.lib")

// ============================================================
// Fix 118: VB6 用户自定义日期/时间格式串 (mmm / mmmm / yyyy / dd / dddd / hh:nn:ss ...)
// ============================================================
// 此前 VT_DATE + 格式串一律落入"数值格式化"分支, 于是
//   Format(DateSerial(2020,1,1), "mmm") → "43831" (日期序列号)
// Charts 2020 的 ucChartBar X 轴月份标签 (Form2: Format(DateSerial(2020,i,1),"mmm"))
// 因此变成一串数字, 旋转后看成了乱码。
static int vb6_fmtIsNamedFormat(BSTR fmt) {
    static const wchar_t* kNames[] = {
        L"general number", L"currency", L"fixed", L"standard", L"percent",
        L"scientific", L"yes/no", L"true/false", L"on/off",
        // Fix 162b-extlist: 补全命名日期/时间格式 (extlist 工程的 Format$ 用到
        // "short date"/"medium date"/"short time"/"medium time"/"long time")。
        // 原表只有 long/short date 与 long/short time —— "medium date"/"medium
        // time" 不被认成命名格式, 也不含 yyyy/dd 等"日期特征字符"齐全度足够,
        // 被当**用户格式串**逐字符解析 (m→月 e→字面量 d→日 i/u→字面量) ⇒
        // 输出 "12e14u12" 这种乱码; 而 short date/short time/long time 虽被
        // vb6_fmtIsDateFormat 排除, 但后面**没有**命名日期分支实现 → 落回
        // 数值格式化输出序列号 "45331"。两者都由 vb6_fmtNamedDateTime 接住。
        L"general date", L"medium date", L"medium time",
        L"long date", L"short date", L"long time", L"short time"};
    int n = fmt ? SysStringLen(fmt) : 0;
    if (n <= 0 || n >= 63) return 0;
    wchar_t buf[64];
    for (int i = 0; i < n; i++) buf[i] = towlower(fmt[i]);
    buf[n] = 0;
    for (size_t k = 0; k < sizeof(kNames) / sizeof(kNames[0]); k++)
        if (wcscmp(buf, kNames[k]) == 0) return 1;
    return 0;
}

static int vb6_fmtNamedDateTimeKind(BSTR fmt) {
    static const wchar_t* kDateNames[] = {
        L"general date", L"medium date", L"medium time",
        L"long date", L"short date", L"long time", L"short time"};
    int n = fmt ? SysStringLen(fmt) : 0;
    if (n <= 0 || n >= 63) return 0;
    wchar_t buf[64];
    for (int i = 0; i < n; i++) buf[i] = towlower(fmt[i]);
    buf[n] = 0;
    for (size_t k = 0; k < sizeof(kDateNames) / sizeof(kDateNames[0]); k++)
        if (wcscmp(buf, kDateNames[k]) == 0) return (int)k + 1;
    return 0;
}

// Fix 162b-extlist: 命名日期/时间格式 → 按各规格取值。
//   general date = 短日期 + 空格 + 长时间 (locale)
//   long date    = locale 长日期        short date = locale 短日期
//   medium date  = d-mmm-yy             (VB6 固定规格, 与 locale 无关)
//   long time    = locale 长时间 (h:mm:ss AM/PM)
//   medium time  = hh:mm AM/PM          short time = HH:mm (24 小时制)
static BSTR vb6_fmtNamedDateTime(double serial, BSTR fmt) {
    int kind = vb6_fmtNamedDateTimeKind(fmt);
    SYSTEMTIME st;
    if (!VariantTimeToSystemTime(serial, &st)) return vb6_BSTR_FromStr(L"");
    wchar_t buf[128];
    buf[0] = 0;
    switch (kind) {
        case 1: {  // general date
            wchar_t d[64], t[64];
            GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, NULL, d, 64);
            GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, NULL, t, 64);
            swprintf(buf, 128, L"%s %s", d, t);
            break;
        }
        case 2: {  // medium date —— locale 相关: zh-CN = yy-MM-dd (参考图 "26-12-14"),
                   // 西文 = d-mmm-yy ("28-Dec-26")。按用户主语言分流。
            wchar_t md[16] = L"d-MMM-yy";
            LANGID lang = GetUserDefaultUILanguage();
            if (PRIMARYLANGID(lang) == 0x04) { // Chinese: 参考图 "26-12-14"
                wcscpy(md, L"yy-MM-dd");
            }
            GetDateFormatW(LOCALE_USER_DEFAULT, 0, &st, md, buf, 128);
            break;
        }
        case 3:  // medium time = hh:mm AM/PM
            GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, L"hh:mm tt", buf, 128);
            break;
        case 4:  // long date
            GetDateFormatW(LOCALE_USER_DEFAULT, DATE_LONGDATE, &st, NULL, buf, 128);
            break;
        case 5:  // short date
            GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, NULL, buf, 128);
            break;
        case 6:  // long time = locale 默认时间格式 (含秒)
            GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, NULL, buf, 128);
            break;
        case 7:  // short time = HH:mm (24 小时制)
            GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, L"HH:mm", buf, 128);
            break;
    }
    return vb6_BSTR_FromStr(buf);
}

// Fix 162b-extlist: vb6_fmtNamedDateTimeKind 定义在上方 (vb6_fmtNamedDateTime 前)。

static int vb6_fmtIsDateFormat(BSTR fmt) {
    int n = fmt ? SysStringLen(fmt) : 0;
    if (n <= 0) return 0;
    if (vb6_fmtIsNamedFormat(fmt)) return 0;   // 命名格式走既有分支
    int hasY = 0, hasD = 0, hasH = 0, hasM = 0, hasS = 0, hasN = 0;
    for (int i = 0; i < n; i++) {
        switch (towlower(fmt[i])) {
            case L'y': hasY = 1; break;
            case L'd': hasD = 1; break;
            case L'h': hasH = 1; break;
            case L'm': hasM = 1; break;   // m 在日期格式里=月, 时间格式里=分 (见下方 prevHour)
            case L's': hasS = 1; break;
            case L'n': hasN = 1; break;
        }
    }
    return hasY || hasD || hasH || hasS || hasN || hasM;
}

static void vb6_fmtAppendW(wchar_t* out, size_t cap, size_t* o, const wchar_t* s) {
    size_t l = s ? wcslen(s) : 0;
    if (*o + l < cap - 1) { wcscpy(out + *o, s); *o += l; }
}
static void vb6_fmtAppendN(wchar_t* out, size_t cap, size_t* o, int value, int digits) {
    wchar_t tmp[16];
    if (digits == 2) swprintf(tmp, 16, L"%02d", value);
    else if (digits == 4) swprintf(tmp, 16, L"%04d", value);
    else swprintf(tmp, 16, L"%d", value);
    vb6_fmtAppendW(out, cap, o, tmp);
}

static BSTR vb6_fmtDateSerial(double serial, BSTR fmt) {
    SYSTEMTIME st;
    if (!VariantTimeToSystemTime(serial, &st)) return vb6_BSTR_FromStr(L"");
    static const wchar_t* kMonLong[12] = {L"January", L"February", L"March", L"April",
        L"May", L"June", L"July", L"August", L"September", L"October", L"November", L"December"};
    static const wchar_t* kMonShort[12] = {L"Jan", L"Feb", L"Mar", L"Apr", L"May", L"Jun",
        L"Jul", L"Aug", L"Sep", L"Oct", L"Nov", L"Dec"};
    static const wchar_t* kDayLong[7] = {L"Sunday", L"Monday", L"Tuesday", L"Wednesday",
        L"Thursday", L"Friday", L"Saturday"};
    static const wchar_t* kDayShort[7] = {L"Sun", L"Mon", L"Tue", L"Wed", L"Thu", L"Fri", L"Sat"};

    int month = st.wMonth, day = st.wDay, year = st.wYear;
    if (month < 1) month = 1; if (month > 12) month = 12;
    if (day < 1) day = 1; if (day > 31) day = 31;
    int dow = st.wDayOfWeek; if (dow < 0 || dow > 6) dow = 0;

    wchar_t out[256]; out[0] = 0; size_t o = 0;
    int n = SysStringLen(fmt);
    int prevHour = 0;   // 紧邻 t 的 m 视为分钟 (VB6: "hh:mm" 的 mm 是分)
    for (int i = 0; i < n; ) {
        wchar_t c = fmt[i];
        int j = i; while (j < n && fmt[j] == c) j++;
        int cnt = j - i;
        wchar_t lc = towlower(c);
        switch (lc) {
            case L'y':
                if (cnt >= 4) vb6_fmtAppendN(out, 256, &o, year, 4);
                else          vb6_fmtAppendN(out, 256, &o, year % 100, 2);
                prevHour = 0;
                break;
            case L'm':
                if (prevHour && cnt <= 2)   vb6_fmtAppendN(out, 256, &o, st.wMinute, cnt == 2 ? 2 : 1);
                else if (cnt >= 4)          vb6_fmtAppendW(out, 256, &o, kMonLong[month - 1]);
                else if (cnt == 3)          vb6_fmtAppendW(out, 256, &o, kMonShort[month - 1]);
                else                        vb6_fmtAppendN(out, 256, &o, month, cnt == 2 ? 2 : 1);
                break;
            case L'd':
                if (cnt >= 4)     vb6_fmtAppendW(out, 256, &o, kDayLong[dow]);
                else if (cnt == 3) vb6_fmtAppendW(out, 256, &o, kDayShort[dow]);
                else if (cnt == 2) vb6_fmtAppendN(out, 256, &o, day, 2);
                else               vb6_fmtAppendN(out, 256, &o, day, 1);
                prevHour = 0;
                break;
            case L'h': {
                int h = st.wHour % 12; if (h == 0) h = 12;
                vb6_fmtAppendN(out, 256, &o, h, cnt == 2 ? 2 : 1);
                prevHour = 1;
                break;
            }
            case L'n':
                vb6_fmtAppendN(out, 256, &o, st.wMinute, cnt == 2 ? 2 : 1);
                prevHour = 0;
                break;
            case L's':
                vb6_fmtAppendN(out, 256, &o, st.wSecond, cnt == 2 ? 2 : 1);
                prevHour = 0;
                break;
            default:
                for (int k = 0; k < cnt; k++)
                    if (o < 254) { out[o++] = c; out[o] = 0; }
                if (lc != L':' && lc != L' ') prevHour = 0;
                break;
        }
        i = j;
    }
    return vb6_BSTR_FromStr(out);
}

BSTR vb6_Format(vb6_VARIANT expr, BSTR fmt, int32_t firstDayOfWeek, int32_t firstWeekOfYear) {
    // <vbeclipse> 扩到 VB6 完整 4 形 (expr[, fmt[, FirstDayOfWeek[, FirstWeekOfYear]]]).
    // 现有格式解析不消费这两个日期参数 (Date 分支走系统默认周)；VBFlexGrid.ctl:
    // 20425/20428 `Format$(Text, Col.Format, vbUseSystemDayOfWeek, vbUseSystem)`
    // 直接 C2197 too many args —— 补上签名即可，语义与旧 2 参实现等价。
    (void)firstDayOfWeek; (void)firstWeekOfYear;
#include "vb6rtl_format_extract.inc"
#include "vb6rtl_format_parse.inc"
#include "vb6rtl_format_string.inc"
#include "vb6rtl_format_numeric_pre.inc"
#include "vb6rtl_format_numeric_body.inc"
#include "vb6rtl_format_numeric_tail.inc"
}


