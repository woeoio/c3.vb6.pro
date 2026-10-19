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

// 「locale 的短日期 / 长日期 / 长时间」这三问在 RTL 里只许住一处 (账 #427 §B201 起，第一百零一刀 §B203 补全)。
// 吃它的: general date 那一档、short date / long date / long time 三档、没有格式串那条兜底臂、
// dddd / ddddd / ttttt 三枚令牌，再加**跨文件**的两位读者 —— vb6rtl_system.c 的 vb6_FormatDateTime 与
// vb6rtl_conv.c 的 vb6_CStrDate（第一百零一刀把它俩接回来；此前每家各抄一份 GetDateFormatW，同一问三个答案）。
// 所以这一族不再是 file-static，声明在 vb6rtl_builtin.h。
void vb6_fmtLocaleShortDate(const SYSTEMTIME* st, wchar_t* out, size_t cap) {
    out[0] = 0;
    GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, st, NULL, out, (int)cap);
}
void vb6_fmtLocaleLongDate(const SYSTEMTIME* st, wchar_t* out, size_t cap) {
    out[0] = 0;
    GetDateFormatW(LOCALE_USER_DEFAULT, DATE_LONGDATE, st, NULL, out, (int)cap);
}
void vb6_fmtLocaleLongTime(const SYSTEMTIME* st, wchar_t* out, size_t cap) {
    out[0] = 0;
    GetTimeFormatW(LOCALE_USER_DEFAULT, 0, st, NULL, out, (int)cap);
}
// 「general date 这一档到底给不给时间」这一问只住这里 (账 #429 ①, 第一百零一刀)。
// 引擎本人的读数 (发问表 .build/k559_olefdt3.c -> k559_out.txt 二十行; 逐点定带 .build/k564_oleband.c ->
// k564_out.txt 十九行 —— **两发合起来**才是这一条, 单看任何一发都会读成别的规则):
//   零头 0 / 1e-11 / 1e-10 / 1e-9 / 1e-8 日, 与 50 / 60 / 70 / 80 / 86.4 / 100 / 150 / 200 / 300 / 400 /
//   499 / 500 毫秒 => 全只给短日期 (`2026/1/5` len=8, **不带尾随空格**)
//   600 毫秒 => `0:00:01`; 59.4 秒 => `0:00:59`; 59.6 秒 => `0:01:00`; 29.9 秒 => `0:00:30`
//                          => 时间那一段是**四入到整秒**交回的 (账 #419 那条「截断」是另一问, 与这里无关)
//   午夜前一毫秒 (46028 - 1 毫秒) => `2026/1/6` len=8 => 四入**越过午夜**那一档整段不给、连日期一起进位
// => 那条闸既不在 0 上、也不是任何浮点容差, 而是「四入之后的当日时分秒全为零就不给时间段」: 判「给不给」
//   与渲染「给什么」吃同一个四入, 而那枚 SYSTEMTIME 已由 VariantTimeToSystemTime 四入过 => 这一问连序列号
//   本体都不需要, 出口的形参跟着少一条。
// 两条旧答案一起作废: Fix 175 留在 conv.c 的 1e-9 在 500 毫秒那一档答带时间; vb6rtl_system.c 那条
//   「零头 != 0」在午夜前一毫秒那一档答带时间 —— 引擎两处都答「只给日期」。
void vb6_fmtLocaleGeneralDate(const SYSTEMTIME* st, wchar_t* out, size_t cap) {
    size_t n;
    vb6_fmtLocaleShortDate(st, out, cap);
    if (st->wHour == 0 && st->wMinute == 0 && st->wSecond == 0) return;
    n = wcslen(out);
    if (n + 2 > cap) return;
    out[n] = L' ';
    out[n + 1] = 0;
    vb6_fmtLocaleLongTime(st, out + n + 1, cap - n - 1);
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
        case 1:  // general date = 短日期，四入后当日时分秒全零就不加时间段 (第一百零一刀)
            vb6_fmtLocaleGeneralDate(&st, buf, 128);
            break;
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
            vb6_fmtLocaleLongDate(&st, buf, 128);
            break;
        case 5:  // short date
            vb6_fmtLocaleShortDate(&st, buf, 128);
            break;
        case 6:  // long time = locale 默认时间格式 (含秒)
            vb6_fmtLocaleLongTime(&st, buf, 128);
            break;
        case 7:  // short time = HH:mm (24 小时制)
            GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, L"HH:mm", buf, 128);
            break;
    }
    return vb6_BSTR_FromStr(buf);
}

// Fix 162b-extlist: vb6_fmtNamedDateTimeKind 定义在上方 (vb6_fmtNamedDateTime 前)。

// 「这枚格式串里有没有上午/下午记号、它怎么写、在哪个位置」这一问收成一处 (账 #426 第九十八刀)。
// 运行期那张表认三形: AM/PM (或写成 AMPM) 交两位记号、A/P 交一位记号，大小写跟随串里写的那一种；
// 串里出现任意一枚就把小时切到 12 小时制，否则 h/hh 是 24 小时制 (语料把「hh:mm:ss」当时间戳用、
// 把「h:nn AMPM」当上下午用，两种写法都存在，所以这一位必须由记号决定而不是固定某一档)。
static int vb6_fmtAmPmLenAt(BSTR fmt, int i, int n) {
    static const wchar_t* const kTokens[] = { L"AM/PM", L"AMPM", L"A/P" };
    for (int k = 0; k < 3; k++) {
        int tl = (int)wcslen(kTokens[k]);
        if (i + tl > n) continue;
        int ok = 1;
        for (int j = 0; j < tl; j++)
            if (towlower(fmt[i + j]) != towlower(kTokens[k][j])) { ok = 0; break; }
        if (ok) return (k == 2) ? -1 : tl;   // 负数 = A/P 那一形 (只交一位字母)
    }
    return 0;
}

static int vb6_fmtHasAmPm(BSTR fmt) {
    int n = fmt ? SysStringLen(fmt) : 0;
    for (int i = 0; i < n; i++)
        if (vb6_fmtAmPmLenAt(fmt, i, n) != 0) return 1;
    return 0;
}

// 「这枚格式串算不算日期/时间格式」这一问被两条臂用 (账 #427 §B201 新认 w / ww / q / ttttt 之后):
//   allowWQT = 1  值真是 VT_DATE (extract.inc 那条臂) —— 新令牌无条件认。
//   allowWQT = 0  值是数值 (parse.inc 的 Fix 118b 兜底) —— 新令牌**只在整串里一个数值占位符都没有**
//                 时才认 (0 与 # 是数值档的唯一记号)。
// 为什么要加这一道闸: 语料里字面量格式串含 w/q/t 的 218 处全是裸写 (没有一处把单位文本用引号包起来),
// 其中「数值 + 单位」那一族 15 处 (#,##0 Twip 3 / Settings 4 / T0 2 / #### Byte 1 / #,# bytes 4 /
// ISO 串那枚字面 T 1) 走的正是数值臂, 而裸字母一视同仁就把它们从「那个数」改成「一个日期」
// (尺 .build/k496_census.txt + k497_wqt_where.txt; 现状读数 .build/k504 / k505 / k511)。
// 为什么单靠「值真是日期」这一档不够: 实测 Date 型实参到 vb6_Format 手里已经是 R8 —— Format(d1,「ttttt」)
// 与 Format(d1,「w」) 走的是数值臂 (而 VarType(d1) 答 7), 所以「值是不是日期」这一问在运行期答不全;
// 那一格另立新账 #431 (要动的是发码那头的装箱档), 本刀用「没有数值占位符」这一条把 Date 局部救回来。
// 游程仍然要紧: t 只有五连 (ttttt)、单枚 t 是字面量 —— ISO 串里那枚字面 T 与「T0」靠这一位保住。
static int vb6_fmtIsDateFormat(BSTR fmt, int allowWQT) {
    int n = fmt ? SysStringLen(fmt) : 0;
    if (n <= 0) return 0;
    if (vb6_fmtIsNamedFormat(fmt)) return 0;   // 命名格式走既有分支
    if (vb6_fmtHasAmPm(fmt)) return 1;   // 上午/下午记号也是日期习惯—— 一枚「A/P」串也得进日期分支 (账 #426)
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
    if (hasY || hasD || hasH || hasS || hasN || hasM) return 1;
    // 数值占位符闸: 串里出现 0 或 # 就是数值档的写法, 数值臂那条不许被新令牌拖进日期臂。
    if (!allowWQT) {
        for (int i = 0; i < n; i++)
            if (fmt[i] == L'0' || fmt[i] == L'#') return 0;
    }
    for (int i = 0; i < n; ) {
        int j = i; while (j < n && fmt[j] == fmt[i]) j++;
        wchar_t lc = towlower(fmt[i]);
        if (lc == L'w' || lc == L'q') return 1;
        if (lc == L't' && j - i >= 5) return 1;
        i = j;
    }
    return 0;
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

static void vb6_fmtAppendAmPm(wchar_t* out, size_t cap, size_t* o, BSTR fmt, int i, int len, int beforeNoon) {
    const wchar_t* s = (len < 0) ? (beforeNoon ? L"A" : L"P") : (beforeNoon ? L"AM" : L"PM");
    wchar_t buf[4];
    size_t k = 0;
    if (towlower(fmt[i]) == fmt[i]) {                 // 串里写的是小写 => 交小写记号
        for (; s[k] && k < 3; k++) buf[k] = (wchar_t)towlower(s[k]);
        buf[k] = 0;
        s = buf;
    }
    vb6_fmtAppendW(out, cap, o, s);
}

static BSTR vb6_fmtDateSerial(double serial, BSTR fmt, int32_t firstDayOfWeek, int32_t firstWeekOfYear) {
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
    int use12 = vb6_fmtHasAmPm(fmt);   // 12/24 小时那一问只在这里定一次

    wchar_t out[256]; out[0] = 0; size_t o = 0;
    int n = SysStringLen(fmt);
    int prevHour = 0;   // 紧邻 t 的 m 视为分钟 (VB6: "hh:mm" 的 mm 是分)
    for (int i = 0; i < n; ) {
        wchar_t c = fmt[i];
        int j = i; while (j < n && fmt[j] == c) j++;
        int cnt = j - i;
        wchar_t lc = towlower(c);
        int ampmLen = vb6_fmtAmPmLenAt(fmt, i, n);
        if (ampmLen != 0) {
            vb6_fmtAppendAmPm(out, 256, &o, fmt, i, ampmLen, st.wHour < 12);
            i += (ampmLen < 0) ? 3 : ampmLen;
            prevHour = 0;
            continue;
        }
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
                // ddddd = locale 短日期 (VB6 那张表里第五枚另是一档), dddd 才是长周名 —— 原先 cnt>=4 一并算走
                if (cnt >= 5) {
                    wchar_t sd[64];
                    vb6_fmtLocaleShortDate(&st, sd, 64);
                    vb6_fmtAppendW(out, 256, &o, sd);
                }
                else if (cnt == 4) vb6_fmtAppendW(out, 256, &o, kDayLong[dow]);
                else if (cnt == 3) vb6_fmtAppendW(out, 256, &o, kDayShort[dow]);
                else if (cnt == 2) vb6_fmtAppendN(out, 256, &o, day, 2);
                else               vb6_fmtAppendN(out, 256, &o, day, 1);
                prevHour = 0;
                break;
            case L'h': {
                int h = use12 ? ((st.wHour % 12) == 0 ? 12 : st.wHour % 12) : st.wHour;
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
            case L'q':
                // 季度: 与 vb6_DatePart 的「q」臂同一算式 (month-1)/3+1, 语料 1 处
                vb6_fmtAppendN(out, 256, &o, (month - 1) / 3 + 1, 1);
                prevHour = 0;
                break;
            case L'w':
                // w = 星期序号 (Sunday=1..7, 与周首无关) / ww = 一年里第几周 (那两条形参真被读)
                if (cnt >= 2) vb6_fmtAppendN(out, 256, &o,
                    vb6_WeekOfYear((int32_t)serial, firstDayOfWeek, firstWeekOfYear), 1);
                else          vb6_fmtAppendN(out, 256, &o, vb6_Weekday(serial, 1), 1);
                prevHour = 0;
                break;
            case L't':
                // 只有 ttttt (五枚) 是「长时间」令牌; 1..4 枚按字面量交 (ISO 串里那枚字面 T 就是这一档)
                if (cnt >= 5) {
                    wchar_t lt[64];
                    vb6_fmtLocaleLongTime(&st, lt, 64);
                    vb6_fmtAppendW(out, 256, &o, lt);
                } else {
                    for (int k = 0; k < cnt; k++)
                        if (o < 254) { out[o++] = c; out[o] = 0; }
                }
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
    // VBFlexGrid.ctl:20425/20428 `Format$(Text, Col.Format, vbUseSystemDayOfWeek, vbUseSystem)`
    // 直接 C2197 too many args —— 补上签名即可。
    // 账 #427 (§B201): 这两条形参从「(void) 丢掉」变成真被读 —— 「ww」那一档把它交给 vb6_WeekOfYear,
    // 与 vb6_DatePart 同一个主人。0 = vbUseSystem / vbUseSystemDayOfWeek 由那一位归到 vbSunday+vbFirstJan1
    // (实测 0 与 1 同档, 见 date.c 那条注释), 这里不再各自兜底。
#include "vb6rtl_format_extract.inc"
#include "vb6rtl_format_parse.inc"
#include "vb6rtl_format_string.inc"
#include "vb6rtl_format_numeric_pre.inc"
#include "vb6rtl_format_numeric_body.inc"
#include "vb6rtl_format_numeric_tail.inc"
}


