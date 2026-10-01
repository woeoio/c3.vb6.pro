// vb6rtl_date.c - VB6 运行时库: 日期时间家族：序列号互转 + Date/Time 函数 + DateAdd/Diff/Part + Array 创建
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 2708~2827 行
//   原第 2828~3044 行
//   原第 3045~3148 行

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
// 日期时间函数
// ============================================================

static int32_t vb6_date_to_serial(int32_t year, int32_t month, int32_t day) {
    // Excel/Lotus日期序列号: 1900-01-01 = 1 (含Lotus bug: 1900-02-29 = 60)
    if (month <= 2) { year--; month += 12; }
    int32_t a = year / 100;
    int32_t b = 2 - a + a / 4;
    return (int32_t)(365.25 * (year + 4716)) + (int32_t)(30.6001 * (month + 1)) + day + b - 1524;
    // 调整到VB6的基准(1899-12-30 = 0)
}

// Excel序列号 → 年月日 (Julian Date Number逆运算)
// 基于 vb6_date_to_serial 的逆运算, 含Lotus 1900-02-29 bug兼容
static void vb6_serial_to_date(int32_t serial, int32_t* year, int32_t* month, int32_t* day) {
    // VB6/OLE Automation日期序列号 → 年月日
    // OLE日期: 1899-12-30=0, 1900-01-01=2, 1900-02-28=60, 1900-03-01=61
    // 注意: OLE日期系统不含Lotus 1900-02-29 bug(那是Excel的)
    // serial = JD - 2415019, 因此 JD = serial + 2415019
    int32_t jd = serial + 2415019;
    int32_t a = jd + 32044;
    int32_t b = (4 * a + 3) / 146097;
    int32_t c = a - (146097 * b) / 4;
    int32_t d = (4 * c + 3) / 1461;
    int32_t e = c - (1461 * d) / 4;
    int32_t m = (5 * e + 2) / 153;
    *day = e - (153 * m + 2) / 5 + 1;
    *month = m + 3 - 12 * (m / 10);
    *year = 100 * b + d - 4800 + m / 10;
}

static double vb6_now_serial(void) {
    time_t t = time(NULL);
    struct tm* lt = localtime(&t);
    // OLE日期: serial = JD(date) - JD(1899-12-30), 无Lotus bug
    int32_t jd_date = vb6_date_to_serial(1900 + lt->tm_year, 1 + lt->tm_mon, lt->tm_mday);
    int32_t jd_base = vb6_date_to_serial(1899, 12, 30);
    int32_t datePart = jd_date - jd_base;
    double timePart = (lt->tm_hour * 3600.0 + lt->tm_min * 60.0 + lt->tm_sec) / 86400.0;
    return (double)datePart + timePart;
}

double vb6_Now(void) { return vb6_now_serial(); }
double vb6_Date(void) { return (double)(int32_t)vb6_now_serial(); }
double vb6_Time(void) { double n = vb6_now_serial(); return n - (double)(int32_t)n; }

void vb6_DateSet(BSTR dateStr) {
    if (!dateStr || SysStringLen(dateStr) == 0) return;
    char buf[32] = {0};
    WideCharToMultiByte(CP_ACP, 0, dateStr, -1, buf, 31, NULL, NULL);
    int m = 0, d = 0, y = 0;
    char sep = strchr(buf, '/') ? '/' : '-';
    char *ctx = NULL;
    char buf2[32]; memcpy(buf2, buf, 32);
    char *p1 = strtok_s(buf2, &sep, &ctx);
    char *p2 = p1 ? strtok_s(NULL, &sep, &ctx) : NULL;
    char *p3 = p2 ? strtok_s(NULL, &sep, &ctx) : NULL;
    if (p1 && p2 && p3) {
        m = atoi(p1); d = atoi(p2); y = atoi(p3);
        SYSTEMTIME st; GetLocalTime(&st);
        st.wYear = (WORD)y; st.wMonth = (WORD)m; st.wDay = (WORD)d;
        SetLocalTime(&st);
    }
}

void vb6_TimeSet(BSTR timeStr) {
    if (!timeStr || SysStringLen(timeStr) == 0) return;
    char buf[32] = {0};
    WideCharToMultiByte(CP_ACP, 0, timeStr, -1, buf, 31, NULL, NULL);
    int h = 0, mi = 0, s = 0;
    char *ctx = NULL;
    char buf2[32]; memcpy(buf2, buf, 32);
    char *p1 = strtok_s(buf2, ":", &ctx);
    char *p2 = p1 ? strtok_s(NULL, ":", &ctx) : NULL;
    char *p3 = p2 ? strtok_s(NULL, ":", &ctx) : NULL;
    if (p1) h = atoi(p1);
    if (p2) mi = atoi(p2);
    if (p3) s = atoi(p3);
    SYSTEMTIME st; GetLocalTime(&st);
    st.wHour = (WORD)h; st.wMinute = (WORD)mi; st.wSecond = (WORD)s; st.wMilliseconds = 0;
    SetLocalTime(&st);
}

int32_t vb6_Year(double date) {
    int32_t y, m, d;
    vb6_serial_to_date((int32_t)date, &y, &m, &d);
    return y;
}

int32_t vb6_Month(double date) {
    int32_t y, m, d;
    vb6_serial_to_date((int32_t)date, &y, &m, &d);
    return m;
}

int32_t vb6_Day(double date) {
    int32_t y, m, d;
    vb6_serial_to_date((int32_t)date, &y, &m, &d);
    return d;
}

int32_t vb6_Hour(double time) {
    double t = time - (double)(int32_t)time;
    if (t < 0) t += 1.0;
    return (int32_t)(t * 24.0) % 24;
}

int32_t vb6_Minute(double time) {
    double t = time - (double)(int32_t)time;
    if (t < 0) t += 1.0;
    return (int32_t)(t * 1440.0) % 60;
}

int32_t vb6_Second(double time) {
    double t = time - (double)(int32_t)time;
    if (t < 0) t += 1.0;
    return (int32_t)(t * 86400.0) % 60;
}

// ============================================================
double vb6_DateSerial(int32_t year, int32_t month, int32_t day) {
    int32_t jd_date = vb6_date_to_serial(year, month, day);
    int32_t jd_base = vb6_date_to_serial(1899, 12, 30);
    return (double)(jd_date - jd_base);
}

// P14.2.4: DateAdd/DateDiff/DatePart 日期函数
// ============================================================

// 解析VB6间隔字符串为枚举
typedef enum {
    vb6_di_year, vb6_di_quarter, vb6_di_month, vb6_di_dayofyear,
    vb6_di_day, vb6_di_weekday, vb6_di_week, vb6_di_hour,
    vb6_di_minute, vb6_di_second
} vb6_date_interval;

static vb6_date_interval vb6_parse_interval(BSTR interval) {
    if (!interval || vb6_BSTR_Len(interval) == 0) return vb6_di_day;
    wchar_t ch = interval[0];
    switch (ch) {
        case L'y': case L'Y':
            if (vb6_BSTR_Len(interval) >= 4)
                return vb6_di_year;  // "yyyy" = year
            return vb6_di_dayofyear;  // "y" = DayOfYear
        case L'q': case L'Q': return vb6_di_quarter;
        case L'm': case L'M': return vb6_di_month;
        case L'd': case L'D': return vb6_di_day;
        case L'w': case L'W':
            if (vb6_BSTR_Len(interval) >= 2 && (interval[1] == L'w' || interval[1] == L'W'))
                return vb6_di_week;
            return vb6_di_weekday;
        case L'h': case L'H': return vb6_di_hour;
        case L'n': case L'N': return vb6_di_minute;
        case L's': case L'S': return vb6_di_second;
        default: return vb6_di_day;
    }
}

double vb6_DateAdd(BSTR interval, double number, double date) {
    vb6_date_interval di = vb6_parse_interval(interval);
    int32_t datePart = (int32_t)date;
    double timePart = date - (double)datePart;
    if (timePart < 0) { timePart += 1.0; datePart--; }
    
    int32_t y, m, d;
    vb6_serial_to_date(datePart, &y, &m, &d);
    int32_t hh = (int32_t)(timePart * 24.0) % 24;
    int32_t mm = (int32_t)(timePart * 1440.0) % 60;
    int32_t ss = (int32_t)(timePart * 86400.0) % 60;
    
    int32_t n = (int32_t)number;
    
    switch (di) {
        case vb6_di_year:
            y += n;
            break;
        case vb6_di_quarter:
            m += n * 3;
            while (m > 12) { m -= 12; y++; }
            while (m < 1)  { m += 12; y--; }
            break;
        case vb6_di_month:
            m += n;
            while (m > 12) { m -= 12; y++; }
            while (m < 1)  { m += 12; y--; }
            break;
        case vb6_di_day:
        case vb6_di_dayofyear:
        case vb6_di_weekday:
            datePart += n;
            goto rebuild;
        case vb6_di_week:
            datePart += n * 7;
            goto rebuild;
        case vb6_di_hour:
            hh += n;
            while (hh >= 24) { hh -= 24; datePart++; }
            while (hh < 0)   { hh += 24; datePart--; }
            goto rebuild;
        case vb6_di_minute:
            mm += n;
            while (mm >= 60) { mm -= 60; hh++; }
            while (mm < 0)   { mm += 60; hh--; }
            while (hh >= 24) { hh -= 24; datePart++; }
            while (hh < 0)   { hh += 24; datePart--; }
            goto rebuild;
        case vb6_di_second:
            ss += n;
            while (ss >= 60) { ss -= 60; mm++; }
            while (ss < 0)   { ss += 60; mm--; }
            while (mm >= 60) { mm -= 60; hh++; }
            while (mm < 0)   { mm += 60; hh--; }
            while (hh >= 24) { hh -= 24; datePart++; }
            while (hh < 0)   { hh += 24; datePart--; }
            goto rebuild;
    }
    // For year/quarter/month: clamp day to valid range for the new month
    {
        static const int32_t daysInMonth[] = {0,31,28,31,30,31,30,31,31,30,31,30,31};
        int32_t maxDay = daysInMonth[m];
        if (m == 2 && (y % 4 == 0 && (y % 100 != 0 || y % 400 == 0))) maxDay = 29;
        if (d > maxDay) d = maxDay;
        // Rebuild serial
        int32_t jd_date = vb6_date_to_serial(y, m, d);
        int32_t jd_base = vb6_date_to_serial(1899, 12, 30);
        datePart = jd_date - jd_base;
    }
rebuild:
    {
        double result = (double)datePart + (hh * 3600.0 + mm * 60.0 + ss) / 86400.0;
        return result;
    }
}

int64_t vb6_DateDiff(BSTR interval, double date1, double date2, int32_t firstDayOfWeek, int32_t firstWeekOfYear) {
    (void)firstDayOfWeek;
    (void)firstWeekOfYear;
    vb6_date_interval di = vb6_parse_interval(interval);
    
    int32_t d1 = (int32_t)date1, d2 = (int32_t)date2;
    
    switch (di) {
        case vb6_di_year: {
            int32_t y1, m1, dd1, y2, m2, dd2;
            vb6_serial_to_date(d1, &y1, &m1, &dd1);
            vb6_serial_to_date(d2, &y2, &m2, &dd2);
            return (int64_t)(y2 - y1);
        }
        case vb6_di_quarter: {
            int32_t y1, m1, dd1, y2, m2, dd2;
            vb6_serial_to_date(d1, &y1, &m1, &dd1);
            vb6_serial_to_date(d2, &y2, &m2, &dd2);
            return (int64_t)((y2 * 4 + (m2 - 1) / 3) - (y1 * 4 + (m1 - 1) / 3));
        }
        case vb6_di_month: {
            int32_t y1, m1, dd1, y2, m2, dd2;
            vb6_serial_to_date(d1, &y1, &m1, &dd1);
            vb6_serial_to_date(d2, &y2, &m2, &dd2);
            return (int64_t)((y2 * 12 + m2) - (y1 * 12 + m1));
        }
        case vb6_di_day:
        case vb6_di_dayofyear:
            return (int64_t)(d2 - d1);
        case vb6_di_weekday:
            return (int64_t)(d2 - d1);
        case vb6_di_week:
            return (int64_t)((d2 - d1) / 7);
        case vb6_di_hour:
            return (int64_t)((date2 - date1) * 24.0);
        case vb6_di_minute:
            return (int64_t)((date2 - date1) * 1440.0);
        case vb6_di_second:
            return (int64_t)((date2 - date1) * 86400.0);
    }
    return 0;
}

int32_t vb6_DatePart(BSTR interval, double date, int32_t firstDayOfWeek, int32_t firstWeekOfYear) {
    (void)firstDayOfWeek;
    (void)firstWeekOfYear;
    vb6_date_interval di = vb6_parse_interval(interval);
    int32_t d = (int32_t)date;
    
    switch (di) {
        case vb6_di_year: {
            int32_t y, m, dd;
            vb6_serial_to_date(d, &y, &m, &dd);
            return y;
        }
        case vb6_di_quarter: {
            int32_t y, m, dd;
            vb6_serial_to_date(d, &y, &m, &dd);
            return (m - 1) / 3 + 1;
        }
        case vb6_di_month: {
            int32_t y, m, dd;
            vb6_serial_to_date(d, &y, &m, &dd);
            return m;
        }
        case vb6_di_day:
        case vb6_di_dayofyear: {
            int32_t y, m, dd;
            vb6_serial_to_date(d, &y, &m, &dd);
            // Day of year: simple approximation (days since start of year)
            int32_t jd_this = vb6_date_to_serial(y, m, dd);
            int32_t jd_jan1 = vb6_date_to_serial(y, 1, 1);
            return jd_this - jd_jan1 + 1;
        }
        case vb6_di_weekday: {
            // Sunday=1, Monday=2, ..., Saturday=7
            // 1899-12-30 (serial 0) was a Saturday (day 7)
            int32_t wd = (d % 7 + 7) % 7;  // 0=Sat,1=Sun,...,6=Fri
            int32_t weekday = wd + 1;        // Sat=1,Sun=2,...,Fri=7
            // Remap: Sun=1,Mon=2,...,Sat=7
            // wd: 0=Sat,1=Sun,2=Mon,3=Tue,4=Wed,5=Thu,6=Fri
            if (wd == 0) return 7;  // Saturday
            return wd;              // Sunday=1, Monday=2, etc.
        }
        case vb6_di_week: {
            // ISO 8601 simplified: week 1 contains Jan 4
            int32_t y, m, dd;
            vb6_serial_to_date(d, &y, &m, &dd);
            int32_t jd_this = vb6_date_to_serial(y, m, dd);
            int32_t jd_jan1 = vb6_date_to_serial(y, 1, 1);
            int32_t dayOfYear = jd_this - jd_jan1 + 1;
            return (dayOfYear - 1) / 7 + 1;
        }
        case vb6_di_hour:
            return vb6_Hour(date);
        case vb6_di_minute:
            return vb6_Minute(date);
        case vb6_di_second:
            return vb6_Second(date);
    }
    return 0;
}
// ============================================================

// ============================================================
// P21-B: Weekday/DateValue/TimeSerial/TimeValue/Array
// ============================================================

int32_t vb6_Weekday(double date, int32_t firstDayOfWeek) {
    /* VB6 Weekday(date[, firstDayOfWeek])
       Returns the day of the week as an integer.
       firstDayOfWeek: 1=vbSunday(default), 2=vbMonday, ..., 7=vbSaturday
       Return: 1-based relative to firstDayOfWeek.
       For default (vbSunday=1): Sunday=1, Monday=2, ..., Saturday=7 */
    int32_t d = (int32_t)date;
    /* 1899-12-30 (serial 0) was a Saturday */
    int32_t wd = (d % 7 + 7) % 7;  /* 0=Sat,1=Sun,2=Mon,3=Tue,4=Wed,5=Thu,6=Fri */
    /* Convert to Sunday=1,Monday=2,...,Saturday=7 */
    int32_t weekday;
    if (wd == 0) weekday = 7;  /* Saturday */
    else weekday = wd;         /* Sunday=1,Monday=2,... */
    
    /* Adjust for firstDayOfWeek */
    if (firstDayOfWeek <= 0) firstDayOfWeek = 1;  /* default vbSunday */
    if (firstDayOfWeek > 7) firstDayOfWeek = 1;
    if (firstDayOfWeek == 1) return weekday;  /* vbSunday: no shift needed */
    /* Shift: e.g. vbMonday(2) -> Monday=1,...,Sunday=7 */
    int32_t shifted = weekday - (firstDayOfWeek - 1);
    if (shifted < 1) shifted += 7;
    return shifted;
}

double vb6_DateValue(BSTR dateStr) {
    /* VB6 DateValue(datestring) -> date serial (double)
       Parses a date string using Windows OLE Automation */
    if (!dateStr) return 0.0;
    DATE result = 0.0;
    if (SUCCEEDED(VarDateFromStr(dateStr, LOCALE_USER_DEFAULT, 0, &result))) {
        return (double)(int32_t)result;
    }
    return 0.0;
}

double vb6_TimeSerial(int32_t hour, int32_t minute, int32_t second) {
    /* VB6 TimeSerial(hour, minute, second) -> date serial (double)
       The date part is 0 (Dec 30, 1899), time part is the fraction of day */
    int32_t totalSeconds = hour * 3600 + minute * 60 + second;
    while (totalSeconds < 0) totalSeconds += 86400;
    int32_t extraDays = totalSeconds / 86400;
    int32_t timeSeconds = totalSeconds % 86400;
    return (double)extraDays + (double)timeSeconds / 86400.0;
}

double vb6_TimeValue(BSTR timeStr) {
    /* VB6 TimeValue(timestring) -> date serial (double)
       Parses a time string using Windows OLE Automation */
    if (!timeStr) return 0.0;
    DATE result = 0.0;
    if (SUCCEEDED(VarDateFromStr(timeStr, LOCALE_USER_DEFAULT, 0, &result))) {
        double datePart = (double)(int32_t)result;
        double timePart = result - datePart;
        if (timePart < 0) timePart += 1.0;
        return timePart;
    }
    return 0.0;
}

vb6_SafeArray1D* vb6_ArrayCreate(int32_t count) {
    /* VB6 Array(arglist) helper: creates a Variant SafeArray with count elements.
       Caller (cgen) sets each element directly via VB6_SA_AT. */
    // Fix <vbeclipse>: `Array()` 零实参在真 VB6 里是**空数组**（用户实测）：
    // UBound = -1、LBound = 0 ⇒ `UBound-LBound+1` 读出 0、For Each 零次。
    // 旧写法把 uBound 夹到 0，于是空数组被当成 1 个元素（多出一个幻影空槽）。
    if (count < 0) count = 0;
    vb6_SafeArray1D* arr = vb6_SafeArrayCreate1D(vb6_sa_variant, 0, count - 1);
    return arr;
}

void vb6_ArraySetLong(vb6_SafeArray1D* arr, int32_t index, int32_t val) {
    if (!arr || index < arr->lBound || index > arr->uBound) return;
    vb6_VARIANT* v = &VB6_SA_AT(vb6_VARIANT, arr, index);
    vb6_VariantClear(v);
    v->vt = vb6_vtLong;
    v->lVal = val;
}

void vb6_ArraySetDouble(vb6_SafeArray1D* arr, int32_t index, double val) {
    if (!arr || index < arr->lBound || index > arr->uBound) return;
    vb6_VARIANT* v = &VB6_SA_AT(vb6_VARIANT, arr, index);
    vb6_VariantClear(v);
    v->vt = vb6_vtDouble;
    v->dblVal = val;
}

void vb6_ArraySetBSTR(vb6_SafeArray1D* arr, int32_t index, BSTR val) {
    if (!arr || index < arr->lBound || index > arr->uBound) return;
    vb6_VARIANT* v = &VB6_SA_AT(vb6_VARIANT, arr, index);
    vb6_VariantClear(v);
    v->vt = vb6_vtBSTR;
    v->bstrVal = vb6_BSTR_FromBSTR(val);
}

void vb6_ArraySetVariant(vb6_SafeArray1D* arr, int32_t index, vb6_VARIANT val) {
    if (!arr || index < arr->lBound || index > arr->uBound) return;
    vb6_VARIANT* v = &VB6_SA_AT(vb6_VARIANT, arr, index);
    vb6_VariantClear(v);
    vb6_VariantCopy(v, &val);
}

