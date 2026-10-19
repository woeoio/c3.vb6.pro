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

// 「一枚年月日折成 VB/OLE 序列号」这一问在这里收成一个出口。
// vb6_date_to_serial 答的是儒略日数，要再减掉基准 1899-12-30 (= 2415019) 才是 VB 的序列号；
// 这个减法原先在本文件抄了四遍，把儒略日数当成序列号直接用会把每个日期整体抬高 2415019
// (账 #422 的第一发读数就是这样：DatePart("ww", …) 答 -2415012 那一档)。
static int32_t vb6_SerialOf(int32_t year, int32_t month, int32_t day) {
    return vb6_date_to_serial(year, month, day) - vb6_date_to_serial(1899, 12, 30);
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
    int32_t datePart = vb6_SerialOf(1900 + lt->tm_year, 1 + lt->tm_mon, lt->tm_mday);
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

// 账 #419 (§B193): 「这一枚 DATE 的双半边怎么切」的唯一出口。DATE 是「整天数 + 一天的分数」，
// 而分数在二进制里落不整 (14:30 = 0.6041666666666666)，所以拆之前必须**四舍五入到整秒** ——
// 直接截断会把这一格少掉一秒的几分之一，实测口径 (本机 VB-family 运行期，尺
// .build/k272_timeparts.vbs + k277_round.vbs + k278_roll.vbs):
//   Minute(CDate("2026/1/1 14:30")) = 30 (截断那一条答 29)；+0.4/+0.5 秒不进位、+0.6/+0.99 进位。
//
// 账 #440 §B209 (第一百零七刀) = 把「一枚序列号拆成日期 + 日内秒」这一问**收成一位主人**，
// 两处各答一半的那笔账一次结掉:
//   ① §B194 (#420): 一天末尾那一格只把日内秒归零、**日期那一半不进位**。运行期两半**一起**进:
//      raw = 46023.9999988426 (= 2026/1/1 23:59:59.9) → H/M/S = 0/0/0 且 **Day = 2**
//      (尺 .build/k420_oracle.vbs + k420b_grid.vbs)。改前答 Day = 1。
//   ② §B195 (#421): epoch 之前的零头取法与运行期**反号**。运行期取**绝对值**, 本仓原先 `+1.0`:
//      raw = -657433.395833333 → 运行期 9:30, 本仓 14:30 (同一瞬间的两种拆法)。
//   两半**符号对称**: -46023.9999988426 → 运行期 1773/12/28 (日期也进了一格) ⇒ 进位不是只给正数。
//   ⚠ 日期那一半用 **(int32_t)date 向零截断**(不是 floor), 运行期就是这一套: 负数那几格
//     Date=2/M=1/Y=100 与本仓改前就一致 —— 反号的只有**零头**那一半, 日期那一半一直是对的。
static void vb6_split_serial(double date, int32_t* dayOut, int32_t* sodOut) {
    double t = date - (double)(int32_t)date;
    if (t < 0) t = -t;                              // ② §B195: 取绝对值, 不再 +1.0
    int32_t sod = (int32_t)(t * 86400.0 + 0.5);
    int32_t day = (int32_t)date;
    if (sod >= 86400) { sod = 0; day += 1; }         // ① §B194: 日期那一半跟着进
    *dayOut = day;
    *sodOut = sod;
}

// 账 #440 §B209: 「日期那一格」这一问的出口 —— **转问上面那一位主人**, 不自己再抄一遍那三行。
// 改前它与 vb6_SecondsOfDay 各答一半: 时间那三枚走 `+1.0` 归一、日期那三枚走裸 `(int32_t)date`
// ⇒ 一天末尾那一格两半**不一致**(时间已归零、日期还没进)。
static inline int32_t vb6_serial_day_half(double date) {
    int32_t day, sod;
    vb6_split_serial(date, &day, &sod);
    return day;
}

int32_t vb6_Year(double date) {
    int32_t y, m, d;
    vb6_serial_to_date(vb6_serial_day_half(date), &y, &m, &d);
    return y;
}

int32_t vb6_Month(double date) {
    int32_t y, m, d;
    vb6_serial_to_date(vb6_serial_day_half(date), &y, &m, &d);
    return m;
}

int32_t vb6_Day(double date) {
    int32_t y, m, d;
    vb6_serial_to_date(vb6_serial_day_half(date), &y, &m, &d);
    return d;
}

// 一条出口，读者 = Hour / Minute / Second 三枚 getter 与 DateAdd 分解那一站 (DatePart 的 h/n/s
// 三档本来就转调这三枚 getter，跟着一起好)。
static int32_t vb6_SecondsOfDay(double date) {
    int32_t day, sod;
    vb6_split_serial(date, &day, &sod);
    (void)day;                                       // 日期那一半归 Year/Month/Day 那一族用
    return sod;
}

int32_t vb6_Hour(double time) {
    return vb6_SecondsOfDay(time) / 3600;
}

int32_t vb6_Minute(double time) {
    return (vb6_SecondsOfDay(time) / 60) % 60;
}

int32_t vb6_Second(double time) {
    return vb6_SecondsOfDay(time) % 60;
}

// ============================================================
double vb6_DateSerial(int32_t year, int32_t month, int32_t day) {
    return (double)vb6_SerialOf(year, month, day);
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
    // 账 #419 (§B193): 这三行原本各抄一份截断 —— DateAdd 是「改完之后还要写回去」的那一站，
    // 截断掉的那一秒会**永久留在结果里** (DateAdd("d", 1, #2026/1/1 14:30#) 少一分钟)。
    int32_t sod = vb6_SecondsOfDay(date);
    int32_t hh = sod / 3600;
    int32_t mm = (sod / 60) % 60;
    int32_t ss = sod % 60;
    
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
        datePart = vb6_SerialOf(y, m, d);
    }
rebuild:
    {
        double result = (double)datePart + (hh * 3600.0 + mm * 60.0 + ss) / 86400.0;
        return result;
    }
}

int64_t vb6_DateDiff(BSTR interval, double date1, double date2, int32_t firstDayOfWeek, int32_t firstWeekOfYear) {
      (void)firstWeekOfYear;
      vb6_date_interval di = vb6_parse_interval(interval);

      // 账 #423 §B197 (第一百零八刀) = 「w / ww 这两个间隔在 **DateDiff** 里各答什么」这一问。
      // 病根 = `vb6_parse_interval` 是 **DateAdd / DateDiff / DatePart 三家共用**的一问，而
      // `w` 与 `ww` 的含义**按 builtin 分岔**:
      //     DatePart("w",  …) = 星期序号(与周首无关)     DateDiff("w",  …) = **整段七天**
      //     DatePart("ww", …) = 一年里第几周             DateDiff("ww", …) = 半开区间 (a,b] 里
      //                                                    落在第 f 天那一格的条数(方向带符号)
      // 共用一枚枚举 ⇒ 改 `parse_interval` 会把另外两家的正确答案一起弄坏(实测 DatePart 的
      // w/ww 今天都是 2、都对)。所以**枚举那一问不动**，DateDiff 自己认这一格。
      //
      // 运行期口径 (cscript 同机代理, 尺 .build/k423_oracle.vbs, 逐格量死):
      //   ① "ww" = floor((b-f)/7) - floor((a-f)/7)，f 参与 ⇒ 改周首**会动**这一格;
      //      拟合面积 162400 格对这条公式零失配(账上 §B197 记的那把尺)。
      //   ② "w"  = trunc((b-a)/7)，**与 f 无关** ⇒ 改周首**不许动**这一格 (81200 格零失配)。
      //      ⚠ 改前本仓这一格发的是 `(d2-d1)/7`，而 d1/d2 是**截断过的整日**、且落进
      //        `case vb6_di_weekday` 的 `d2 - d1` ⇒ 答的是**天数**(1/5→1/12 答 7、1/5→2/5 答 31)。
      // ⚠ 这一格用 `int` 而不是 `bool`: 用户 TU 那一遍编译把 `bool` 宏成了 `int`,
      //   `const bool` 会展开成 `const int bool` = 两个类型词 ⇒ cl C2086/C2146(本刀撞过)。
      const int ivIsWw423 =
          interval && vb6_BSTR_Len(interval) >= 1 && (interval[0] == L'w' || interval[0] == L'W')
          && vb6_BSTR_Len(interval) >= 2 && (interval[1] == L'w' || interval[1] == L'W');
      const int ivIsW423 =
          interval && vb6_BSTR_Len(interval) >= 1 && (interval[0] == L'w' || interval[0] == L'W')
          && !(vb6_BSTR_Len(interval) >= 2 && (interval[1] == L'w' || interval[1] == L'W'));

      int32_t d1 = (int32_t)date1, d2 = (int32_t)date2;

      // ② "w": 整段七天，trunc 向零，与 firstDayOfWeek **无关** —— 先答，两条形参都不看。
      if (ivIsW423) {
          return (int64_t)((date2 - date1) / 7.0);
      }
      // ① "ww": 半开区间 (a,b] 里落在第 f 天那一格的条数。f = 1 (vbSunday) 是本仓序列号口径。
      if (ivIsWw423) {
          const double f = (double)firstDayOfWeek;
          return (int64_t)(floor((date2 - f) / 7.0) - floor((date1 - f) / 7.0));
      }

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

// 账 #422 (§B196): 「这一年里第几周」这一问的主人 (原先这问在 vb6_DatePart 里就地抄了一条公式)。
// 账 #427 (§B201) 第一百刀: 第二个读者到了 —— Format 的「ww」档也问它, 所以这一问不再住 date.c 内部
// (定义不再 static, 声明在 vb6rtl_builtin.h)。星期序号那一问同样收成一条: DatePart 的「w」臂改走 vb6_Weekday。
// 对齐周由 firstDayOfWeek 定死: b = 年内第一个该星期几之前已有的天数 (0..6)，也就是开头那条
// 不满整周的长度。firstWeekOfYear 只决定「哪条周算第 1 周」和年两端那几条不满整周的归谁:
//   1 = 开头那条不满的周就是第 1 周;
//   2 = 它满 4 天才算第 1 周，不满就改从第一个 fdow 起算，开头那几天交回上一年的最后一周;
//   3 = 第一条满整周才算第 1 周。
// 年末那一半只有档 2 会被来年抢走 (抢走的那些天答 1)，而且**来年正好是闰年时，抢走的起点要往后挪
// 一天** —— 这一条同样是实测读数 (2023/2027 两年各三格)，别当笔误删。
// 口径 = 本机 VB-family 运行期 14 年 × 21 组合 = 107394 天逐日拟合，零失配 (尺
// .build/k321_ww_grid.vbs + k324_ww_grid2.vbs + k330_ww_grid3.vbs，拟合与复核 .build/k323_ww_fit.py，
// 读数报告 .build/k329_fit_report.txt + k330_fit_report.txt)。
static int32_t vb6_LeadingPartialDays(int32_t year, int32_t fdow) {
    int32_t wd = vb6_Weekday((double)vb6_SerialOf(year, 1, 1), 1);  // 1=Sun..7=Sat，与 fdow 同一套编号
    int32_t b = (fdow - wd) % 7;
    return b < 0 ? b + 7 : b;
}

int32_t vb6_WeekOfYear(int32_t serial, int32_t fdow, int32_t fwoy) {
    int32_t y, mo, dd;
    vb6_serial_to_date(serial, &y, &mo, &dd);
    if (fdow < 1 || fdow > 7) fdow = 1;   // 0 = vbGeneral
        //
        // 账 #424 §B198 (第一百零九刀) —— 这一格的口径由**用户拍板**，不是实现默默选的：
        //   `firstDayOfWeek = 0` (vbGeneral) 在运行期是**交系统**的（本机 = 周一），
        //   任何「跟随系统」的发法都**跨机器不可复现**。本仓两条 DatePart/DateDiff 路
        //   统一 clamp 到 **vbSunday(1)**，作为一条**显式的已知偏离**：与运行期在
        //   「系统周首 ≠ 周日」的机器上不同，但换来的是读数跨机器稳定 —— 这正是本仓
        //   整套 emit-manifest / 夹具逐字节对账赖以成立的前提。
        //   ⚠ 本机量不出这一格的差别（尺 .build/k424_oracle.vbs：DatePart("ww") 的
        //     f=0/1/2/7 同答 2、DateDiff("ww") 的 f=0/1/7 同答 8）⇒ 它是一处
        //     **有意保留**的偏离，不是漏改；要跟着系统走，得先把可复现性那套一起想清楚。
    //
    // 同一笔账的另一半（firstWeekOfYear）：运行期对 >= 4 **直接报错**（cscript 实测 fwy=4/5
    //   两格当场抛，尺 .build/k424_oracle.vbs）；改前这里静默按 1 档答 2。口径同样由用户拍板
    //   = **对齐运行期报 err 5**，与 #436「越界 namedFormat 报 err 5」同一口径。
    if (fwoy >= 4) {
        vb6_ErrRaiseNumber(5);
        return 0;
    }
    if (fwoy != 2 && fwoy != 3) fwoy = 1;  // 实测 0 与 1 同档
    int32_t doy = serial - vb6_SerialOf(y, 1, 1) + 1;
    int32_t n = vb6_SerialOf(y, 12, 31) - vb6_SerialOf(y, 1, 1) + 1;
    int32_t b = vb6_LeadingPartialDays(y, fdow);
    if (fwoy == 2) {
        int32_t bn = vb6_LeadingPartialDays(y + 1, fdow);
        if (bn >= 4) {
            int32_t ny = y + 1;
            int32_t leap_next = (ny % 4 == 0 && ny % 100 != 0) || (ny % 400 == 0);
            if (doy > n - (7 - bn) + (leap_next ? 1 : 0)) return 1;
        }
    }
    if (fwoy == 1 || (fwoy == 2 && b >= 4)) {
        if (b == 0) return 1 + (doy - 1) / 7;
        if (doy <= b) return 1;
        return 2 + (doy - b - 1) / 7;
    }
    if (doy <= b) return vb6_WeekOfYear(vb6_SerialOf(y - 1, 12, 31), fdow, fwoy);
    return 1 + (doy - b - 1) / 7;
}

int32_t vb6_DatePart(BSTR interval, double date, int32_t firstDayOfWeek, int32_t firstWeekOfYear) {
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
        case vb6_di_day: {
            int32_t y, m, dd;
            vb6_serial_to_date(d, &y, &m, &dd);
            return dd;              // 「d」= 月内第几天，与「一年里第几天」是两问 (账 #422)
        }
        case vb6_di_dayofyear: {
            int32_t y, m, dd;
            vb6_serial_to_date(d, &y, &m, &dd);
            return vb6_SerialOf(y, m, dd) - vb6_SerialOf(y, 1, 1) + 1;
        }
        case vb6_di_weekday:
            // 账 #427 (§B201): 这一问的主人早就是 vb6_Weekday (Sunday=1..Saturday=7), 这里原先是第二份抄写。
            // DatePart 的两条形参对「w」不参与 (与运行期同口径: 星期序号与周首无关), 所以固定问 vbSunday 那一档。
            return vb6_Weekday((double)d, 1);
        case vb6_di_week:
            return vb6_WeekOfYear(d, firstDayOfWeek, firstWeekOfYear);
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

// 账 #417 (§B191): 「这串文本是哪一天」这一问原先有三家各自抄写 (vb6_DateValue、
// vb6_TimeValue、vb6rtl_conv.c 的 CStr(Date 形文本) 入口)，「解析不出算几号」那句又各自
// 沉默。解析政策继续交给 oleaut32 (§B142 同口径)，这里只留一条 VarDateFromStr 站点；
// 「认不出是不是错」不再由这一层答，交给上面两位主人各按自己的档定 (VB6 实测两档不同)。
// 参数位次同 vb6_NumVal：slot2 是 LCID、slot3 是 dwFlags (尺 .build/k207_olearg.c)。
static int32_t vb6_TryDateFromStr(BSTR s, DATE* out) {
    return (s && SUCCEEDED(VarDateFromStr(s, LOCALE_USER_DEFAULT, 0, out))) ? -1 : 0;
}

// DateValue / TimeValue / CStr(日期形文本) 那一档的主人：只认日期时间写法，
// 落不到报 13。VB6 实测 (尺 .build/k236_datefail.vbs) DateValue("abc")、DateValue("46000")
// 全是 13 —— 纯数字串不收，那是 CDate 独有的兜底 (见 vb6_DateVal)。
double vb6_DateParse(BSTR s) {
    DATE d = 0.0;
    if (vb6_TryDateFromStr(s, &d)) return d;
    vb6_RaiseError(13, vb6_BSTR_FromStr(L"Type mismatch"));
    return 0.0;
}

// CDate 那一档 = 日期写法 ∪ 严格数值，且**保留时间小数**。VB6 实测 (尺 .build/k234_cdate.vbs
// 与 k235_datescan.vbs)：CDate("2026/1/1 14:30") = 46023.60416667、CDate("46000") = 46000、
// CDate("&H1") = 1、CDate("1e2") = 100，而 CDate("123abc") 与 CDate("abc") 都是 13 ——
// 兜底不是 Val 那种吞尾巴的宽松档，正是账 #416 收成一位主人的 vb6_NumVal。
double vb6_DateVal(BSTR s) {
    DATE d = 0.0;
    if (vb6_TryDateFromStr(s, &d)) return d;
    return vb6_NumVal(s);
}

// 账 #418 (§B192): IsDate 那一问要的是「这串文本像不像日期」= 上面那条站点的**成功本身**，
// 不是「折出来是几号」更不是「折不出报几号」。实测 (尺 .build/k253_isfam.vbs) 运行期
// IsDate("2026/1/1") 真、IsDate("46000") 假 —— 收下集合与 vb6_TryDateFromStr 一字不差，
// 所以这一档不许再问第二遍 oleaut32，也不许报错。
int32_t vb6_DateCanParse(BSTR s) {
    DATE d = 0.0;
    return vb6_TryDateFromStr(s, &d);
}

double vb6_DateValue(BSTR dateStr) {
    /* VB6 DateValue(datestring) -> 只留日期那一位。实测 DateValue("2026/1/1 14:30") = 46023、
       DateValue("12:00") = 0；认不出由 vb6_DateParse 报 13，不再静默 0。 */
    return floor(vb6_DateParse(dateStr));
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
    /* VB6 TimeValue(timestring) -> 只留一天里的那一小段。实测 TimeValue("2026/1/1 14:30") =
       0.60416667、TimeValue("2026/1/1") = 0；小数恒在 [0,1)，负序列 (1899-12-30 之前) 也一样。
       认不出由 vb6_DateParse 报 13，不再静默 0。 */
    double d = vb6_DateParse(timeStr);
    return d - floor(d);
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

