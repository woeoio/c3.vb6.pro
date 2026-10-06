// vb6rtl_conv.c - VB6 运行时库: 转换家族：MsgBox/数值函数/类型转换/类型检查/随机数/Debug 对象
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 894~902 行
//   原第 903~918 行
//   原第 919~953 行
//   原第 1272~1327 行
//   原第 1328~1439 行
//   原第 2680~2707 行
//   原第 3149~3182 行

#include "vb6rtl.h"
#include "vb6forms.h"   /* Fix 112: 宿主对象模型 (窗体/控件/集合/字体) */
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
// COM互操作前向声明 (实现在vb6com.c中，避免vb6_VARIANT类型冲突)
// ============================================================
extern void* vb6_CreateObject(const wchar_t* progId);
extern void* vb6_NewBuiltinObject(const wchar_t* className);
extern void* vb6_GetObject(const wchar_t* pathName, const wchar_t* progId);
extern int32_t vb6_IsNothing(void* obj);
extern void vb6_ReleaseObject(void** objPtr);
extern void* vb6_ComCall(void* disp, const wchar_t* methodName, void* args, int32_t argc);
extern void* vb6_ComCallByDispid(void* disp, int32_t dispid, void* args, int32_t argc);
extern void* vb6_ComGetProp(void* disp, const wchar_t* propName);
extern void vb6_ComSetProp(void* disp, const wchar_t* propName, void* value);
extern void vb6_ComSetRef(void* disp, const wchar_t* propName, void* objRef);
extern void vb6_ComInit(void);
extern void vb6_ComExit(void);

// ============================================================
// MsgBox
// ============================================================

int32_t vb6_MsgBox(BSTR prompt, int32_t buttons, BSTR title) {
    /* Win32 MessageBox */
    return (int32_t)MessageBoxW(NULL, prompt ? prompt : L"", title ? title : L"", (UINT)buttons);
}

// ============================================================
// 数值函数
// ============================================================

double vb6_Abs(double x) { return fabs(x); }
int32_t vb6_Sgn(double x) { return (x > 0) ? 1 : (x < 0) ? -1 : 0; }
double vb6_Sqr(double x) { return sqrt(x); }
double vb6_Round(double x, int32_t decimals) {
    double factor = pow(10.0, (double)decimals);
    return round(x * factor) / factor;
}

// Fix 117: VB6 默认「数值 → 字符串」(CStr(x) / Format(x, "") 无格式串) 语义。
// VB6 打印的是**最短且能往返解析**的十进制表示, 且 Single 最多 7 位有效数字:
//   CStr(21.1!)  = "21.1"      (而非 "21.1000003814697")
//   CStr(32.7!)  = "32.7"
//   CStr(0.1#)   = "0.1"
// 此前一律用 "%.15g" 打印 double, 于是 Single 的二进制误差被完整暴露 ——
// Charts 2020 的 ucPieChart 标签 ("{P}%" ← Percent As Single) 因此显示成
// 21.100000381 / 32.700000763。
BSTR vb6_NumToBSTRDefault(double v, int isSingle) {
    if (v != v) return vb6_BSTR_FromStr(L"");   // NaN → 空串 (VB6 不会出现 NaN)
    // Fix 117d: Double 若恰好是 Single 精度值 (由 Single 运算/赋值/Round 提升而来),
    // 按 VB6 的 Single 规则打印 7 位有效数字。否则 Round(v!, 1)/Single 除法等会把
    // float 二进制噪声整个暴露出来 (如 "10.3999996185"、"36.9000015")。
    if (!isSingle && v == (double)(float)v) isSingle = 1;
    int maxPrec = isSingle ? 7 : 15;
    char nbuf[64] = "0";
    double av = fabs(v);
    // VB6 只在极大/极小时用指数记法; 其余一律定点 (CStr(30!) = "30" 而非 "3e+01")
    int useExp = (av >= 1e15) || (av != 0.0 && av < 1e-4);
    for (int prec = 1; prec <= maxPrec; prec++) {
        snprintf(nbuf, sizeof(nbuf), "%.*g", prec, v);
        if (!useExp && (strchr(nbuf, 'e') || strchr(nbuf, 'E'))) {
            // "%.*g" 选了指数 → 换算成对应的定点小数位再试
            int exp10 = (av > 0.0) ? (int)floor(log10(av)) : 0;
            int dec = prec - 1 - exp10;
            if (dec < 0) dec = 0;
            if (dec > 20) dec = 20;
            snprintf(nbuf, sizeof(nbuf), "%.*f", dec, v);
        }
        double back = strtod(nbuf, NULL);
        if (isSingle) { if ((float)back == (float)v) break; }
        else          { if (back == v) break; }
    }
    wchar_t wbuf[64];
    MultiByteToWideChar(CP_ACP, 0, nbuf, -1, wbuf, 64);
    return vb6_BSTR_FromStr(wbuf);
}
// Fix 135: VB6 Rnd 的真实算法 —— 24-bit LCG (与 VBA 一致):
//   state = (state * 1140671485 + 12820163) Mod 2^24
//   Rnd   = state / 2^24
// 初值 (未 Randomize) = 0x50000。序列与 VB6 逐值一致, 图表 demo 的数据
// (Charts 2020 各控件 Random(Min,Max)) 因此与 VB6 参考图完全相同。
// 此前用 C rand(): 序列不同 → 柱高/饼图占比/TreeMaps 布局全都对不上参考图。
static int32_t g_rndState135 = 0x50000;

float vb6_Rnd(int32_t seed) {
    if (seed < 0) {
        g_rndState135 = (int32_t)(seed & 0xFFFFFF);   /* 负参: 重播种 */
    } else if (seed == 0) {
        return (float)g_rndState135 / 16777216.0f;    /* 0: 重复上一个 */
    }
    g_rndState135 = (int32_t)(((int64_t)g_rndState135 * 1140671485LL + 12820163LL) & 0xFFFFFF);
    return (float)g_rndState135 / 16777216.0f;
}

// ============================================================
// 转换函数
// ============================================================

// ============================================================
// 溢出检查 (Error 6)
// ============================================================
// VB6 对窄整型 (Byte/Integer/Long) 的**收窄赋值**是查表的: 值落在目标范围外就抛
// run-time error 6 "Overflow", 由 `On Error Resume Next` / `On Error GoTo` 正常捕获。
// 此前这一族是纯 C 截断 (`uint8_t a = 0; a = (-1);` ⇒ a=255), 即 ai/009 §5.10
// 记的那条 P3 缺口 —— 用户写 `a = -1` 得到 255 而不是报错, 属静默错编。
// 范围值取自 VB6 手册: Byte 0~255 / Integer -32768~32767 / Long -2147483648~2147483647。
//
// 口径说明 (与 VB6 的两处差别, 都是**已知取舍**, 不是漏):
//   (1) 取整仍沿用既有的 round() (半值远离零), 不改 —— CInt 的银行家舍入是另一条账,
//       改它会动到全部既有读数。
//   (2) `On Error Resume Next` 下 vb6_RaiseError 提前返回, 此时本函数返回**被截断**的
//       值并真的存进目标; VB6 那边是**放弃这次赋值、目标保持原值**。表达式的形状
//       (而不是语句块) 决定了这个差别: 同一个 vb6_ChkXxx 要同时服务赋值 / 函数返回 /
//       For 界 / 数组下标, 没法在那些位置发多语句块。相对"完全不检查", 这已经
//       收窄到只差 Resume Next 下的存值。
// Fix <vbeclipse> rev37: **撤净 P48 探针** (定位 Charts2020 x86 的 error 6)。
//   探针结论: 越界值恒为 `v=-2147483648 lo=0 hi=255`。lo/hi 唯一来自 vb6_CByte
//   内部那道闸门 (x86 生成码 vb6_ChkByte 零命中), 而 `(Abs(Opacity)/100)*255`
//   交出 -2147483648 只有一种成因: **Opacity 形参里装的是 OLE_COLOR 值**
//   (&H80000000 = 系统色标志位, RGBtoARGB 自己第一行就 `If (RGBColor And
//   &H80000000)` 判它) ⇒ 真正的缺陷是**调用点实参错位**, 不是 CByte 的检查。
//   本函数是全局溢出安全网, 一律按原样保留, 不因这一次定位而放宽。
static int32_t vb6_OvfChk(int64_t v, int64_t lo, int64_t hi) {
    if (v < lo || v > hi) {
        vb6_RaiseError(6, vb6_BSTR_FromStr(L"Overflow"));
    }
    return (int32_t)v;
}
uint8_t vb6_ChkByte(int32_t v) { return (uint8_t)vb6_OvfChk(v, 0, 255); }
int16_t vb6_ChkInt(int32_t v)  { return (int16_t)vb6_OvfChk(v, -32768, 32767); }
int32_t vb6_ChkLong(int64_t v) {
    return (int32_t)vb6_OvfChk(v, -2147483647LL - 1, 2147483647LL);
}

#ifdef vb6_CInt
#undef vb6_CInt      // vb6rtl_builtin.h 的 _Generic 宏在定义处必须关闭
#endif
// CInt/CLng/CByte 在 VB6 里同样查表 (`CByte(300)` = Error 6), 不是截断。
int16_t vb6_CInt(double x) { return (int16_t)vb6_OvfChk((int64_t)round(x), -32768, 32767); }
#ifdef vb6_CLng
#undef vb6_CLng
#endif
int32_t vb6_CLng(double x) {
    return (int32_t)vb6_OvfChk((int64_t)round(x), -2147483647LL - 1, 2147483647LL);
}
#ifdef vb6_CDbl
#undef vb6_CDbl
#endif
double vb6_CDbl(double x) {
    return x;
}
// Fix 158r: CLng 的 BSTR 实参 (UDT 单元格 .Text 等) 需先 vb6_Val 解析为 double
int32_t vb6_CLngBSTR(BSTR s) { return (int32_t)round(vb6_Val(s)); }
// Fix 158q: CDbl/CInt/CCur 的 BSTR 实参 (UDT 单元格 .Text 等) —— 同 CLngBSTR.
double  vb6_CDblBSTR(BSTR s) { return vb6_Val(s); }
int16_t vb6_CIntBSTR(BSTR s) { return (int16_t)round(vb6_Val(s)); }
double  vb6_CCurBSTR(BSTR s) { double d = vb6_Val(s); return round(d * 10000.0) / 10000.0; }

BSTR vb6_CStr(vb6_VARIANT x) {
    return vb6_Format(x, NULL, 1, 1);
}

// M22: typed CStr overloads (C has no overloading, use suffix)
#ifdef vb6_CStrLong
#undef vb6_CStrLong   // vb6rtl_builtin.h 的 _Generic 宏在这里必须关闭, 否则函数定义被宏改写
#endif
BSTR vb6_CStrLong(int32_t x) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v)); v.vt = (vb6_vartype)VT_I4; v.lVal = x;
    return vb6_Format(v, NULL, 1, 1);
}
// Fix 158q: 生成代码可能产出 vb6_CStrLong(<vb6_VARIANT 表达式>) —— With 后端 COM 属性
// (vb6_VariantFromComResult(vb6_ComGetProp(...))) 原生返回 VARIANT 却被当 Long 转 BSTR
// (VBFlexGridDemo PPVBFlexGridGeneral.c). msbuild_vs C2440 (无法从 vb6_VARIANT 转换到
// int32_t). vb6rtl_builtin.h 的 _Generic 直通到这里先解包.
BSTR vb6_CStrLongFromVariant(vb6_VARIANT v) {
    return vb6_CStrLong(vb6_VariantToLong(v));
}
#ifdef vb6_CStrDbl
#undef vb6_CStrDbl   // vb6rtl_builtin.h 的 _Generic 宏在这里必须关闭, 否则函数定义被宏改写
#endif
BSTR vb6_CStrDbl(double x) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v)); v.vt = (vb6_vartype)VT_R8; v.dblVal = x;
    return vb6_Format(v, NULL, 1, 1);
}
// vbeclipse: 后期绑定 COM 读到的数值属性, codegen 会发成
//   vb6_CStrDbl(vb6_ComGetStringProp(_vb6_with_N, L"Ratio"))
// (BSTR 实参), 而本函数原签名只收 double → msbuild_vs C2440 ×4
// (ucPerspective.c 3333/3384/3453/3504; Folder.cls `Property Get Ratio() As Double`).
// 与 Fix 158q 的 vb6_CStrLongFromVariant 同款思路: 补一个 BSTR 入口, 由
// vb6rtl_builtin.h 的 _Generic 按实参 C 类型分派. BSTR → double 走 vb6_Val
// (与 vb6_CDblBSTR 同一条口径).
BSTR vb6_CStrDblFromBSTR(BSTR s) {
    return vb6_CStrDbl(vb6_Val(s));
}
BSTR vb6_CStrDblFromVariant(vb6_VARIANT v) {
    return vb6_CStrDbl(vb6_VariantToDouble(v));
}
// Fix 084m: LongLong (恒 64 位有符号) → String。
//   为什么另开一个函数而不复用 vb6_CStrLong: vb6_Format 不认识 VT_I8 (它只覆盖
//   vtInteger/vtLong/vtDouble/... 这一族), 走 vb6_CStrLong 会先被截成 int32_t
//   —— 实测 BigAdd(2e9,2e9)=4000000000 打成 -294967296。故这里直接按 64 位格式化,
//   不经过 VARIANT。VB6 的 CStr 对整数就是十进制无前导零的短形式, 与 %lld 一致。
BSTR vb6_CStrLongLong(int64_t x) {
    wchar_t buf[32];
    _snwprintf_s(buf, 32, _TRUNCATE, L"%lld", (long long)x);
    return vb6_BSTR_FromStr(buf);
}
// Fix 117c: Single → String 必须保留 VT_R4 (7 位有效数字 + 最短往返), 否则
// CSng(21.1) 会像 Double 一样打印成 "21.1000003814697"。
BSTR vb6_CStrSingle(float x) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v)); v.vt = (vb6_vartype)VT_R4; v.fltVal = x;
    return vb6_Format(v, NULL, 1, 1);
}
BSTR vb6_CStrBool(int16_t x) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v)); v.vt = (vb6_vartype)VT_BOOL; v.boolVal = x;
    return vb6_Format(v, NULL, 1, 1);
}
BSTR vb6_CStrByte(uint8_t x) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v)); v.vt = (vb6_vartype)VT_UI1; v.bVal = x;
    return vb6_Format(v, NULL, 1, 1);
}
BSTR vb6_CStrDate(double x) {
    // Fix 175: VB6 的 CStr(Date) 在**时间分量为 0** 时只给短日期, 而共享的
    // vb6_Format(VT_DATE, NULL) 无条件拼 "短日期 + 空格 + 时间"
    // (detail/vb6rtl_format_extract.inc 的 !fmt 分支) → demo 的日期列会变成
    // "2026/1/1 0:00:00"。Format 有自己的语义, 故只在这里分叉。
    double frac = x - floor(x);
    if (frac < 0) frac = -frac;
    if (frac < 1e-9) {
        SYSTEMTIME st;
        wchar_t dateBuf[64];
        if (VariantTimeToSystemTime(x, &st)
            && GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, NULL, dateBuf, 64) > 0) {
            return vb6_BSTR_FromStr(dateBuf);
        }
    }
    vb6_VARIANT v; memset(&v, 0, sizeof(v)); v.vt = (vb6_vartype)VT_DATE; v.dblVal = x;
    return vb6_Format(v, NULL, 1, 1);
}
// ============================================================
// 类型检查
// ============================================================

int32_t vb6_IsNumeric(vb6_VARIANT v) {
    switch (v.vt) {
        case vb6_vtInteger: case vb6_vtLong: case vb6_vtSingle:
        case vb6_vtDouble: case vb6_vtCurrency: case vb6_vtByte:
        case vb6_vtBoolean:
            return -1;  // VB6 True
        default:
            return 0;
    }
}

int32_t vb6_IsNull(vb6_VARIANT v) { return v.vt == vb6_vtNull ? -1 : 0; }
int32_t vb6_IsEmpty(vb6_VARIANT v) { return v.vt == vb6_vtEmpty ? -1 : 0; }
int32_t vb6_IsObject(vb6_VARIANT v) { return (v.vt == vb6_vtDispatch && v.pdispVal != NULL) ? -1 : 0; }
int32_t vb6_IsArray(vb6_VARIANT v) { return (v.vt & 0x2000) ? -1 : 0; }  // VT_ARRAY=0x2000
int32_t vb6_IsDate(vb6_VARIANT v) { return v.vt == vb6_vtDate ? -1 : 0; }
int32_t vb6_IsError(vb6_VARIANT v) { return v.vt == vb6_vtError ? -1 : 0; }
// P21-09: CVErr — create VT_ERROR Variant
vb6_VARIANT vb6_CVErr(int32_t errorNumber) {
    vb6_VARIANT v;
    memset(&v, 0, sizeof(v));
    v.vt = vb6_vtError;
    v.lVal = errorNumber;
    return v;
}


// P8.4: VarType - 返回Variant的VT类型码
int32_t vb6_VarType(vb6_VARIANT v) { return (int32_t)v.vt; }

// P8.4: TypeName - 返回Variant类型的VB6类型名
BSTR vb6_TypeName(vb6_VARIANT v) {
    const wchar_t* name = L"Empty";
    /* Fix 112: 宿主对象 (窗体/控件 HWND, Controls 集合, Font) 返回 VB6 类型名 */
    if (v.vt == vb6_vtDispatch && v.pdispVal) {
        const wchar_t* hostName = vb6_Host_TypeNameOf(v.pdispVal);
        if (hostName) return vb6_BSTR_FromStr(hostName);
    }
    switch (v.vt) {
        case vb6_vtEmpty:    name = L"Empty"; break;
        case vb6_vtNull:     name = L"Null"; break;
        case vb6_vtInteger:  name = L"Integer"; break;
        case vb6_vtLong:     name = L"Long"; break;
        case vb6_vtSingle:   name = L"Single"; break;
        case vb6_vtDouble:   name = L"Double"; break;
        case vb6_vtCurrency: name = L"Currency"; break;
        case vb6_vtDate:     name = L"Date"; break;
        case vb6_vtBSTR:     name = L"String"; break;
        case vb6_vtDispatch: name = L"Object"; break;
        case vb6_vtError:    name = L"Error"; break;
        case vb6_vtBoolean:  name = L"Boolean"; break;
        case vb6_vtByte:     name = L"Byte"; break;
        default:             name = L"Variant"; break;
    }
    return vb6_BSTR_FromStr(name);
}

// ============================================================
// Debug对象
// ============================================================

void vb6_Debug_Print(BSTR s) {
    if (s) {
        wprintf(L"%ls\n", s);
    } else {
        wprintf(L"\n");
    }
    fflush(stdout);
}

void vb6_Debug_PrintInt(int32_t n) {
    wprintf(L"%d\n", n);
    fflush(stdout);
}

void vb6_Debug_PrintDouble(double d) {
    wprintf(L"%g\n", d);
    fflush(stdout);
}

// ============================================================
// 对象操作 (占位)
// ============================================================

// ============================================================
// 对象操作 (P6: COM互操作)
// ============================================================

void* vb6_NewObject(const wchar_t* className) {
    // Fix 112c: Collection 是 VB6 内建类, 不能走 CLSIDFromProgID (必失败 → 429 弹窗)
    if (className && _wcsicmp(className, L"Collection") == 0) {
        return vb6_Collection_New();
    }
    // Fix 112: StdFont → RTL 内建字体对象 (With m_TitleFont: .Size/.Bold 直接写结构体)
    if (className && _wcsicmp(className, L"StdFont") == 0) {
        return vb6_UC_NewFont();
    }
    // 对于未知类名，尝试通过COM创建 (Dim x As New ClassName，className不在已知类中)
    // VB6中如果className不是项目内的类模块，则尝试COM创建
    // Fix 103: VB6 内建对象 (Collection 等) 实现于运行时内部, 不注册 ProgID —— 直接查
    // 注册表必然失败并抛 429 (实测 New Collection: ProgID: Collection, 0x800401F3)。
    // 先探测 C3 运行时自带的内建实现, 未命中才回退到 COM 注册表。
    void* builtin = vb6_NewBuiltinObject(className);
    if (builtin) return builtin;
    return vb6_CreateObject(className);
}

int32_t vb6_TypeOf(void* obj, const wchar_t* typeName) {
    if (!obj) return 0;  // Nothing不匹配任何类型
    // TypeOf的完整实现需要IDispatch/ITypeInfo，在vb6com.c中
    // 简化版: 始终返回False (后续P6.2完善)
    (void)typeName;
    return 0;
}

void* vb6_DictAccess(void* obj, const wchar_t* key) {
    (void)obj; (void)key;
    return NULL;
}

// ============================================================
// Debug.Print 变参版 (cgen生成用)
// ============================================================

// 控制台输出: 句柄是**控制台**就走 WriteConsoleW —— 控制台对宽字符的渲染与 chcp 无关,
// 中文 cmd(936)、chcp 65001 的终端、Windows Terminal(ConPTY) 一样正确; 重定向(文件/管道)
// 时写 UTF-8 字节, 保持"文件里是 UTF-8"这条既有口径。
// 旧写法是 wprintf: 重定向时 CRT 按 C locale 转窄 ⇒ 中文全变 '?'; 控制台上虽然能显示,
// 但一旦用户改过代码页/在 ConPTY 里就跟着变。
static void vb6_ConWriteHandle(HANDLE h, const wchar_t* s, int len) {
    DWORD mode = 0;
    if (!s || len <= 0) return;
    if (GetConsoleMode(h, &mode)) {
        const wchar_t* p = s;
        int left = len;
        while (left > 0) {
            DWORD chunk = (DWORD)(left > 4096 ? 4096 : left);
            DWORD done = 0;
            if (!WriteConsoleW(h, p, chunk, &done, NULL) || done == 0) break;
            p += done;
            left -= (int)done;
        }
        return;
    }
    {
        /* 非控制台 (重定向到文件 / 被 PowerShell 之类接到管道): 写**控制台代码页**的字节 ——
         * 这正是 cmd 重定向与 PowerShell([Console]::OutputEncoding 默认跟控制台代码页,
         * 实测 PS 5.1 与 pwsh 7 都是 gb2312) 期待的编码; 装不下 (如 437 代码页要写中文)
         * 就退回 UTF-8, 不丢字。 */
        UINT cp = GetConsoleOutputCP();
        int need;
        char* buf;
        FILE* f = (h == GetStdHandle(STD_ERROR_HANDLE)) ? stderr : stdout;
        BOOL usedDefault = FALSE;
        if (!cp) cp = GetACP();
        need = WideCharToMultiByte(cp, 0, s, len, NULL, 0, NULL, &usedDefault);
        if (need <= 0 || usedDefault) {
            cp = CP_UTF8;
            need = WideCharToMultiByte(CP_UTF8, 0, s, len, NULL, 0, NULL, NULL);
        }
        if (need <= 0) return;
        buf = (char*)malloc((size_t)need);
        if (!buf) return;
        if (WideCharToMultiByte(cp, 0, s, len, buf, need, NULL, NULL) == need) {
            fwrite(buf, 1, (size_t)need, f);
            fflush(f);
        }
        free(buf);
    }
}

void vb6_ConWriteOutW(const wchar_t* s, int len) {
    vb6_ConWriteHandle(GetStdHandle(STD_OUTPUT_HANDLE), s, len);
}

void vb6_ConWriteErrW(const wchar_t* s, int len) {
    vb6_ConWriteHandle(GetStdHandle(STD_ERROR_HANDLE), s, len);
}

void vb6_DebugPrintStr(BSTR s) {
    if (s) {
        vb6_ConWriteOutW((const wchar_t*)s, (int)vb6_BSTR_Len(s));
    }
    vb6_ConWriteOutW(L"\n", 1);
    fflush(stdout);
}

// Debug.Print 分项输出
void vb6_DebugWriteBSTR(BSTR s) {
    if (s) vb6_ConWriteOutW((const wchar_t*)s, (int)vb6_BSTR_Len(s));
    fflush(stdout);
}

void vb6_DebugWriteLong(int32_t n) {
    wchar_t buf[32];
    swprintf(buf, 32, L"%d", (int)n);
    vb6_ConWriteOutW(buf, (int)wcslen(buf));
    fflush(stdout);
}

void vb6_DebugWriteDouble(double d) {
    wchar_t buf[64];
    swprintf(buf, 64, L"%g", d);
    vb6_ConWriteOutW(buf, (int)wcslen(buf));
    fflush(stdout);
}

void vb6_DebugWriteNewline(void) {
    vb6_ConWriteOutW(L"\n", 1);
    fflush(stdout);
}

void vb6_DebugOutputFmt(const char* fmt, ...) {
    va_list args;
    va_start(args, fmt);
    for (const char* p = fmt; *p; p++) {
        switch (*p) {
            case 's': {
                BSTR s = va_arg(args, BSTR);
                if (s) wprintf(L"%ls", s);
                break;
            }
            case 'd': {
                int32_t n = va_arg(args, int32_t);
                wprintf(L"%d", n);
                break;
            }
            case 'f': {
                double d = va_arg(args, double);
                wprintf(L"%g", d);
                break;
            }
            default:
                break;
        }
    }
    wprintf(L"\n");
    fflush(stdout);
    va_end(args);
}

// ============================================================

double vb6_Sin(double x) { return sin(x); }
double vb6_Cos(double x) { return cos(x); }
double vb6_Tan(double x) { return tan(x); }
double vb6_Atn(double x) { return atan(x); }
double vb6_Log(double x) { return log(x); }
double vb6_Exp(double x) { return exp(x); }
double vb6_Fix(double x) { return (x >= 0) ? floor(x) : ceil(x); }
double vb6_Int(double x) { return floor(x); }

void vb6_Randomize(double seed) {
    /* Fix 135: VB6 Randomize 用 Timer 计时值作种 (省略参数时)。 */
    if (seed == 0.0) {
        g_rndState135 = (int32_t)((unsigned)(GetTickCount() & 0xFFFFFF));
    } else {
        g_rndState135 = (int32_t)((unsigned)(int32_t)seed & 0xFFFFFF);
    }
}

float vb6_Rnd_Full(int32_t seed) {
    /* Fix 135: 与 vb6_Rnd 同一 LCG。 */
    return vb6_Rnd(seed);
}

// 类型转换 (补充)
// ============================================================

int16_t vb6_CBool(double v) {
    return (v != 0.0) ? -1 : 0;  // VB6 True = -1
}

uint8_t vb6_CByte(double v) {
    // 范围按既有的**截断**值算 (取整本身是另一条账, 见文件头口径 (1));
    // 截断前若已越界, 截断后又会绕回合法值 (-1 -> 255), 所以只能查截断前。
    int32_t t = (int32_t)v;
    (void)vb6_OvfChk(t, 0, 255);
    return (uint8_t)t;
}

float vb6_CSng(double v) {
    return (float)v;
}

double vb6_CDate(vb6_VARIANT v) {
    // <vbeclipse>: 这一档的判据以前只列 Double/Single/Long/Integer —— 于是
    // CDate(一个 VT_DATE 的 Variant) 返回 0 (哨兵 VB-date-cdate 实测)。装箱表现在会产
    // VT_DATE (VB6 的 Date 在 Variant 里就是 VT_DATE=7), 消费端必须认它。
    // 数值面统一走 vb6_VariantToDouble (它已含 vtDate/vtCurrency/vtByte 各档),
    // 这里只排除"非数值"的几档, 免得又写一份平行表。
    if (v.vt == vb6_vtEmpty || v.vt == vb6_vtNull
        || v.vt == vb6_vtDispatch || v.vt == vb6_vtError) return 0.0;
    if (v.vt == vb6_vtBSTR) return vb6_Val(v.bstrVal);   /* 字符串按日期解析 */
    return vb6_VariantToDouble(v);
}

BSTR vb6_Hex(int32_t n) {
    wchar_t buf[16];
    swprintf(buf, 16, L"%X", n);
    return vb6_BSTR_FromStr(buf);
}

BSTR vb6_Oct(int32_t n) {
    wchar_t buf[16];
    swprintf(buf, 16, L"%o", n);
    return vb6_BSTR_FromStr(buf);
}

