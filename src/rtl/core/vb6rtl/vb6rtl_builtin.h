#pragma once
// vb6rtl_builtin.h - 内置函数与系统函数声明（字符串/数学/日期/财务/文件等）
// 由 vb6rtl.h 伞头 include；生成代码不要直接 include 本文件
#include "vb6rtl_base.h"

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 内置函数 (最小子集)
// ============================================================

// 字符串函数
int32_t vb6_Len(BSTR s);
int32_t vb6_LenB_BSTR(BSTR s);  // Fix 048: LenB for BSTR — byte length of string
// Fix 048: LenB — _Generic macro: BSTR → byte length, UDT → sizeof
// VB6 LenB() works for both strings (byte length) and UDTs (structure size)
#define vb6_LenB(x) _Generic((x), \
    BSTR: vb6_LenB_BSTR(x), \
    const BSTR: vb6_LenB_BSTR(x), \
    default: ((int32_t)sizeof(x)) \
)
BSTR vb6_Left(BSTR s, int32_t n);
BSTR vb6_Right(BSTR s, int32_t n);
BSTR vb6_Mid(BSTR s, int32_t start, int32_t len);
int32_t vb6_InStr(int32_t start, BSTR haystack, BSTR needle);
// <vbeclipse>: vbTextCompare 共用核 (定义在 vb6rtl_string.c, InStr/InStrRev/Replace/
// Split/Filter/StrComp 共用)。compare: 0 = vbBinaryCompare, 非 0 = vbTextCompare。
// 位置都是 0 基; 找不到返回 -1 (InStr 系列在外面换算 1 基与 0)。
int32_t vb6_TextFind(BSTR haystack, BSTR needle, int32_t from0, int32_t compare);
int32_t vb6_TextFindRev(BSTR haystack, BSTR needle, int32_t from0, int32_t compare);
int32_t vb6_TextMatchAt(BSTR haystack, int32_t at0, BSTR needle, int32_t compare);
// <vbeclipse>: 带 compare 的 InStr (codegen 在实参是 4 个、或 3 个而首参是字符串时用它)
int32_t vb6_InStrC(int32_t start, BSTR haystack, BSTR needle, int32_t compare);
int32_t vb6_InStrVarC(int32_t start, BSTR haystack, vb6_VARIANT needle, int32_t compare);
// Fix 093a: InStrB — 字节版 InStr. 实参可为 Byte() 一维数组 (vb6_SafeArray1D*)
// 或 BSTR; 返回 1 基字节位置 (0=未找到). 此前 RTL 无此符号 → LNK2019.
int32_t vb6_InStrB(int32_t start, void* haystack, void* needle);
BSTR vb6_UCase(BSTR s);
BSTR vb6_LCase(BSTR s);
BSTR vb6_Trim(BSTR s);
BSTR vb6_LTrim(BSTR s);
// Fix 160w: vb6_LTrim 实参是 vb6_VARIANT 时 (如 Common.c:530 `LTrim(SplitToArray(..)[i])`
// — 下标读出的元素是 Variant → 生成代码裸发 vb6_LTrim(vb6_VARIANT) → C2440
// "无法从 vb6_VARIANT 转换到 BSTR"). 同款 _Generic 分派 (Fix 158s 先例):
// Variant → vb6_LTrimVar (先 vb6_VariantToString 解包), 其余原样. 宏在自身展开
// 期间被禁用, 故 default: 指向真函数; vb6rtl_string.c 的 vb6_LTrim 定义处 #undef.
BSTR vb6_RTrim(BSTR s);
BSTR vb6_Chr(int32_t code);
int32_t vb6_Asc(BSTR s);
double vb6_Val(BSTR s);
// Fix 090d: VB6 类型转换函数的字符串解析 (支持 &H/&O 前缀, 见 vb6rtl.c)
double vb6_NumVal(BSTR s);
BSTR vb6_Str(int32_t n);
BSTR vb6_Format(vb6_VARIANT expr, BSTR fmt, int32_t firstDayOfWeek, int32_t firstWeekOfYear);
#ifndef __cplusplus
#define vb6_LTrim(x) _Generic((x), \
    vb6_VARIANT: vb6_LTrimVar, \
    default: vb6_LTrim)((x))
#endif
BSTR vb6_LTrimVar(vb6_VARIANT s);

// 消息框
int32_t vb6_MsgBox(BSTR prompt, int32_t buttons, BSTR title);
// P7.7: single-arg convenience (MsgBox "text" -> vb6_MsgBox1(text))
static inline int32_t vb6_MsgBox1(BSTR prompt) {
    return vb6_MsgBox(prompt, 0, NULL);
}

// 数值函数
double vb6_Abs(double x);
int32_t vb6_Sgn(double x);
double vb6_Sqr(double x);
double vb6_Round(double x, int32_t decimals);
// Fix 117: VB6 默认数值→字符串 (最短往返表示; Single 最多 7 位有效数字)
BSTR vb6_NumToBSTRDefault(double v, int isSingle);
float vb6_Rnd(int32_t seed);

// 转换函数
int16_t vb6_CInt(double x);
int32_t vb6_CLng(double x);
double vb6_CDbl(double x);
BSTR vb6_CStr(vb6_VARIANT x);
// M22: typed CStr overloads
BSTR vb6_CStrLong(int32_t x);
BSTR vb6_CStrLongFromVariant(vb6_VARIANT x);   // Fix 158q: _Generic 兜底的 Variant 解包入口
// Fix 084m: LongLong → String (64 位, 不截断; vb6_Format 不认 VT_I8, 故独立实现)
BSTR vb6_CStrLongLong(int64_t x);
BSTR vb6_CStrDbl(double x);
BSTR vb6_CStrDblFromBSTR(BSTR s);      // vbeclipse: BSTR 实参 (后期绑定 COM 数值属性读) → vb6_Val 解析
BSTR vb6_CStrDblFromVariant(vb6_VARIANT v);  // Fix 158q 同款: _Generic 的 Variant 解包入口
// Fix 117c: Single 专用 (VT_R4, 7 位有效数字 + 最短往返)
BSTR vb6_CStrSingle(float x);
BSTR vb6_CStrBool(int16_t x);
BSTR vb6_CStrByte(uint8_t x);
BSTR vb6_CStrDate(double x);
// P8.4: Variant版转换函数
// Fix 090d: 前置声明 (定义在下方"类型转换"区, 供 inline V 变体引用)
int16_t vb6_CBool(double v);
uint8_t vb6_CByte(double v);
float vb6_CSng(double v);
static inline int16_t vb6_CIntV(vb6_VARIANT v) { return vb6_CInt(vb6_VariantToDouble(v)); }
static inline int32_t vb6_CLngV(vb6_VARIANT v) { return vb6_CLng(vb6_VariantToDouble(v)); }
static inline double vb6_CDblV(vb6_VARIANT v) { return vb6_VariantToDouble(v); }
static inline uint8_t vb6_CByteV(vb6_VARIANT v) { return vb6_CByte(vb6_VariantToDouble(v)); }
static inline float vb6_CSngV(vb6_VARIANT v) { return (float)vb6_VariantToDouble(v); }
static inline int16_t vb6_CBoolV(vb6_VARIANT v) { return (v.vt == vb6_vtBoolean) ? v.boolVal : vb6_CBool(vb6_VariantToDouble(v)); }

// 类型检查
int32_t vb6_IsNumeric(vb6_VARIANT v);
int32_t vb6_IsNull(vb6_VARIANT v);
int32_t vb6_IsEmpty(vb6_VARIANT v);
int32_t vb6_IsObject(vb6_VARIANT v);
int32_t vb6_IsArray(vb6_VARIANT v);
int32_t vb6_IsDate(vb6_VARIANT v);
int32_t vb6_IsError(vb6_VARIANT v);
// P21-09: CVErr — create VT_ERROR Variant
vb6_VARIANT vb6_CVErr(int32_t errorNumber);

BSTR vb6_TypeName(vb6_VARIANT v);
int32_t vb6_VarType(vb6_VARIANT v);

// 内置对象
void vb6_Debug_Print(BSTR s);
void vb6_Debug_PrintInt(int32_t n);
void vb6_Debug_PrintDouble(double d);

// Debug.Print 变参版 (cgen保留接口, 暂未使用)
// 内部使用stdarg, fmt字符: s=BSTR, d=int32_t, f=double
void vb6_DebugOutputFmt(const char* fmt, ...);

// 简化版: 单BSTR输出+换行
void vb6_DebugPrintStr(BSTR s);

// Debug.Print 分项输出: 逐参数输出不换行, 最后统一换行
void vb6_DebugWriteBSTR(BSTR s);       // 输出BSTR片段(不换行)
void vb6_DebugWriteLong(int32_t n);    // 输出整数片段(不换行)
void vb6_DebugWriteDouble(double d);   // 输出浮点片段(不换行)
void vb6_DebugWriteNewline(void);      // 输出换行

// 控制台输出: 宽字符走控制台原生路径 (WriteConsoleW, 与 chcp 无关); 重定向时写 UTF-8 字节。
// 中文 cmd 默认代码页 936, 而 RTL 内部是 UTF-16 —— 直接把宽字符交给 CRT 的窄路径会在
// 非 CJK 代码页/重定向下把中文丢成 '?'。所有"给人看的"控制台输出都该走这两个函数。
void vb6_ConWriteOutW(const wchar_t* s, int len);
void vb6_ConWriteErrW(const wchar_t* s, int len);

// Fix 161: 控制台输入原语 (stdin 侧) —— 与上面的输出侧对称:
//   真控制台(键盘) → ReadConsoleW  (本就是 UTF-16)
//   管道/重定向     → ReadFile 字节 → MultiByteToWideChar(GetConsoleCP() ?: ACP)
// 原 vb6_InputBox 用裸 fgetws(stdin) 会把管道字节当 UTF-16 解析 ⇒ ASCII/中文
// 全部读错(不是显示乱码)。VB6 侧无对应关键字, 供宿主/夹具 Declare 或 RTL 内部调用。
int32_t vb6_ConInputIsConsole(void);  // 1=控制台键盘, 0=管道/重定向
BSTR vb6_ConReadLine(void);           // 读一行(不含行尾); 无输入时返回空 BSTR
BSTR vb6_ConReadKey(void);            // 读一个字符(取行首字符)

// Fix 161: Console 对象 (对齐 twinBASIC) —— VB6 侧 `Console.ReadLine()` 等
void vb6_Console_Write(BSTR s);
void vb6_Console_WriteLine(BSTR s);
BSTR vb6_Console_ReadLine(void);
BSTR vb6_Console_ReadKey(void);

// 整除
int32_t vb6_IntDiv(int32_t a, int32_t b);

// Task #44: VB6 '/' 与 Mod 的字面量 0 除数 → 运行期错误 11
// (常量折叠 C2124 规避 + On Error Resume Next 语义, 见 vb6rtl.c 实现)
double vb6_Num_Div(double a, double b);
int32_t vb6_Num_Mod(int32_t a, int32_t b);

// 幂运算
double vb6_Pow(double base, double exp);

// 对象操作
void* vb6_NewObject(const wchar_t* className);
int32_t vb6_TypeOf(void* obj, const wchar_t* typeName);
void* vb6_DictAccess(void* obj, const wchar_t* key);

// 字符串函数 (补充)
BSTR vb6_Replace(BSTR expr, BSTR find, BSTR rep, int32_t start, int32_t count, int32_t compare);
BSTR vb6_Space(int32_t n);
BSTR vb6_String(int32_t n, int32_t charCode);
int32_t vb6_StrComp(BSTR s1, BSTR s2, int32_t compare);
BSTR vb6_StrReverse(BSTR s);
int32_t vb6_InStrRev(BSTR haystack, BSTR needle, int32_t start, int32_t compare);
BSTR vb6_LCase_str(BSTR s);  // LCase$别名
BSTR vb6_UCase_str(BSTR s);
int16_t vb6_Like(BSTR source, BSTR pattern);  // Like运算符
// <vbeclipse>: 带比较模式的 Like (mode: 0=vbBinaryCompare, 非0=文本)。
// `Option Compare Text` 模块里 codegen 发改写调用点到这里。
int16_t vb6_LikeC(BSTR source, BSTR pattern, int32_t mode);
// <vbeclipse>: 文本比较的唯一口径 (CompareStringW 区域语言序, 见 vb6rtl_string.c);
// vb6_StrComp 的 vbTextCompare 形与 Option Compare Text 模块的运算符共用它。
int vb6_TextCmp(const wchar_t* a, const wchar_t* b);

// ============================================================
// P14.2.2: 系统函数
// ============================================================
BSTR vb6_Dir(BSTR pathname, int32_t attributes);
BSTR vb6_CurDir(BSTR drive);
int32_t vb6_Shell(BSTR pathname, int32_t windowstyle);
BSTR vb6_Environ(BSTR envstring);
BSTR vb6_Command(void);










// 数学函数 (补充)
double vb6_Sin(double x);
double vb6_Cos(double x);
double vb6_Tan(double x);
double vb6_Atn(double x);
double vb6_Log(double x);
double vb6_Exp(double x);
double vb6_Fix(double x);
double vb6_Int(double x);
void vb6_Randomize(double seed);
float vb6_Rnd_Full(int32_t seed);

// 日期时间函数
double vb6_Now(void);
double vb6_Date(void);
double vb6_Time(void);
void vb6_DateSet(BSTR dateStr);
void vb6_TimeSet(BSTR timeStr);
int32_t vb6_Year(double date);
int32_t vb6_Month(double date);
int32_t vb6_Day(double date);
int32_t vb6_Hour(double time);
int32_t vb6_Minute(double time);
int32_t vb6_Second(double time);

// P14.2.4: DateAdd/DateDiff/DatePart/DateSerial
double vb6_DateSerial(int32_t year, int32_t month, int32_t day);
// P14.2.4: DateAdd/DateDiff/DatePart
double vb6_DateAdd(BSTR interval, double number, double date);
int64_t vb6_DateDiff(BSTR interval, double date1, double date2, int32_t firstDayOfWeek, int32_t firstWeekOfYear);
int32_t vb6_DatePart(BSTR interval, double date, int32_t firstDayOfWeek, int32_t firstWeekOfYear);

// P21-B: Weekday/DateValue/TimeSerial/TimeValue
int32_t vb6_Weekday(double date, int32_t firstDayOfWeek);
double vb6_DateValue(BSTR dateStr);
double vb6_TimeSerial(int32_t hour, int32_t minute, int32_t second);
double vb6_TimeValue(BSTR timeStr);

// P14.3.4: App全局对象属性
BSTR vb6_App_Path(void);     // App.Path - EXE所在目录
BSTR vb6_App_EXEName(void);  // App.EXEName - EXE文件名(不含扩展名)
intptr_t vb6_App_hInstance(void); // App.hInstance - 模块实例句柄 (Fix 179: 必须是指针宽度)
BSTR vb6_App_HelpFile(void); // App.HelpFile - 帮助文件名(无App COM对象, 返回空串)

// P18-C: Clipboard 对象
void   vb6_Clipboard_SetText(BSTR text);
BSTR   vb6_Clipboard_GetText(void);
void   vb6_Clipboard_Clear(void);
void   vb6_Clipboard_SetData(void* pPicture, int32_t format);  // format: 0=auto, 2=vbCFBitmap, 14=vbCFEMetafile
int32_t vb6_Clipboard_GetFormat(int32_t format);  // 1=vbCFText, 2=vbCFBitmap, etc.

// P18-C: Screen 对象
int32_t vb6_Screen_Width(void);      // Screen.Width (twips)
int32_t vb6_Screen_Height(void);     // Screen.Height (twips)
int32_t vb6_Screen_MouseX(void);     // Mouse position X (twips)
int32_t vb6_Screen_MouseY(void);     // Mouse position Y (twips)
void*  vb6_Screen_ActiveControl(void);  // Active control HWND
void*  vb6_Screen_ActiveForm(void);     // Active form HWND
int32_t vb6_Screen_TwipsPerPixelX(void);
int32_t vb6_Screen_TwipsPerPixelY(void);
// Fix 161f: Screen.MousePointer 全局读/写 (非控件级)
int32_t vb6_Screen_MousePointer(void);
void    vb6_Screen_SetMousePointer(int32_t pointer);

// P18-C: Printer 对象
void   vb6_Printer_Print(BSTR text);
void   vb6_Printer_EndDoc(void);
void   vb6_Printer_NewPage(void);
int32_t vb6_Printer_Width(void);
int32_t vb6_Printer_Height(void);
int32_t vb6_Printer_CurrentX(void);
int32_t vb6_Printer_CurrentY(void);
void   vb6_Printer_SetCurrentX(int32_t x);
void   vb6_Printer_SetCurrentY(int32_t y);
// Task #39: Printer / Printers 内置全局对象 (cDlg.cls ShowFont / IsPrinter)
// vb6_Printer_Object: 默认打印机 DC 指针作哨兵 (懒创建, 无打印机时 NULL →
//   `Printer Is Nothing` 判真, 调用方自然跳过 hDC 取用);
// vb6_Printer_hDC: HDC 句柄 (Printer.hDC, ChooseFont.hDC 等 API 用);
// vb6_Printers_Collection: 空 RTL Collection (vb6_ForEach_Init 原生支持,
//   For Each 循环零次; 真实 EnumPrinters 枚举属后续增强)。
void*  vb6_Printer_Object(void);

// Fix <vbeclipse>: vb6_Screen_Object — Screen 全局对象哨兵 (NULL)。
// .Width/.MouseX 等已登记成员在 MemberAccess 分支直接发 vb6_Screen_Xxx();
// 未登记成员赋值落 NULL 槽语义, 无害。
void*  vb6_Screen_Object(void);
extern void* vb6_Screen_MouseIcon;         // Fix <vbeclipse>: Screen.MouseIcon 读写槽
void*  vb6_Printer_hDC(void);
void*  vb6_Printers_Collection(void);

// P18-C: Forms 集合
int32_t vb6_Forms_Count(void);
void*  vb6_Forms_Item(int32_t index);  // 0-based
void   vb6_Forms_Register(void* hwnd);   // 窗体创建时注册
void   vb6_Forms_Unregister(void* hwnd); // 窗体销毁时注销
void   vb6_Forms_LoopDepth(int delta);   // Fix 188: 消息循环进出 (最后一个窗体卸载才投 WM_QUIT)
void*  vb6_Forms_GetActive(void);        // Fix 146: 当前活动窗体 (Screen.ActiveForm)

// P14.2.4: IIf / InputBox
BSTR vb6_IIfBSTR(int32_t cond, BSTR truepart, BSTR falsepart);
int32_t vb6_IIfLong(int32_t cond, int32_t truepart, int32_t falsepart);
double vb6_IIfDouble(int32_t cond, double truepart, double falsepart);
vb6_VARIANT vb6_IIfVariant(int32_t cond, vb6_VARIANT truepart, vb6_VARIANT falsepart);
BSTR vb6_InputBox(BSTR prompt, BSTR title, BSTR defaultstr, int32_t xpos, int32_t ypos, BSTR helpfile, int32_t context);

// 类型转换 (补充)
int16_t vb6_CBool(double v);
uint8_t vb6_CByte(double v);
// 溢出检查收窄 (VB6 Error 6): 值越界时 vb6_RaiseError(6, "Overflow"),
// 可被 On Error 捕获; 范围内原样窄化。ai/009 §5.10 P3。
uint8_t vb6_ChkByte(int32_t v);
int16_t vb6_ChkInt(int32_t v);
int32_t vb6_ChkLong(int64_t v);
float vb6_CSng(double v);
double vb6_CDate(vb6_VARIANT v);
BSTR vb6_Hex(int32_t n);
BSTR vb6_Oct(int32_t n);

// P18-A: 兼容性填平 — 新增RTL函数
double vb6_CCur(double v);             // Fix 126: CCur 值语义 (4 位小数)
vb6_VARIANT vb6_CDec(vb6_VARIANT v);     // P20-07: CDec返回真实DECIMAL (vt=14)
long vb6_RGB(int32_t r, int32_t g, int32_t b);  // RGB: OLE color
long vb6_QBColor(int32_t n);        // QBColor: 16-color lookup
double vb6_FileDateTime(BSTR pathname); // FileDateTime: 文件修改时间→VB6 date serial
int32_t vb6_FileLen(BSTR pathname);    // FileLen: 文件大小(字节)
// <vbeclipse>: GetAttr/SetAttr 只有定义没有声明 (vb6rtl_misc.c:99/107) ⇒ 生成的 C 走
// 隐式 int 声明, 每个用到它们的工程都吃一条 cl C4013; 补上原型消除该族警告。
int32_t vb6_GetAttr(BSTR pathname);    // GetAttr: 文件属性 (GetFileAttributesW)
void    vb6_SetAttr(BSTR pathname, int32_t attributes); // SetAttr: 设置文件属性
void   vb6_SendKeys(BSTR keys, int32_t wait);    // SendKeys: 发送按键
void   vb6_AppActivate(BSTR title, int32_t wait); // AppActivate: 激活窗口(标题或数字PID)
void   vb6_AppActivateByPid(int32_t pid, int32_t wait); // AppActivate: 按进程ID激活窗口
void   vb6_MidSet(BSTR* target, int32_t start, int32_t len, BSTR replacement); // Mid$ statement赋值

// P18-D: 兼容性填平 — 字符串/指针/格式化函数
int32_t vb6_AscW(BSTR s);           // AscW: Unicode code point
BSTR   vb6_ChrW(int32_t code);      // ChrW: Unicode character
int32_t vb6_AscB(BSTR s);           // AscB: first byte value
BSTR   vb6_ChrB(int32_t code);      // ChrB: single-byte string
double  vb6_Timer(void);            // Timer: seconds since midnight (fractional)
BSTR   vb6_StrConv(BSTR text, int32_t conversion, int32_t localeID); // StrConv
struct vb6_SafeArray1D; // forward declaration
// StrConv 字节数组语义 (twinbasic 兼容):
//   vb6_StringToByteArray    : String -> Byte()，复制 BSTR 原始 UTF-16LE 内存字节
//   vb6_StrConvToByteArray   : StrConv(s, 128/64) -> Byte()
//                               128 (vbFromUnicode): Unicode -> 系统 ANSI/GBK 字节
//                               64  (vbUnicode)    : 原始 BSTR 内存按 ANSI 解码后重建 UTF-16LE 字节
//   vb6_StrConvFromByteArray : 反向, Byte() -> BSTR (供 s = StrConv(ba, vbUnicode))
struct vb6_SafeArray1D* vb6_StringToByteArray(BSTR text);
struct vb6_SafeArray1D* vb6_StrConvToByteArray(BSTR text, int32_t conversion, int32_t localeID);
BSTR   vb6_StrConvFromByteArray(struct vb6_SafeArray1D* arr, int32_t conversion, int32_t localeID);
// Fix 140: Byte() -> BSTR (UTF-16LE 字节直读) 与 Byte() 数据指针.
//   vb6_ByteArrayToString(arr): 把 arr->data 当作 UTF-16LE 宽字符串还原为 BSTR.
//     与 vb6_StringToByteArray 互逆 (原始内存字节, 非 ANSI 转码).
//   vb6_SafeArrayDataPtr(arr) : StrPtr(Byte()) 语义 — 返回 arr->data (非结构体指针).
BSTR   vb6_ByteArrayToString(struct vb6_SafeArray1D* arr);
void*  vb6_SafeArrayDataPtr(struct vb6_SafeArray1D* arr);
struct vb6_SafeArray1D* vb6_Filter(struct vb6_SafeArray1D* source, BSTR match, int32_t include, int32_t compare);
void*    vb6_StrPtr(BSTR s);          // StrPtr: address of string data
uintptr_t vb6_ObjPtr(void* obj);       // ObjPtr: address of object
BSTR   vb6_LSet(BSTR str, int32_t length);  // LSet: left-justify
BSTR   vb6_RSet(BSTR str, int32_t length);  // RSet: right-justify
BSTR   vb6_WeekdayName(int32_t weekday, int32_t abbreviate, int32_t firstDayOfWeek);
BSTR   vb6_MonthName(int32_t month, int32_t abbreviate);
BSTR   vb6_FormatCurrency(double value, int32_t numDigits, int32_t incLeading, int32_t useParens, int32_t groupDigits);
BSTR   vb6_FormatNumber(double value, int32_t numDigits, int32_t incLeading, int32_t useParens, int32_t groupDigits);
BSTR   vb6_FormatPercent(double value, int32_t numDigits, int32_t incLeading, int32_t useParens, int32_t groupDigits);

// P18-E: 兼容性填平 — 金融函数+文件锁定+Partition
double vb6_SLN(double cost, double salvage, double life);
double vb6_SYD(double cost, double salvage, double life, double period);
double vb6_DDB(double cost, double salvage, double life, double period, double factor);
double vb6_FV(double rate, double nper, double pmt, double pv, int32_t type);
double vb6_PV(double rate, double nper, double pmt, double fv, int32_t type);
double vb6_Pmt(double rate, double nper, double pv, double fv, int32_t type);
double vb6_IPmt(double rate, double per, double nper, double pv, double fv, int32_t type);
double vb6_PPmt(double rate, double per, double nper, double pv, double fv, int32_t type);
double vb6_RATE(double nper, double pmt, double pv, double fv, int32_t type, double guess);
double vb6_NPV(double rate, struct vb6_SafeArray1D* values);
void   vb6_Lock(int32_t filenum, int64_t start, int64_t end);
void   vb6_Unlock(int32_t filenum, int64_t start, int64_t end);
void   vb6_Reset(void);
BSTR   vb6_Partition(int64_t number, int64_t start, int64_t stop, int64_t interval);

// ============================================================
// Fix 158q: CStrLong/CLng/CDbl/CInt/CCur 的 Variant·BSTR 实参分派
// ============================================================
// 生成代码里会出现 vb6_CStrLong(<vb6_VARIANT>) (With 后端 COM 属性
// vb6_VariantFromComResult(vb6_ComGetProp(...)) 被当 Long 转 BSTR → C2440
// "无法从 vb6_VARIANT 转换到 int32_t"), 以及 vb6_CLng/vb6_CDbl/vb6_CInt/
// vb6_CCur 收 vb6_VARIANT 或 BSTR 实参 (UDT 单元格 .Text → C2440 BSTR→double).
// 用 _Generic (MSVC 2019 16.8+, /std:c11) 在**调用点**分派:
//   vb6_VARIANT → 既有 vb6_*V 解包版; BSTR → 新增 vb6_*BSTR 解析版; 其余原样.
// 位置必须在上面所有 vb6_CInt/CLng/CDbl/CCur 声明**之后**: 宏一旦定义, 其后
// 出现同名声明/定义都会被改写 (头内 79~81 行的 CIntV/CLngV 内联体因此已在
// 宏之前解析, 走的是真函数, 不会自我展开). 各 .c 的定义处用 #undef 自行关闭
// (vb6rtl_conv.c / vb6rtl_misc.c), 同 vb6_LenB (Fix 048) 的既有先例.
int32_t vb6_CLngBSTR(BSTR s);
double vb6_CDblBSTR(BSTR s);
int16_t vb6_CIntBSTR(BSTR s);
double vb6_CCurBSTR(BSTR s);
static inline double vb6_CCurV(vb6_VARIANT v) { return vb6_CCur(vb6_VariantToDouble(v)); }
// Fix 160w: CLng(指针) 分派目标 — 截断为 int32_t (VB6 Long = 地址低半).
static inline int32_t vb6_CLngPtr(void* p) { return (int32_t)(intptr_t)p; }
#ifndef __cplusplus
#define vb6_CStrLong(x) _Generic((x), \
    vb6_VARIANT: vb6_CStrLongFromVariant, \
    default: vb6_CStrLong)((x))
// vbeclipse: CStrDbl 补 BSTR/Variant 分派 (默认分支不变, 老调用点行为一致).
// 背景: 后期绑定 COM 数值属性在拼接语境下被 codegen 读成 vb6_ComGetStringProp
// (BSTR), 而 vb6_CStrDbl 原签名只收 double → C2440 (ucPerspective.c 3333/3384/
// 3453/3504).
#define vb6_CStrDbl(x) _Generic((x), \
    BSTR: vb6_CStrDblFromBSTR, \
    vb6_VARIANT: vb6_CStrDblFromVariant, \
    default: vb6_CStrDbl)((x))
#define vb6_CLng(x) _Generic((x), \
    vb6_VARIANT: vb6_CLngV, \
    BSTR: vb6_CLngBSTR, \
    void*: vb6_CLngPtr, \
    default: vb6_CLng)((x))
#define vb6_CDbl(x) _Generic((x), \
    vb6_VARIANT: vb6_CDblV, \
    BSTR: vb6_CDblBSTR, \
    default: vb6_CDbl)((x))
#define vb6_CInt(x) _Generic((x), \
    vb6_VARIANT: vb6_CIntV, \
    BSTR: vb6_CIntBSTR, \
    default: vb6_CInt)((x))
#define vb6_CCur(x) _Generic((x), \
    vb6_VARIANT: vb6_CCurV, \
    BSTR: vb6_CCurBSTR, \
    default: vb6_CCur)((x))
#endif

// Fix 158s: vb6_InStr 的 needle 实参是 vb6_VARIANT 时 (For 循环里对每个 Buffer 判
// 定包含关系, Value 形参是 Variant) → 生成代码裸发 vb6_InStr(1, Buffer, Value),
// VBFlexGrid.c 一系 C2440 "无法从 vb6_VARIANT 转换到 BSTR"。同款 _Generic 分派:
// Variant → vb6_InStrVar (先 vb6_VariantToString 解包), 其余原样。两侧返回类型同为
// int32_t, 分派自洽。宏在自身展开期间被禁用 (与 vb6_CLng 同技巧), 故 default: 指向
// 真函数, 无递归; vb6rtl_string.c 的定义处自带 #undef。
int32_t vb6_InStrVar(int32_t start, BSTR haystack, vb6_VARIANT needle);
#ifndef __cplusplus
#define vb6_InStr(start, haystack, needle) _Generic((needle), \
    vb6_VARIANT: vb6_InStrVar, \
    default: vb6_InStr)((start), (haystack), (needle))
#endif

// <vbeclipse>: 4 参形的同款分派 (Fix 158s 的机制, 只是多一个 compare 实参) ——
// InStr 的"字符串优先三参形"与四参形都发到这里, needle 是 Variant 时先解包。
#ifndef __cplusplus
#define vb6_InStrC(start, haystack, needle, compare) _Generic((needle), \
    vb6_VARIANT: vb6_InStrVarC, \
    default: vb6_InStrC)((start), (haystack), (needle), (compare))
#endif

#ifdef __cplusplus
}
#endif
