// vb6rtl_file.c - VB6 运行时库: 文件家族：文件 I/O 通道表 + 目录/文件操作 + Get/Put + 文件锁
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 3618~3918 行
//   原第 4115~4181 行
//   原第 4595~4666 行

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
#else
#include <sys/stat.h>  // mkdir
#include <unistd.h>    // rmdir / chdir
#endif

// P24-08: MessageBoxW (user32) + GetConsoleWindow (kernel32)
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "kernel32.lib")

// ============================================================
// Fix 197: 路径与文本的平台编码 —— 不再把 UTF-16 硬截成 char
// ============================================================
// C3 的 BSTR 在 Windows 下就是 wchar_t* (UTF-16), 本文件原先却用
// vb6_bstr_to_narrow() 把每个 UTF-16 码元直接截成 1 个 char —— 连 ACP 转换都不是,
// 于是 "新建文件夹" 变成 8 个垃圾字节, Open/Kill/MkDir/RmDir/ChDir/Name/FileCopy
// 在中文 (以及任何非 Latin-1) 路径下一律失败。
//
// 正解不是"转成 ACP 再用窄 API": GBK 编不出韩文/俄文里的字母, 转换会永久丢字
// —— 那是 VB6 自己的局限, 我们复现它的**语义**, 不复现它的**局限**。文件家族的其他
// 成员 (FileLen / GetAttr / SetAttr / Dir / CurDir / Shell / Environ / App.Path) 早
// 就是宽字符直通, 这里只是把它们拉齐:
//   * 路径 —— Windows 把 BSTR 原样交给 _wfopen/_wremove/_wrename/_wmkdir/_wrmdir/
//     _wchdir; 其他平台文件名是字节串, 转 UTF-8。
//   * 文本 (Print# / Write# / LineInput# / Input$) —— 按本平台文本编码编解码
//     (Windows=ACP, 正是 VB6 写文件用的编码), 而不是按码元截断。

// 平台路径持有者: Windows 零拷贝直指 BSTR 本体; 其他平台持有一份 UTF-8 副本。
typedef struct {
#if defined(_WIN32)
    const wchar_t* p;
#else
    char* p;
#endif
} vb6_path;

#if defined(_WIN32)

static vb6_path vb6_path_from(BSTR bstr) {
    vb6_path pp;
    pp.p = bstr ? (const wchar_t*)bstr : L"";  // NULL → 空串, 交给 API 去报"找不到"
    return pp;
}
static void vb6_path_release(vb6_path* pp) { (void)pp; }  // 零拷贝, 无需释放

// mode 都是 ASCII 字面量 ("rb"/"w+b"...), 逐字符加宽即可
static FILE* vb6_path_open(const vb6_path* pp, const char* mode) {
    wchar_t wmode[8];
    int i = 0;
    for (; mode[i] != '\0' && i < 7; i++) wmode[i] = (wchar_t)(unsigned char)mode[i];
    wmode[i] = L'\0';
    return _wfopen(pp->p, wmode);
}
static int vb6_path_remove(const vb6_path* pp) { return _wremove(pp->p); }
static int vb6_path_rename(const vb6_path* a, const vb6_path* b) { return _wrename(a->p, b->p); }
static int vb6_path_mkdir(const vb6_path* pp) { return _wmkdir(pp->p); }
static int vb6_path_rmdir(const vb6_path* pp) { return _wrmdir(pp->p); }
static int vb6_path_chdir(const vb6_path* pp) { return _wchdir(pp->p); }

#else

// UTF-16 → UTF-8。自备编码器: wcstombs 受 C locale 限制, 默认 locale 下非 ASCII 全丢。
static size_t vb6_utf8_put(char* out, uint32_t cp) {
    if (cp < 0x80) {
        out[0] = (char)cp;
        return 1;
    }
    if (cp < 0x800) {
        out[0] = (char)(0xC0 | (cp >> 6));
        out[1] = (char)(0x80 | (cp & 0x3F));
        return 2;
    }
    if (cp < 0x10000) {
        out[0] = (char)(0xE0 | (cp >> 12));
        out[1] = (char)(0x80 | ((cp >> 6) & 0x3F));
        out[2] = (char)(0x80 | (cp & 0x3F));
        return 3;
    }
    out[0] = (char)(0xF0 | (cp >> 18));
    out[1] = (char)(0x80 | ((cp >> 12) & 0x3F));
    out[2] = (char)(0x80 | ((cp >> 6) & 0x3F));
    out[3] = (char)(0x80 | (cp & 0x3F));
    return 4;
}

static vb6_path vb6_path_from(BSTR bstr) {
    vb6_path pp;
    pp.p = NULL;
    int32_t len = vb6_BSTR_Len(bstr);
    char* buf = (char*)malloc((size_t)len * 4 + 1);  // 每个码元最多 4 字节
    if (!buf) return pp;
    size_t n = 0;
    for (int32_t i = 0; i < len; i++) {
        uint32_t cp = (uint32_t)bstr[i];
        if (cp >= 0xD800 && cp <= 0xDBFF && i + 1 < len) {  // 代理对
            uint32_t lo = (uint32_t)bstr[i + 1];
            if (lo >= 0xDC00 && lo <= 0xDFFF) {
                cp = 0x10000 + ((cp - 0xD800) << 10) + (lo - 0xDC00);
                i++;
            }
        } else if (cp >= 0xD800 && cp <= 0xDFFF) {
            cp = 0xFFFD;  // 孤立代理项
        }
        n += vb6_utf8_put(buf + n, cp);
    }
    buf[n] = '\0';
    pp.p = buf;
    return pp;
}
static void vb6_path_release(vb6_path* pp) {
    free(pp->p);
    pp->p = NULL;
}
static FILE* vb6_path_open(const vb6_path* pp, const char* mode) { return fopen(pp->p, mode); }
static int vb6_path_remove(const vb6_path* pp) { return remove(pp->p); }
static int vb6_path_rename(const vb6_path* a, const vb6_path* b) { return rename(a->p, b->p); }
static int vb6_path_mkdir(const vb6_path* pp) { return mkdir(pp->p, 0755); }
static int vb6_path_rmdir(const vb6_path* pp) { return rmdir(pp->p); }
static int vb6_path_chdir(const vb6_path* pp) { return chdir(pp->p); }

#endif

// ---- 文本编解码 (Print# / Write# / LineInput# / Input$) ----
// Windows: ACP —— 与 VB6 写文件完全一致 (中文 Windows 上就是 GBK); 码页编不出的字符
//          由系统默认字符 (通常是 '?') 顶上, 也与 VB6 一致。
// 其他平台: UTF-8。
// 逐码元转换, 所以跨调用拆开的代理对 (emoji) 会退化成 '?' —— VB6 同样处理不了它们。
static void vb6_text_putc(FILE* f, wchar_t ch) {
#if defined(_WIN32)
    char buf[8];
    int n = WideCharToMultiByte(CP_ACP, 0, &ch, 1, buf, (int)sizeof(buf), NULL, NULL);
    if (n <= 0) {
        fputc('?', f);
        return;
    }
    fwrite(buf, 1, (size_t)n, f);
#else
    char buf[4];
    fwrite(buf, 1, vb6_utf8_put(buf, (uint32_t)ch), f);
#endif
}

// 读一个字符; 文件尾或尾巴上的半截序列返回 -1
static int32_t vb6_text_getc(FILE* f) {
#if defined(_WIN32)
    int b = fgetc(f);
    if (b == EOF) return -1;
    char buf[2];
    int n = 1;
    buf[0] = (char)b;
    if (IsDBCSLeadByteEx(CP_ACP, (BYTE)b)) {  // ACP 是双字节码页 (936/932/949...) 时
        int b2 = fgetc(f);
        if (b2 == EOF) return -1;             // 尾巴上只剩半个字符
        buf[1] = (char)b2;
        n = 2;
    }
    wchar_t wc = 0;
    if (MultiByteToWideChar(CP_ACP, 0, buf, n, &wc, 1) != 1) return -1;
    return (int32_t)wc;
#else
    /* UTF-8 增量解码 */
    int b = fgetc(f);
    if (b == EOF) return -1;
    if (b < 0x80) return (int32_t)b;
    if (b < 0xC0) return -1; /* 孤立续字节: 丢弃 */
    int extra = (b >= 0xF0) ? 3 : ((b >= 0xE0) ? 2 : 1);
    int32_t cp = b & ((1 << (6 - extra)) - 1);
    for (int i = 0; i < extra; i++) {
        int c = fgetc(f);
        if (c == EOF) return -1;
        cp = (cp << 6) | (c & 0x3F);
    }
    return cp;
#endif
}

// ============================================================
// 文件 I/O (MVP)
// ============================================================

// VB6文件I/O使用通道号(1-511), 我们用文件指针表实现
#define VB6_MAX_FILES 32
static FILE* vb6_file_table[VB6_MAX_FILES] = {0};
static int32_t vb6_file_mode[VB6_MAX_FILES] = {0};  // 1=Input, 2=Output, 4=Random, 8=Append, 16=Binary
static int32_t vb6_file_reclen[VB6_MAX_FILES] = {0}; // P8.2: 记录长度 (Random模式)
static int32_t vb6_width_table[VB6_MAX_FILES] = {0}; // P22-08: Width# 行宽 (0=不限)
static int32_t vb6_col_table[VB6_MAX_FILES] = {0};   // P22-08: 当前列位置

int32_t vb6_FreeFile(void) {
    for (int32_t i = 1; i < VB6_MAX_FILES; i++) {
        if (!vb6_file_table[i]) return i;
    }
    return -1;  // 无可用通道
}

// P21-17: FileAttr — 查询已打开文件的属性
//   attribute=1 → 文件模式 (1=Input, 2=Output, 4=Random, 8=Append, 16=Binary)
//   attribute=2 → 操作系统文件句柄
// 声明见 vb6rtl_class_com.h. 此前只有声明没有定义 → 用到即 LNK2019.
// 通道未打开时按 VB6 语义报错 52 (Bad file name or number).
int32_t vb6_FileAttr(int32_t filenumber, int32_t attribute) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) {
        vb6_RaiseError(52, vb6_BSTR_FromStr(L"Bad file name or number"));
        return 0;
    }
    if (attribute == 1) return vb6_file_mode[filenumber];
    if (attribute == 2) return (int32_t)_fileno(vb6_file_table[filenumber]);
    return 0;
}

int32_t vb6_Open(BSTR pathname, int32_t mode, int32_t access, int32_t filenumber, int32_t reclength) {
    (void)access;  // 简化: 忽略access参数
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES) return 0;
    if (vb6_file_table[filenumber]) return 0;  // 已打开

    const char* modeStr = "";
    switch (mode) {
        case 1: modeStr = "rb"; break;   // Input (binary: VB6 uses CRLF explicitly)
        case 2: modeStr = "wb"; break;   // Output
        case 4: modeStr = "r+b"; break; // Random
        case 8: modeStr = "ab"; break;   // Append
        case 16: modeStr = "r+b"; break; // Binary (读写)
        default: return 0;
    }

    vb6_path path = vb6_path_from(pathname);

    // Random/Binary模式需要文件存在才能r+b, 否则先创建
    FILE* f = NULL;
    if (mode == 4 || mode == 16) {
        // Random/Binary模式需要读写, 尝试打开已有文件, 不存在则创建
        f = vb6_path_open(&path, "r+b");
        if (!f) f = vb6_path_open(&path, "w+b");
    } else {
        f = vb6_path_open(&path, modeStr);
    }
    vb6_path_release(&path);

    if (!f) {
        vb6_RaiseError(53, vb6_BSTR_FromStr(L"File not found"));
        return 0;
    }
    vb6_file_table[filenumber] = f;
    vb6_file_mode[filenumber] = mode;
    vb6_file_reclen[filenumber] = (reclength > 0) ? reclength : 128;  // P8.2: 默认128
    vb6_width_table[filenumber] = 0;  // P22-08: reset width on open
    vb6_col_table[filenumber] = 0;    // P22-08: reset column on open
    return -1;  // True
}

int32_t vb6_Close(int32_t filenumber) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES) return 0;
    if (vb6_file_table[filenumber]) {
        fclose(vb6_file_table[filenumber]);
        vb6_file_table[filenumber] = NULL;
        vb6_file_mode[filenumber] = 0;
        vb6_file_reclen[filenumber] = 0;
        vb6_width_table[filenumber] = 0;  // P22-08
        vb6_col_table[filenumber] = 0;    // P22-08
    }
    return -1;
}

int32_t vb6_CloseAll() {
    int count = 0;
    for (int i = 1; i < VB6_MAX_FILES; i++) {
        if (vb6_file_table[i]) {
            fclose(vb6_file_table[i]);
            vb6_file_table[i] = NULL;
            vb6_file_mode[i] = 0;
            vb6_file_reclen[i] = 0;
            vb6_width_table[i] = 0;  // P22-08
            vb6_col_table[i] = 0;    // P22-08
            count++;
        }
    }
    return count;
}

int32_t vb6_EOF(int32_t filenumber) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return -1;
    FILE* f = vb6_file_table[filenumber];
    int c = fgetc(f);
    if (c == EOF) return -1;  // True
    ungetc(c, f);
    return 0;  // False
}

int32_t vb6_LOF(int32_t filenumber) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return 0;
    FILE* f = vb6_file_table[filenumber];
    long cur = ftell(f);
    fseek(f, 0, SEEK_END);
    long size = ftell(f);
    fseek(f, cur, SEEK_SET);
    return (int32_t)size;
}

int32_t vb6_Loc(int32_t filenumber) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return 0;
    // 简化: 返回当前字节位置 / 128 (VB6 Random模式)
    return (int32_t)(ftell(vb6_file_table[filenumber]) / 128) + 1;
}

// P21-13: Seek function — return current file position
int32_t vb6_SeekFunc(int32_t filenumber) {
    // Fix 197: 上界是表长 VB6_MAX_FILES, 原写 255 → filenumber=100 越界读文件表
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return 0;
    return (int32_t)(ftell(vb6_file_table[filenumber]) + 1);  // VB6 is 1-based
}

// P21-13: Seek statement — set file position
void vb6_SeekStmt(int32_t filenumber, int32_t position) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return;
    fseek(vb6_file_table[filenumber], (long)(position - 1), SEEK_SET);  // VB6 is 1-based
}


// P22-08: Width# — 设置文件输出行宽
void vb6_Width(int32_t filenumber, int32_t width) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES) return;
    vb6_width_table[filenumber] = width;
}

void vb6_Print(int32_t filenumber, BSTR s) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return;
    FILE* f = vb6_file_table[filenumber];
    int32_t w = vb6_width_table[filenumber];  // P22-08: Width# supported line width
    if (s) {
        int32_t len = vb6_BSTR_Len(s);
        for (int32_t i = 0; i < len; i++) {
            // Fix 197: 按**字符** (不是按字节) 判断与计列宽 —— 编成 ACP/UTF-8 后一个
            // 汉字占 2~3 字节, 原先的 (char) 截断既写错内容, 也让列宽算错。
            wchar_t ch = (wchar_t)s[i];
            if (ch == L'\r' || ch == L'\n') {
                // VB6: string content CR/LF written as-is (binary mode, no text-mode conversion)
                vb6_text_putc(f, ch);
                vb6_col_table[filenumber] = 0;
            } else {
                if (w > 0 && vb6_col_table[filenumber] >= w) {
                    vb6_text_putc(f, L'\r');
                    vb6_text_putc(f, L'\n');
                    vb6_col_table[filenumber] = 0;
                }
                vb6_text_putc(f, ch);
                vb6_col_table[filenumber]++;
            }
        }
    }
    // VB6 Print always terminates with CRLF
    vb6_text_putc(f, L'\r');
    vb6_text_putc(f, L'\n');
    vb6_col_table[filenumber] = 0;
    fflush(f);
}
void vb6_Write(int32_t filenumber, BSTR s) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return;
    FILE* f = vb6_file_table[filenumber];
    fputc('"', f);
    if (s) {
        int32_t len = vb6_BSTR_Len(s);
        for (int32_t i = 0; i < len; i++) vb6_text_putc(f, (wchar_t)s[i]);
    }
    fputc('"', f);
    fputc(',', f);  // VB6 Write用逗号分隔
    fflush(f);
}

BSTR vb6_LineInput(int32_t filenumber) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber])
        return vb6_BSTR_Empty();
    FILE* f = vb6_file_table[filenumber];

    // Fix 197: 按本平台文本编码逐字符读。原先 fgets + 逐字节零扩展, 中文文件读回来
    // 是乱码; 而且单行被 4096 字节的栈缓冲硬截断。现在无行长上限, 编码也对。
    int32_t cap = 256, len = 0;
    wchar_t* buf = (wchar_t*)malloc((size_t)cap * sizeof(wchar_t));
    if (!buf) return vb6_BSTR_Empty();

    for (;;) {
        int32_t ch = vb6_text_getc(f);
        if (ch < 0) break;  // EOF (或尾巴上的半截序列)
        if (ch == L'\r' || ch == L'\n') {
            if (ch == L'\r') {
                // CRLF 只算一个换行: 吃掉成对的 LF, 不是 LF 就只推回这一个字节
                int nx = fgetc(f);
                if (nx != '\n' && nx != EOF) ungetc(nx, f);
            }
            break;
        }
        if (len + 1 >= cap) {
            cap *= 2;
            wchar_t* grown = (wchar_t*)realloc(buf, (size_t)cap * sizeof(wchar_t));
            if (!grown) break;
            buf = grown;
        }
        buf[len++] = (wchar_t)ch;
    }
    buf[len] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

int32_t vb6_Input(int32_t filenumber, BSTR* outVar) {
    // 简化: 读取一行
    BSTR result = vb6_LineInput(filenumber);
    if (outVar) *outVar = result;
    return (result != NULL) ? -1 : 0;
}

// P15.4: Input function - reads count characters from file
// Fix <vbeclipse>: 形参序改成 (count, filenumber) —— cgen 按 VB6 源码序发射
// `Input$(3, ch)` → vb6_InputString(3, ch) (2026-10-05 t6.bas 实测, 旧序把 3 当
// 通道号、ch 当字符数, 通道表越界恒返 NULL, Input$ 从来就没读到过东西)。
// InputB 同口径。
BSTR vb6_InputString(int32_t count, int32_t filenumber) {
    // Fix 197: 通道号上界是 VB6_MAX_FILES (表长), 不是 511 —— 原先 filenumber=100
    // 会越界读 vb6_file_table[100]。
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return NULL;
    if (count <= 0) return vb6_BSTR_Empty();
    wchar_t* buf = (wchar_t*)malloc(((size_t)count + 1) * sizeof(wchar_t));
    if (!buf) return NULL;
    int32_t read = 0;
    for (int32_t i = 0; i < count; i++) {
        // Fix 197: count 是**字符**数, 按文本编码解码, 不再按字节零扩展
        int32_t ch = vb6_text_getc(vb6_file_table[filenumber]);
        if (ch < 0) break;
        buf[read++] = (wchar_t)ch;
    }
    buf[read] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

// Fix <vbeclipse>: InputB$ — 按字节读文件 (不经过文本解码), 与 Input$ 的
// "按字符/文本编码"相对。VB6 里 InputB 常配 Binary 通道读协议头/定长记录。
// 载体口径同 LeftB/RightB/MidB: 一个 wchar 存一个字节 (0-255)。
BSTR vb6_InputB(int32_t count, int32_t filenumber) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return NULL;
    if (count <= 0) return vb6_BSTR_Empty();
    wchar_t* buf = (wchar_t*)malloc(((size_t)count + 1) * sizeof(wchar_t));
    if (!buf) return NULL;
    int32_t read = 0;
    for (int32_t i = 0; i < count; i++) {
        int32_t ch = fgetc(vb6_file_table[filenumber]);
        if (ch == EOF) break;
        buf[read++] = (wchar_t)(ch & 0xFF);
    }
    buf[read] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

int32_t vb6_Kill(BSTR pathname) {
    vb6_path path = vb6_path_from(pathname);
    int rc = vb6_path_remove(&path);
    vb6_path_release(&path);
    return (rc == 0) ? -1 : 0;
}

int32_t vb6_MkDir(BSTR pathname) {
    vb6_path path = vb6_path_from(pathname);
    int rc = vb6_path_mkdir(&path);
    vb6_path_release(&path);
    return (rc == 0) ? -1 : 0;
}

int32_t vb6_RmDir(BSTR pathname) {
    vb6_path path = vb6_path_from(pathname);
    int rc = vb6_path_rmdir(&path);
    vb6_path_release(&path);
    return (rc == 0) ? -1 : 0;
}

int32_t vb6_ChDir(BSTR pathname) {
    vb6_path path = vb6_path_from(pathname);
    int rc = vb6_path_chdir(&path);
    vb6_path_release(&path);
    return (rc == 0) ? -1 : 0;
}

int32_t vb6_ChDrive(BSTR drive) {
    // P20-49: 尝试切换驱动器, 失败也返回成功(空操作兼容)
#ifdef _WIN32
    if (drive && SysStringLen(drive) > 0) {
        int driveNum = towupper(drive[0]) - 'A' + 1;
        _chdrive(driveNum);
    }
#endif
    return 0;  // VB6 ChDrive是Sub, 无返回值, 始终返回0
}

int32_t vb6_Name(BSTR oldPath, BSTR newPath) {
    vb6_path a = vb6_path_from(oldPath);
    vb6_path b = vb6_path_from(newPath);
    int rc = vb6_path_rename(&a, &b);
    vb6_path_release(&a);
    vb6_path_release(&b);
    return (rc == 0) ? -1 : 0;
}

int32_t vb6_FileCopy(BSTR source, BSTR destination) {
    vb6_path src = vb6_path_from(source);
    vb6_path dst = vb6_path_from(destination);
    FILE* sf = vb6_path_open(&src, "rb");
    if (!sf) {
        vb6_path_release(&src);
        vb6_path_release(&dst);
        return 0;
    }
    FILE* df = vb6_path_open(&dst, "wb");
    vb6_path_release(&src);
    vb6_path_release(&dst);
    if (!df) { fclose(sf); return 0; }
    char buf[4096];
    size_t n;
    while ((n = fread(buf, 1, sizeof(buf), sf)) > 0) {
        fwrite(buf, 1, n, df);
    }
    fclose(sf);
    fclose(df);
    return -1;
}

// ============================================================
// P8.2: 随机/二进制文件访问 (Get/Put)
// ============================================================

int32_t vb6_Get(int32_t filenumber, int32_t recnumber, void* varPtr, int32_t varSize) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return 0;
    FILE* f = vb6_file_table[filenumber];
    int32_t mode = vb6_file_mode[filenumber];

    if (mode == 4) {
        // Random模式: recnumber是1-based记录号, 按reclength定位
        int32_t reclen = vb6_file_reclen[filenumber];
        if (reclen <= 0) reclen = 128;
        long pos = (long)(recnumber - 1) * reclen;
        fseek(f, pos, SEEK_SET);
        // 读取min(varSize, reclen)字节
        int32_t readLen = (varSize < reclen) ? varSize : reclen;
        size_t n = fread(varPtr, 1, readLen, f);
        // 不足部分填零
        if ((int32_t)n < varSize) {
            memset((char*)varPtr + n, 0, varSize - n);
        }
    } else if (mode == 16) {
        // Binary模式: recnumber是1-based字节位置
        if (recnumber > 0) {
            fseek(f, (long)(recnumber - 1), SEEK_SET);
        }
        fread(varPtr, 1, varSize, f);
    } else {
        return 0;  // 不支持的模式
    }
    return -1;  // True
}

int32_t vb6_Put(int32_t filenumber, int32_t recnumber, void* varPtr, int32_t varSize) {
    if (filenumber < 1 || filenumber >= VB6_MAX_FILES || !vb6_file_table[filenumber]) return 0;
    FILE* f = vb6_file_table[filenumber];
    int32_t mode = vb6_file_mode[filenumber];

    if (mode == 4) {
        // Random模式: recnumber是1-based记录号, 按reclength定位
        int32_t reclen = vb6_file_reclen[filenumber];
        if (reclen <= 0) reclen = 128;
        long pos = (long)(recnumber - 1) * reclen;
        fseek(f, pos, SEEK_SET);
        // 写入min(varSize, reclen)字节, 不足部分填零
        int32_t writeLen = (varSize < reclen) ? varSize : reclen;
        fwrite(varPtr, 1, writeLen, f);
        if (writeLen < reclen) {
            // 记录剩余部分填零
            char zero = 0;
            for (int32_t i = writeLen; i < reclen; i++) fwrite(&zero, 1, 1, f);
        }
    } else if (mode == 16) {
        // Binary模式: recnumber是1-based字节位置
        if (recnumber > 0) {
            fseek(f, (long)(recnumber - 1), SEEK_SET);
        }
        fwrite(varPtr, 1, varSize, f);
    } else {
        return 0;  // 不支持的模式
    }
    fflush(f);
    return -1;  // True
}


// ============================================================
// P18-E: File Locking
// ============================================================

void vb6_Lock(int32_t filenum, int64_t start, int64_t end) {
#ifdef _WIN32
    if (filenum < 1 || filenum >= VB6_MAX_FILES) return;
    FILE* f = vb6_file_table[filenum];
    if (!f) return;

    intptr_t osfhandle = _get_osfhandle(_fileno(f));
    if (osfhandle == -1) return;
    HANDLE h = (HANDLE)osfhandle;

    DWORD64 offset, length;
    if (start <= 0 && end <= 0) {
        offset = 0;
        length = 0x7FFFFFFF;
    } else if (end <= 0) {
        offset = (DWORD64)start;
        length = 0x7FFFFFFF;
    } else {
        offset = (DWORD64)start;
        length = (DWORD64)(end - start + 1);
    }

    OVERLAPPED ov;
    ZeroMemory(&ov, sizeof(ov));
    ov.Offset = (DWORD)offset;
    ov.OffsetHigh = (DWORD)(offset >> 32);
    LockFileEx(h, LOCKFILE_EXCLUSIVE_LOCK, 0, (DWORD)length, (DWORD)(length >> 32), &ov);
#else
    (void)filenum; (void)start; (void)end;
#endif
}

void vb6_Unlock(int32_t filenum, int64_t start, int64_t end) {
#ifdef _WIN32
    if (filenum < 1 || filenum >= VB6_MAX_FILES) return;
    FILE* f = vb6_file_table[filenum];
    if (!f) return;

    intptr_t osfhandle = _get_osfhandle(_fileno(f));
    if (osfhandle == -1) return;
    HANDLE h = (HANDLE)osfhandle;

    DWORD64 offset, length;
    if (start <= 0 && end <= 0) {
        offset = 0;
        length = 0x7FFFFFFF;
    } else if (end <= 0) {
        offset = (DWORD64)start;
        length = 0x7FFFFFFF;
    } else {
        offset = (DWORD64)start;
        length = (DWORD64)(end - start + 1);
    }

    OVERLAPPED ov;
    ZeroMemory(&ov, sizeof(ov));
    ov.Offset = (DWORD)offset;
    ov.OffsetHigh = (DWORD)(offset >> 32);
    UnlockFileEx(h, 0, (DWORD)length, (DWORD)(length >> 32), &ov);
#else
    (void)filenum; (void)start; (void)end;
#endif
}

void vb6_Reset(void) {
    vb6_CloseAll();
}

