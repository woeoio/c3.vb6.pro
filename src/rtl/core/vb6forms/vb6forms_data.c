// vb6forms_data.c — VB6 Data 控件的原生复刻 (ai/029 C29-Data)。
//
// ⚠ 计划书 D4 曾写死 "Data 不做", 理由是 "要么引外部引擎 (装不起零预编译依赖),
//   要么自研 Jet 兼容层 (另一条产品线)"。**这条被推翻** (与 C29-OLE 同一判例):
//   后端走 **ODBC** —— odbc32.dll 是 Windows 自带组件、双架构原生 (System32 /
//   SysWOW64), 用 LoadLibraryW 动态取函数, **零新增 import lib**。Jet 兼容层不需要:
//   SQL 遍历那一小层 (MoveFirst/Next/Prior/Last + BOF/EOF + Fields) 就是 Recordset
//   语义的全部判据面。引擎差异留在连接串里 (见下)。
//
// 引擎选择 (2026-09-26 用户定调): **provider 由用户通过 Connect 属性自己给** ——
// 控件的意义就是方便用户选择。Connect 填**完整 ODBC 连接串**, 例:
//   Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\db\test.mdb
//   Driver={Microsoft Access Text Driver (*.txt, *.csv)};DefaultDir=C:\data
// RTL **不替用户挑引擎**; Connect 为空时才落默认候选 (Text ISAM, 按 32/64 位
// 驱动名差异顺序 fallback —— 仅判据/开箱便利, 不是推荐用法)。
//   x64 默认: `Microsoft Access Text Driver (*.txt, *.csv)` (ACE)
//   x86 默认: `Microsoft Text Driver (*.txt; *.csv)` (系统 Jet, 32 位列表自带)
//   (32/64 位 ACE 不能双装, 所以默认候选按位宽各有一份驱动名)
//
// 判据: tests/c29data (Text Driver + people.csv, 零建库准备)。
// 事件: Reposition (函数指针注册; RTL 在行变化后调用 —— 不可见控件没有 WM_NOTIFY
// 可借, 函数指针是最直接的一层, cgen 在 handler 存在时发射注册)。
//
// W 字符纪律: ODBC 全走 W 版 (SQLDriverConnectW/SQLExecDirectW/SQLDescribeColW/
// SQLGetData with SQL_C_WCHAR); 窗口层全 W; 字体无关 (不可见控件)。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6forms_prop_ctrl.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sql.h>
#include <sqlext.h>

// ===================== 日志 =====================
static int dataTrace(void) {
    static int t = -1;
    if (t < 0) t = (GetEnvironmentVariableW(L"C3_DATA_TRACE", NULL, 0) > 0) ? 1 : 0;
    return t;
}
#define DT_LOG(...) do { if (dataTrace()) { fprintf(stderr, "[C3DATA] "); \
    fprintf(stderr, __VA_ARGS__); fflush(stderr); } } while (0)

// ===================== 实例 =====================
#define VB6_DATA_CLASS L"VB6_DATA"

typedef struct Vb6Data {
    wchar_t databaseName[MAX_PATH * 2];  // DatabaseName (Text Driver = 目录)
    wchar_t recordSource[2048];          // RecordSource (SQL 语句)
    wchar_t connect[512];                // Connect (可空 —— 空则拼默认)

    // ODBC 句柄 (打开后有效)
    HMODULE      odbcLib;
    SQLHENV      hEnv;
    SQLHDBC      hDbc;
    SQLHSTMT     hStmt;
    int32_t      open;
    int32_t      colCount;
    int32_t      curRow;     // 1 基; 0 = BOF 之前, rowCount+1 = EOF 之后
    int32_t      rowCount;
    wchar_t      colNames[32][64];
    // 当前行整行缓存 —— ODBC 的 SQLGetData 有**列号递增**约束 (同一行乱序读低列
    // 会失败, 实测绑定读 col2 后 Fields 读 col1 全空)。行变化时按升序一次读全。
    wchar_t      rowCache[32][1024];
    int32_t      bof, eof;

    // Reposition 事件 (函数指针; cgen 在 handler 存在时发射注册)
    void (*repositionFn)(void);

    // 绑定 (TextBox.DataSource/DataField → 控件句柄 + 字段名)
    struct {
        HWND     ctl;
        wchar_t  field[64];
    } binds[16];
    int32_t      bindCount;
} Vb6Data;

static Vb6Data* dataFromHwnd(HWND hwnd) {
    return (Vb6Data*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
}

// ===================== ODBC 动态加载 =====================
typedef SQLRETURN  (SQL_API *FN_SQLALLOC)(SQLSMALLINT, SQLHANDLE*, SQLHANDLE);
typedef SQLRETURN  (SQL_API *FN_SQLFREE)(SQLSMALLINT, SQLHANDLE);
/* 标准 9 略 8 参签名: (dbc, hwnd, conn, connLen, out, outMax, outLen*, completion)
 * ⚠ 形参顺序/数量必须与 odbc32 导出严格一致 —— 曾把 conn/len 位置颠倒又漏了
 * completion, 实参错位后 &outLen 落在 SMALLINT 形参上, 驱动写 *outLen 时往被截断的
 * 野地址写 → 0xC0000005 (探针 typedef 正确所以探针不崩, 同一 dll 同一序列)。 */
typedef SQLRETURN  (SQL_API *FN_SQLDRIVERCONNW)(SQLHDBC, SQLHWND, SQLWCHAR*, SQLSMALLINT,
                                                SQLWCHAR*, SQLSMALLINT, SQLUSMALLINT*,
                                                SQLUSMALLINT);
typedef SQLRETURN  (SQL_API *FN_SQLDISCONNECT)(SQLHDBC);
typedef SQLRETURN  (SQL_API *FN_SQLEXECDIRECTW)(SQLHSTMT, SQLWCHAR*, SQLINTEGER);
typedef SQLRETURN  (SQL_API *FN_SQLNUMRESULTCOLS)(SQLHSTMT, SQLSMALLINT*);
typedef SQLRETURN  (SQL_API *FN_SQLDESCRIBECOLW)(SQLHSTMT, SQLUSMALLINT, SQLWCHAR*,
                                                 SQLSMALLINT, SQLSMALLINT*, SQLSMALLINT*,
                                                 SQLULEN*, SQLSMALLINT*, SQLSMALLINT*);
typedef SQLRETURN  (SQL_API *FN_SQLFETCH)(SQLHSTMT);
typedef SQLRETURN  (SQL_API *FN_SQLFETCHSCROLL)(SQLHSTMT, SQLSMALLINT, SQLLEN);
typedef SQLRETURN  (SQL_API *FN_SQLGETDATAW)(SQLHSTMT, SQLUSMALLINT, SQLSMALLINT,
                                             SQLPOINTER, SQLLEN, SQLLEN*);
typedef SQLRETURN  (SQL_API *FN_SQLCLOSECURSOR)(SQLHSTMT);
typedef SQLRETURN  (SQL_API *FN_SQLSETCONNECTATTR)(SQLHDBC, SQLINTEGER, SQLPOINTER, SQLINTEGER);
typedef SQLRETURN  (SQL_API *FN_SQLSETSTMTATTR)(SQLHSTMT, SQLINTEGER, SQLPOINTER, SQLINTEGER);
typedef SQLRETURN  (SQL_API *FN_SQLSETENVATTR)(SQLHENV, SQLINTEGER, SQLPOINTER, SQLINTEGER);

static FN_SQLALLOC              p_SQLAllocHandle;
static FN_SQLFREE               p_SQLFreeHandle;
static FN_SQLDRIVERCONNW        p_SQLDriverConnectW;
static FN_SQLDISCONNECT         p_SQLDisconnect;
static FN_SQLEXECDIRECTW        p_SQLExecDirectW;
static FN_SQLNUMRESULTCOLS      p_SQLNumResultCols;
static FN_SQLDESCRIBECOLW       p_SQLDescribeColW;
static FN_SQLFETCH              p_SQLFetch;
static FN_SQLFETCHSCROLL        p_SQLFetchScroll;
static FN_SQLGETDATAW           p_SQLGetData;
static FN_SQLCLOSECURSOR        p_SQLCloseCursor;
static FN_SQLSETCONNECTATTR     p_SQLSetConnectAttrW;
static FN_SQLSETSTMTATTR        p_SQLSetStmtAttr;
static FN_SQLSETENVATTR         p_SQLSetEnvAttr;

static int dataLoadOdbc(Vb6Data* s) {
    if (s->odbcLib) return 1;
    s->odbcLib = LoadLibraryW(L"odbc32.dll");
    if (!s->odbcLib) { DT_LOG("odbc32.dll not loadable\n"); return 0; }
    #define DTA_GET(name) p_##name = (void*)GetProcAddress(s->odbcLib, #name); \
        if (!p_##name) { DT_LOG("missing " #name "\n"); return 0; }
    DTA_GET(SQLAllocHandle)
    DTA_GET(SQLFreeHandle)
    DTA_GET(SQLDriverConnectW)
    DTA_GET(SQLDisconnect)
    DTA_GET(SQLExecDirectW)
    DTA_GET(SQLNumResultCols)
    DTA_GET(SQLDescribeColW)
    DTA_GET(SQLFetch)
    DTA_GET(SQLFetchScroll)
    DTA_GET(SQLGetData)
    DTA_GET(SQLCloseCursor)
    DTA_GET(SQLSetConnectAttrW)
    DTA_GET(SQLSetStmtAttr)
    DTA_GET(SQLSetEnvAttr)
    #undef DTA_GET
    return 1;
}

// ===================== ODBC 生命周期 =====================
static void dataClose(Vb6Data* s) {
    if (s->hStmt) { p_SQLFreeHandle(SQL_HANDLE_STMT, s->hStmt); s->hStmt = NULL; }
    if (s->hDbc)  { p_SQLDisconnect(s->hDbc); p_SQLFreeHandle(SQL_HANDLE_DBC, s->hDbc); s->hDbc = NULL; }
    if (s->hEnv)  { p_SQLFreeHandle(SQL_HANDLE_ENV, s->hEnv); s->hEnv = NULL; }
    s->open = 0;
    s->colCount = 0;
    s->curRow = 0;
    s->rowCount = 0;
    s->bof = s->eof = 0;
}

// 连接串: Text Driver 的两版驱动名不同 (见文件头), 按候选顺序 fallback。
// .mdb/.accdb 场景用户在 Connect 里自己写全 (Connect 优先级最高)。
static int dataOpenConnection(Vb6Data* s) {
    if (!dataLoadOdbc(s)) return 0;

    wchar_t cands[3][512];
    int ncand = 0;
    if (s->connect[0]) {
        _snwprintf(cands[ncand++], 511, L"%ls", s->connect);
    } else {
        _snwprintf(cands[ncand++], 511,
                   L"Driver={Microsoft Access Text Driver (*.txt, *.csv)};DefaultDir=%ls;Extensions=csv",
                   s->databaseName);
        _snwprintf(cands[ncand++], 511,
                   L"Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=%ls;Extensions=csv",
                   s->databaseName);
        _snwprintf(cands[ncand++], 511,
                   L"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%ls",
                   s->databaseName);
    }

    SQLRETURN rc;
    rc = p_SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &s->hEnv);
    DT_LOG("alloc env rc=%d\n", (int)rc);
    if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO) return 0;
    // ODBC 3 环境版本必须显式声明, 否则部分驱动管理器拒发后续调用
    rc = p_SQLSetEnvAttr(s->hEnv, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0);
    DT_LOG("setenvver rc=%d\n", (int)rc);
    rc = p_SQLAllocHandle(SQL_HANDLE_DBC, s->hEnv, &s->hDbc);
    DT_LOG("alloc dbc rc=%d\n", (int)rc);
    if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO) return 0;

    for (int i = 0; i < ncand; i++) {
        SQLWCHAR out[1024];
        SQLSMALLINT outLen = 0;
        DT_LOG("driverconnect cand=%d len=%d\n", i, (int)wcslen(cands[i]));
        rc = p_SQLDriverConnectW(s->hDbc, (SQLHWND)NULL, (SQLWCHAR*)cands[i], (SQLSMALLINT)wcslen(cands[i]),
                                 out, 1024, &outLen, SQL_DRIVER_NOPROMPT);
        DT_LOG("driverconnect returned rc=%d\n", (int)rc);
        if (rc == SQL_SUCCESS || rc == SQL_SUCCESS_WITH_INFO) {
            DT_LOG("open ok cand=%d [%ls]\n", i, cands[i]);
            return 1;
        }
        DT_LOG("open fail cand=%d [%ls]\n", i, cands[i]);
    }
    p_SQLFreeHandle(SQL_HANDLE_DBC, s->hDbc); s->hDbc = NULL;
    p_SQLFreeHandle(SQL_HANDLE_ENV, s->hEnv); s->hEnv = NULL;
    return 0;
}

// ===================== 绑定灌值 =====================
// 行缓存: ODBC 的 SQLGetData 有**列号递增**约束 (同一行乱序读低列会失败 —— 实测
// 绑定先读 col2 后 Fields 读 col1 全空)。行变化时按升序一次读全, 读值全走缓存。
static void dataLoadRowCache(Vb6Data* s) {
    if (s->curRow < 1 || s->curRow > s->rowCount) return;
    for (int c = 1; c <= s->colCount; c++) {
        SQLLEN got = 0;
        SQLRETURN rc = p_SQLGetData(s->hStmt, (SQLUSMALLINT)c, SQL_C_WCHAR,
                                    s->rowCache[c - 1], 1022, &got);
        if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO) s->rowCache[c - 1][0] = 0;
        if (dataTrace()) fprintf(stderr, "[C3DATA] rowcache col%d=[%ls]\n", c, s->rowCache[c - 1]);
    }
}

static void dataFeedBinds(Vb6Data* s) {
    if (s->curRow < 1 || s->curRow > s->rowCount) return;   // BOF/EOF 之间无当前行
    for (int b = 0; b < s->bindCount; b++) {
        for (int c = 0; c < s->colCount; c++) {
            if (_wcsicmp(s->colNames[c], s->binds[b].field) == 0) {
                SetWindowTextW(s->binds[b].ctl, s->rowCache[c]);
                break;
            }
        }
    }
}

static void dataNotifyReposition(Vb6Data* s) {
    if (s->repositionFn) s->repositionFn();
}

// 行定位: 用 SQLFetchScroll 的静态游标 (SQL_CURSOR_STATIC=3)。
// curRow 语义: 1 基; 0 = BOF, rowCount+1 = EOF。
static int dataGoto(Vb6Data* s, SQLSMALLINT how, int32_t targetRow) {
    if (!s->open || !s->hStmt) return 0;
    SQLRETURN rc = p_SQLFetchScroll(s->hStmt, how, 0);
    if (rc == SQL_SUCCESS || rc == SQL_SUCCESS_WITH_INFO) {
        s->curRow = targetRow;
        s->bof = 0; s->eof = 0;
        dataLoadRowCache(s);
        dataFeedBinds(s);
        return 1;
    }
    // NO_DATA: 越过头/尾
    if (how == SQL_FETCH_PRIOR) { s->curRow = 0; s->bof = 1; }
    else if (how == SQL_FETCH_NEXT) { s->curRow = s->rowCount + 1; s->eof = 1; }
    return 0;
}

// ===================== 窗口过程 (属性宿主) =====================
static LRESULT CALLBACK vb6_DataWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    Vb6Data* s = dataFromHwnd(hwnd);
    switch (msg) {
    case WM_DESTROY:
        if (s) { DT_LOG("WM_DESTROY hwnd=%p\n", (void*)hwnd); dataClose(s); if (s->odbcLib) { FreeLibrary(s->odbcLib); s->odbcLib = NULL; } }
        break;
    default:
        break;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

int vb6_RegisterDataClass(void* hInstance) {
    static int done = 0;
    WNDCLASSW wc;
    if (done) return 1;
    done = 1;
    memset(&wc, 0, sizeof(wc));
    wc.lpfnWndProc   = vb6_DataWndProc;
    wc.hInstance     = (HINSTANCE)hInstance;
    wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
    wc.lpszClassName = VB6_DATA_CLASS;
    if (!RegisterClassW(&wc)) return 0;
    return 1;
}

// ===================== 初始化 / 属性 =====================
void vb6_Data_Init(void* hwnd, const wchar_t* databaseName, const wchar_t* recordSource,
                   const wchar_t* connect) {
    Vb6Data* s;
    HWND h = (HWND)hwnd;
    if (!h) return;
    if (!dataFromHwnd(h)) {
        s = (Vb6Data*)calloc(1, sizeof(Vb6Data));
        if (!s) return;
        SetWindowLongPtrW(h, GWLP_USERDATA, (LONG_PTR)s);
    } else {
        s = dataFromHwnd(h);
    }
    if (databaseName) wcsncpy(s->databaseName, databaseName, MAX_PATH * 2 - 1);
    if (recordSource) wcsncpy(s->recordSource, recordSource, 2047);
    if (connect)      wcsncpy(s->connect, connect, 511);
    DT_LOG("Init hwnd=%p db=%ls src=%.96ls\n", (void*)h, s->databaseName, s->recordSource);
}

void vb6_Data_SetDatabaseName(void* hwnd, const wchar_t* v) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (s && v) wcsncpy(s->databaseName, v, MAX_PATH * 2 - 1);
}
void vb6_Data_SetRecordSource(void* hwnd, const wchar_t* v) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (s && v) wcsncpy(s->recordSource, v, 2047);
}
void vb6_Data_SetConnect(void* hwnd, const wchar_t* v) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (s && v) wcsncpy(s->connect, v, 511);
}
wchar_t* vb6_Data_GetDatabaseName(void* hwnd) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    return s ? s->databaseName : NULL;
}
wchar_t* vb6_Data_GetRecordSource(void* hwnd) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    return s ? s->recordSource : NULL;
}
wchar_t* vb6_Data_GetConnect(void* hwnd) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    return s ? s->connect : NULL;
}

// ===================== Recordset 操作 =====================
int vb6_Data_Refresh(void* hwnd) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    DT_LOG("refresh enter hwnd=%p s=%p src0=%d\n", hwnd, (void*)s, s ? (int)s->recordSource[0] : -1);
    if (!s || !s->recordSource[0]) { DT_LOG("refresh early-return s=%p src0=%d\n", (void*)s, s ? (int)s->recordSource[0] : -1); return 0; }
    dataClose(s);
    if (!dataOpenConnection(s)) return 0;

    SQLRETURN rc = p_SQLAllocHandle(SQL_HANDLE_STMT, s->hDbc, &s->hStmt);
    if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO) return 0;
    // 静态游标: SQLFetchScroll 的 PRIOR/LAST 才可用 (forward-only 默认只 NEXT)
    SQLUINTEGER cursor = SQL_CURSOR_STATIC;
    p_SQLSetStmtAttr(s->hStmt, SQL_ATTR_CURSOR_TYPE, (SQLPOINTER)(uintptr_t)cursor, 0);

    rc = p_SQLExecDirectW(s->hStmt, (SQLWCHAR*)s->recordSource, (SQLINTEGER)wcslen(s->recordSource));
    if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO) {
        DT_LOG("exec fail [%ls]\n", s->recordSource);
        p_SQLFreeHandle(SQL_HANDLE_STMT, s->hStmt); s->hStmt = NULL;
        return 0;
    }
    SQLSMALLINT n = 0;
    p_SQLNumResultCols(s->hStmt, &n);
    s->colCount = n;
    for (int c = 1; c <= n && c < 32; c++) {
        SQLSMALLINT nameLen = 0, type = 0, scale = 0, nullable = 0;
        SQLULEN size = 0;
        p_SQLDescribeColW(s->hStmt, (SQLUSMALLINT)c, s->colNames[c - 1], 64, &nameLen,
                          &type, &size, &scale, &nullable);
    }
    // 数行: 重开一次游标数完 (驱动无关、无歧义), 再重开定位第一行。
    // (SQLRowCount 对 SELECT 不可靠; SQL_ATTR_ROW_NUMBER 各驱动实现不齐。)
    {
        int32_t cnt = 0;
        for (;;) {
            SQLRETURN rf = p_SQLFetch(s->hStmt);
            if (rf == SQL_SUCCESS || rf == SQL_SUCCESS_WITH_INFO) cnt++;
            else break;
        }
        s->rowCount = cnt;
        p_SQLCloseCursor(s->hStmt);
        rc = p_SQLAllocHandle(SQL_HANDLE_STMT, s->hDbc, &s->hStmt);
        if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO) return 0;
        p_SQLSetStmtAttr(s->hStmt, SQL_ATTR_CURSOR_TYPE, (SQLPOINTER)(uintptr_t)cursor, 0);
        rc = p_SQLExecDirectW(s->hStmt, (SQLWCHAR*)s->recordSource, (SQLINTEGER)wcslen(s->recordSource));
        if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO) {
            p_SQLFreeHandle(SQL_HANDLE_STMT, s->hStmt); s->hStmt = NULL;
            return 0;
        }
    }
    s->open = 1;   /* ⚠ 忘了这行 = Refresh 成功但 open 恒 0, 后续全被 early-return 静默吞掉 */
    s->curRow = 1;
    s->bof = s->eof = 0;
    if (s->rowCount > 0) dataGoto(s, SQL_FETCH_FIRST, 1);
    dataNotifyReposition(s);
    DT_LOG("refresh rows=%d cols=%d\n", s->rowCount, s->colCount);
    return 1;
}

// VB6 语义: True = -1 —— 判据里有 `Not EOF` (位非), 返回 1 会算错 (~1 = -2 仍真,
// 但语义是脏的; -1/0 与 VB6 全等)。
int32_t vb6_Data_BOF(void* hwnd)          { Vb6Data* s = dataFromHwnd((HWND)hwnd); return s ? (s->bof ? -1 : 0) : 0; }
int32_t vb6_Data_EOF(void* hwnd)          { Vb6Data* s = dataFromHwnd((HWND)hwnd); return s ? (s->eof ? -1 : 0) : 0; }
int32_t vb6_Data_RecordCount(void* hwnd)  { Vb6Data* s = dataFromHwnd((HWND)hwnd); return s ? s->rowCount : 0; }
int32_t vb6_Data_FieldCount(void* hwnd)   { Vb6Data* s = dataFromHwnd((HWND)hwnd); return s ? s->colCount : 0; }
int32_t vb6_Data_CurrentRow(void* hwnd)   { Vb6Data* s = dataFromHwnd((HWND)hwnd); return s ? s->curRow : 0; }

void vb6_Data_MoveFirst(void* hwnd) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (!s || !s->open || s->rowCount == 0) return;
    if (dataGoto(s, SQL_FETCH_FIRST, 1)) dataNotifyReposition(s);
}
void vb6_Data_MoveLast(void* hwnd) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (!s || !s->open || s->rowCount == 0) return;
    if (dataGoto(s, SQL_FETCH_LAST, s->rowCount)) dataNotifyReposition(s);
}
void vb6_Data_MoveNext(void* hwnd) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (!s || !s->open || s->rowCount == 0 || s->eof) return;
    if (dataGoto(s, SQL_FETCH_NEXT, s->curRow + 1)) dataNotifyReposition(s);
}
void vb6_Data_MovePrevious(void* hwnd) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (!s || !s->open || s->rowCount == 0 || s->bof) return;
    if (dataGoto(s, SQL_FETCH_PRIOR, s->curRow - 1)) dataNotifyReposition(s);
}

wchar_t* vb6_Data_FieldName(void* hwnd, int32_t idx) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (!s || idx < 1 || idx > s->colCount) return L"";
    return s->colNames[idx - 1];
}

// Fields("name").Value / Fields(i).Value
void vb6_Data_FieldValue(void* hwnd, const wchar_t* nameOrIndex, wchar_t* out, int32_t outCap) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    out[0] = 0;
    if (!s || !s->open || s->curRow < 1 || s->curRow > s->rowCount) return;
    int col = 0;
    const wchar_t* p = nameOrIndex;
    int allDigit = (*p != 0);
    for (const wchar_t* q = p; *q; q++) if (!(*q >= L'0' && *q <= L'9')) { allDigit = 0; break; }
    if (allDigit) {
        col = (int)wcstol(p, NULL, 10);
    } else {
        for (int c = 0; c < s->colCount; c++)
            if (_wcsicmp(s->colNames[c], nameOrIndex) == 0) { col = c + 1; break; }
    }
    if (col < 1 || col > s->colCount) return;
    SQLLEN got = 0;
    SQLRETURN rc = p_SQLGetData(s->hStmt, (SQLUSMALLINT)col, SQL_C_WCHAR,
                                out, (SQLLEN)outCap - 2, &got);
    if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO) out[0] = 0;
}

// 账 #278 §B127: 这里从前有两枚出口 —— 一枚把 Recordset 求值透传成控件自己(供发码侧认前缀)，
// 一枚把 Fields("x").Value 折成静态缓冲的宽串。第十五刀量到发码侧那套认前缀的判据恒假并已撤，
// 两枚出口从此全仓 0 个调用者 ⇒ 一起撤。Recordset 成员面唯一的路是 vb6_Data_RecordsetObj
// (下面 20 行) → RTL memberobj 的真 IDispatch；Fields("x") / BOF / RecordCount 都由那张成员名表应答。
// 按下标取字段值 (Field.Value 的后端) —— 走整行缓存 (SQLGetData 列序约束)
void vb6_Data_FieldValueByIdx(void* hwnd, int32_t idx, wchar_t* out, int32_t outCap) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    out[0] = 0;
    if (!s || !s->open || idx < 1 || idx > s->colCount) return;
    if (s->curRow < 1 || s->curRow > s->rowCount) return;
    wcsncpy(out, s->rowCache[idx - 1], outCap - 1);
    out[outCap - 1] = 0;
}

// C29-Data: Recordset 真 IDispatch (实现见 vb6forms_memberobj.c 的 VB6_MEMK_RS 族)
void* vb6_Data_RecordsetObj(void* hwnd) {
    extern void* vb6_MemberObj_NewRs(void* hwnd);
    if (!hwnd) return NULL;
    return vb6_MemberObj_NewRs(hwnd);
}

// ===================== 绑定 / 事件 =====================
void vb6_Data_Bind(void* hwnd, void* ctlHwnd, const wchar_t* fieldName) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (!s || !ctlHwnd || !fieldName || s->bindCount >= 16) return;
    for (int i = 0; i < s->bindCount; i++)
        if (s->binds[i].ctl == (HWND)ctlHwnd) return;   // 一控件一字段
    wcsncpy(s->binds[s->bindCount].field, fieldName, 63);
    s->binds[s->bindCount].ctl = (HWND)ctlHwnd;
    s->bindCount++;
    // 若已打开, 立刻灌当前行
    if (s->open) dataFeedBinds(s);
}

void vb6_Data_SetRepositionHandler(void* hwnd, void* fn) {
    Vb6Data* s = dataFromHwnd((HWND)hwnd);
    if (s) s->repositionFn = (void (*)(void))fn;
}

