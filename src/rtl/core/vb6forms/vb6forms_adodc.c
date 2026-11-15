/* vb6forms_adodc.c —— Adodc（Microsoft ADO Data Control 6.0，MSADODC.OCX）复刻。C30-5。
 *
 * 底座：**ACE（Microsoft Access Database Engine）**，走它的 **ODBC** 那一层。
 *
 * 为什么不是 OLEDB —— 本机实测（cscript 64 位）：
 *   CreateObject("ADODB.Connection")        → ok
 *   conn.Open("Provider=Microsoft.ACE.OLEDB.12.0;...") → ok（State=1）
 *   rs.Open / conn.Execute                  → 0x80040154 Class not registered
 * ACE 装在 `...\Microsoft Office\root\VFS\ProgramFilesCommonX64\Microsoft Shared\Office16\ACEOLEDB.DLL`
 * （Office Click-to-Run 的虚拟路径），provider 自己能被建起来、连接也通，但 provider 内部再要
 * 一档 Recordset 就落空。同引擎的 **ODBC** 驱动是 64 位正常注册的，实测
 *   Driver={Microsoft Access Text Driver (*.txt, *.csv)};Dbq=<目录>;Extensions=csv,txt;HDR=Yes
 *   SELECT id, name, dept FROM [people.csv] ORDER BY id  ⇒ 3 行
 * 所以本文件把 ConnectionString 里的 OLEDB 串**翻译**成等价的 ODBC 串，"底座"仍是 ACE，
 * 只是换了一层可用的入口。32 位没有 ACE（cscript x86: 3706 未找到提供程序），如实保留。
 *
 * 与 Data 控件的关系：真引擎不重写 —— 本文件给每枚 Adodc 建**一个隐藏的 VB6_DATA 窗口**，
 * 所有 ODBC 活在 vb6forms_data.c 里（连接、游标、行缓存、字段读）。这样
 *   ① 只有一套 ODBC 代码要维护；
 *   ② `Adodc1.Recordset` 直接交 `vb6_Data_RecordsetObj(engine)` —— Recordset 成员对象
 *      （VB6_MEMK_RS：Fields/MoveFirst/RecordCount…）原样可用，晚绑定链不用再写一遍。
 */

#include <windows.h>
#include <oleauto.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6forms_prop_ctrl.h"

#define VB6_ADODC_CLASS L"VB6_ADODC"
#define VB6_ADODC_MAX   32

// 诊断：C3_ADO_TRACE=1 时把 Refresh 的每一步打到 stderr。默认全关（发码/性能不受影响）。
static int adodcTrace(void) {
    static int t = -1;
    if (t < 0) { const char* e = getenv("C3_ADO_TRACE"); t = (e && *e == '1') ? 1 : 0; }
    return t;
}
#define AT_LOG(...) do { if (adodcTrace()) { fprintf(stderr, __VA_ARGS__); fflush(stderr); } } while (0)

// 事件号（与 VB6 的 WillMove/MoveComplete 对齐；C30-5 首版只发 MoveComplete）
#define VB6_ADODC_EV_MOVECOMPLETE 0
#define VB6_ADODC_EV_COUNT        1

struct vb6_Adodc {
    HWND      hwnd;                 // 身份窗；NULL = 空槽
    HWND      engine;               // 隐藏的 VB6_DATA 窗口（真引擎 + Recordset 成员对象宿主）
    wchar_t   connStr[1024];        // ConnectionString（原样回读，不做规范化）
    wchar_t   recordSource[2048];   // RecordSource
    int32_t   commandType;          // 0 adCmdUnknown / 1 adCmdText / 2 adCmdTable / 4 adCmdStoredProc
    int32_t   cursorLocation;       // 2 adUseServer / 3 adUseClient
    int32_t   cursorType;           // 0 ForwardOnly / 1 Keyset / 2 Dynamic / 3 Static
    int32_t   lockType;             // 1 ReadOnly / 2 Pessimistic / 3 Optimistic / 4 BatchOptimistic
    int32_t   mode;                 // 1 只读 / 2 只写 / 3 读写
    int32_t   maxRecords;           // 0 = 不限（本实现存值，不截断；见 ai/029 C30-5 的"未做"）
    int32_t   bofAction;            // 0 移首 / 1 停留
    int32_t   eofAction;            // 0 移末 / 1 停留 / 2 自动 AddNew
    int32_t   opened;               // 上一次 Refresh 成功？
    int32_t   lastRows;             // 上一次 Refresh 拿到的行数（夹具证人；不是 VB6 属性）
    void     (*eventFn[VB6_ADODC_EV_COUNT])(void);
};

static struct vb6_Adodc g_adodc[VB6_ADODC_MAX];

static struct vb6_Adodc* adodcOf(HWND hwnd) {
    int i;
    if (!hwnd) return NULL;
    for (i = 0; i < VB6_ADODC_MAX; i++)
        if (g_adodc[i].hwnd == hwnd) return &g_adodc[i];
    return NULL;
}

// 取槽；没有就占一个。占不到（>32 枚）返回 NULL —— 静默不建，调用方各自早退。
static struct vb6_Adodc* adodcSlot(HWND hwnd) {
    struct vb6_Adodc* a = adodcOf(hwnd);
    int i;
    if (a || !hwnd) return a;
    for (i = 0; i < VB6_ADODC_MAX; i++) {
        if (!g_adodc[i].hwnd) {
            memset(&g_adodc[i], 0, sizeof(g_adodc[i]));
            g_adodc[i].hwnd = hwnd;
            g_adodc[i].commandType = 1;      // VB6 缺省 adCmdText
            g_adodc[i].cursorLocation = 2;   // adUseServer
            g_adodc[i].cursorType = 0;       // adOpenForwardOnly
            g_adodc[i].lockType = 3;         // adLockOptimistic（VB6 缺省）
            g_adodc[i].mode = 3;             // adModeReadWrite
            return &g_adodc[i];
        }
    }
    return NULL;
}

/* ConnectionString → ODBC 连接串。认三种写法：
 *   1) 已经是 ODBC 串（`Driver=` 开头）—— 原样交出去（用户自己写全的优先）；
 *   2) ACE/Jet 的 OLEDB 串且 Extended Properties 里有 text ⇒ ACE 的 Text 驱动；
 *   3) ACE/Jet 的 OLEDB 串（其它）⇒ ACE 的 Access 驱动，Dbq = Data Source。
 * 翻译不了（没有 Provider、也不是 Driver=）返回 0 —— 打开失败，不猜。 */
static int adodcValue(const wchar_t* conn, const wchar_t* key, wchar_t* out, int cap) {
    const wchar_t* p = conn;
    size_t klen = wcslen(key);
    if (cap <= 0) return 0;
    out[0] = 0;
    while (p && *p) {
        const wchar_t* eq;
        const wchar_t* end;
        const wchar_t* seg = p;
        // 段尾：分号（引号内的分号不算 —— "text;HDR=Yes" 是常见写法）
        int inq = 0;
        while (*p && (inq || *p != L';')) {
            if (*p == L'"') inq = !inq;
            p++;
        }
        end = p;
        eq = seg;
        while (eq < end && *eq != L'=') eq++;
        if (eq < end) {
            size_t n = (size_t)(eq - seg);
            // key 两端可能有空格
            while (n && seg[0] == L' ') { seg++; n--; }
            while (n && seg[n - 1] == L' ') n--;
            if (n == klen && _wcsnicmp(seg, key, klen) == 0) {
                const wchar_t* v = eq + 1;
                const wchar_t* ve = end;
                int j = 0;
                while (v < ve && *v == L' ') v++;
                while (ve > v && ve[-1] == L' ') ve--;
                // 去掉值外面的引号
                if (ve - v >= 2 && *v == L'"' && ve[-1] == L'"') { v++; ve--; }
                for (; v < ve && j < cap - 1; v++) out[j++] = *v;
                out[j] = 0;
                return j > 0;
            }
        }
        if (*p == L';') p++;
    }
    return 0;
}

static int adodcTranslate(const wchar_t* conn, wchar_t* out, int cap) {
    wchar_t provider[128], src[1024], ext[256];
    if (!out || cap < 32) return 0;
    out[0] = 0;
    if (!conn || !conn[0]) return 0;
    // 1) 已是 ODBC 串
    {
        const wchar_t* p = conn;
        while (*p == L' ') p++;
        if (_wcsnicmp(p, L"Driver=", 7) == 0) {
            _snwprintf(out, (size_t)cap - 1, L"%ls", p);
            return 1;
        }
    }
    if (!adodcValue(conn, L"Provider", provider, 128)) return 0;
    if (!adodcValue(conn, L"Data Source", src, 1024)) return 0;
    adodcValue(conn, L"Extended Properties", ext, 256);

    if (_wcsicmp(provider, L"Microsoft.ACE.OLEDB.12.0") == 0 ||
        _wcsicmp(provider, L"Microsoft.ACE.OLEDB.16.0") == 0 ||
        _wcsicmp(provider, L"Microsoft.ACE.OLEDB.15.0") == 0 ||
        _wcsicmp(provider, L"Microsoft.Jet.OLEDB.4.0") == 0 ||
        _wcsicmp(provider, L"Microsoft.Jet.OLEDB.3.51") == 0) {
        if (wcsstr(ext, L"text") || wcsstr(ext, L"Text") || wcsstr(ext, L"TEXT") ||
            wcsstr(ext, L"csv") || wcsstr(ext, L"CSV")) {
            // Text 驱动：Dbq 是**目录**，表名用 [文件名]；HDR/分隔符沿用 Extended Properties
            int hdr = 1;
            if (wcsstr(ext, L"HDR=No") || wcsstr(ext, L"HDR=no")) hdr = 0;
            _snwprintf(out, (size_t)cap - 1,
                       L"Driver={Microsoft Access Text Driver (*.txt, *.csv)};"
                       L"DefaultDir=%ls;Extensions=csv,txt;HDR=%ls;FMT=Delimited",
                       src, hdr ? L"Yes" : L"No");
            return 1;
        }
        _snwprintf(out, (size_t)cap - 1,
                   L"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=%ls", src);
        return 1;
    }
    return 0;
}

static LRESULT CALLBACK vb6_AdodcWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_DESTROY: {
        struct vb6_Adodc* a = adodcOf(hwnd);
        if (a) {
            if (a->engine) { DestroyWindow(a->engine); a->engine = NULL; }
            memset(a, 0, sizeof(*a));
        }
        break;
    }
    default: break;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void vb6_RegisterAdodcClass(void* hInstance) {
    static int done = 0;
    WNDCLASSW wc;
    if (done) return;
    done = 1;
    memset(&wc, 0, sizeof(wc));
    wc.lpfnWndProc   = vb6_AdodcWndProc;
    wc.hInstance     = (HINSTANCE)hInstance;
    wc.hCursor       = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);   // IDC_ARROW 在无 UNICODE 档是 LPSTR，显式转掉 C4133
    wc.lpszClassName = VB6_ADODC_CLASS;
    RegisterClassW(&wc);   // 已注册（第二个 Adodc）时失败也无妨
}

// 隐藏引擎窗：类 VB6_DATA（vb6forms_data.c 注册），看不见、不参与布局，只当 ODBC 句柄的宿主。
static int adodcEnsureEngine(struct vb6_Adodc* a) {
    if (!a) return 0;
    if (a->engine) return 1;
    vb6_RegisterDataClass(GetModuleHandleW(NULL));
    a->engine = CreateWindowExW(0, L"VB6_DATA", L"", WS_POPUP, 0, 0, 0, 0,
                                NULL, NULL, GetModuleHandleW(NULL), NULL);
    // ⚠ 必须紧跟一句 vb6_Data_Init：Data 那一族的每个出口都是 dataFromHwnd(hwnd) 起头，
    // 槽（GWLP_USERDATA 里那枚 Vb6Data）不存在时 SetConnect/SetRecordSource/Refresh
    // **全部静默空转**。实测过：只建窗不 Init ⇒ Refresh 恒 0、报表“数据 refresh=0”。
    if (a->engine) vb6_Data_Init(a->engine, NULL, NULL, NULL);
    return a->engine != NULL;
}

void vb6_Adodc_Init(void* hwnd) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd);
    if (a) adodcEnsureEngine(a);
}

void vb6_Adodc_SetEventHandler(void* hwnd, int32_t kind, void* fn) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd);
    if (a && kind >= 0 && kind < VB6_ADODC_EV_COUNT) a->eventFn[kind] = (void (*)(void))fn;
}

// ============================ 属性 ============================
const void* vb6_Adodc_GetConnectionString(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd);
    return a ? SysAllocString(a->connStr) : SysAllocString(L"");
}
void vb6_Adodc_SetConnectionString(void* hwnd, const void* bstr) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd);
    const wchar_t* v = (const wchar_t*)bstr;
    if (a && v) { wcsncpy(a->connStr, v, 1023); a->connStr[1023] = 0; }
}

const void* vb6_Adodc_GetRecordSource(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd);
    return a ? SysAllocString(a->recordSource) : SysAllocString(L"");
}
void vb6_Adodc_SetRecordSource(void* hwnd, const void* bstr) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd);
    const wchar_t* v = (const wchar_t*)bstr;
    if (a && v) { wcsncpy(a->recordSource, v, 2047); a->recordSource[2047] = 0; }
}

int32_t vb6_Adodc_GetCommandType(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? a->commandType : 1;
}
void vb6_Adodc_SetCommandType(void* hwnd, int32_t v) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd); if (a) a->commandType = v;
}

int32_t vb6_Adodc_GetCursorLocation(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? a->cursorLocation : 2;
}
void vb6_Adodc_SetCursorLocation(void* hwnd, int32_t v) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd); if (a) a->cursorLocation = v;
}

int32_t vb6_Adodc_GetCursorType(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? a->cursorType : 0;
}
void vb6_Adodc_SetCursorType(void* hwnd, int32_t v) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd); if (a) a->cursorType = v;
}

int32_t vb6_Adodc_GetLockType(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? a->lockType : 3;
}
void vb6_Adodc_SetLockType(void* hwnd, int32_t v) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd); if (a) a->lockType = v;
}

int32_t vb6_Adodc_GetMode(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? a->mode : 3;
}
void vb6_Adodc_SetMode(void* hwnd, int32_t v) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd); if (a) a->mode = v;
}

int32_t vb6_Adodc_GetMaxRecords(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? a->maxRecords : 0;
}
void vb6_Adodc_SetMaxRecords(void* hwnd, int32_t v) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd); if (a) a->maxRecords = v;
}

int32_t vb6_Adodc_GetBOFAction(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? a->bofAction : 0;
}
void vb6_Adodc_SetBOFAction(void* hwnd, int32_t v) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd); if (a) a->bofAction = v;
}

int32_t vb6_Adodc_GetEOFAction(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? a->eofAction : 0;
}
void vb6_Adodc_SetEOFAction(void* hwnd, int32_t v) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd); if (a) a->eofAction = v;
}

// 夹具证人（C3 扩展，不是 VB6 属性）：上一次 Refresh 真拿到几行。
int32_t vb6_Adodc_GetLastRowCount(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? a->lastRows : 0;
}
int32_t vb6_Adodc_GetOpened(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd); return a ? (a->opened ? -1 : 0) : 0;
}

// ============================ 方法 ============================
// Refresh：翻译连接串 → 交给引擎 → 真开游标。成功返回 1（cgen 发成语句，不取返回值）。
int32_t vb6_Adodc_Refresh(void* hwnd) {
    struct vb6_Adodc* a = adodcSlot((HWND)hwnd);
    wchar_t odbc[1200];
    wchar_t sql[2100];
    int ok;
    if (!a) return 0;
    a->opened = 0;
    a->lastRows = 0;
    if (!adodcEnsureEngine(a)) { AT_LOG("adodc: engine create failed\n"); return 0; }
    if (!adodcTranslate(a->connStr, odbc, 1200)) {
        AT_LOG("adodc: translate failed conn=[%ls]\n", a->connStr);
        // 连接串翻不出来 = 打开失败。**不静默降级**：opened 保持 0，RecordCount 会是 0，
        // 夹具一眼看得出。VB6 在这种情形是 raise 一个 "未找到提供程序" 类错误。
        return 0;
    }
    // adCmdTable：RecordSource 是**表名**，引擎只吃 SQL，这里补成 SELECT *。
    if (a->commandType == 2 && a->recordSource[0]) {
        const wchar_t* t = a->recordSource;
        while (*t == L' ') t++;
        if (_wcsnicmp(t, L"SELECT", 6) == 0 || _wcsnicmp(t, L"select", 6) == 0) {
            _snwprintf(sql, 2099, L"%ls", t);
        } else {
            _snwprintf(sql, 2099, L"SELECT * FROM [%ls]", t);
        }
    } else {
        _snwprintf(sql, 2099, L"%ls", a->recordSource);
    }
    vb6_Data_SetConnect(a->engine, odbc);
    vb6_Data_SetRecordSource(a->engine, sql);
    AT_LOG("adodc: odbc=[%ls] sql=[%ls]\n", odbc, sql);
    ok = vb6_Data_Refresh(a->engine);
    AT_LOG("adodc: data refresh=%d\n", ok);
    a->opened = ok;
    a->lastRows = ok ? vb6_Data_RecordCount(a->engine) : 0;
    if (ok && a->eventFn[VB6_ADODC_EV_MOVECOMPLETE]) a->eventFn[VB6_ADODC_EV_MOVECOMPLETE]();
    return ok;
}

// Adodc1.Recordset：交引擎的真 IDispatch（VB6_MEMK_RS）。晚绑定链（Fields/MoveNext/RecordCount）照走。
void* vb6_Adodc_Recordset(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd);
    if (!a || !adodcEnsureEngine(a)) return NULL;
    return vb6_Data_RecordsetObj(a->engine);
}

// 绑定用：DataGrid 拿到的是**引擎句柄**（Data 那一族 API 的入参）。
void* vb6_Adodc_Engine(void* hwnd) {
    struct vb6_Adodc* a = adodcOf((HWND)hwnd);
    if (!a) return NULL;
    return adodcEnsureEngine(a) ? a->engine : NULL;
}
