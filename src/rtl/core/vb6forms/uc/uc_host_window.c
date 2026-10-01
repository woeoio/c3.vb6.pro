// uc_host_window.c - vb6forms_uc 拆分片：宿主窗口过程 + 类注册 + GDI+ 初始化
//
// 内容 = 拆分前 vb6forms_uc.c 第 459~565 / 567~607 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_uc_internal.h

#include "vb6forms_uc_internal.h"
#include <stdio.h>  /* rev10: C3_UC_TRACE 的 fprintf */

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 宿主窗口过程
// ============================================================

static LRESULT CALLBACK vb6_uc_wndproc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    vb6_UCRec* r = (vb6_UCRec*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    switch (msg) {
        case WM_NCCREATE: {
            CREATESTRUCTW* cs = (CREATESTRUCTW*)lParam;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)cs->lpCreateParams);
            break;
        }
        case WM_SIZE: {
            if (!r || !r->ready) break;
            r->scaleWidth = (int32_t)LOWORD(lParam);
            r->scaleHeight = (int32_t)HIWORD(lParam);
            if (r->desc && r->desc->resize) {
                vb6_UCSaved saved;
                vb6_uc_push(r, &saved);
                vb6_uc_trace("resize.begin", r->desc->typeName, r->me);
                r->desc->resize(r->me);
                vb6_uc_trace("resize.end", r->desc->typeName, r->me);
                vb6_uc_pop(&saved);
            }
            InvalidateRect(hwnd, NULL, FALSE);
            return 0;
        }
        case WM_SHOWWINDOW:
            if (r && r->ready && wParam && r->desc && r->desc->show) {
                vb6_UCSaved saved;
                vb6_uc_push(r, &saved);
                vb6_uc_trace("show.begin", r->desc->typeName, r->me);
                r->desc->show(r->me);
                vb6_uc_trace("show.end", r->desc->typeName, r->me);
                vb6_uc_pop(&saved);
            }
            return 0;
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_RBUTTONUP:
        case WM_MOUSEMOVE: {
            // czUI fix: 把宿主窗口收到的鼠标消息转成 UserControl_MouseDown/Up/Move
            // (此前完全没有转发, 控件无法点击/拖动 — czUI 标题栏拖动、按钮、开关全死)。
            if (!r || !r->ready || !r->desc || !r->me) break;
            void (*hook)(void*, int32_t, int32_t, float, float) = NULL;
            int32_t button = 0;
            switch (msg) {
                case WM_LBUTTONDOWN: hook = r->desc->mouseDown;  button = 1; break;
                case WM_RBUTTONDOWN: hook = r->desc->mouseDown;  button = 2; break;
                case WM_LBUTTONUP:   hook = r->desc->mouseUp;    button = 1; break;
                case WM_RBUTTONUP:   hook = r->desc->mouseUp;    button = 2; break;
                case WM_MOUSEMOVE:   hook = r->desc->mouseMove;  button = 0; break;
            }
            if (!hook) break;
            // 捕获鼠标, 保证按下后拖出窗口仍能收到 UP/MOVE (VB6 隐式行为)
            if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN) SetCapture(hwnd);
            else if (msg == WM_LBUTTONUP || msg == WM_RBUTTONUP) ReleaseCapture();
            float sx = (float)(short)LOWORD(lParam);
            float sy = (float)(short)HIWORD(lParam);
            // 坐标换算到控件当前 ScaleMode (1=Twip 3=Pixel, 其他按像素)
            if (vb6_UserControl_ScaleMode == 1) { sx *= 15.0f; sy *= 15.0f; }
            vb6_UCSaved saved;
            vb6_uc_push(r, &saved);
            // czUI fix: 回调只更新状态; 视觉刷新统一走 WM_PAINT 双缓冲
            // (直接 GetDC 画屏幕与擦除/重绘交错会造成闪烁)
            vb6_UserControl_hDC = NULL;
            hook(r->me, button, 0, sx, sy);
            vb6_uc_pop(&saved);
            InvalidateRect(hwnd, NULL, FALSE);
            UpdateWindow(hwnd);
            return 0;
        }
        case WM_CAPTURECHANGED:
            break;
        case WM_PAINT: {
            if (!r || !r->ready) break;
            if (!r->desc || !r->desc->paint) break;
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            // Fix 113h: 调试捕获 (见上方 vb6_UCDib 说明) — C3_UC_DUMPDIR 未设置时
            // 走原路径 (直接绘制到窗口 DC).
            const char* dumpDir113h = getenv("C3_UC_DUMPDIR");
            vb6_UCDib dib113h;
            int useDib113h = 0;
            if (1) { /* czUI fix: 始终离屏双缓冲, 一次 BitBlt 上屏 (消闪烁) */
                RECT crc;
                GetClientRect(hwnd, &crc);
                int cw113h = (int)(crc.right - crc.left);
                int ch113h = (int)(crc.bottom - crc.top);
                if (cw113h > 0 && ch113h > 0) {
                    vb6_uc_dibCreate(&dib113h, hdc, cw113h, ch113h);
                    if (dib113h.memDC && dib113h.bits) {
                        RECT full113h = { 0, 0, cw113h, ch113h };
                        COLORREF crefFill = (vb6_Ambient_BackColor & 0x80000000L)
                            ? GetSysColor(vb6_Ambient_BackColor & 0xFF)
                            : (COLORREF)vb6_Ambient_BackColor;
                        HBRUSH wb113h = CreateSolidBrush(crefFill);
                        FillRect(dib113h.memDC, &full113h, wb113h);
                        DeleteObject(wb113h);
                        useDib113h = 1;
                    }
                }
            }
            r->hdc = useDib113h ? dib113h.memDC : hdc;
            vb6_UCSaved saved;
            vb6_uc_push(r, &saved);
            vb6_uc_trace("paint.begin", r->desc->typeName, r->me);
            r->desc->paint(r->me);
            vb6_uc_trace("paint.end", r->desc->typeName, r->me);
            vb6_uc_pop(&saved);
            if (useDib113h) {
                BitBlt(hdc, 0, 0, dib113h.w, dib113h.h, dib113h.memDC, 0, 0, SRCCOPY);
                char path113h[1024];
                const int doDump = (dumpDir113h && *dumpDir113h);
                (void)doDump;
                _snprintf(path113h, sizeof(path113h), "%s\\%s_%d_%p.bmp", dumpDir113h,
                          r->desc->typeName, (int)++g_uc_dumpSeq, hwnd);
                path113h[sizeof(path113h) - 1] = '\0';
                if (doDump) {
                    vb6_uc_dibSaveBmp(&dib113h, path113h);
                    // Fix 123: 同步输出整窗合成图 (含 z 序/相对位置)
                    vb6_uc_dumpFormComposite((HWND)GetAncestor(hwnd, GA_ROOT), dumpDir113h);
                }
                vb6_uc_dibDestroy(&dib113h);
            }
            r->hdc = NULL;
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_ERASEBKGND:
            if (r && r->ready) {
                // 用容器底色填充, 避免重绘闪烁 (固定白色会在深色 UI 上闪白)
                RECT rc; GetClientRect(hwnd, &rc);
                COLORREF cref = (vb6_Ambient_BackColor & 0x80000000L)
                                    ? GetSysColor(vb6_Ambient_BackColor & 0xFF)
                                    : (COLORREF)vb6_Ambient_BackColor;
                HBRUSH b = CreateSolidBrush(cref);
                FillRect((HDC)wParam, &rc, b);
                DeleteObject(b);
            }
            return 1;
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC: {
            // czUI fix: 设计器子控件 (txtEmbed 等) 的 BackColor/ForeColor 落在
            // VB6_BackColor/ForeColor 窗口属性上 (vb6_SetControlBackColor), 这里
            // 应用之 — 否则深色 UI 上子编辑框是默认白底黑字。
            HWND child = (HWND)lParam;
            if (!child) break;
            HDC hdc = (HDC)wParam;
            COLORREF fg = (COLORREF)vb6_GetControlForeColor((void*)child);
            COLORREF bg = (COLORREF)vb6_GetControlBackColor((void*)child);
            if (bg & 0x80000000L) bg = GetSysColor(bg & 0xFF);
            SetTextColor(hdc, fg);
            SetBkColor(hdc, bg);
            // 刷子缓存进窗口属性, 避免每条消息泄漏 GDI 句柄
            HBRUSH br = (HBRUSH)GetPropW(child, L"VB6_BgBrush");
            if (!br) {
                br = CreateSolidBrush(bg);
                SetPropW(child, L"VB6_BgBrush", (HANDLE)br);
            }
            return (LRESULT)br;
        }
        case WM_DESTROY:
            // Fix <vbeclipse> rev10 (C3_UC_TRACE): 临时 —— 定位
            // "ShowPerspective 报 No perspective id found" 是否因为本控件的
            // WM_DESTROY 提前跑过 UserControl_Terminate (源码里它会
            // `Set m_Perspectives = Nothing`)。
            if (getenv("C3_UC_TRACE")) {
                fprintf(stderr, "[UC] WM_DESTROY hwnd=%p me=%p hasTerminate=%d\n",
                        (void*)hwnd, r ? r->me : NULL,
                        (r && r->desc && r->desc->terminate) ? 1 : 0);
                fflush(stderr);
            }
            if (r && r->desc && r->desc->terminate) {
                vb6_UCSaved saved;
                vb6_uc_push(r, &saved);
                r->desc->terminate(r->me);
                vb6_uc_pop(&saved);
            }
            break;
        default:
            break;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// Fix 112b: GDI+ 进程级初始化. 图表控件的绘制走 GdipCreateFromHDC, 而
// Charts 2020 自带的 ManageGDIToken 在打完补丁后会 GdiplusShutdown 掉它自己
// 启动的 token —— RTL 这里保一个常驻 token, 保证绘制始终可用.
// gdiplus.h 是 C++ 专用头, 这里按桩文件的做法用 LoadLibrary + GetProcAddress.
typedef struct vb6_GdiplusStartupInput {
    uint32_t GdiplusVersion;
    void*    DebugEventCallback;
    int32_t  SuppressBackgroundThread;
    int32_t  SuppressExternalCodecs;
} vb6_GdiplusStartupInput;

void vb6_uc_gdiplusInit(void) {
    static int done = 0;
    if (done) return;
    done = 1;
    // czUI fix: 环境字体默认名 — Bag 重放 ReadProperty("Font", Ambient.Font)
    // 会把此对象设为控件字体; Name=NULL 时所有 GDI+ 文字静默消失
    if (!g_vb6_UserControl_FontObj.Name)
        g_vb6_UserControl_FontObj.Name = SysAllocString(L"Segoe UI");
    HMODULE mod = LoadLibraryA("gdiplus.dll");
    if (!mod) return;
    long (__stdcall *pStartup)(ULONG_PTR*, const void*, void*) =
        (long (__stdcall *)(ULONG_PTR*, const void*, void*))GetProcAddress(mod, "GdiplusStartup");
    if (!pStartup) return;
    vb6_GdiplusStartupInput si;
    memset(&si, 0, sizeof(si));
    si.GdiplusVersion = 1;
    ULONG_PTR token = 0;
    pStartup(&token, &si, NULL);   /* 常驻 token, 不配对 Shutdown */
}

void vb6_uc_registerClass(HINSTANCE hInst) {
    static int done = 0;
    if (done) return;
    WNDCLASSW wc;
    memset(&wc, 0, sizeof(wc));
    wc.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;  // czUI fix: UserControl_DblClick 需要
    wc.lpfnWndProc = vb6_uc_wndproc;
    wc.hInstance = hInst;
    wc.hCursor = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);
    wc.hbrBackground = NULL;
    wc.lpszClassName = L"VB6_UserControlHost";
    RegisterClassW(&wc);
    done = 1;
}

#ifdef __cplusplus
} // extern "C"
#endif
