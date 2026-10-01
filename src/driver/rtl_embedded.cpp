// P10 (restored): RTL runtime embedded resource management - implementation
// Extract 7 .h headers + 25 .c sources from c3.exe RCDATA resources.
// The P11.3 pre-compiled .lib scheme was reverted (c3 is open source);
// RTL sources are compiled by MSVC together with the generated code.

#include "driver/rtl_embedded.hpp"
#include "common/encoding.hpp"

#include <iostream>
#include <fstream>
#include <filesystem>
#include <chrono>
#include <cstdlib>
#include <vector>
#include <string>

#ifdef _WIN32
#include <windows.h>
#endif

namespace vb6c3 {

// ============================================================
// Helper: load RCDATA resource from current EXE
// ============================================================
static bool loadRtlResource(int resId, std::vector<char>& outData) {
#ifdef _WIN32
    HRSRC hrsrc = FindResourceW(nullptr, MAKEINTRESOURCEW(resId), RT_RCDATA);
    if (!hrsrc) {
        // Fallback: try loading from current module (when loaded as DLL)
        HMODULE hMod = nullptr;
        GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
                           reinterpret_cast<LPCWSTR>(&loadRtlResource), &hMod);
        if (hMod) {
            hrsrc = FindResourceW(hMod, MAKEINTRESOURCEW(resId), RT_RCDATA);
        }
    }
    if (!hrsrc) return false;

    HGLOBAL hGlobal = LoadResource(nullptr, hrsrc);
    if (!hGlobal) return false;

    DWORD size = SizeofResource(nullptr, hrsrc);
    const void* ptr = LockResource(hGlobal);
    if (!ptr || size == 0) return false;

    outData.assign(static_cast<const char*>(ptr), static_cast<const char*>(ptr) + size);
    return true;
#else
    // Non-Windows: resources not available, fallback to file system
    (void)resId;
    (void)outData;
    return false;
#endif
}

// ============================================================
// SessionManager implementation
// ============================================================

SessionManager::SessionManager() = default;

SessionManager::~SessionManager() {
    cleanup();
}

std::string SessionManager::getSessionRoot() {
#ifdef _WIN32
    // Fix 196: 环境变量必须按**宽字符**读。CRT 的 getenv 返回的是 ACP 字节, 而 C3
    // 内部统一 UTF-8 —— %TEMP% 一旦含非 ASCII (典型: 中文用户名 C:\Users\张三\...),
    // 这些 ACP 字节会被 executeCommand 的 MultiByteToWideChar(CP_UTF8) 当 UTF-8 解码,
    // 临时目录在命令行里变成乱码 -> cmd 的重定向目标非法 -> **cl 根本没执行**,
    // 只报 exit code 1 且零输出 (实测)。这里转成 UTF-8, 与其余路径口径一致。
    auto envW = [](const wchar_t* name) -> std::string {
        wchar_t buf[32768];
        DWORD n = GetEnvironmentVariableW(name, buf, 32768);
        if (n == 0 || n >= 32768) return std::string();
        return wideToUtf8(std::wstring(buf, n));
    };
    std::string tmp = envW(L"TMP");
    if (tmp.empty()) tmp = envW(L"TEMP");
    if (tmp.empty()) {
        wchar_t tbuf[32768];
        DWORD n = GetTempPathW(32768, tbuf);
        if (n > 0) tmp = wideToUtf8(std::wstring(tbuf, n));
    }
    // GetTempPathW 的返回值以 '\' 结尾; 手写 TMP 可能也带尾分隔符
    while (!tmp.empty() && (tmp.back() == '\\' || tmp.back() == '/')) tmp.pop_back();
    if (tmp.empty()) tmp = "C:\\Temp";
    return tmp + "\\C3C";
#else
    const char* tmp = std::getenv("TMP");
    if (!tmp || tmp[0] == '\0') tmp = std::getenv("TEMP");
    if (!tmp || tmp[0] == '\0') tmp = "C:\\Temp";
    return std::string(tmp) + "\\C3C";
#endif
}

// ============================================================
// RTL 解包清单: resId -> 解包后的文件名
// 解包目录是平铺的 (与源码树的 axsite/ uc/ 子目录无关), 故这里记 basename.
// create() 与 restoreMissing() 必须共用同一份清单 —— 两处各写一份就会出现
// "解包了但没编译" / "编译了但没解包" 的错配, 后果是 cl 的 C1083.
// ============================================================
struct RtlFileEntry { int id; const char* name; };

static const RtlFileEntry kRtlFiles[] = {
    { RTL_VB6RTL_H,             "vb6rtl.h" },
    { RTL_VB6RTL_C,             "vb6rtl.c" },
    { RTL_VB6COM_H,             "vb6com.h" },
    { RTL_VB6COM_C,             "vb6com.c" },
    { RTL_VB6COMSERVER_H,       "vb6comserver.h" },
    { RTL_VB6COMSERVER_C,       "vb6comserver.c" },
    { RTL_VB6FORMS_H,           "vb6forms.h" },
    { RTL_VB6FORMS_C,           "vb6forms.c" },
    { RTL_VB6_DI_STUBS_C,       "vb6_di_stubs.c" },
    { RTL_VB6_DI_WIN32_STUBS_C, "vb6_di_win32_stubs.c" },
    { RTL_VB6_DI_USER32_STUBS_C,   "vb6_di_user32_stubs.c" },
    { RTL_VB6_DI_GDIPLUS_DRAW_STUBS_C,  "vb6_di_gdiplus_draw_stubs.c" },
    { RTL_VB6_DI_GDIPLUS_TEXT_STUBS_C,  "vb6_di_gdiplus_text_stubs.c" },
    { RTL_VB6_DI_GDIPLUS_IMAGE_STUBS_C, "vb6_di_gdiplus_image_stubs.c" },
    { RTL_VB6_DI_GDIPLUS_BRUSH_STUBS_C, "vb6_di_gdiplus_brush_stubs.c" },
    { RTL_VB6_DI_GDIPLUS_PATH_STUBS_C,  "vb6_di_gdiplus_path_stubs.c" },
    { RTL_VB6_DI_CRYPTO_STUBS_C,   "vb6_di_crypto_stubs.c" },
    { RTL_VB6_DI_COM_STUBS_C,      "vb6_di_com_stubs.c" },
    { RTL_VB6_DI_NET_STUBS_C,      "vb6_di_net_stubs.c" },
    { RTL_VB6_DI_SHELL_STUBS_C,    "vb6_di_shell_stubs.c" },
    { RTL_VB6_DI_UNKNOWN_STUBS_C,  "vb6_di_unknown_stubs.c" },
    { RTL_VB6COMSERVER_INTERNAL_H, "vb6comserver_internal.h" },
    { RTL_VB6COMSERVER_OBJ_C,      "vb6comserver_obj.c" },
    { RTL_VB6COMSERVER_FACTORY_C,  "vb6comserver_factory.c" },
    { RTL_VB6COMSERVER_CP_C,       "vb6comserver_cp.c" },
    { RTL_VB6COMSERVER_PCI_C,      "vb6comserver_pci.c" },
    { RTL_VB6COM_INTERNAL_H,       "vb6com_internal.h" },
    { RTL_VB6COM_INVOKE_C,         "vb6com_invoke.c" },
    { RTL_VB6COM_PACK_C,           "vb6com_pack.c" },
    { RTL_VB6COM_WRAP_C,           "vb6com_wrap.c" },
    { RTL_VB6COM_SINK_C,           "vb6com_sink.c" },
    { RTL_VB6COM_FOREACH_C,        "vb6com_foreach.c" },
    { RTL_VB6COM_COLLECTION_C,     "vb6com_collection.c" },
    { RTL_VB6COM_COLLECTION_ENUM_C, "vb6com_collection_enum.c" },
    { RTL_VB6FORMS_INTERNAL_H,     "vb6forms_internal.h" },
    { RTL_VB6FORMS_CTRL_C,         "vb6forms_ctrl.c" },
    { RTL_VB6FORMS_LIST_C,         "vb6forms_list.c" },
    { RTL_VB6FORMS_STYLE_C,        "vb6forms_style.c" },
    { RTL_VB6FORMS_SCROLL_C,       "vb6forms_scroll.c" },
    { RTL_VB6FORMS_PICTURE_C,      "vb6forms_picture.c" },
    { RTL_VB6FORMS_CTRLARR_C,      "vb6forms_ctrlarr.c" },
    { RTL_VB6FORMS_WEBVIEW_C,      "vb6forms_webview.c" },
    { RTL_VB6FORMS_WIDGET_C,       "vb6forms_widget.c" },
    { RTL_VB6FORMS_SHAPE_C,        "vb6forms_shape.c" },
    { RTL_VB6FORMS_PROGRESS_C,     "vb6forms_progress.c" },
    { RTL_VB6FORMS_IMAGELIST_C,    "vb6forms_imagelist.c" },
    { RTL_VB6FORMS_STATUSBAR_C,    "vb6forms_statusbar.c" },
    { RTL_VB6FORMS_SSTAB_C,        "vb6forms_sstab.c" },
    { RTL_VB6FORMS_OLEDD_C,        "vb6forms_oledd.c" },
    { RTL_VB6FORMS_LISTVIEW_C,     "vb6forms_listview.c" },
    { RTL_VB6FORMS_MEMBEROBJ_C,    "vb6forms_memberobj.c" },
    { RTL_VB6FORMS_TREEVIEW_C,     "vb6forms_treeview.c" },
    { RTL_VB6FORMS_OLECON_C,       "vb6forms_olecon.c" },
    { RTL_VB6FORMS_DATA_C,         "vb6forms_data.c" },
    { RTL_VB6FORMS_TOOLBAR_C,      "vb6forms_toolbar.c" },
    { RTL_VB6FORMS_DTPICKER_C,     "vb6forms_dtpicker.c" },
    { RTL_VB6FORMS_MONTHVIEW_C,    "vb6forms_monthview.c" },
    { RTL_VB6FORMS_RICHTEXTBOX_C,  "vb6forms_richtextbox.c" },
    { RTL_VB6FORMS_WINSOCK_C,      "vb6forms_winsock.c" },
    { RTL_VB6FORMS_SLIDER_C,       "vb6forms_slider.c" },
    { RTL_VB6FORMS_AXSITE_C,       "vb6forms_axsite.c" },
    { RTL_VB6FORMS_AXCONTAINER_C,  "vb6forms_axcontainer.c" },
    // vb6forms_axsite.c 按功能家族拆分 (2026-09-20): 内部头 + axsite/ 下 5 个族编译单元
    // 注意: 解包后是平铺目录, files[] 记的是解包后的 basename, 与源码树子目录无关
    { RTL_VB6FORMS_AXSITE_INTERNAL_H, "vb6forms_axsite_internal.h" },
    { RTL_AX_SITE_C,                  "ax_site.c" },
    { RTL_AX_SITE_EXT_C,              "ax_site_ext.c" },
    { RTL_AX_PROPBAG_C,               "ax_propbag.c" },
    { RTL_AX_LOAD_C,                  "ax_load.c" },
    { RTL_AX_HOST_C,                  "ax_host.c" },
    { RTL_VB6FORMS_UC_C,           "vb6forms_uc.c" },

    // vb6rtl.h 按家族拆分 (2026-09-17)
    { RTL_VB6RTL_BASE_H,                  "vb6rtl_base.h" },
    { RTL_VB6RTL_BSTR_H,                  "vb6rtl_bstr.h" },
    { RTL_VB6RTL_VARIANT_H,               "vb6rtl_variant.h" },
    { RTL_VB6RTL_BUILTIN_H,               "vb6rtl_builtin.h" },
    { RTL_VB6RTL_ARRAY_H,                 "vb6rtl_array.h" },
    { RTL_VB6RTL_CLASS_COM_H,             "vb6rtl_class_com.h" },
    { RTL_VB6RTL_RUNTIME_H,               "vb6rtl_runtime.h" },

    // vb6forms 家族再细分 (2026-09-17)
    { RTL_VB6FORMS_PICTURE_PROP_C,        "vb6forms_picture_prop.c" },
    { RTL_VB6FORMS_WIDGET_PROP_C,         "vb6forms_widget_prop.c" },

    // vb6forms.h 按功能域拆分 (2026-09-17)
    { RTL_VB6FORMS_WINDOW_H,              "vb6forms_window.h" },
    { RTL_VB6FORMS_PROP_H,                "vb6forms_prop.h" },
    { RTL_VB6FORMS_PROP_CTRL_H,           "vb6forms_prop_ctrl.h" },
    { RTL_VB6FORMS_PROP_PIC_H,            "vb6forms_prop_pic.h" },
    { RTL_VB6FORMS_PROP_FORM_H,           "vb6forms_prop_form.h" },
    { RTL_VB6FORMS_CTRLARR_H,             "vb6forms_ctrlarr.h" },
    { RTL_VB6FORMS_MDI_H,                 "vb6forms_mdi.h" },
    { RTL_VB6FORMS_WEBVIEW_H,             "vb6forms_webview.h" },
    { RTL_VB6FORMS_CONTROLS_H,            "vb6forms_controls.h" },

    // vb6rtl.c 按家族拆分 (2026-09-17)
    { RTL_VB6RTL_STRING_C,                "vb6rtl_string.c" },
    { RTL_VB6RTL_FORMAT_C,                "vb6rtl_format.c" },
    { RTL_VB6RTL_FORMAT_EXTRACT_INC,      "vb6rtl_format_extract.inc" },
    { RTL_VB6RTL_FORMAT_PARSE_INC,        "vb6rtl_format_parse.inc" },
    { RTL_VB6RTL_FORMAT_NUMERIC_PRE_INC,  "vb6rtl_format_numeric_pre.inc" },
    { RTL_VB6RTL_FORMAT_NUMERIC_BODY_INC, "vb6rtl_format_numeric_body.inc" },
    { RTL_VB6RTL_FORMAT_NUMERIC_TAIL_INC, "vb6rtl_format_numeric_tail.inc" },
    { RTL_VB6RTL_FORMAT_STRING_INC,       "vb6rtl_format_string.inc" },
    { RTL_VB6RTL_CONV_C,                  "vb6rtl_conv.c" },
    { RTL_VB6RTL_MISC_C,                  "vb6rtl_misc.c" },
    { RTL_VB6RTL_SYSTEM_C,                "vb6rtl_system.c" },
    { RTL_VB6RTL_COMPAT_C,                "vb6rtl_compat.c" },
    { RTL_VB6RTL_DATE_C,                  "vb6rtl_date.c" },
    { RTL_VB6RTL_ARRAY_C,                 "vb6rtl_array.c" },
    { RTL_VB6RTL_FILE_C,                  "vb6rtl_file.c" },
    { RTL_VB6RTL_PARAMARRAY_C,            "vb6rtl_paramarray.c" },
    { RTL_VB6RTL_FINANCIAL_C,             "vb6rtl_financial.c" },
    { RTL_VB6RTL_COM_C,                   "vb6rtl_com.c" },
    { RTL_VB6RTL_REGISTRY_C,              "vb6rtl_registry.c" },
    { RTL_VB6RTL_USERCTL_H,               "vb6rtl_userctl.h" },  // Fix 105

    // vb6forms_uc 按族细分 (2026-09-19): 内部头 + 6 族编译单元 + 7 个片段
    { RTL_VB6FORMS_UC_INTERNAL_H,         "vb6forms_uc_internal.h" },
    { RTL_UC_HOST_C,                      "uc_host.c" },
    { RTL_UC_HOST_WINDOW_C,               "uc_host_window.c" },
    { RTL_UC_HOSTMODEL_C,                 "uc_hostmodel.c" },
    { RTL_UC_CONTROLS_C,                  "uc_controls.c" },
    { RTL_UC_COLLECTION_C,                "uc_collection.c" },
    { RTL_UC_DEBUG_C,                     "uc_debug.c" },
    { RTL_UC_HOST_CREATE_INC,             "uc_host_create.inc" },
    { RTL_UC_HOSTMODEL_GETPROP_INC,       "uc_hostmodel_getprop.inc" },
    { RTL_UC_HOSTMODEL_SETPROP_INC,       "uc_hostmodel_setprop.inc" },
    { RTL_UC_HOSTMODEL_CALL_INC,          "uc_hostmodel_call.inc" },
    { RTL_UC_COLLECTION_API_INC,          "uc_collection_api.inc" },
    { RTL_UC_DEBUG_DIB_INC,               "uc_debug_dib.inc" },
    { RTL_UC_DEBUG_COMPOSITE_INC,         "uc_debug_composite.inc" },
    { RTL_UC_PROPBAG_C,                   "uc_propbag.c" },
};

static const size_t kRtlFileCount = sizeof(kRtlFiles) / sizeof(kRtlFiles[0]);

// 会话目录名格式 "<steady_clock ticks>_<pid>"; 取不出 pid 返回 0
static DWORD sessionDirOwnerPid(const std::string& name) {
    size_t us = name.rfind('_');
    if (us == std::string::npos || us + 1 >= name.size()) return 0;
    try {
        unsigned long long v = std::stoull(name.substr(us + 1));
        if (v == 0 || v > 0xFFFFFFFFull) return 0;
        return static_cast<DWORD>(v);
    } catch (...) {
        return 0;
    }
}

// 属主进程是否仍存活. 打不开且非权限问题时视为已退出.
static bool isProcessAlive(DWORD pid) {
#ifdef _WIN32
    HANDLE h = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (!h) {
        // 权限不足: 保守当作存活 —— 宁可不删, 也不误删正在编译的会话目录
        return GetLastError() == ERROR_ACCESS_DENIED;
    }
    DWORD code = 0;
    bool alive = (GetExitCodeProcess(h, &code) != 0) && (code == STILL_ACTIVE);
    CloseHandle(h);
    return alive;
#else
    (void)pid;
    return false;
#endif
}

std::string SessionManager::create() {
    // 1. Clean up old sessions
    cleanupOldSessions();

    // 2. Create unique session directory
    auto now = std::chrono::steady_clock::now().time_since_epoch().count();
    // 并行编译竞态防护 (T1 CI 16 worker 实证): 多个 C3 进程同时启动时 steady_clock
    // 可能同 tick 撞值 -> 共享同一 %TEMP%\C3C 目录互相删对方文件 (D8022 rsp 消失).
    // 目录名追加当前进程 PID 天然唯一; cleanupOldSessions 靠后缀 PID 判定属主是否存活.
    DWORD pid = GetCurrentProcessId();
    std::string root = getSessionRoot();
    sessionDir_ = root + "\\" + std::to_string(now) + "_" + std::to_string(pid);
    rtlDir_ = sessionDir_ + "\\rtl";

    std::error_code ec;
    std::filesystem::create_directories(utf8ToPath(rtlDir_), ec);
    if (ec) {
        std::cerr << "C3: cannot create session directory: " << rtlDir_ << " (" << ec.message() << ")" << std::endl;
        rtlDir_.clear();
        sessionDir_.clear();
        return "";
    }

    // 3. Extract RTL resource files (.h headers + .c sources, all arch-neutral)
    //    清单 = 文件顶部的 kRtlFiles (与 restoreMissing() 共用)
    for (size_t fi = 0; fi < kRtlFileCount; fi++) {
        const RtlFileEntry& entry = kRtlFiles[fi];
        if (!extractResource(entry.id, entry.name, rtlDir_)) {
            std::cerr << "C3: cannot extract RTL resource: " << entry.name << std::endl;
            cleanup();
            return "";
        }
    }
    return rtlDir_;
}

bool SessionManager::extractResource(int resId, const std::string& fileName, const std::string& targetDir) {
    std::vector<char> data;
    if (!loadRtlResource(resId, data)) {
        return false;
    }

    std::string filePath = targetDir + "\\" + fileName;
    std::ofstream ofs(utf8ToPath(filePath), std::ios::out | std::ios::trunc | std::ios::binary);
    if (!ofs) return false;

    ofs.write(data.data(), static_cast<std::streamsize>(data.size()));
    return ofs.good();
}

void SessionManager::cleanup() {
    if (sessionDir_.empty()) return;

    std::error_code ec;
    std::filesystem::remove_all(utf8ToPath(sessionDir_), ec);
    // Ignoring errors - best effort cleanup

    sessionDir_.clear();
    rtlDir_.clear();
}

// 编译前自愈: 会话 rtl/ 下每个 RTL 文件都必须存在且非空, 缺的当场重新解包.
// 背景 (T1 CI 35517635219 实证): 兄弟进程的 cleanupOldSessions 曾把一个**仍在编译**
// 的会话目录按"过期"整棵删掉 —— cl 带 /MP 多源并行编译, 被删的是尚未编译的那一批,
// 于是报 c1 fatal error C1083: Cannot open source file '...\rtl/ax_site.c'.
// 两个修法都要有: cleanupOldSessions 不再误删 (治本), 本函数兜住已经发生的缺失.
bool SessionManager::restoreMissing() {
    if (rtlDir_.empty()) return false;

    std::error_code ec;
    std::filesystem::create_directories(utf8ToPath(rtlDir_), ec);   // 目录本身被删也要能恢复

    int repaired = 0, failed = 0;
    for (size_t i = 0; i < kRtlFileCount; i++) {
        const std::string path = rtlDir_ + "\\" + kRtlFiles[i].name;
        ec.clear();
        auto size = std::filesystem::file_size(utf8ToPath(path), ec);
        if (!ec && size > 0) continue;                  // 齐备

        if (extractResource(kRtlFiles[i].id, kRtlFiles[i].name, rtlDir_)) {
            repaired++;
        } else {
            failed++;
            std::cerr << "C3: cannot restore RTL file: " << kRtlFiles[i].name << std::endl;
        }
    }
    if (repaired > 0) {
        std::cerr << "C3: warning: restored " << repaired << " missing RTL file(s) in "
                  << rtlDir_ << std::endl;
    }
    return failed == 0;
}

void SessionManager::cleanupOldSessions() {
    std::string root = getSessionRoot();
    if (!std::filesystem::exists(utf8ToPath(root))) return;

    // 陈旧阈值 30 分钟. 单次编译要跑 cl /MP 并行编译 100+ 个 RTL 源文件, 在 CI 16 路
    // 并发下实测可超过 300 秒 —— 旧的 300s 阈值会把**仍在编译**的兄弟会话目录判成过期
    // 删除. 真正的护栏是下面的"属主进程存活就不删"; 阈值只需兜住崩溃/被杀的残留.
    const auto maxAge = std::chrono::minutes(30);

    std::error_code ec;
    for (auto& entry : std::filesystem::directory_iterator(utf8ToPath(root), ec)) {
        if (!entry.is_directory()) continue;

        // 1) 属主进程仍存活 -> 绝不删除 (会话目录名 = "<ticks>_<pid>")
        const std::string dirName = entry.path().filename().string();
        if (isProcessAlive(sessionDirOwnerPid(dirName))) continue;

        // 2) 陈旧判定用目录自身最后写时间 (文件系统时钟), 不再解析目录名里的 ticks:
        //    名称格式与 steady_clock 单位耦合, 一旦变化就会误判成"很久以前".
        ec.clear();
        auto mt = std::filesystem::last_write_time(entry.path(), ec);
        if (ec) continue;
        auto age = std::filesystem::file_time_type::clock::now() - mt;
        if (age > maxAge) {
            std::filesystem::remove_all(entry.path(), ec);
        }
    }
}

} // namespace vb6c3
