// P10 (restored): RTL runtime embedded resource management
// Extract RTL .h headers + .c sources from c3.exe RCDATA resources.
// The P11.3 pre-compiled .lib scheme was reverted (c3 is open source);
// RTL sources are compiled by MSVC together with the generated code.
// Compiles and cleans up temp session dir automatically

#ifndef VB6C3_RTL_EMBEDDED_HPP
#define VB6C3_RTL_EMBEDDED_HPP

#include <string>
#include <vector>

namespace vb6c3 {

// RTL embedded resource IDs (must match c3rtl.rc)
enum RtlResourceID {
    // Headers (for #include in generated code) — arch-neutral
    RTL_VB6RTL_H        = 100,
    RTL_VB6RTL_C        = 101,
    RTL_VB6COM_H        = 102,
    RTL_VB6COM_C        = 103,
    RTL_VB6COMSERVER_H  = 104,
    RTL_VB6COMSERVER_C  = 105,
    RTL_VB6FORMS_H      = 106,
    RTL_VB6FORMS_C      = 107,
    RTL_VB6_DI_STUBS_C  = 108,
    RTL_VB6_DI_WIN32_STUBS_C = 109,

    // vb6comserver 按 COM 接口家族拆分 (P6.6): 内部共享头 + 各族实现
    RTL_VB6COMSERVER_INTERNAL_H = 110,
    RTL_VB6COMSERVER_OBJ_C      = 111,
    RTL_VB6COMSERVER_FACTORY_C  = 112,
    RTL_VB6COMSERVER_CP_C       = 113,
    RTL_VB6COMSERVER_PCI_C      = 114,

    // vb6com 按 COM 调用层次拆分 (P6): 内部共享头 + 各层实现
    RTL_VB6COM_INTERNAL_H = 115,
    RTL_VB6COM_INVOKE_C   = 116,
    RTL_VB6COM_PACK_C     = 117,
    RTL_VB6COM_WRAP_C     = 118,
    RTL_VB6COM_SINK_C     = 119,
    RTL_VB6COM_FOREACH_C  = 120,

    // vb6forms 按控件/窗体功能家族拆分 (P7): 内部共享头 + 各家族实现
    RTL_VB6FORMS_INTERNAL_H = 121,
    RTL_VB6FORMS_CTRL_C     = 122,
    RTL_VB6FORMS_LIST_C     = 123,
    RTL_VB6FORMS_STYLE_C    = 124,
    RTL_VB6FORMS_SCROLL_C   = 125,
    RTL_VB6FORMS_PICTURE_C  = 126,
    RTL_VB6FORMS_CTRLARR_C  = 127,
    RTL_VB6FORMS_WEBVIEW_C  = 128,
    RTL_VB6FORMS_WIDGET_C   = 129,
    RTL_VB6FORMS_SHAPE_C    = 130,
    RTL_VB6FORMS_AXSITE_C   = 131,
    RTL_VB6FORMS_UC_C       = 175,  // Fix 112: 工程内 UserControl 实例宿主

    // vb6rtl.h 按家族拆分 (2026-09-17): 伞头 + 7 个子头
    RTL_VB6RTL_BASE_H = 132,
    RTL_VB6RTL_BSTR_H = 133,
    RTL_VB6RTL_VARIANT_H = 134,
    RTL_VB6RTL_BUILTIN_H = 135,
    RTL_VB6RTL_ARRAY_H = 136,
    RTL_VB6RTL_CLASS_COM_H = 137,
    RTL_VB6RTL_RUNTIME_H = 138,

    // vb6forms 家族再细分 (2026-09-17)
    RTL_VB6FORMS_PICTURE_PROP_C = 139,
    RTL_VB6FORMS_WIDGET_PROP_C  = 140,

    // DI 转发桩按 Lib 家族拆分 (2026-09-17); gdiplus 2026-09-20 再按对象域拆 5 份 (di/gdiplus/)
    RTL_VB6_DI_USER32_STUBS_C    = 141,
    RTL_VB6_DI_GDIPLUS_DRAW_STUBS_C = 142,
    RTL_VB6_DI_CRYPTO_STUBS_C    = 143,
    RTL_VB6_DI_COM_STUBS_C       = 144,
    RTL_VB6_DI_NET_STUBS_C       = 145,
    RTL_VB6_DI_SHELL_STUBS_C     = 146,

    // vb6forms.h 按功能域拆分 (2026-09-17): 伞头 + 9 个子头
    RTL_VB6FORMS_WINDOW_H      = 147,
    RTL_VB6FORMS_PROP_H        = 148,
    RTL_VB6FORMS_PROP_CTRL_H   = 149,
    RTL_VB6FORMS_PROP_PIC_H    = 150,
    RTL_VB6FORMS_PROP_FORM_H   = 151,
    RTL_VB6FORMS_CTRLARR_H     = 152,
    RTL_VB6FORMS_MDI_H         = 153,
    RTL_VB6FORMS_WEBVIEW_H     = 154,
    RTL_VB6FORMS_CONTROLS_H    = 155,

    // vb6rtl.c 按家族拆分 (2026-09-17): 主文件 + 13 个族实现
    RTL_VB6RTL_STRING_C        = 156,
    RTL_VB6RTL_FORMAT_C        = 157,
    // format 族按函数体片段拆: 6 个 .inc（不单独编译, 由 format.c include）
    RTL_VB6RTL_FORMAT_EXTRACT_INC      = 158,
    RTL_VB6RTL_FORMAT_PARSE_INC        = 159,
    RTL_VB6RTL_FORMAT_NUMERIC_PRE_INC  = 160,
    RTL_VB6RTL_FORMAT_NUMERIC_BODY_INC = 161,
    RTL_VB6RTL_FORMAT_NUMERIC_TAIL_INC = 162,
    RTL_VB6RTL_FORMAT_STRING_INC       = 163,
    RTL_VB6RTL_CONV_C          = 164,
    RTL_VB6RTL_MISC_C          = 165,
    RTL_VB6RTL_SYSTEM_C        = 166,
    RTL_VB6RTL_COMPAT_C        = 167,
    RTL_VB6RTL_DATE_C          = 168,
    RTL_VB6RTL_ARRAY_C         = 169,
    RTL_VB6RTL_FILE_C          = 170,
    RTL_VB6RTL_PARAMARRAY_C    = 171,
    RTL_VB6RTL_FINANCIAL_C     = 172,
    RTL_VB6RTL_COM_C           = 173,
    RTL_VB6RTL_REGISTRY_C      = 174,

    RTL_VB6RTL_USERCTL_H       = 176,  // Fix 105: UserControl/Ambient/Extender/PropertyPage 宿主

    // Fix 103: VB6 内建对象 Collection 的 C 实现 (IDispatch + IEnumVARIANT)
    RTL_VB6COM_COLLECTION_C    = 177,
    RTL_VB6COM_COLLECTION_ENUM_C = 178,
    // vb6forms_uc 按族细分 (2026-09-19): 内部头 + 6 族编译单元 + 7 个函数体片段
    RTL_VB6FORMS_UC_INTERNAL_H             = 179,
    RTL_UC_HOST_C                          = 180,
    RTL_UC_HOST_WINDOW_C                   = 181,
    RTL_UC_HOSTMODEL_C                     = 182,
    RTL_UC_CONTROLS_C                      = 183,
    RTL_UC_COLLECTION_C                    = 184,
    RTL_UC_DEBUG_C                         = 185,
    RTL_UC_HOST_CREATE_INC                 = 186,
    RTL_UC_HOSTMODEL_GETPROP_INC           = 187,
    RTL_UC_HOSTMODEL_SETPROP_INC           = 188,
    RTL_UC_HOSTMODEL_CALL_INC              = 189,
    RTL_UC_COLLECTION_API_INC              = 190,
    RTL_UC_DEBUG_DIB_INC                   = 191,
    RTL_UC_DEBUG_COMPOSITE_INC             = 192,
    RTL_UC_PROPBAG_C                       = 193,

    // DI gdiplus 二级拆分 (2026-09-20): text/image/brush/path (draw 见 142)
    RTL_VB6_DI_GDIPLUS_TEXT_STUBS_C        = 194,
    RTL_VB6_DI_GDIPLUS_IMAGE_STUBS_C       = 195,
    RTL_VB6_DI_GDIPLUS_BRUSH_STUBS_C       = 196,
    RTL_VB6_DI_GDIPLUS_PATH_STUBS_C        = 197,

    // Fix 148: VB6 容器对象模型 (Extender/Container/Controls) — OCX 真宿主 (axcontainer.c)
    // 注意: 179~197 已被 vb6forms_uc 拆分 + gdiplus 拆分占满, 故顺延到 198 避免 ID 冲突
    RTL_VB6FORMS_AXCONTAINER_C = 198,

    // vb6forms_axsite.c 按功能家族拆分 (2026-09-20): 内部头 + 5 个族编译单元
    RTL_VB6FORMS_AXSITE_INTERNAL_H         = 199,
    RTL_AX_SITE_C                          = 200,
    RTL_AX_SITE_EXT_C                      = 201,
    RTL_AX_PROPBAG_C                       = 202,
    RTL_AX_LOAD_C                          = 203,
    RTL_AX_HOST_C                          = 204,

    // Fix 160y: DI 转发桩未带 vb6_di_lib 标记的杂项符号 (Imm 输入法 / version /
    // TransparentBlt / msvbvm60 运行时) — gen_di_stubs.ps1 归入 unknown 族
    RTL_VB6_DI_UNKNOWN_STUBS_C             = 205,

    // P20-38: VB6 ProgressBar 控件 (msctls_progress32 复刻, 不加载 mscomctl.ocx)
    RTL_VB6FORMS_PROGRESS_C                = 206,
    RTL_VB6FORMS_IMAGELIST_C               = 207,

    // P20-40: VB6 StatusBar 控件 (msctls_status32 复刻, 不加载 mscomctl.ocx)
    RTL_VB6FORMS_STATUSBAR_C               = 208,

    // P20-42: VB6 SSTab 控件 (SysTabControl32 复刻, 不加载 TABCTL32.OCX)
    RTL_VB6FORMS_SSTAB_C                   = 209,

    // P20-44: OLE 拖放 (IDataObject/IDropSource/IDropTarget 复刻)
    RTL_VB6FORMS_OLEDD_C                   = 210,

    // P20-45: VB6 ListView 控件 (SysListView32 复刻, 不加载 MSCOMCTL.OCX)
    RTL_VB6FORMS_LISTVIEW_C                = 211,
    // C29-3: 控件"成员对象/成员集合"的真 IDispatch (ListImages 立样)
    RTL_VB6FORMS_MEMBEROBJ_C               = 212,
    // ai/029 C29-8a: TreeView 标量属性面 (dev 原用 212, 合并后顺延让位)
    RTL_VB6FORMS_TREEVIEW_C                = 213,
    // OLE 容器 (本地判据, 不进 CI —— 嵌入对象依赖目标机器的 OLE 服务器)
    RTL_VB6FORMS_OLECON_C                  = 214,
    // Data 控件 (ODBC 后端)
    RTL_VB6FORMS_DATA_C                    = 216,

    // ai/029 C29-5a: VB6 Toolbar 控件 (ToolbarWindow32 复刻, 不加载 MSCOMCTL.OCX)
    RTL_VB6FORMS_TOOLBAR_C                 = 215,

    // ai/029 C29-DT-a: VB6 DTPicker 控件 (SysDateTimePick32 复刻, 不加载 MSCOMCT2.OCX)
    RTL_VB6FORMS_DTPICKER_C                = 217,

    // ai/029 C29-MV-a: VB6 MonthView 控件 (SysMonthCal32 复刻, 同样不加载 MSCOMCT2.OCX)
    RTL_VB6FORMS_MONTHVIEW_C               = 218,

    // ai/029 C29-RT-a: VB6 RichTextBox 控件 (原生 RICHEDIT50W, 不加载 RICHTX32.OCX)
    RTL_VB6FORMS_RICHTEXTBOX_C             = 219,
    RTL_VB6FORMS_WINSOCK_C                 = 220,

    // ai/029 C29-SL-a: VB6 Slider 控件 (原生 msctls_trackbar32, 不加载 MSCOMCTL.OCX)
    RTL_VB6FORMS_SLIDER_C                  = 221,
};

// Session directory manager
// Creates %TMP%\C3C\{timestamp}\ temp dir, extracts RTL files, cleans up after compile
class SessionManager {
public:
    SessionManager();
    ~SessionManager();

    // Create new session directory and extract RTL files (.h + .c, arch-neutral)
    // Returns: RTL directory path (contains .h + .c)
    // Empty string on failure
    std::string create();

    // Get current session's RTL directory path
    const std::string& rtlDir() const { return rtlDir_; }

    // 校验会话 rtl/ 下每个 RTL 文件是否存在且非空, 缺失的当场重新解包.
    // 用途: 编译前调用. 兄弟进程的 cleanupOldSessions 可能已把本会话目录里的
    // 部分文件删掉 (见 cleanupOldSessions 注释), 此时 cl 会对尚未编译的源文件报
    // C1083 —— 这里自愈, 而不是让整次编译失败.
    // 返回: 所有文件最终齐备 -> true
    bool restoreMissing();

    // P11.2: session root dir (for intermediates .c/.h/.obj)
    const std::string& sessionDir() const { return sessionDir_; }

    // Clean up session directory (auto-called by destructor)
    void cleanup();

    // Release session without cleanup (keep intermediates for debugging)
    void release() { sessionDir_.clear(); rtlDir_.clear(); }

    // Clean up old session dirs (owner process already gone + idle > 30 min)
    // Auto-called on each create()
    static void cleanupOldSessions();

    // Check if session is active
    bool isActive() const { return !rtlDir_.empty(); }

private:
    std::string sessionDir_;  // session root dir (e.g. %TMP%\C3C\{ts})
    std::string rtlDir_;      // rtl subdir (sessionDir_\rtl)

    // Extract file from RCDATA resource
    bool extractResource(int resId, const std::string& fileName, const std::string& targetDir);

    // Get session root directory (%TMP%\C3C)
    static std::string getSessionRoot();
};

} // namespace vb6c3

#endif // VB6C3_RTL_EMBEDDED_HPP
