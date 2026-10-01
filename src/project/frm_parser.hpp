#pragma once
// VB6 窗体文件 (.frm) 解析器
// .frm 文件结构:
//   VERSION 5.00
//   Begin VB.Form FormName
//      Caption = "Hello"
//      ClientHeight = 3000
//      ...
//      Begin VB.CommandButton Command1
//         Caption = "OK"
//         ...
//      End
//   End
//   Attribute VB_Name = "Form1"
//   <VB代码段>

#include <string>
#include <vector>
#include <map>
#include <filesystem>

namespace vb6c3 {

// ============================================================
// 窗体属性值 — 支持多种类型
// ============================================================

enum class FrmValueType {
    Integer,        // 整数: 3000, -1
    Float,          // 浮点: 8.25
    String,         // 字符串: "Hello"
    Identifier,     // 标识符: vbModal, vbUpperCase
    FrxReference,   // .frx引用: "Form1.frx":10CA
};

struct FrmValue {
    FrmValueType type;
    std::string rawText;        // 原始文本（字符串含引号）
    long long intValue = 0;     // type==Integer时有效
    double floatValue = 0.0;    // type==Float时有效
    std::string frxFile;        // type==FrxReference时: .frx文件名
    size_t frxOffset = 0;       // type==FrxReference时: .frx文件内偏移

    // 便捷构造
    static FrmValue fromInt(long long v, const std::string& raw) {
        FrmValue r; r.type = FrmValueType::Integer; r.rawText = raw; r.intValue = v; return r;
    }
    static FrmValue fromFloat(double v, const std::string& raw) {
        FrmValue r; r.type = FrmValueType::Float; r.rawText = raw; r.floatValue = v; return r;
    }
    static FrmValue fromString(const std::string& raw) {
        FrmValue r; r.type = FrmValueType::String; r.rawText = raw; return r;
    }
    static FrmValue fromIdent(const std::string& raw) {
        FrmValue r; r.type = FrmValueType::Identifier; r.rawText = raw; return r;
    }
    static FrmValue fromFrxRef(const std::string& file, size_t offset, const std::string& raw) {
        FrmValue r; r.type = FrmValueType::FrxReference; r.rawText = raw;
        r.frxFile = file; r.frxOffset = offset; return r;
    }
};

// ============================================================
// 复合属性块 (BeginProperty ... EndProperty)
// ============================================================

struct FrmPropertyBlock {
    std::string blockName;          // "Font", "Images", "ListImage1", 等
    std::string blockGuid;          // GUID (含花括号): "{2C247F25-8591-11D1-B16A-00C0F0283628}", 无GUID时为空
    std::map<std::string, FrmValue> properties;   // Name="MS Sans Serif", Size=8.25, Picture="Form1.frx":1A89, 等
    std::vector<FrmPropertyBlock> nestedBlocks;   // 嵌套 BeginProperty 块 (如 ListImage1..N 嵌套在 Images 内)
};

// ============================================================
// 控件描述
// ============================================================

enum class FrmControlType {
    // VB6标准控件 (VB.xxx)
    Form,               // VB.Form — 窗体自身
    MDIForm,             // VB.MDIForm — MDI父窗体
    CommandButton,      // VB.CommandButton
    TextBox,            // VB.TextBox
    Label,              // VB.Label
    CheckBox,           // VB.CheckBox
    OptionButton,       // VB.OptionButton
    ListBox,            // VB.ListBox
    ComboBox,           // VB.ComboBox
    Frame,              // VB.Frame
    PictureBox,         // VB.PictureBox
    Timer,              // VB.Timer
    HScrollBar,         // VB.HScrollBar
    VScrollBar,         // VB.VScrollBar
    Image,              // VB.Image
    Shape,              // VB.Shape
    Line,               // VB.Line
    Data,               // VB.Data
    OLE,                // VB.OLE
    DriveListBox,       // VB.DriveListBox
    DirListBox,         // VB.DirListBox
    FileListBox,        // VB.FileListBox
    Menu,               // VB.Menu (菜单项)
    WebBrowser,         // SHDocVw.WebBrowser / WebBrowser (WebView2宿主)
    Toolbar,            // MSComctlLib.Toolbar 等 (第三方)
    StatusBar,          // MSComctlLib.StatusBar
    CommonDialog,       // MSComDlg.CommonDialog
    ImageList,          // MSComctlLib.ImageList (ActiveX, COM后期绑定)
    ProgressBar,        // MSComctlLib.ProgressBar — Win32 原生复刻 (msctls_progress32)
    SSTab,              // TabDlg.SSTab — Win32 原生复刻 (SysTabControl32), P20-42
    ListView,           // MSComctlLib.ListView — Win32 原生复刻 (SysListView32), P20-45
    TreeView,           // MSComctlLib.TreeView — Win32 原生复刻 (SysTreeView32), P20-46
    DTPicker,           // MSComCtl2.DTPicker — Win32 原生复刻 (SysDateTimePick32), C29-DT
    MonthView,          // MSComCtl2.MonthView — Win32 原生复刻 (SysMonthCal32), C29-MV
    RichTextBox,        // MSFT_TextBox.RichTextBox — Win32 原生复刻 (RICHEDIT50W), C29-RT
    Winsock,            // MSWinsockLib.Winsock — Winsock2 原生复刻 (无外观控件), C29-WS
    Slider,             // MSComctlLib.Slider — Win32 原生复刻 (msctls_trackbar32), C29-SL
    Unknown,            // 未识别的控件类型
};

struct FrmControl {
    FrmControlType controlType = FrmControlType::Unknown;
    std::string controlTypeName;       // 原始类型名: "VB.CommandButton", "MSComctlLib.Toolbar"
    std::string controlName;           // 控件实例名: "Command1", "Text1"
    int index = -1;                    // 控件数组索引 (-1=非数组)

    // 简单属性: Caption="OK", Left=120, Top=240, ...
    std::map<std::string, FrmValue> properties;

    // 复合属性块: Font, MouseIcon, etc.
    std::vector<FrmPropertyBlock> propertyBlocks;

    // 嵌套子控件 (Frame内的控件, Menu子项等)
    std::vector<FrmControl> children;

    // 事件处理桩标记 (从属性推断)
    // VB6中控件的事件处理器格式: Sub ControlName_EventName()
};

// ============================================================
// 窗体描述 (解析结果)
// ============================================================

struct FrmFormDesc {
    std::string formName;              // 窗体名: "Form1"
    bool isMDIChild = false;            // 是否为MDI子窗体 (MDIChild=-1属性)
    FrmControl formControl;            // 窗体自身的属性 (Caption, ClientHeight, 等)
                                       // formControl.controlType == FrmControlType::Form

    // 顶层控件列表 (直接放在窗体上的控件)
    // formControl.children 就是顶层控件

    // .frx 二进制资源引用 (图片等)
    // 格式: 属性值中出现的 $begin...$end 块 或 :X 形式的偏移引用
    std::map<std::string, std::string> frxReferences;  // key=偏移, value=资源描述
    // P24: .frx 文件路径 (与 .frm 同目录)
    std::filesystem::path frxFilePath;
};

// ============================================================
// .frm 文件完整解析结果
// ============================================================

struct FrmFile {
    std::string version;               // "5.00"
    FrmFormDesc form;                  // 窗体描述
    std::string codeSection;           // VB代码段 (Attribute和VB代码, 不含Begin...End)
    std::filesystem::path frmFilePath; // 文件路径
    // Fix 196: 文件根本没读进来 (不存在 / 路径错 / 权限)。调用方**必须**据此报错 ——
    // 否则整张窗体 (设计器 + 代码段) 会静默变成空模块, 编出的 exe 什么都不做,
    // 却报"编译成功"。中文路径下曾经正是这个症状。
    bool readFailed = false;
};

// ============================================================
// Fix 195: 外部二进制资源引用 (.frx/.ctx/.pgx)
// ============================================================
// VB6 的行式设计器格式 (Begin/End + `键 = 值`) 无法表达多行文本与二进制图片,
// 于是把这类属性值甩进同名资源文件, .frm 里只留 `Text = "Form1.frx":0000` 的
// 字节偏移引用。资源缺失时这些属性的设计期取值会被静默丢弃 —— 所以解析与
// 定位这两件事必须只有一份实现 (前端告警、--extract-frx 都走这里)。

struct FrmResourceRef {
    std::string where;      // 属性路径, 如 "Text1.Text" / "List1.Images.ListImage1.Picture"
    std::string fileName;   // 属性行里写的文件名, 如 "Form1.frx"
    size_t offset = 0;      // 该属性数据在资源文件里的字节偏移
    std::string rawText;    // 属性行右侧原文, 如 "\"Form1.frx\":0000" (供导出器列出待删行)
};

// 递归收集 (含子控件 / BeginProperty 块 / 嵌套块) 的所有资源引用。
// 路径按 VB6 的控件寻址习惯拼: 控件名 + 属性名 (不叠父控件名, Frame 里的控件
// 在 VB6 里也是直接按自己的名字访问的)。
void collectFrmResourceRefs(const FrmControl& ctrl,
                            std::vector<FrmResourceRef>& out);

// 定位资源文件: 优先用属性行里的引用名 (权威, 且能容忍 .frm 被改名),
// 再退回 <窗体基名> + fallbackExt。都找不到返回空 path。
std::filesystem::path resolveFrmResourceFile(const FrmFile& frm,
                                             const std::string& fallbackExt);

// ============================================================
// .frm 解析器
// ============================================================

class FrmParser {
public:
    // 解析 .frm 文件
    static FrmFile parse(const std::string& frmFilePath);

    // 解析 .frm 内容字符串
    static FrmFile parseString(const std::string& content, const std::string& frmFilePath = "");

    // 控件类型名 → FrmControlType
    static FrmControlType parseControlType(const std::string& typeName);

    // FrmControlType → Win32窗口类名 (用于代码生成)
    static const char* controlTypeToWin32Class(FrmControlType type);

    // FrmControlType → VB6运行时类型名 (用于符号表)
    static const char* controlTypeToVb6Name(FrmControlType type);

private:
    // 行级解析辅助
    static std::string trim(const std::string& s);
    static std::string unquote(const std::string& s);
    static FrmValue parseValue(const std::string& s);

    // 解析 Begin...End 块 (控件或窗体)
    static FrmControl parseControlBlock(
        const std::vector<std::string>& lines,
        size_t& lineIdx,
        const std::string& parentIndent = "");

    // 解析控件内集合块的项行 `.ListImage(1, "k", "x.frx":0000)` → 一个 FrmPropertyBlock,
    // 参数按 VB6 的位置语义挂成 properties (ListImage: Index/Key/Picture/Key2/Picture2)。
    static FrmPropertyBlock parseCollectionItem(const std::string& line);

    // 解析 BeginProperty...EndProperty 块
    static FrmPropertyBlock parsePropertyBlock(
        const std::vector<std::string>& lines,
        size_t& lineIdx);

    // 解析属性行: Key = Value 或 Key = Value ' Comment
    static bool parsePropertyLine(const std::string& line,
        std::string& key, FrmValue& value);

    // 提取代码段 (Begin...End之后的内容)
    static std::string extractCodeSection(
        const std::vector<std::string>& lines,
        size_t startIdx);
};

} // namespace vb6c3
