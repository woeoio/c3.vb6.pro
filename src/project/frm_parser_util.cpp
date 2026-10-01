// vb6c3 - VB6窗体文件(.frm)解析器 — 字符串工具 + 控件类型映射
// 由 src/project/frm_parser.cpp 拆出（2026-09-17），纯搬移、零行为改动。

#include "project/frm_parser.hpp"
#include "common/encoding.hpp"
#include "common/source_manager.hpp"
#include <sstream>
#include <algorithm>
#include <cctype>

namespace vb6c3 {

std::string FrmParser::trim(const std::string& s) {
    size_t start = s.find_first_not_of(" \t\r\n");
    if (start == std::string::npos) return "";
    size_t end = s.find_last_not_of(" \t\r\n");
    return s.substr(start, end - start + 1);
}

std::string FrmParser::unquote(const std::string& s) {
    if (s.size() >= 2 && s.front() == '"' && s.back() == '"') {
        return s.substr(1, s.size() - 2);
    }
    return s;
}

FrmValue FrmParser::parseValue(const std::string& s) {
    std::string t = trim(s);
    if (t.empty()) return FrmValue::fromInt(0, "");

    // .frx 二进制资源引用: "filename.frx":HEXOFFSET
    // 格式: "Form1.frx":0000  或  "Form1.frx":10CA
    if (t.size() >= 8 && t.front() == '"') {
        size_t colon = t.find("\":");
        if (colon != std::string::npos && colon + 2 < t.size()) {
            std::string fileName = t.substr(1, colon - 1);  // 去掉前引号
            std::string offsetStr = t.substr(colon + 2);
            // 验证偏移是十六进制
            bool validHex = !offsetStr.empty();
            for (char c : offsetStr) {
                if (!std::isxdigit(static_cast<unsigned char>(c))) { validHex = false; break; }
            }
            if (validHex) {
                size_t offset = std::stoull(offsetStr, nullptr, 16);
                return FrmValue::fromFrxRef(fileName, offset, t);
            }
        }
    }

    // 字符串字面量
    if (t.size() >= 2 && t.front() == '"') {
        return FrmValue::fromString(t);
    }

    // VB6 十六进制/八进制整数字面量: &H00FF8021&, &HFF& (尾随 &/%%/@ 为类型符)
    if (t.size() > 3 && (t[0] == '&' || t[0] == '#') &&
        (t[1] == 'H' || t[1] == 'h' || t[1] == 'O' || t[1] == 'o')) {
        int base = (t[1] == 'H' || t[1] == 'h') ? 16 : 8;
        size_t b = 2, e = t.size();
        // 去掉尾随类型符 & / % / @ / ! / # 以及行内注释
        while (e > b && (t[e-1] == '&' || t[e-1] == '%' || t[e-1] == '@'
                         || t[e-1] == '!' || t[e-1] == '#')) e--;
        std::string hex = t.substr(b, e - b);
        bool validHex = !hex.empty();
        for (char c : hex) {
            if (base == 16 ? !std::isxdigit(static_cast<unsigned char>(c))
                           : (c < '0' || c > '7')) { validHex = false; break; }
        }
        if (validHex) {
            long long v = (long long)std::stoull(hex, nullptr, base);
            // &H80000000 及以上是 32 位负数 (VB6 Long)
            if (v > 0xFFFFFFFFLL) v &= 0xFFFFFFFFLL;
            return FrmValue::fromInt(v, t);
        }
        return FrmValue::fromInt(0, t);
    }

    // 负数
    bool negative = false;
    size_t numStart = 0;
    if (t[0] == '-' || t[0] == '+') {
        negative = (t[0] == '-');
        numStart = 1;
    }

    // 检查是否为数字
    bool hasDot = false;
    bool allDigit = true;
    for (size_t i = numStart; i < t.size(); i++) {
        char c = t[i];
        if (c == '.' && !hasDot) {
            hasDot = true;
        } else if (!std::isdigit(static_cast<unsigned char>(c))) {
            allDigit = false;
            break;
        }
    }

    if (allDigit && numStart < t.size()) {
        if (hasDot) {
            double v = std::stod(t);
            return FrmValue::fromFloat(v, t);
        } else {
            long long v = std::stoll(t);
            return FrmValue::fromInt(v, t);
        }
    }

    // 标识符 (vbModal, vbUpperCase, True, False 等)
    return FrmValue::fromIdent(t);
}

bool FrmParser::parsePropertyLine(const std::string& line,
    std::string& key, FrmValue& value) {
    // 格式: Key = Value 或 Key = Value ' Comment
    // VB6 .frm属性行格式: 键名(可能含空格) = 值
    // 例如: "Caption         =   "Form1""
    // 注意: 等号前后可能有大量空格(对齐)

    auto eqPos = line.find('=');
    if (eqPos == std::string::npos) return false;

    key = trim(line.substr(0, eqPos));
    if (key.empty()) return false;

    // 跳过属性值后的注释: ' Windows Default
    std::string valuePart = line.substr(eqPos + 1);
    auto commentPos = valuePart.find('\'');
    if (commentPos != std::string::npos) {
        valuePart = valuePart.substr(0, commentPos);
    }

    value = parseValue(valuePart);
    return true;
}

// ============================================================
// 控件类型映射
// ============================================================

FrmControlType FrmParser::parseControlType(const std::string& typeName) {
    // typeName格式: "VB.CommandButton", "MSComctlLib.Toolbar", 等
    std::string lower = typeName;
    for (auto& c : lower) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));

    // VB6标准控件 (VB.xxx)
    if (lower.find("vb.mdiform") != std::string::npos) return FrmControlType::MDIForm;
    if (lower.find("vb.form") != std::string::npos) return FrmControlType::Form;
    if (lower.find("vb.commandbutton") != std::string::npos) return FrmControlType::CommandButton;
    if (lower.find("vb.textbox") != std::string::npos) return FrmControlType::TextBox;
    if (lower.find("vb.label") != std::string::npos && lower.find("vb.labellistbox") == std::string::npos) return FrmControlType::Label;
    if (lower.find("vb.checkbox") != std::string::npos) return FrmControlType::CheckBox;
    if (lower.find("vb.optionbutton") != std::string::npos) return FrmControlType::OptionButton;
    if (lower.find("vb.listbox") != std::string::npos) return FrmControlType::ListBox;
    if (lower.find("vb.combobox") != std::string::npos) return FrmControlType::ComboBox;
    if (lower.find("vb.frame") != std::string::npos) return FrmControlType::Frame;
    if (lower.find("vb.picturebox") != std::string::npos) return FrmControlType::PictureBox;
    if (lower.find("vb.timer") != std::string::npos) return FrmControlType::Timer;
    if (lower.find("vb.hscrollbar") != std::string::npos) return FrmControlType::HScrollBar;
    if (lower.find("vb.vscrollbar") != std::string::npos) return FrmControlType::VScrollBar;
    if (lower.find("vb.image") != std::string::npos) return FrmControlType::Image;
    if (lower.find("vb.shape") != std::string::npos) return FrmControlType::Shape;
    if (lower.find("vb.line") != std::string::npos) return FrmControlType::Line;
    if (lower.find("vb.data") != std::string::npos) return FrmControlType::Data;
    if (lower.find("vb.ole") != std::string::npos) return FrmControlType::OLE;
    if (lower.find("vb.drivelistbox") != std::string::npos) return FrmControlType::DriveListBox;
    if (lower.find("vb.dirlistbox") != std::string::npos) return FrmControlType::DirListBox;
    if (lower.find("vb.filelistbox") != std::string::npos) return FrmControlType::FileListBox;
    if (lower.find("vb.menu") != std::string::npos) return FrmControlType::Menu;

    // WebBrowser (WebView2宿主)
    if (lower.find("webbrowser") != std::string::npos) return FrmControlType::WebBrowser;
    if (lower.find("shdocvw") != std::string::npos) return FrmControlType::WebBrowser;

    // 常见第三方控件
    if (lower.find("progressbar") != std::string::npos) return FrmControlType::ProgressBar;
    if (lower.find("toolbar") != std::string::npos) return FrmControlType::Toolbar;
    if (lower.find("statusbar") != std::string::npos) return FrmControlType::StatusBar;
    if (lower.find("commondialog") != std::string::npos) return FrmControlType::CommonDialog;
    if (lower.find("imagelist") != std::string::npos) return FrmControlType::ImageList;
    // P20-42: SSTab (TabDlg.SSTab)。必须排在 imagelist 之后、Unknown 之前;
    // "sstab" 是 TabDlg.SSTab / SSTab 两种写法的公共子串。
    if (lower.find("sstab") != std::string::npos) return FrmControlType::SSTab;
    // P20-45/46: ListView / TreeView。按字符串长短排: 先 listview 再 treeview,
    // 两者互不为子串, 但都排在 sstab 之后以免误吃。
    if (lower.find("listview") != std::string::npos) return FrmControlType::ListView;
    if (lower.find("treeview") != std::string::npos) return FrmControlType::TreeView;
    // C29-DT: DTPicker (MSComCtl2.DTPicker)。"dtpicker" 与上面任何一条互不为子串,
    // 且 VB6 那侧只有这一种写法 (ProgID 也是 MSComCtl2.DTPicker.1), 排在 Unknown 前即可。
    if (lower.find("dtpicker") != std::string::npos) return FrmControlType::DTPicker;
    // C29-MV: MonthView (MSComCtl2.MonthView)。与上面几条互不为子串（"monthview" 里既没有
    // "listview" 也没有 "treeview"），排在 Unknown 前即可。
    if (lower.find("monthview") != std::string::npos) return FrmControlType::MonthView;
    // C29-RT: RichTextBox (MSFT_TextBox.RichTextBox)。"richtextbox" 不含 "vb.textbox"
    // （库里名前缀是 msft_textbox，中间没有 "vb."），所以放在 textbox 之后也安全；
    // 排在 Unknown 前即可。VB6 里另有写法 RichTextBox.RichTextBox.6 之类，同样命中这条。
    if (lower.find("richtextbox") != std::string::npos) return FrmControlType::RichTextBox;
    // C29-WS: Winsock (MSWinsockLib.Winsock / Winsock. 各种写法都收)。"winsock" 与上面任何一条
    // 互不为子串，排在 Unknown 前即可。VB6 侧另有 ProgID 写法 MSWinsockLib.Winsock.1，同样命中。
    if (lower.find("winsock") != std::string::npos) return FrmControlType::Winsock;
    // C29-SL: Slider (MSComctlLib.Slider)。"slider" 与上面任何一条互不为子串
    // （没有 "sstab"/"listview"/"treeview"/"progressbar" 之类的重叠），排在 Unknown 前即可。
    // VB6 侧另有 ProgID 写法 MSComctlLib.Slider.1，同样命中这条。
    if (lower.find("slider") != std::string::npos) return FrmControlType::Slider;

    return FrmControlType::Unknown;
}

const char* FrmParser::controlTypeToWin32Class(FrmControlType type) {
    switch (type) {
        case FrmControlType::MDIForm:       return "#32770";     // MDI父窗体 (用RegisterClass+MDICLIENT)
        case FrmControlType::Form:         return "#32770";     // 对话框类 (实际用RegisterClass)
        case FrmControlType::CommandButton: return "BUTTON";
        case FrmControlType::TextBox:      return "EDIT";
        case FrmControlType::Label:        return "STATIC";
        case FrmControlType::CheckBox:     return "BUTTON";     // BS_AUTOCHECKBOX样式
        case FrmControlType::OptionButton: return "BUTTON";     // BS_AUTORADIOBUTTON样式
        case FrmControlType::ListBox:      return "LISTBOX";
        case FrmControlType::ComboBox:     return "COMBOBOX";
        case FrmControlType::Frame:        return "BUTTON";     // BS_GROUPBOX样式
        case FrmControlType::PictureBox:   return "STATIC";     // SS_BITMAP样式
        case FrmControlType::HScrollBar:   return "SCROLLBAR";  // SBS_HORZ样式
        case FrmControlType::VScrollBar:   return "SCROLLBAR";  // SBS_VERT样式
        // C29-T: Timer 以前是"无窗口"控件 ⇒ 生成代码里的 vb6_hwnd_<timer> 恒 NULL，
        // 于是 `Timer1.Enabled = True` 这类运行期写法落到 SetPropW(NULL,...) 被静默丢。
        // 现在给它一枚自注册的**不可见**窗口当身份（注册见 vb6forms.c 的
        // vb6_RegisterTimerClass），Enabled/Interval 才找得回自己那一格计时器。
        case FrmControlType::Timer:        return "VB6_TIMER";
        case FrmControlType::Image:        return "STATIC";     // SS_BITMAP
        // ProgressBar: comctl32 通用控件类 (msctls_progress32)。
        // 不加载 mscomctl.ocx —— 本机 InprocServer32 缺失, 且 32 位 inproc
        // OCX 无法进 x64 进程; 改用 Win32 等价类复刻。
        case FrmControlType::ProgressBar:  return "msctls_progress32";
        // P20-40: StatusBar 用 comctl32 的 msctls_status32 复刻 (不加载 mscomctl.ocx)。
        // 漏掉这个 case 会被当成"不可见控件"直接 continue, 窗口根本不创建。
        case FrmControlType::StatusBar:    return "msctls_status32";
        // P20-42: SSTab —— 与 StatusBar 不同, **SysTabControl32 是 comctl32 注册好的**
        // (实测 x64/x86 进程里 GetClassInfoW 直接成功, InitCommonControlsEx 之前就在),
        // 所以这里不需要 RTL 自注册兜底。
        case FrmControlType::SSTab:        return "SysTabControl32";
        // P20-45/46: ListView/TreeView —— 与 SSTab 同型, **comctl32 已注册**
        // (实测 GetClassInfoW 直接成功), 不需要 StatusBar 那套自注册兜底。
        case FrmControlType::ListView:     return "SysListView32";
        case FrmControlType::TreeView:     return "SysTreeView32";
        // C29-DT: DTPicker —— comctl32 注册的 SysDateTimePick32 (头 6653 行
        // DATETIMEPICK_CLASSW)。类已由 vb6forms.c:67 那次 InitCommonControlsEx 的
        // ICC_DATE_CLASSES 请求过 ⇒ 与 TreeView/SSTab 同型, 不需要 RTL 自注册兜底。
        case FrmControlType::DTPicker:     return "SysDateTimePick32";
        // C29-MV: MonthView —— comctl32 注册的 SysMonthCal32（头里 MONTHCAL_CLASSW），
        // 与 DTPicker 同一族，类同样由 vb6forms.c 那次 InitCommonControlsEx 的
        // ICC_DATE_CLASSES 请求过 ⇒ 不需要 RTL 自注册兜底。
        case FrmControlType::MonthView:    return "SysMonthCal32";
        // C29-RT: RichTextBox —— 类不在 comctl32 里，是 Msftedit.dll 注册的 RICHEDIT50W
        // （richedit.h:41 MSFTEDIT_CLASS）。⇒ vb6_ComCtl_Init 里多了一步 LoadLibrary，
        // 否则 CreateWindowEx 直接失败（实测：类不在时 GetClassInfoW 问不出）。
        case FrmControlType::RichTextBox:  return "RICHEDIT50W";
        // C29-5a: Toolbar 同样是 comctl32 注册好的类 (ICC_BAR_CLASSES 在 vb6_ComCtl_Init
        // 里早就请求过)。以前这格缺着 + 被"ImageList || Toolbar 走 CoCreateInstance"那一组
        // 扣住 => 控件根本没窗口，读一个 tb1.Visible 就是 C2065: vb6_hwnd_tb1 未声明。
        case FrmControlType::Toolbar:      return "ToolbarWindow32";
        // C29-SL-a: Slider —— comctl32 注册的 msctls_trackbar32 (头里 TRACKBAR_CLASSW)，
        // 与 ProgressBar/Toolbar 同一族：类由 vb6forms.c 那次 InitCommonControlsEx 的
        // ICC_BAR_CLASSES 请求过 ⇒ 不需要 RTL 自注册兜底（探针 .build/slprobe 实测
        // CreateWindowEx 直接拿到句柄、类名读回就是 msctls_trackbar32）。
        case FrmControlType::Slider:       return "msctls_trackbar32";
        // C29-1b: 文件系统三控件在 VB6 里本来就是公共控件的薄封装 —— Drive 是
        // CBS_DROPDOWNLIST 的组合框、Dir / File 是列表框。以前这三格缺映射, 创建流程
        // 把它们当"不可见控件"跳过, 于是 RTL 里那套 P20-37 填充 helper 从来没被喂过
        // 一个真句柄 (与 Shape/Line 同一类洞, 见 029 §二-2)。
        case FrmControlType::DriveListBox:   return "COMBOBOX";
        case FrmControlType::DirListBox:     return "LISTBOX";
        case FrmControlType::FileListBox:    return "LISTBOX";
        // C29-1a: Shape/Line 是 VB6 的"轻量图形控件"，RTL 里自注册了这两个类
        // (vb6forms_shape.c: vb6_RegisterShapeLineClasses → VB6_SHAPE / VB6_LINE)，
        // 但这里一直缺映射 ⇒ 创建流程把控件当"不可见控件"跳过，句柄永远是 NULL。
        case FrmControlType::Shape:        return "VB6_SHAPE";
        case FrmControlType::Line:         return "VB6_LINE";
        // D6 / C29-9: CommonDialog **不再走 MSComDlg.OCX** —— 那控件只有 32 位，x64 里
        // CoCreateInstance 直接失败，今天整枚控件是静默空转（读数全空、Show* 不出现、
        // 退出码照旧 0，见 029 §九）。这里给它一枚自注册的**不可见**类当属性宿主：
        // 有句柄才谈得上 SetPropW 存属性、GetParent 拿模态父窗（注册见 vb6forms_ctrl.c）。
        case FrmControlType::CommonDialog: return "VB6_COMMONDIALOG";
        // OLE 容器 (C29-OLE): RTL 自注册类 (vb6forms_olecon.c vb6_RegisterOleConClass)。
        // 嵌入对象依赖目标机器的 OLE 服务器 ⇒ 判据只本地跑 (用户指示), 不进 CI。
        case FrmControlType::OLE:          return "VB6_OLECONTAINER";
        // C29-Data: ODBC 后端的 Data 控件 (自注册不可见类, CommonDialog 属性宿主同款)。
        case FrmControlType::Data:         return "VB6_DATA";
        // C29-WS: Winsock —— 无外观，但**这枚的窗口要真收 WSAAsyncSelect 的 FD_\***，
        // 不像 Timer / CommonDialog 只是属性袋（类注册在 vb6forms_winsock.c）。
        case FrmControlType::Winsock:       return "VB6_WINSOCK";
        case FrmControlType::Menu:         return nullptr;       // 菜单, 非窗口
        case FrmControlType::WebBrowser:  return nullptr;       // WebView2, 运行时动态创建
        default:                           return nullptr;
    }
}

const char* FrmParser::controlTypeToVb6Name(FrmControlType type) {
    switch (type) {
        case FrmControlType::MDIForm:       return "MDIForm";
        case FrmControlType::Form:         return "Form";
        case FrmControlType::CommandButton: return "CommandButton";
        case FrmControlType::TextBox:      return "TextBox";
        case FrmControlType::Label:        return "Label";
        case FrmControlType::CheckBox:     return "CheckBox";
        case FrmControlType::OptionButton: return "OptionButton";
        case FrmControlType::ListBox:      return "ListBox";
        case FrmControlType::ComboBox:     return "ComboBox";
        case FrmControlType::Frame:        return "Frame";
        case FrmControlType::PictureBox:   return "PictureBox";
        case FrmControlType::Timer:        return "Timer";
        case FrmControlType::HScrollBar:   return "HScrollBar";
        case FrmControlType::VScrollBar:   return "VScrollBar";
        case FrmControlType::Image:        return "Image";
        case FrmControlType::ProgressBar:  return "ProgressBar";
        case FrmControlType::StatusBar:    return "StatusBar";
        case FrmControlType::SSTab:        return "SSTab";
        case FrmControlType::ListView:     return "ListView";
        case FrmControlType::TreeView:     return "TreeView";
        case FrmControlType::DTPicker:     return "DTPicker";
        case FrmControlType::MonthView:    return "MonthView";
        case FrmControlType::RichTextBox:  return "RichTextBox";
        case FrmControlType::Toolbar:      return "Toolbar";
        case FrmControlType::Shape:        return "Shape";
        case FrmControlType::Line:         return "Line";
        case FrmControlType::DriveListBox: return "DriveListBox";
        case FrmControlType::DirListBox:     return "DirListBox";
        case FrmControlType::FileListBox:    return "FileListBox";
        case FrmControlType::Menu:         return "Menu";
        case FrmControlType::WebBrowser:  return "WebBrowser";
        default:                           return "Control";
    }
}

} // namespace vb6c3
