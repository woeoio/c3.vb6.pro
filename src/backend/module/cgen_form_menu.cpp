// cgen_form_menu.cpp — 窗体菜单构建与 C 字符串转义
// 2026-09-17 由 backend/module/cgen_form.cpp 第 1999~2213 行纯搬移拆出，逐行未改。
// 搬移原因：cgen_form.cpp 的 emitFormFramework 单函数 1945 行占满全文件，
//           把与之无关的菜单(Win32 菜单构建 / WM_COMMAND 派发)与转义成员函数移出，
//           主文件只留窗体框架，便于按「函数体片段」继续切分。
#include "backend/cgen.hpp"

#include <cstdio>

namespace vb6c3 {

// P7.8: 递归生成菜单项
void CCodeGen::emitMenuItem(const std::string& parentVar, const FrmControl& menuCtrl, int& menuId) {
    std::string caption = menuCtrl.controlName;
    auto capIt = menuCtrl.properties.find("Caption");
    if (capIt != menuCtrl.properties.end() && capIt->second.type == FrmValueType::String) {
        std::string raw = capIt->second.rawText;
        if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') {
            caption = raw.substr(1, raw.size() - 2);
        }
    }

    // Check for separator: Caption = "-"
    if (caption == "-") {
        c_.emitLine("AppendMenuW(" + parentVar + ", MF_SEPARATOR, 0, NULL);");
        return;
    }

    // Extract menu properties
    bool checked = false;
    bool enabled = true;
    bool visible = true;

    auto chkIt = menuCtrl.properties.find("Checked");
    if (chkIt != menuCtrl.properties.end() && chkIt->second.type == FrmValueType::Integer) {
        checked = (chkIt->second.intValue != 0);
    }
    auto enIt = menuCtrl.properties.find("Enabled");
    if (enIt != menuCtrl.properties.end() && enIt->second.type == FrmValueType::Identifier) {
        if (enIt->second.rawText == "0" || enIt->second.rawText == "False") enabled = false;
    }
    // Also check integer Enabled = 0
    if (enIt != menuCtrl.properties.end() && enIt->second.type == FrmValueType::Integer) {
        if (enIt->second.intValue == 0) enabled = false;
    }
    auto visIt = menuCtrl.properties.find("Visible");
    if (visIt != menuCtrl.properties.end() && visIt->second.type == FrmValueType::Identifier) {
        if (visIt->second.rawText == "0" || visIt->second.rawText == "False") visible = false;
    }
    if (visIt != menuCtrl.properties.end() && visIt->second.type == FrmValueType::Integer) {
        if (visIt->second.intValue == 0) visible = false;
    }

    if (!visible) {
        // Invisible menu item: skip entirely, but still assign ID
        menuId++;
        return;
    }

    long flags = 0;  // MF_STRING = 0
    if (checked) flags |= 0x0008;  // MF_CHECKED
    if (!enabled) flags |= 0x0002;  // MF_GRAYED

    if (!menuCtrl.children.empty()) {
        // Submenu: create a popup
        std::string popupVar = "hPopup_" + cIdent(menuCtrl.controlName);
        c_.emitLine("HMENU " + popupVar + " = CreatePopupMenu();");

        for (const auto& child : menuCtrl.children) {
            emitMenuItem(popupVar, child, menuId);
        }

        c_.emitLine("AppendMenuW(" + parentVar + ", 0x0010L | " + std::to_string(flags) + ", (UINT_PTR)" + popupVar + ", L\"" + escapeWideCString(caption) + "\");");
        // 0x0010 = MF_POPUP
    } else {
        // Leaf menu item
        c_.emitLine("AppendMenuW(" + parentVar + ", " + std::to_string(flags) + ", " + std::to_string(menuId) + ", L\"" + escapeWideCString(caption) + "\");");
        menuId++;
    }
}

// P7.8: 递归生成菜单点击事件派发 (WM_COMMAND中)
void CCodeGen::emitMenuClickDispatch(const FrmControl& menuCtrl, int& menuId) {
    // Fix <vbeclipse>: 菜单项在 .frm 里存在但源文件**没有**对应 Click 过程
    // (frmToolWin.frm 的 mnuMaximize/mnuClose 只有 Begin VB.Menu, 无
    // `Private Sub mnuMaximize_Click`) — VB6 语义是"点了没反应"。此前无条件发
    // `{ extern void vb6_<f>_<mnu>_Click(); ... }` → 该过程无定义 → LNK2001
    // "无法解析的外部符号". 只有真实存在过程时才派发 (菜单本身仍创建, 保留 ID).
    auto hasClickHandler = [&](const std::string& ctrlName) -> bool {
        Symbol* sym = symTab_.lookupModule(ctrlName + "_Click");
        return sym && (sym->kind == SymbolKind::Sub || sym->kind == SymbolKind::Function);
    };
    // M22-Issue3: 如果顶层Menu控件没有children, 它自身就是叶菜单项, 需要生成Click处理
    if (menuCtrl.children.empty()) {
        std::string caption = menuCtrl.controlName;
        auto capIt = menuCtrl.properties.find("Caption");
        if (capIt != menuCtrl.properties.end() && capIt->second.type == FrmValueType::String) {
            std::string raw = capIt->second.rawText;
            if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') {
                caption = raw.substr(1, raw.size() - 2);
            }
        }
        if (caption != "-") {
            bool visible = true;
            auto visIt = menuCtrl.properties.find("Visible");
            if (visIt != menuCtrl.properties.end()) {
                if (visIt->second.type == FrmValueType::Identifier &&
                    (visIt->second.rawText == "0" || visIt->second.rawText == "False")) {
                    visible = false;
                }
                if (visIt->second.type == FrmValueType::Integer && visIt->second.intValue == 0) {
                    visible = false;
                }
            }
            if (visible && hasClickHandler(menuCtrl.controlName)) {
                std::string clickFn = cProcName(menuCtrl.controlName + "_Click", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(menuId) + ") {");
                c_.indent();
                c_.emitLine("{ extern void " + clickFn + "(); " + clickFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            menuId++;
        }
        return;
    }
    // 顶层Menu控件 (如mnuFile) 只是popup容器, 只对叶子菜单项和子菜单项生成WM_COMMAND派发
    for (const auto& child : menuCtrl.children) {
        std::string caption = child.controlName;
        auto capIt = child.properties.find("Caption");
        if (capIt != child.properties.end() && capIt->second.type == FrmValueType::String) {
            std::string raw = capIt->second.rawText;
            if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') {
                caption = raw.substr(1, raw.size() - 2);
            }
        }

        if (caption == "-") {
            // 分隔线没有点击事件，不分配ID
            continue;
        }

        // 检查Visible属性
        bool visible = true;
        auto visIt = child.properties.find("Visible");
        if (visIt != child.properties.end()) {
            if (visIt->second.type == FrmValueType::Identifier &&
                (visIt->second.rawText == "0" || visIt->second.rawText == "False")) {
                visible = false;
            }
            if (visIt->second.type == FrmValueType::Integer && visIt->second.intValue == 0) {
                visible = false;
            }
        }

        if (!child.children.empty()) {
            // 子菜单: 递归处理子项
            emitMenuClickDispatch(child, menuId);
        } else {
            // 叶子菜单项
            if (visible && hasClickHandler(child.controlName)) {
                std::string clickFn = cProcName(child.controlName + "_Click", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(menuId) + ") {");
                c_.indent();
                c_.emitLine("{ extern void " + clickFn + "(); " + clickFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            menuId++;
        }
    }
}
// M22: 转义宽C字符串（UTF-8多字节→Unicode \xNNNN转义）
std::string CCodeGen::escapeWideCString(const std::string& s) {
    std::string result;
    result.reserve(s.size() + 16);
    for (size_t j = 0; j < s.size(); ) {
        unsigned char ch = (unsigned char)s[j];
        if (ch == '"') {
            result += "\\\""; j++;
        } else if (ch == '\\') {
            result += "\\\\"; j++;
        } else if (ch == '\n') {
            result += "\\n"; j++;
        } else if (ch == '\r') {
            result += "\\r"; j++;
        } else if (ch == '\t') {
            result += "\\t"; j++;
        } else if (ch < 0x80) {
            result += (char)ch; j++;
        } else {
            // Fix 021: UTF-8多字节解码 Unicode 码点 (与 cgen_expr.cpp 同样修复)
            // (a) 4-byte UTF-8 lead 掩码错 (0xF8→0xF0)
            // (b) \x%04X -> \u%04X: \x 会贪婪吃掉后续十六进制字符触发 MSVC C7744
            uint32_t cp = 0;
            int bytes = 0;
            if ((ch & 0xE0) == 0xC0) { cp = ch & 0x1F; bytes = 2; }
            else if ((ch & 0xF0) == 0xE0) { cp = ch & 0x0F; bytes = 3; }
            else if ((ch & 0xF8) == 0xF0) { cp = ch & 0x07; bytes = 4; }
            else { cp = ch; bytes = 1; }
            for (int b = 1; b < bytes && j + b < s.size(); b++) {
                cp = (cp << 6) | ((unsigned char)s[j + b] & 0x3F);
            }
            j += bytes;
            char hex[16];
            if (cp <= 0xFFFF) {
                snprintf(hex, sizeof(hex), "\\u%04X", cp);
            } else {
                // Supplementary plane: UTF-16 surrogate pair
                uint32_t v = cp - 0x10000;
                uint16_t hi = 0xD800 + (v >> 10);
                uint16_t lo = 0xDC00 + (v & 0x3FF);
                snprintf(hex, sizeof(hex), "\\u%04X\\u%04X", hi, lo);
            }
            result += hex;
        }
    }
    return result;
}

// P7.8: 转义C字符串
std::string CCodeGen::escapeCString(const std::string& s) {
    std::string result;
    for (char c : s) {
        switch (c) {
            case '"':  result += "\\\""; break;
            case '\\': result += "\\\\"; break;
            case '\n': result += "\\n"; break;
            case '\r': result += "\\r"; break;
            case '\t': result += "\\t"; break;
            default:   result += c; break;
        }
    }
    return result;
}

} // namespace vb6c3
