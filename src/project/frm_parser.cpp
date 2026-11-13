// vb6c3 - VB6窗体文件(.frm)解析器
// .frm 文件结构: VERSION行 + Begin...End窗体描述块 + Attribute行 + VB代码段
// 窗体描述块包含: 窗体属性 + 嵌套控件(递归Begin...End) + 复合属性(BeginProperty...EndProperty)

#include "project/frm_parser.hpp"
#include "common/encoding.hpp"
#include "common/source_manager.hpp"
#include <sstream>
#include <algorithm>
#include <cctype>

namespace vb6c3 {

// ============================================================
// 工具函数
// ============================================================


// ============================================================
// 复合属性块解析 (BeginProperty...EndProperty)
// ============================================================

FrmPropertyBlock FrmParser::parsePropertyBlock(
    const std::vector<std::string>& lines,
    size_t& lineIdx) {
    // 当前行格式:
    //   BeginProperty Font
    //   BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628}
    //     NumListImages = 5
    //     BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628}
    //       Picture = "Form1.frx":1A89
    //       Key = ""
    //     EndProperty
    //   EndProperty

    FrmPropertyBlock block;

    // 解析 BeginProperty 行: "BeginProperty Name" 或 "BeginProperty Name {GUID}"
    std::string line = trim(lines[lineIdx]);
    // "BeginProperty Font" → rest = "Font"
    // "BeginProperty Images {GUID}" → rest = "Images {GUID}"
    auto spacePos = line.find(' ');
    if (spacePos != std::string::npos) {
        std::string rest = trim(line.substr(spacePos));
        // 分离 blockName 和 blockGuid
        auto bracePos = rest.find('{');
        if (bracePos != std::string::npos) {
            block.blockName = trim(rest.substr(0, bracePos));
            auto endBrace = rest.find('}', bracePos);
            if (endBrace != std::string::npos) {
                block.blockGuid = rest.substr(bracePos, endBrace - bracePos + 1);
            }
        } else {
            block.blockName = rest;
        }
    }

    lineIdx++;  // 进入块体

    while (lineIdx < lines.size()) {
        std::string curLine = trim(lines[lineIdx]);

        if (curLine.empty()) {
            lineIdx++;
            continue;
        }

        // EndProperty — 结束
        if (curLine.find("EndProperty") == 0) {
            lineIdx++;
            break;
        }

        // 嵌套 BeginProperty — 递归解析
        if (curLine.find("BeginProperty") == 0) {
            auto nested = parsePropertyBlock(lines, lineIdx);
            block.nestedBlocks.push_back(std::move(nested));
            continue;
        }

        // 属性行: Key = Value (包括 "Object.Tag" 等带点号的键)
        std::string key;
        FrmValue value;
        if (parsePropertyLine(curLine, key, value)) {
            block.properties[key] = value;
        }

        lineIdx++;
    }

    return block;
}

// ============================================================
// 控件内集合块的项行 (P20-39)
// ============================================================
// 形态: `.ListImage(1, "dt1", "CtrlImageList.frx":00000345)`
//       `.Button(1, "New", "New", "New", 3)` 等
// 参数按 VB6 的位置语义命名 —— 各控件集合的签名不同, 这里按最常见的 ImageList 口径
// 命名, 其余控件 (.Button/.ListItem/.Node/.Panel) 后续按需加映射即可。

static const char* kCollArgNames[5] = {
    "Index", "Key", "Picture", "Key2", "Picture2"
};

FrmPropertyBlock FrmParser::parseCollectionItem(const std::string& line) {
    FrmPropertyBlock item;

    // 项名: '.' 与 '(' 之间那一截
    size_t dot = line.find('.');
    size_t open = line.find('(');
    if (dot != std::string::npos && open != std::string::npos && open > dot) {
        item.blockName = trim(line.substr(dot + 1, open - dot - 1));
    } else {
        item.blockName = "Item";
    }

    // 实参: 括号里的东西按**顶层逗号**切 (值里可能带逗号/frx 引用里没有, 但字符串里可能有)。
    // **必须到第一个 ')' 为止**, 不能按行尾取 —— StatusBar 的 `Panels(1) = "Ready"` 行尾是
    // 引号, 按行尾截会把 `1)      =   "Ready"` 整个当成实参, Index 于是解析不出来。
    std::string args;
    if (open != std::string::npos) {
        size_t cp = line.find(')', open);
        if (cp != std::string::npos && cp > open)
            args = line.substr(open + 1, cp - open - 1);
    }

    // 尾随 `= "..."` 的默认实参 (VB6 的 StatusBar 就这么写 `Panels(1) = "Ready"`,
    // 语义等价于 Text)。不接住的话这条文本会整条丢掉 —— 面板显示成空串。
    // 从**闭括号之后**找第一个 '=': 这样实参内部的等号 (`"a=b"`) 不会被误当分隔符。
    size_t closeParen = (open != std::string::npos) ? line.find(')', open) : std::string::npos;
    size_t eq = (closeParen != std::string::npos) ? line.find('=', closeParen) : std::string::npos;
    if (eq != std::string::npos && eq > open) {
        FrmValue dv = parseValue(trim(line.substr(eq + 1)));
        if (dv.type == FrmValueType::String) item.properties["Text"] = dv;
    }

    size_t start = 0;
    int slot = 0;
    while (start <= args.size() && slot < 5) {
        size_t comma = args.find(',', start);
        std::string one = (comma == std::string::npos)
            ? args.substr(start) : args.substr(start, comma - start);
        std::string t = trim(one);
        if (!t.empty()) {
            std::string keyName = kCollArgNames[slot];
            item.properties[keyName] = parseValue(t);
        }
        if (comma == std::string::npos) break;
        start = comma + 1;
        slot++;
    }
    return item;
}

// ============================================================
// 控件块解析 (Begin...End)
// ============================================================

FrmControl FrmParser::parseControlBlock(
    const std::vector<std::string>& lines,
    size_t& lineIdx,
    const std::string& parentIndent) {
    // 当前行: Begin VB.CommandButton Command1
    //   Caption = "OK"
    //   ...
    //   Begin VB.Label Label1
    //      ...
    //   End
    // End

    FrmControl ctrl;

    // 解析 Begin 行: "Begin TypeName InstanceName" 或 "Begin TypeName InstanceName(Index)"
    std::string line = trim(lines[lineIdx]);

    // 去掉 "Begin " 前缀
    auto spacePos = line.find(' ');
    if (spacePos != std::string::npos) {
        std::string rest = trim(line.substr(spacePos));

        // TypeName 和 InstanceName 分离
        // 例如: "VB.CommandButton Command1" 或 "VB.CommandButton Command1(0)"
        auto namePos = rest.find(' ');
        if (namePos != std::string::npos) {
            ctrl.controlTypeName = rest.substr(0, namePos);
            std::string instancePart = trim(rest.substr(namePos));

            // 检查控件数组: Command1(0)
            auto parenPos = instancePart.find('(');
            if (parenPos != std::string::npos) {
                ctrl.controlName = instancePart.substr(0, parenPos);
                auto closeParen = instancePart.find(')', parenPos);
                if (closeParen != std::string::npos) {
                    std::string idxStr = instancePart.substr(parenPos + 1, closeParen - parenPos - 1);
                    try { ctrl.index = std::stoi(idxStr); } catch (...) { ctrl.index = -1; }
                }
            } else {
                ctrl.controlName = instancePart;
            }
        } else {
            ctrl.controlTypeName = rest;
        }
    }

    ctrl.controlType = parseControlType(ctrl.controlTypeName);

    // P20-39: 控件内的**集合块**。VB6 的 .frm 里它不是 Begin..End 也不是 BeginProperty,
    // 而是「一条裸标题 + 若干 `.项(...)` 行 + 一个 End」:
    //      ListImages
    //         .ListImage(1, "dt1", "CtrlImageList.frx":00000000)
    //      End
    // 只认 Begin/BeginProperty 的话, 这里第一个 End 就把控件块提前收掉, 外层的 End
    // 于是变成 `VB2002: unexpected token at module level: End`。
    // 这条不只是 ImageList —— Toolbar.Buttons / ListView.ListItems+ColumnHeaders /
    // TreeView.Nodes / StatusBar.Panels 全是这个形态, 后面的控件都要用。
    FrmPropertyBlock coll;         // 正在累积的集合块
    std::string     collName;      // 非空 = 已打开一个集合块, 等它的 End
    size_t          collIndent = (size_t)-1;  // 集合块的缩进; (size_t)-1 = 尚未打开

    lineIdx++;  // 进入块体

    while (lineIdx < lines.size()) {
        const std::string& rawLine = lines[lineIdx];
        std::string curLine = trim(rawLine);
        // 缩进量 —— 判定 End 到底收的是集合还是控件, 只能靠它 (见下方 End 分支)。
        size_t lineIndent = rawLine.size() - curLine.size();

        if (curLine.empty()) {
            lineIdx++;
            continue;
        }

        // 集合块标题行: 控件体里不含 '=' 的**单个裸标识符** (ListImages/Buttons/Nodes/...)。
        // 判定必须够窄, 宽松一点就会把 `End` / `Begin xxx` 这类块边界当成集合标题:
        // 那会让控件块一路吃到 EOF, 代码段被整段吞掉 (生成物里 Form_Load 凭空消失)。
        if (collName.empty() && curLine.find('=') == std::string::npos
            && curLine.find('(') == std::string::npos
            && curLine.find(' ') == std::string::npos
            && curLine.find('\t') == std::string::npos
            && curLine != "End"
            && curLine.compare(0, 5, "Begin") != 0) {
            collName = curLine;
            coll = FrmPropertyBlock();
            coll.blockName = curLine;
            collIndent = lineIndent;
            lineIdx++;
            continue;
        }

        // 集合块的项行: `.ListImage(1, "dt1", "CtrlImageList.frx":00000000)`
        // 也收**没有前导点号**的那种 —— VB6 的 StatusBar/Toolbar 就这么写
        // `Panels(1) = "Ready"` / `Buttons(1) = "Open"`, 少了这条分支这些行会当成控件属性
        // 挂到 StatusBar1 自己身上, 集合恒空且不报任何错。
        // 收尾判定不能用 `back() == ')'` —— StatusBar 的 `Panels(1) = "Ready"` 以引号结尾。
        // 只要括号里有实参就算项行 (尾部可以跟 `= "值"` 这种默认实参)。
        // ⚠ 括号必须出现在 '=' **之前**（括号是键名上的下标）。只看"行里有 ( 且有 ="
        // 会把**值里带括号**的普通属性行也吃进来：实测 `Filter = "文本 (*.txt)|*.txt"`
        // 被判成集合项行 ⇒ Filter 从此不进 ctrl.properties ⇒ 设计期不写、读回空
        // （ctrldlg 的 DL5=N 就是这么来的）。VB6 的 CommonDialog Filter/DialogTitle
        // 里带括号极常见，属真 bug。合并引入，2026-09-26 修。
        // ⚠ Task #44 (SSTabEx): `TabCaption(0) = "Theme"` / `TabVisible(2)` /
        // `TabPic16(0)` 是 **控件自身的带下标属性**，`Tab(0).ControlCount = 1` /
        // `Tab(0).Control(0) = "Picture2(0)"` 是 **点号分级的下标属性** —— 文本形态与
        // StatusBar `Panels(1) = "Ready"` 的集合项行同形，但语义相反：吃成集合项会让
        // 设计期 caption 全进 "Items" 集合 (SSTabEx frmTest: 标签条 7 页全空、页归属
        // 只能靠 Left 反推)。排除判据：
        //   ①键名(到'='前)含 '.' → Tab(n).Control* 形态；集合项行键名从不含点
        //     (子属性行以 '.' **开头**，由下方专门的分支收)。
        //   ②键名命中 SSTab 家族已知下标属性前缀 (TabCaption/TabVisible/TabPicNN)。
        if (curLine.size() > 1
            && curLine.find('(') != std::string::npos
            && (curLine.find('=') == std::string::npos
                || curLine.find('(') < curLine.find('='))
            && (curLine.front() == '.' || isalpha((unsigned char)curLine.front()))
            && (curLine.back() == ')' || curLine.find('=') != std::string::npos)) {
            {
                std::string keyPart44 = curLine;
                size_t eq44 = keyPart44.find('=');
                if (eq44 != std::string::npos) keyPart44 = keyPart44.substr(0, eq44);
                // trim (与 parsePropertyLine 的键口径一致)
                size_t b44 = keyPart44.find_first_not_of(" \t\r\n");
                size_t e44 = keyPart44.find_last_not_of(" \t\r\n");
                keyPart44 = (b44 == std::string::npos) ? "" : keyPart44.substr(b44, e44 - b44 + 1);
                // ⚠ 修正 (#44 收窄, 同日): 排除只该打 **标识符开头** 的复合键
                // (Tab(0).Control(0) — 键名自身带点号分级)。**以 '.' 开头** 的行是
                // VB6 集合块的成员条目 (`.ListImage(1, "dt1", …frx…)` 隶属上一行
                // ListImages 集合头), 必须照旧走集合项路径 —— 上面的项行判据首句
                // 就放行点开头行; "含点即排除" 曾把 ctrlimagelist 的两张设计期图
                // 读成 0 (IL1-DTCOUNT=0, 全量回归实抓)。
                bool dotted44 = (keyPart44.find('.') != std::string::npos)
                             && (!keyPart44.empty() && keyPart44.front() != '.');
                bool sstabIdx44 = false;
                static const char* kSstabIdxProps44[] = {
                    "TabCaption", "TabVisible", "TabPic16", "TabPic20", "TabPic24",
                    "TabToolTipText"  // 预留; 现无下标形态, 命中也无害
                };
                if (!dotted44) {
                    for (auto* p44 : kSstabIdxProps44) {
                        size_t n44 = strlen(p44);
                        if (_strnicmp(keyPart44.c_str(), p44, n44) == 0
                            && keyPart44.size() > n44 && keyPart44[n44] == '(') {
                            sstabIdx44 = true;
                            break;
                        }
                    }
                }
                if (!dotted44 && !sstabIdx44) {
                    if (collName.empty()) {
                        // 没有标题行也照样收下 —— 否则这些行 parsePropertyLine 判不出来,
                        // 会被静默丢掉 (表现为"设计期图片凭空少一张")。
                        collName = "Items";
                        coll = FrmPropertyBlock();
                        coll.blockName = collName;
                        collIndent = lineIndent;
                    }
                    FrmPropertyBlock item = parseCollectionItem(curLine);
                    coll.nestedBlocks.push_back(std::move(item));
                    lineIdx++;
                    continue;
                }
                // dotted44 / sstabIdx44 → 落到普通属性行分支 (下方 parsePropertyLine)
            }
        }

        // 集合块内 **子属性行**: `.Key = "k"` / `.Style = 1` / `.Text = "x"`。
        // 必须挂到**上一项**身上, 挂到控件上就错了 (VB6 里这些属于 Panels(1) 而不是 StatusBar1)。
        // 后续 Toolbar.Buttons / ListView.ListItems / TreeView.Nodes 全是这个形态, 现在不修
        // 后面每个控件都要重踩一遍。
        if (!collName.empty() && !coll.nestedBlocks.empty()
            && curLine.size() > 1 && curLine.front() == '.' && curLine.find('=') != std::string::npos
            // 同款: 括号在 '=' 之后是**值**的一部分 (`.Text = "a (b)"`), 不该排除。
            && (curLine.find('(') == std::string::npos
                || curLine.find('(') > curLine.find('='))) {
            FrmValue sv;
            std::string sk;
            if (parsePropertyLine(curLine, sk, sv) && !sk.empty()) {
                // parsePropertyLine 连前导点号一起返回 (`.Key` → ".Key"), 而生成侧按
                // **不带点**的名字取 (与项行实参名表 kCollArgNames 的口径一致)。
                if (!sk.empty() && sk.front() == '.') sk.erase(0, 1);
                if (!sk.empty()) {
                    coll.nestedBlocks.back().properties[sk] = sv;
                    lineIdx++;
                    continue;
                }
            }
        }

        // End — 收掉当前集合块, 或结束控件块。
        // **只能靠缩进区分**: 一个控件里可以有多个集合 (StatusBar 的 Panels 就能有
        // 好几项各自的 End), 只认 collName 的话第二组 End 会被当成"收集合", 控件块的
        // 收尾 End 于是被吃掉 → 控件块一路 break 到 EOF → 后面的代码段整体被当代码解析,
        // 报一连串 `VB2002: unexpected token at module level: End`。
        if (curLine == "End") {
            // collIndent 是"打开集合的那一行"的缩进 (标题行, 没有标题行时取第一个项行)。
            // VB6 把集合的收尾 End 写在与首行**同一列**上, 项与子属性再往里缩 3 格。
            if (!collName.empty() && lineIndent == collIndent) {
                if (coll.blockName.empty()) coll.blockName = collName;
                ctrl.propertyBlocks.push_back(std::move(coll));
                coll = FrmPropertyBlock();
                collName.clear();
                collIndent = (size_t)-1;
                lineIdx++;
                continue;
            }
            lineIdx++;
            break;
        }

        // Begin — 嵌套子控件
        if (curLine.find("Begin ") == 0 || curLine == "Begin") {
            auto child = parseControlBlock(lines, lineIdx, parentIndent + "   ");
            ctrl.children.push_back(std::move(child));
            continue;
        }

        // BeginProperty — 复合属性块
        if (curLine.find("BeginProperty") == 0) {
            auto propBlock = parsePropertyBlock(lines, lineIdx);
            ctrl.propertyBlocks.push_back(std::move(propBlock));
            continue;
        }

        // 普通属性行
        std::string key;
        FrmValue value;
        if (parsePropertyLine(curLine, key, value)) {
            ctrl.properties[key] = value;
            // P7.6: Index属性设置控件数组索引
            if (key == "Index" && value.type == FrmValueType::Integer) {
                ctrl.index = (int)value.intValue;
            }
        }

        lineIdx++;
    }

    return ctrl;
}

// ============================================================
// 代码段提取
// ============================================================

std::string FrmParser::extractCodeSection(
    const std::vector<std::string>& lines,
    size_t startIdx) {
    // 从 startIdx 开始收集所有后续行作为VB代码段
    // 这些行包含: Attribute语句, Option语句, Sub/Function定义等
    std::ostringstream oss;
    for (size_t i = startIdx; i < lines.size(); i++) {
        oss << lines[i] << "\n";
    }
    return oss.str();
}

// ============================================================
// 主解析入口
// ============================================================

FrmFile FrmParser::parse(const std::string& frmFilePath) {
    FrmFile frmFile;
    frmFile.frmFilePath = std::filesystem::absolute(utf8ToPath(frmFilePath));

    // M22: 使用编码检测+转换读取, 确保GBK等非UTF-8文件正确解码
    auto readResult = SourceBuffer::readAndConvertToUtf8(frmFilePath);
    if (readResult.content.empty()) {
        // Fix 196: 明确标记失败。合法 .frm 至少有 VERSION 行, 内容为空只可能是
        // 打不开 (路径错/不存在/权限)。调用方报错, 不再静默产出空模块。
        frmFile.readFailed = true;
        return frmFile;
    }

    return parseString(readResult.content, frmFilePath);
}

FrmFile FrmParser::parseString(const std::string& content, const std::string& frmFilePath) {
    FrmFile frmFile;
    if (!frmFilePath.empty()) {
        frmFile.frmFilePath = std::filesystem::absolute(utf8ToPath(frmFilePath));
    }

    // 按行分割
    std::vector<std::string> lines;
    std::istringstream stream(content);
    std::string line;
    while (std::getline(stream, line)) {
        // 统一换行: 去掉行尾 \r
        if (!line.empty() && line.back() == '\r') {
            line.pop_back();
        }
        lines.push_back(line);
    }

    size_t lineIdx = 0;

    // 1. VERSION 行
    if (lineIdx < lines.size()) {
        std::string verLine = trim(lines[lineIdx]);
        if (verLine.find("VERSION") == 0) {
            auto spacePos = verLine.find(' ');
            if (spacePos != std::string::npos) {
                frmFile.version = trim(verLine.substr(spacePos));
            }
            lineIdx++;
        }
    }

    // 1.5 Skip Object= lines (ActiveX control references in .frm header)
    // e.g. Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
    while (lineIdx < lines.size()) {
        std::string objLine = trim(lines[lineIdx]);
        if (objLine.empty() ||
            objLine.find("Object ") == 0 ||
            objLine.find("Object=") == 0) {
            lineIdx++;
            continue;
        }
        break;
    }

    // 2. Begin VB.Form ... End 窗体描述块
    if (lineIdx < lines.size()) {
        std::string beginLine = trim(lines[lineIdx]);
        if (beginLine.find("Begin ") == 0) {
            // 解析窗体控件块 (递归解析包含嵌套控件)
            auto formCtrl = parseControlBlock(lines, lineIdx);

            // 提取窗体名
            if (formCtrl.controlType == FrmControlType::Form) {
                frmFile.form.formName = formCtrl.controlName;
            } else {
                // 可能不是VB.Form (如VB.MDIForm), 仍保留
                frmFile.form.formName = formCtrl.controlName;
            }

            frmFile.form.formControl = std::move(formCtrl);

            // P7.7: 检测MDIChild属性
            auto mdiChildIt = frmFile.form.formControl.properties.find("MDIChild");
            if (mdiChildIt != frmFile.form.formControl.properties.end() &&
                mdiChildIt->second.type == FrmValueType::Integer && mdiChildIt->second.intValue != 0) {
                frmFile.form.isMDIChild = true;
            }
        }
    }

    // 3. 代码段 (Attribute + VB代码)
    frmFile.codeSection = extractCodeSection(lines, lineIdx);

    // 从代码段的 Attribute VB_Name 中提取窗体名 (优先)
    // Attribute VB_Name = "Form1"
    std::istringstream codeStream(frmFile.codeSection);
    while (std::getline(codeStream, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        std::string t = trim(line);
        if (t.find("Attribute") == 0 && t.find("VB_Name") != std::string::npos) {
            auto eqPos = t.find('=');
            if (eqPos != std::string::npos) {
                std::string nameVal = unquote(trim(t.substr(eqPos + 1)));
                if (!nameVal.empty()) {
                    frmFile.form.formName = nameVal;
                }
            }
        }
    }

    return frmFile;
}

// ============================================================
// Fix 195: 外部二进制资源 (.frx/.ctx/.pgx) 的收集与定位
// ============================================================

namespace {

void collectRefsInBlocks(const std::vector<FrmPropertyBlock>& blocks,
                         const std::string& prefix,
                         std::vector<FrmResourceRef>& out) {
    for (const auto& blk : blocks) {
        std::string blkPrefix = prefix + "." + blk.blockName;
        for (const auto& kv : blk.properties) {
            if (kv.second.type != FrmValueType::FrxReference) continue;
            out.push_back({blkPrefix + "." + kv.first, kv.second.frxFile, kv.second.frxOffset,
                           kv.second.rawText});
        }
        collectRefsInBlocks(blk.nestedBlocks, blkPrefix, out);
    }
}

} // namespace

void collectFrmResourceRefs(const FrmControl& ctrl, std::vector<FrmResourceRef>& out) {
    // 控件数组元素在 VB 里写作 `Label1(0)`, 路径里带上索引, 这样 where 本身
    // 就是一句合法的 VB 左值 (导出器直接照抄)。
    const std::string self = ctrl.controlName
        + (ctrl.index >= 0 ? "(" + std::to_string(ctrl.index) + ")" : std::string());
    for (const auto& kv : ctrl.properties) {
        if (kv.second.type != FrmValueType::FrxReference) continue;
        out.push_back({self + "." + kv.first, kv.second.frxFile, kv.second.frxOffset,
                       kv.second.rawText});
    }
    collectRefsInBlocks(ctrl.propertyBlocks, self, out);
    // 子控件 (含 Frame 内的) 一律按自己的名字拼路径 —— VB6 就是这么寻址的。
    for (const auto& child : ctrl.children) collectFrmResourceRefs(child, out);
}

std::filesystem::path resolveFrmResourceFile(const FrmFile& frm,
                                             const std::string& fallbackExt) {
    auto candidate = frm.frmFilePath;
    candidate.replace_extension(fallbackExt);

    std::vector<FrmResourceRef> refs;
    collectFrmResourceRefs(frm.form.formControl, refs);

    // 1) 属性行里的引用名是权威 —— 且能容忍 .frm 被改名的场景
    //    (Form1.frm 改名成 Form1_copy.frm 后, 属性里写的仍是 "Form1.frx")
    if (!refs.empty() && !refs.front().fileName.empty()) {
        // FGZ-ICE-001: refs.front().fileName 是 UTF-8 窄串, 直接喂给
        // path::operator/(std::string) 会被 MSVC STL 按 ACP(936) 解释;
        // UTF-8 多字节序列在 ACP 下非法 → _Convert_narrow_to_wide 抛
        // filesystem_error → 无人接 → ICE。统一转宽路径后再拼接。
        auto named = candidate.parent_path() / utf8ToPath(refs.front().fileName);
        std::error_code ec;
        if (std::filesystem::exists(named, ec)) return named;
    }
    // 2) 退回"同基名"
    std::error_code ec2;
    if (std::filesystem::exists(candidate, ec2)) return candidate;
    return {};
}

} // namespace vb6c3
