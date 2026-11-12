#include "backend/cgen.hpp"
#include "backend/cgen_pch.hpp"
#include "common/event_param_abi.hpp"
#include "common/invisible_at_runtime.hpp"   // 账 #305 258: InvisibleAtRuntime 的唯一判据     // 账 #305 256: 事件形参按不按地址的唯一主人（两处 P6.5 包装站）
#include "driver/detail/phase_timing.hpp"   // ai/030 续: C3_TIMING=1 的相位计时
#include <algorithm>
#include <cstdio>
#include <cctype>
#include <iostream>
#include <functional>
#include <unordered_set>

namespace vb6c3 {

// --- cgen_base.cpp: 辅助函数 + CodeEmitter + 构造 + generate（骨架，含 11 个函数体片段） ---
// 2026-09-17 两步法拆分：原 1753 行单文件（22 个顶层函数，主导 CCodeGen::generate 占 1067 行）
//   第 1 步（纯搬移）：cgen_base_type.cpp（类型映射/常量折叠）+ cgen_base_naming.cpp（默认值/标识符/常量/过程名）
//   第 2 步（函数体片段）：generate 函数体 91~1152 行切为 11 个 .inc，在 generate() 函数体内被 #include（C++ 允许）——
//     detail/cgen_base_generate_prologue.inc     入口初始化 + Fix 090q 返回 UDT 预扫（原 91~131 行）
//     detail/cgen_base_generate_state_scan.inc   P6.11 类成员 BSTR 扫描 + Fix 010n 模块级 UDT 持久化（原 132~255 行）
//     detail/cgen_base_generate_header_open.inc  .h 头文件开头：guard、include 顺序(P7)、标准宏取消(Fix 020)（原 256~314 行）
//     detail/cgen_base_generate_crossmod.inc     跨模块注册与定义顺序：010r-11 / 010r-12 / P8.7（原 315~384 行）
//     detail/cgen_base_generate_c_open.inc       .c 源文件头与类模块结构体定义（原 385~497 行）
//     detail/cgen_base_generate_decl_pass.inc    第一遍：声明（Const/变量/Declare/Event/过程前向声明）（原 498~596 行）
//     detail/cgen_base_generate_evt_decl.inc     事件处理器包装声明与 vtable source interface 前向声明（原 597~672 行）
//     detail/cgen_base_generate_body_pass.inc    第二遍：过程体 + 类模块工厂函数 + P6.5 事件处理器包装（原 673~795 行）
//     detail/cgen_base_generate_evt_impl.inc     COM 事件回调 + vtable 接收器实现 + 模块级变量延迟初始化（原 796~954 行）
//     detail/cgen_base_generate_entry.inc        入口点生成（EXE main / ActiveX DLL）（原 955~1101 行）
//     detail/cgen_base_generate_epilogue.inc     保存生成结果 + COM 接口 typedef 前向声明（原 1102~1152 行）
// 11 个 .inc 是「函数体片段」，不是独立编译单元（单独 include 会编译不过），片段内容逐行未改 → 零行为改动。

FrmControlType controlTypeFromName(const std::string& name) {
    static const std::unordered_map<std::string, FrmControlType> map = {
        {"commandbutton", FrmControlType::CommandButton},
        {"textbox", FrmControlType::TextBox},
        {"label", FrmControlType::Label},
        {"checkbox", FrmControlType::CheckBox},
        {"optionbutton", FrmControlType::OptionButton},
        {"listbox", FrmControlType::ListBox},
        {"combobox", FrmControlType::ComboBox},
        {"hscrollbar", FrmControlType::HScrollBar},
        {"vscrollbar", FrmControlType::VScrollBar},
        {"frame", FrmControlType::Frame},
        {"timer", FrmControlType::Timer},
        {"picturebox", FrmControlType::PictureBox},
        {"image", FrmControlType::Image},
    };
    std::string lower = name;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    auto it = map.find(lower);
    return (it != map.end()) ? it->second : FrmControlType::Unknown;
}


// ============================================================
// CodeEmitter
// ============================================================

// Fix 109: 生成 C 里 `UserControl.Font` 的成员访问基址改写.
// VB6 中 UserControl.Font 是对象, 生成代码写作 `vb6_UserControl_Font.<成员>`
// (VB6 点语法). RTL 里该宿主对象是 vb6_ComIface_Font* 指针, 指针 + '.' 触发 C2224.
// 生成端散落在读/写两处, 这里在写出前统一修正: 仅当该标识符后面紧跟 '.' 时
// 改写为 '->', 不触碰其它表达式与注释.
static void fixupAmbientFontMemberAccess(std::string& line) {
    static const char kBase[] = "vb6_UserControl_Font";
    const size_t kLen = sizeof(kBase) - 1;
    size_t pos = 0;
    while ((pos = line.find(kBase, pos)) != std::string::npos) {
        size_t after = pos + kLen;
        // 排除更长标识符 (如 vb6_UserControl_FontObj)
        bool longer = (after < line.size()) &&
                      (std::isalnum(static_cast<unsigned char>(line[after])) || line[after] == '_');
        bool prevOk = (pos == 0) ||
                      !(std::isalnum(static_cast<unsigned char>(line[pos - 1])) || line[pos - 1] == '_');
        if (!longer && prevOk && after < line.size() && line[after] == '.') {
            line.replace(after, 1, "->");
            pos = after + 2;
        } else {
            pos = after;
        }
    }
}

// Fix 133u: UserControl.Parent 链改写 (czUI.ctl).
// VB6 UserControl.Parent 是容器窗体对象; cgen 按 "对象成员" 生成
//   `vb6_UserControl_Parent.hWnd` / `.Move(...)` / `.Icon.Handle` /
//   `(void*)vb6_UserControl_Parent` — RTL 不提供该结构 (窗体运行时对象),
//   而是提供语义等价的函数. 这里在写出前把生成文本改写为函数调用:
//    .Icon.Handle  → vb6_UC_ParentIconHandle()
//    .hWnd         → vb6_UC_ParentHwnd()
//    .Move(        → vb6_UC_ParentMove(
//    其余裸符号   → vb6_UC_ParentObject()  (With 绑定 / 对象引用)
// 顺序: 先长匹配 (.Icon.Handle) 再短 (.hWnd), 最后剩余裸引用.
// Fix <vbeclipse>: Screen.MouseIcon 读写 — Screen 是内置全局哨兵 (vb6_Screen_Object()),
// 成员写侧落 `.MouseIcon` 字段访问必 C2224。改写为 RTL 全局槽 vb6_Screen_MouseIcon
// (IPicture 指针槽; VbEclipse ucPerspective 设鼠标光标用)。
static void fixupScreenMouseIcon(std::string& line) {
    static const char kPat[] = "vb6_Screen_Object().MouseIcon";
    size_t pos = 0;
    while ((pos = line.find(kPat, pos)) != std::string::npos) {
        line.replace(pos, sizeof(kPat) - 1, "vb6_Screen_MouseIcon");
        pos += 19;
    }
}

static void fixupUserControlParentChain(std::string& line) {
    static const char kBase[] = "vb6_UserControl_Parent";
    const size_t kLen = sizeof(kBase) - 1;
    bool hit = false;
    size_t pos = 0;
    while ((pos = line.find(kBase, pos)) != std::string::npos) {
        size_t after = pos + kLen;
        bool longer = (after < line.size()) &&
                      (std::isalnum(static_cast<unsigned char>(line[after])) || line[after] == '_');
        bool prevOk = (pos == 0) ||
                      !(std::isalnum(static_cast<unsigned char>(line[pos - 1])) || line[pos - 1] == '_');
        if (longer || !prevOk) { pos = after; continue; }
        if (line.compare(after, 12, ".Icon.Handle") == 0) {
            line.replace(pos, kLen + 12, "vb6_UC_ParentIconHandle()");
            pos += 27;
        } else if (line.compare(after, 5, ".hWnd") == 0
                   && (after + 5 >= line.size() ||
                       !(line[after + 5] == '.'))) {
            line.replace(pos, kLen + 5, "vb6_UC_ParentHwnd()");
            pos += 20;
        } else if (line.compare(after, 6, ".Move(") == 0) {
            line.replace(pos, kLen + 6, "vb6_UC_ParentMove(");
            pos += 20;
        } else if (line.compare(after, 10, ".BackColor") == 0) {
            // ucTreeView: ucScrollbar.ctl `UserControl.BackColor = UserControl.Parent.BackColor`
            // 生成 `vb6_UserControl_BackColor = vb6_UC_ParentObject().BackColor;`.
            // Parent 的 BackColor 就是容器的 ambient 底色 (RTL 已有这枚全局)。
            line.replace(pos, kLen + 10, "vb6_Ambient_BackColor");
            pos += 23;
        } else {
            line.replace(pos, kLen, "vb6_UC_ParentObject()");
            pos += 22;
        }
        hit = true;
    }
    (void)hit;
}

// ucTreeView: ucScrollbar.ctl 里有
//   Public Property Get Extender() As Object: Set Extender = UserControl.Extender
//   ... RunMode = Extender.Parent.RunMode
// 生成 C 把 `Extender` 发成 `vb6_ucScrollbar_prop_get_Extender((void*)me)`,
// 再拼 `.Parent.RunMode` —— 函数调用结果后面接 `.Parent` 不是合法 C, 而且
// Extender.Parent 的 RunMode 在运行期恒为 1 (vbFormRuntime)。RTL 的
// vb6_Ambient_UserMode 正是这个值 (-1 = 运行期)。这里把整段改写过去。
static void fixupExtenderParentRunMode(std::string& line) {
    static const char kFn[] = "_prop_get_Extender(";
    static const char kTail[] = ").Parent.RunMode";
    size_t pos = 0;
    while ((pos = line.find(kFn, pos)) != std::string::npos) {
        size_t fnStart = pos;
        while (fnStart > 0 &&
               (std::isalnum(static_cast<unsigned char>(line[fnStart - 1])) ||
                line[fnStart - 1] == '_')) {
            fnStart--;
        }
        size_t tail = line.find(kTail, pos);
        if (tail == std::string::npos) { pos += sizeof(kFn) - 1; continue; }
        size_t end = tail + sizeof(kTail) - 1;
        line.replace(fnStart, end - fnStart, "vb6_Ambient_UserMode");
        pos = fnStart + 22;
    }
}

void CodeEmitter::emitLine(const std::string& line) {
    C3ScopeTimer _c3t(0);   // ai/030 续: emitLine 累计统计 (slot 0)
    flushPending();
    for (int i = 0; i < indentLevel_; i++) {
        oss_ << "    ";  // 4空格缩进
    }
    if (line.find("vb6_UserControl_Font") != std::string::npos) {
        // Fix 109: 仅在该标识符出现时做一次修正 (见上).
        std::string fixed = line;
        fixupAmbientFontMemberAccess(fixed);
        oss_ << fixed << "\n";
    } else if (line.find("vb6_Screen_Object().") != std::string::npos) {
        std::string fixed = line;
        fixupScreenMouseIcon(fixed);
        oss_ << fixed << "\n";
    } else if (line.find("vb6_UserControl_Parent") != std::string::npos) {
        // Fix 133u: UserControl.Parent 链改写 (见上).
        std::string fixed = line;
        fixupUserControlParentChain(fixed);
        oss_ << fixed << "\n";
    } else if (line.find("_prop_get_Extender(") != std::string::npos &&
               line.find(").Parent.RunMode") != std::string::npos) {
        std::string fixed = line;
        fixupExtenderParentRunMode(fixed);
        oss_ << fixed << "\n";
    } else {
        oss_ << line << "\n";
    }
    // Fix 133: 落地"本语句之后须回写"的行 (Variant 实参传 Declare 标量 ByRef
    // 出参: 先拷临时、调用后写回 Variant). 见 addPostLine 注释.
    for (auto& pl : postLines_) {
        for (int i = 0; i < indentLevel_; i++) {
            oss_ << "    ";
        }
        oss_ << pl << "\n";
    }
    postLines_.clear();
}

void CodeEmitter::emit(const std::string& text) {
    oss_ << text;
}

void CodeEmitter::emitBlank() {
    flushPending();
    oss_ << "\n";
}

// Fix 090q: 落地待输出声明行(带当前缩进). 供 As Any ByRef 实参的 UDT 临时
// 变量声明在表达式行之前落位 — 保证 "声明先于引用" (C 要求).
void CodeEmitter::flushPending() {
    for (auto& pl : pendingLines_) {
        for (int i = 0; i < indentLevel_; i++) {
            oss_ << "    ";  // 4空格缩进
        }
        oss_ << pl << "\n";
    }
    pendingLines_.clear();
}

// ============================================================
// CCodeGen 构造
// ============================================================

CCodeGen::CCodeGen(Diagnostics& diag, const SymbolTable& symTab,
                   const TypeSystem& typeSys,
                   const std::unordered_map<std::string, std::set<std::string>>* classVoidFieldMap,
                   const std::unordered_map<std::string, std::unordered_map<std::string, std::string>>* classTypedFieldMap,
                   const std::unordered_set<std::string>* variantReturnFuncs,
                   bool verbose)
    : diag_(diag), symTab_(symTab), typeSys_(typeSys),
      classVoidFieldMap_(classVoidFieldMap), classTypedFieldMap_(classTypedFieldMap),
      variantReturnFuncs_(variantReturnFuncs), verbose_(verbose) {}

// ============================================================
// 主入口: 生成 .h + .c
// ============================================================

bool CCodeGen::generate(Module& module, const std::string& baseName,
                         const std::vector<std::string>& externalModules,
                         bool isDll, const std::string& dllProgId,
                         const FrmFormDesc* frmDesc) {
#include "backend/detail/base/cgen_base_generate_prologue.inc"
    c3PhaseMark("g:prologue");
#include "backend/detail/base/cgen_base_generate_state_scan.inc"
    c3PhaseMark("g:state_scan");
#include "backend/detail/base/cgen_base_generate_header_open.inc"
    c3PhaseMark("g:header_open");
#include "backend/detail/base/cgen_base_generate_crossmod.inc"
    c3PhaseMark("g:crossmod");
#include "backend/detail/base/cgen_base_generate_c_open.inc"
    c3PhaseMark("g:c_open");
#include "backend/detail/base/cgen_base_generate_decl_pass.inc"
    c3PhaseMark("g:decl_pass");
#include "backend/detail/base/cgen_base_generate_evt_decl.inc"
    c3PhaseMark("g:evt_decl");
#include "backend/detail/base/cgen_base_generate_body_pass.inc"
    c3PhaseMark("g:body_pass");
#include "backend/detail/base/cgen_base_generate_evt_impl.inc"
    c3PhaseMark("g:evt_impl");
#include "backend/detail/base/cgen_base_generate_entry.inc"
    c3PhaseMark("g:entry");
#include "backend/detail/base/cgen_base_generate_epilogue.inc"
}

// ============================================================
// 表达式求值 (累加器模式: lastExpr_接收结果)
// ============================================================


// ============================================================
// 账 §56.52: 跨模块 include 瘦身
// ============================================================
//
// 为什么要有这一刀 (账 §56.51 的读数): 每个模块的 .h 都把**全工程**每个模块的头包进来
// (PD 539 个模块 ⇒ 每个 .h 约 470 行 include, 539 份头共 10.4 MB) ⇒ 623 个 TU 合计
// 重解析 6.5 GB 头文本 + 29 万次头文件打开。本机单 TU 实测: 带 include 0.94 s vs
// 去掉 0.06 s ⇒ 94% 的 cl 时间是重复劳动。
//
// 裁的判据是「本模块真正引用到哪些外部模块」, 三路取并集 (只往「多留一个头」偏,
// 绝不往「少留」偏 —— 少留一个头就是 C2065 编译断):
//   1) usedExternalModules_ —— 发码时登记 (过程/类/枚举三处收口), 构造上完备;
//   2) 文本兜底 —— 本模块文本里出现的**带模块名**的标识符 (vb6_<Mod>_* /
//      vb6_cls_<Mod>_* / vb6_form_ctrl_<Mod>_* / vb6_form_hwnd_<Mod>);
//   3) 符号表兜底 —— 本模块文本里出现过的标识符 ∩ 外部符号的名字。三档口径不同:
//      * 公共变量/常量 ⇒ 整个 token 匹配 (发的是**裸名**, 靠定义方 .h 的 #define / extern);
//      * 类 / 接口 / 枚举 ⇒ **组件**匹配 (名字**不带模块名**, 且嵌在长标识符里:
//        vb6_cls_<C>_<member> / vb6_iv_thin_of_<C> / vb6_ivref_<I> / vb6_iface_<I> /
//        vb6_iv_test_iid_<C> / vb6_enum_<T>_<Member> …);
//      * 模块 ⇒ 只认**带模块名**的形态 (组件匹配在这儿不行: 每个模块的 .c 都有
//        `vb6_mod_<X>_init()` ⇒ 模块名必然作为组件出现 ⇒ 那样判等于「全留」)。
//
// 裁掉一行 include 时必须补一行 init 前向声明: 每个模块的 .c 都会调**每个**别的模块的
// `vb6_mod_<X>_init();` ⇒ include 掉了, 声明必须还在。
//
// ⚠ 开着时发码形状会变 ⇒ 形状门要重新登记。
//   **固定行为, 没有用户可见的开关** (用户 2026-11-12 裁定: 「副作用选项没啥好处」)。
//   瘦身不影响调试信息: 裁掉的是**整行替换** (行号不变), 被引用符号的原型/类型
//   全部照旧可见, 只有「没用到」的模块的头不再进这个 TU。
void CCodeGen::trimCrossModuleIncludes() {
    if (externalModules_.empty()) return;

    const std::string all = header_ + "\n" + source_;

    // ---- 文本工具: 标识符 token 集 + 前缀集 (小写) ----
    //   token  = 一个完整标识符 (字母/数字/下划线)
    //   prefix = token 里每个 '_' 之前的截断 (如 "vb6_cls_cFoo_Bar" → "vb6_cls_cFoo")
    //            ⇒ 用来认「以 X 开头的更长标识符出现过」
    // ⚠ 必须跳过 #include 行: 那些行里全是**别的模块名**, 若入 token 集, 裸名判据
    //   会把每个模块都误判成「被引用」⇒ 瘦身等于没做。
    std::unordered_set<std::string> tokens;
    std::unordered_set<std::string> prefixes;
    std::unordered_set<std::string> components;
    {
        size_t pos = 0;
        while (pos < all.size()) {
            size_t eol = all.find('\n', pos);
            if (eol == std::string::npos) eol = all.size();
            const std::string line = all.substr(pos, eol - pos);
            const size_t lead = line.find_first_not_of(" \t");
            const bool isInclude =
                (lead != std::string::npos) && line.compare(lead, 9, "#include ") == 0;
            if (!isInclude) {
                size_t i = 0;
                while (i < line.size()) {
                    const unsigned char ch = static_cast<unsigned char>(line[i]);
                    if (std::isalpha(ch) || ch == '_') {
                        size_t j = i + 1;
                        while (j < line.size()) {
                            const unsigned char c2 = static_cast<unsigned char>(line[j]);
                            if (std::isalnum(c2) || c2 == '_') ++j; else break;
                        }
                        std::string tok = line.substr(i, j - i);
                        std::transform(tok.begin(), tok.end(), tok.begin(),
                            [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                        tokens.insert(tok);
                        for (size_t k = tok.find('_'); k != std::string::npos;
                             k = tok.find('_', k + 1)) {
                            prefixes.insert(tok.substr(0, k));
                            components.insert(tok.substr(0, k));
                        }
                        components.insert(tok);
                        i = j;
                    } else {
                        ++i;
                    }
                }
            }
            pos = (eol >= all.size()) ? all.size() : eol + 1;
        }
    }
    auto mentions = [&](const std::string& s) {
        return tokens.count(s) > 0 || prefixes.count(s) > 0;
    };

    // ---- 1) 需要的模块集合 (小写) ----
    std::unordered_set<std::string> needed = usedExternalModules_;

    // ---- 2) 文本兜底: 带模块名的形态 ----
    //   标准模块过程 vb6_<Mod>_<proc> / 类成员 vb6_cls_<Class>_<member> /
    //   跨窗体控件 vb6_form_ctrl_<Mod>_<ctrl> / 窗体句柄 vb6_form_hwnd_<Mod> ...
    static const char* kModPrefixes[] = {
        "vb6_", "vb6_cls_", "vb6_form_ctrl_", "vb6_form_hwnd_"
    };
    for (const auto& extMod : externalModules_) {
        const std::string cx = Symbol::toLower(cIdent(extMod));
        for (const char* pfx : kModPrefixes) {
            if (mentions(std::string(pfx) + cx)) {
                needed.insert(Symbol::toLower(extMod));
                break;
            }
        }
    }

    // ---- 3) 符号表兜底: 裸名引用 ----
    //   跨模块 Public 常量/变量发的是**裸名** (定义方 .h 里才有那个 #define / extern
    //   标识符) ⇒ 少了那个头就是 C2065; 而枚举 (vb6_enum_<Type>_<Member>) 与类
    //   (vb6_cls_<Class>) 的名字**不带模块名**, 上面那路按模块名认不到 ⇒ 这里按
    //   符号自己的名字认。
    //   (UDT 不在此列: 账 §G1 让每个 .h 自带需要的 UDT 定义块, 且定义块自带
    //    #ifndef VB6_TYPE_<X>_DEFINED ⇒ 重复包含也不会 C2011。)
    if (auto* ms = symTab_.moduleScope()) {
        for (const auto& kv : ms->symbols()) {
            const Symbol* sym = kv.second.get();
            if (!sym || !sym->isExternal || sym->sourceModule.empty()) continue;
            if (sym->kind != SymbolKind::Variable && sym->kind != SymbolKind::Constant &&
                sym->kind != SymbolKind::EnumType && sym->kind != SymbolKind::Class) continue;
            const std::string owner = Symbol::toLower(sym->sourceModule);
            if (needed.count(owner)) continue;
            if (sym->name.empty()) continue;
            const std::string bare = Symbol::toLower(cIdent(sym->name));
            if (bare.empty()) continue;
            bool hit = false;
            if (sym->kind == SymbolKind::Variable || sym->kind == SymbolKind::Constant) {
                // 公共变量/常量发的是**裸名** (定义方 .h 里的 #define / extern 才有那个标识符)
                // ⇒ 按**整个 token** 认 (不做组件匹配: 组件匹配会把同名局部变量也算进来)。
                hit = tokens.count(bare) > 0;
            } else if (sym->kind == SymbolKind::EnumType) {
                // 枚举成员发成 vb6_enum_<Type>_<Member> —— **类型名**嵌在长标识符里。
                hit = components.count(bare) > 0;
            } else {
                // 类/接口: 名字**不带模块名**, 且嵌在一堆形态里 ——
                //   vb6_cls_<C>_<member> / vb6_iv_thin_of_<C> / vb6_ivref_<I> /
                //   vb6_iface_<I> / vb6_iv_test_iid_<C> / vb6_iv_from_iv_<C> ...
                // ⇒ 必须按**组件**认 (token 按 '_' 切开后含 <C> 这一段)。
                // ⚠ 不能拿组件匹配去认**模块**: 每个模块的 .c 都有 `vb6_mod_<X>_init()`
                //    ⇒ 模块名必然作为组件出现 ⇒ 那样判等于「全留」。
                hit = components.count(bare) > 0;
            }
            if (hit) needed.insert(owner);
        }
    }

    // ---- 2) 逐行裁 include, 换成 init 前向声明 ----
    auto trimText = [&](std::string& text, bool isHeader) {
        std::string out;
        out.reserve(text.size());
        size_t pos = 0;
        while (pos < text.size()) {
            const size_t eol = text.find('\n', pos);
            const bool hasEol = (eol != std::string::npos);
            const size_t lineEnd = hasEol ? eol : text.size();
            std::string line = text.substr(pos, lineEnd - pos);
            std::string body = line;
            const bool hadCR = !body.empty() && body.back() == '\r';
            if (hadCR) body.pop_back();
            const size_t lead = body.find_first_not_of(" \t");
            const std::string indent = (lead == std::string::npos) ? std::string() : body.substr(0, lead);
            const std::string rest  = (lead == std::string::npos) ? std::string() : body.substr(lead);
            if (rest.rfind("#include \"", 0) == 0 && rest.size() > 13 &&
                rest.compare(rest.size() - 3, 3, ".h\"") == 0) {
                const std::string name = rest.substr(10, rest.size() - 10 - 3);
                // 自己那一行 (`.c` 里的 #include "<Base>.h") 一律不动。
                if (Symbol::toLower(name) != Symbol::toLower(baseName_)) {
                    auto it = std::find_if(externalModules_.begin(), externalModules_.end(),
                        [&](const std::string& m) {
                            return Symbol::toLower(m) == Symbol::toLower(name);
                        });
                    if (it != externalModules_.end() && needed.count(Symbol::toLower(*it)) == 0) {
                        line = indent + (isHeader ? "extern void " : "void ")
                             + "vb6_mod_" + cIdent(*it) + "_init(void);";
                    }
                }
            }
            out += line;
            if (hadCR) out += '\r';
            if (hasEol) out += '\n';
            pos = hasEol ? eol + 1 : text.size();
        }
        text.swap(out);
    };
    trimText(header_, true);
    trimText(source_, false);
}


} // namespace vb6c3
