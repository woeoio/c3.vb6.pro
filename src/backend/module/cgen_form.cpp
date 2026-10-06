#include "backend/cgen.hpp"
#include "common/float_literal.hpp"  // 账 #188: 浮点字面量的单一出口
#include <cstdio>
#include <cstdlib>
#include "project/frx_reader.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>
#include <sstream>
#include <iomanip>

namespace vb6c3 {

// P24: Convert binary data to C hex array string
static std::string bytesToHexArray(const uint8_t* data, size_t size, const std::string& varName) {
    std::ostringstream ss;
    ss << "static const unsigned char " << varName << "[] = {\n";
    for (size_t i = 0; i < size; i++) {
        if (i % 16 == 0) ss << "    ";
        ss << "0x" << std::setfill('0') << std::setw(2) << std::hex << (int)data[i];
        if (i + 1 < size) ss << ",";
        if (i % 16 == 15 || i + 1 == size) ss << "\n";
        else ss << " ";
    }
    ss << "};\n";
    ss << "static const int " << varName << "_size = " << std::dec << size << ";";
    return ss.str();
}

// P24: Escape string for C string literal (handles backslash, quote, newlines, tabs, etc.)
static std::string escapeCString(const std::string& s) {
    std::string result;
    for (char c : s) {
        switch (c) {
            case '\\': result += "\\\\"; break;
            case '"': result += "\\\""; break;
            case '\n': result += "\\n"; break;
            case '\r': result += "\\r"; break;
            case '\t': result += "\\t"; break;
            default:
                if ((unsigned char)c < 0x20) {
                    char buf[8];
                    snprintf(buf, sizeof(buf), "\\x%02x", (unsigned char)c);
                    result += buf;
                } else {
                    result += c;
                }
        }
    }
    return result;
}

// --- cgen_form.cpp: 窗体框架（emitFormFramework）---
// 2026-09-17 拆分：原 2222 行（emitFormFramework 单函数 1945 行）先纯搬移走 4 个成员函数
//   （emitMenuItem / emitMenuClickDispatch / escapeWideCString / CCodeGen::escapeCString
//     → backend/module/cgen_form_menu.cpp，已登记 CMakeLists），
//   余下函数体按既有分节注释切为 8 个「函数体片段」，在 emitFormFramework() 函数体内 #include：
//   detail/cgen_form_prelude.inc            —— 初始化（frx 加载、菜单ID收集）与窗体属性提取（原 54~144 行）
//   detail/cgen_form_ctrl_registry.inc      —— 控件名映射 / 控件数组检测 与 .h 声明（原 145~261 行）
//   detail/cgen_form_wndproc_subclass.inc   —— .c 实现开头：Form_Unload trampoline 与需子类化控件的 WndProc（原 262~597 行）
//   detail/cgen_form_wndproc_create.inc     —— WndProc 前半：WM_CREATE / 延迟 Form_Load / WM_COMMAND / WithEvents 派发（原 598~921 行）
//   detail/cgen_form_wndproc_dispatch.inc   —— WndProc 后半：焦点 / 滚动 / 键盘 / 关闭 / 清理 / default（原 922~1272 行）
//   detail/cgen_form_create_controls.inc    —— CreateControls：控件树 CreateWindow（原 1273~1772 行）
//   detail/cgen_form_frame_menu.inc         —— Frame 子控件 / 控件子类化安装 / Win32 菜单构建（原 1773~1937 行）
//   detail/cgen_form_show.inc               —— Show 函数（原 1938~1996 行）
// 八个 .inc 是「函数体片段」，在函数体内被 #include（C++ 允许），故不用 .cpp/.hpp 后缀 ——
// 它们不是独立编译单元，单独编译会报错，也不登记 CMakeLists。片段内局部 lambda 与块作用域原样不动，
// 逐行未改 → 零行为改动；切点全落在原函数体的分节注释处（相对花括号深度 0）。

void CCodeGen::emitFormFramework(const FrmFormDesc& frmDesc, Module& module) {
    // C29-7: **必须在生成任何代码之前登记**哪些控件是 ListView —— 下面各片段会查
    // listViewVars_ 来决定 `ListView1.ListItems` 走"真集合对象"还是退化路径。
    // ⚠ 别把这个登记挪进 emitDesignerControlDecls: 那个函数是在
    // cgen_base_generate_decl_pass.inc 的**第 46 行**调用的, 而本函数在第 40 行 ——
    // 也就是"窗体代码全生成完了才登记", listViewVars_ 永远是空的。
    // (实测症状: 生成出来还是 `vb6_ComGetObjectProp(vb6_hwnd_ListView1, …)`,
    //  对 HWND 当 IDispatch 用 → 运行期读数全空。)
    {
        std::function<void(const FrmControl&)> regLV = [&](const FrmControl& c) {
            if (c.controlType == FrmControlType::ListView) listViewVars_.insert(c.controlName);
            // C29-OLE: OLE 容器同 ListView 一样是真窗口, 方法/属性都按 HWND 槽认。
            if (c.controlType == FrmControlType::OLE) oleConVars_.insert(c.controlName);
            // C29-Data: Data 控件同批登记。
            if (c.controlType == FrmControlType::Data) dataVars_.insert(c.controlName);
            // C29-DT-c: DTPicker 同批登记 —— `dt1.SimChange` 那三条判据方法靠 dtpickerVars_
            // 认出宿主是真窗口槽, 才能改道到 vb6_DTP_Sim*(HWND)。不认就退化成
            // "拿 HWND 当 IDispatch 问它要 SimChange 属性": 编得过、跑起来什么都不发。
            if (c.controlType == FrmControlType::DTPicker) dtpickerVars_.insert(c.controlName);
            // C29-MV-c: MonthView 同批登记（判据方法 SimDateClick 的宿主槽是真窗口）。
            if (c.controlType == FrmControlType::MonthView) monthviewVars_.insert(c.controlName);
            // C29-SL-c: Slider 同批登记（判据方法 SimNotify 的宿主槽是真窗口）。
            if (c.controlType == FrmControlType::Slider) sliderVars_.insert(c.controlName);
            // C29-8b: TreeView 同批登记 —— `tv1.Nodes` 那条链靠 treeViewVars_ 认出宿主,
            // 才能改道到 vb6_TreeView_Nodes( 的真 IDispatch 集合 (不认就发
            // vb6_ComGetObjectProp(vb6_hwnd_tv1, L"Nodes") = 拿 HWND 当 IDispatch 用)。
            if (c.controlType == FrmControlType::TreeView) treeViewVars_.insert(c.controlName);
            for (const auto& ch : c.children) regLV(ch);
        };
        regLV(frmDesc.formControl);
    }
#include "backend/detail/module/cgen_form_prelude.inc"
#include "backend/detail/module/cgen_form_ctrl_registry.inc"
#include "backend/detail/module/cgen_form_wndproc_subclass.inc"
#include "backend/detail/module/cgen_form_wndproc_create.inc"
#include "backend/detail/module/cgen_form_wndproc_dispatch.inc"
#include "backend/detail/module/cgen_form_create_controls.inc"
#include "backend/detail/module/cgen_form_ctrl_style_apply.inc"
#include "backend/detail/module/cgen_form_frame_menu.inc"
#include "backend/detail/module/cgen_form_show.inc"
}

// ============================================================
// P6.6: 单独生成 DLL 入口文件 (dll_entry.c)
// 当DLL工程只有类模块(无标准模块)时使用
// ============================================================


// ============================================================
// Fix 110: 设计期子控件句柄变量声明 (窗体 / UserControl / PropertyPage 共用)
// ============================================================
// 规则 (与 .frm 旧行为兼容) 并补上旧行为漏掉的类别:
//   - ActiveX 控件 (ImageList/Toolbar/StatusBar/CommonDialog) → vb6_com_<name> (IDispatch*)
//   - 同名控件在设计器出现多次, 或带 Index= 的控件数组 → vb6_arr_<name> (vb6_CtrlArr)
//   - 其余 (如 Timer / WebBrowser / 工程内 UserControl 实例 / Unknown) → vb6_hwnd_<name>
//
// 旧实现有三个缺口, 均导致生成代码里出现未声明的 vb6_<ctrl>_<prop> (C2065):
//   1) 只遍历顶层 children, 不递归 → Frame 内的 Picture2/TxtARGB 等没有符号;
//   2) controlTypeToWin32Class()==nullptr 且不属于 Timer/WebBrowser/ActiveX →
//      直接 continue → 工程内 UserControl 实例 (ucPieChart1 等) 没有符号;
//   3) 早期版本 Timer 也被跳过.
void CCodeGen::emitControlHandleDecls(const FrmFormDesc& frmDesc) {
    std::unordered_set<std::string> emitted;
    std::function<void(const FrmControl&)> emitRec = [&](const FrmControl& ctrl) {
        std::string ctrlLower = ctrl.controlName;
        std::transform(ctrlLower.begin(), ctrlLower.end(), ctrlLower.begin(), ::tolower);
        if (!ctrlLower.empty() && emitted.insert(ctrlLower).second) {
            // 合并(P20-45 + C29-9): **无窗口控件只有 ImageList / Toolbar** —— 它们的槽是
            // 实例指针 vb6_com_X。其余控件一律 vb6_hwnd_X, 含 StatusBar (P20-40 起是
            // msctls_status32 原生复刻) 与 CommonDialog (C29-9 起是自注册不可见类
            // VB6_COMMONDIALOG 属性宿主); 把这两个归回 vb6_com_ 家族会 C2065
            // (声明成 vb6_com_X, 用出来却是 vb6_hwnd_X)。
            // C29-5a: Toolbar 摘出 —— 它有真窗口，槽一律回到 vb6_hwnd_X 那一族。
            if (ctrl.controlType == FrmControlType::ImageList) {
                c_.emitLine("static void* vb6_com_" + cIdent(ctrl.controlName) + " = NULL;  /* IDispatch* */");
            } else if (knownControlArrays_.count(ctrlLower)) {
                c_.emitLine("static vb6_CtrlArr vb6_arr_" + cIdent(ctrl.controlName) + ";");
            } else {
                c_.emitLine("static void* vb6_hwnd_" + cIdent(ctrl.controlName) + " = NULL;");
            }
        }
        for (const auto& child : ctrl.children) emitRec(child);
    };
    for (const auto& ctrl : frmDesc.formControl.children) emitRec(ctrl);
}

// ============================================================
// Fix 110: .ctl / .pag 设计期子控件符号支撑
// ============================================================
// VB6 的 UserControl/PropertyPage 可以在设计器里放子控件 (Begin VB.Timer Timer1,
// Begin VB.PictureBox Picture2, ...). 模块代码里 `Timer1.Interval = 100` /
// `Picture1.ScaleWidth` 在 VB6 语义上是"访问设计期控件的属性".
//
// 若这些子控件没有注册到 knownFormControls_, cgen 的两处回退会把它当作
// "模块级成员" 处理:
//   - 读 →  vb6_<Ctrl>_<Prop>
//   - 写 →  vb6_<Ctrl>_<Prop> = ...
// 这些标识符在生成代码与 RTL 中都不存在 → C2065.
//
// 这里登记控件并发射句柄变量声明. 不发射窗体窗口框架: UserControl /
// PropertyPage 对外是类, 其可见内容由代码绘制到 UserControl.hDC / hwnd,
// 子控件句柄保持 NULL (RTL 的属性 setter 对 NULL 句柄是安全空操作).
void CCodeGen::emitDesignerControlDecls(const FrmFormDesc& frmDesc, DocumentKind kind) {
    // Fix 110f: 记录设计器种类 (.pag 为 PropertyPage, .ctl 为 UserControl),
    // 二者在宿主内建成员前缀上不同: vb6_PropertyPage_* / vb6_UserControl_*.
    // 种类读 Module::docKind (driver 按扩展名一处写入, 账 #217 第二刀) —— 以前这里从
    // controlTypeName 里找 "PropertyPage" 猜一遍, 与语义层的判据是两份.
    isPropertyPageDesigner_ = (kind == DocumentKind::PropertyPage);
    // 1) 登记控件名映射 (与 emitFormFramework 的 P7.5/P7.6 块保持一致)
    std::string ownerLower = frmDesc.formName;
    std::transform(ownerLower.begin(), ownerLower.end(), ownerLower.begin(), ::tolower);
    if (!ownerLower.empty()) {
        knownFormControlOriginalNames_[ownerLower] = frmDesc.formName;
    }
    std::unordered_set<std::string> ctrlArraySeen;
    std::function<void(const FrmControl&)> registerCtrlRec = [&](const FrmControl& ctrl) {
        std::string ctrlLower = ctrl.controlName;
        std::transform(ctrlLower.begin(), ctrlLower.end(), ctrlLower.begin(), ::tolower);
        if (!ctrlLower.empty()) {
            const bool duplicateName = !ctrlArraySeen.insert(ctrlLower).second;
            if (ctrl.index >= 0 || duplicateName) {
                knownControlArrays_[ctrlLower] = true;
            }
            knownFormControls_[ctrlLower] = ctrl.controlType;
            knownFormControlOriginalNames_[ctrlLower] = ctrl.controlName;
        }
        for (const auto& child : ctrl.children) registerCtrlRec(child);
    };
    for (const auto& ctrl : frmDesc.formControl.children) registerCtrlRec(ctrl);

    // 2) 发射句柄/数组/COM 变量声明
    emitControlHandleDecls(frmDesc);
    // czUI fix: 设计器子控件句柄改为按实例槽位 — 全局句柄被最后创建的实例覆盖,
    // 导致 11 个实例只有最后一个的 timer/textbox 生效 (开关动画死、文本框错乱)。
    // Fix VbEclipse: 必须传 `me` — UC 实例方法多由外部模块直接 C 调用发起,
    // 此时全局 g_uc_current 已 pop 成 NULL, 只按上下文的旧签名会退化成共享
    // orphan 槽 (恒 NULL), 面板 SetParent/Move 全部落空。
    c_.emitLine("extern void** vb6_UC_DesignSlotOf(void* inst, const char* name);");
    for (const auto& child : frmDesc.formControl.children) {
        std::string n114 = cIdent(child.controlName);
        c_.emitLine("#undef vb6_hwnd_" + n114);
        c_.emitLine("#define vb6_hwnd_" + n114
                  + " (*vb6_UC_DesignSlotOf(me, \"" + child.controlName + "\"))");
    }
    // Fix <vbeclipse> rev20: ScaleWidth/ScaleHeight 同样必须**按实例**取, 否则
    // 多实例共享进程级全局 (vb6_uc_push/pop 维护) ⇒ 子控件 Move 拿到别的实例的
    // 尺寸。ucFolder.ctl 的 `ViewArea.Move 20,20,ScaleWidth-30,ScaleHeight-30`
    // 是这类布局的主力, 读错就整块面板停在设计期尺寸。
    //
    // 用 #undef/#define 而不是改赋值侧: `ScaleWidth = x` 这种写法在 VB6 里
    // 只出现在设计期属性块 (由 driver 消费, 不进 cgen 的赋值路径), 代码里的
    // ScaleWidth/ScaleHeight 一律是**读** —— 故重定义为按 me 取值的函数调用是安全的。
    // (唯一会写它的是 RTL 的 vb6_UserControl_Size / vb6_uc_push, 都在 RTL 侧,
    //  不经过本宏。)
    //
    // `me` 的可得性: 本函数只对 .ctl/.pag 的**实例方法**发射, cgen 为每个实例
    // 方法都加了 `me` 形参 (类实例指针) —— 与上面 vb6_hwnd_* 宏依赖 `me` 是
    // 同一个前提。实测 VbEclipse 全部生成 .c 里 ScaleWidth/Height 的 71 处引用
    // 100% 落在带 `me` 的函数体内 (无一处落在模块级初始化表达式)。
    if (!isPropertyPageDesigner_) {
        c_.emitLine("extern int32_t vb6_UC_ScaleWidthOf(void* inst);");
        c_.emitLine("extern int32_t vb6_UC_ScaleHeightOf(void* inst);");
        c_.emitLine("#undef vb6_UserControl_ScaleWidth");
        c_.emitLine("#define vb6_UserControl_ScaleWidth vb6_UC_ScaleWidthOf((void*)me)");
        c_.emitLine("#undef vb6_UserControl_ScaleHeight");
        c_.emitLine("#define vb6_UserControl_ScaleHeight vb6_UC_ScaleHeightOf((void*)me)");
        // 账 #177/#178: UserControl.TextWidth/.TextHeight 同一个坑、同一味药 —— 量出来是
        // 设备像素, 折算单位必须按**这一枚控件**声明的 ScaleMode 取。读进程级
        // vb6_UserControl_ScaleMode 不够: 容器直调控件公共成员时没人换入宿主上下文
        // (实测 ve_units: 控件自己 Initialize 里 600 缇, 容器调同一个函数拿到 40 像素)。
        // 只在 .ctl 里重定向, 所以取的是带 `me` 的实例方法体 —— 与上面两条同一前提。
        c_.emitLine("extern int32_t vb6_UC_TextWidthOf(void* inst, BSTR text);");
        c_.emitLine("extern int32_t vb6_UC_TextHeightOf(void* inst, BSTR text);");
        c_.emitLine("#undef vb6_UserControl_TextWidth");
        c_.emitLine("#define vb6_UserControl_TextWidth(t) vb6_UC_TextWidthOf((void*)me, (t))");
        c_.emitLine("#undef vb6_UserControl_TextHeight");
        c_.emitLine("#define vb6_UserControl_TextHeight(t) vb6_UC_TextHeightOf((void*)me, (t))");
    }
    c_.emitBlank();

    // 3) Fix 112: UserControl (.ctl) 宿主描述 — 让窗体可以把本控件实例挂到子窗口.
    //    PropertyPage 是设计期窗体, 不是可实例化控件, 跳过.
    if (!isPropertyPageDesigner_) {
        std::string ctl = cIdent(moduleName_);
        // 生命周期处理器是否存在 (Private Sub UserControl_Initialize/Paint/...)
        auto hasProc = [&](const char* n) -> bool {
            Symbol* sym = symTab_.lookupModule(n);
            return sym && (sym->kind == SymbolKind::Sub || sym->kind == SymbolKind::Function);
        };
        const bool hasInit = hasProc("UserControl_Initialize");
        const bool hasPaint = hasProc("UserControl_Paint");
        const bool hasShow = hasProc("UserControl_Show");
        const bool hasResize = hasProc("UserControl_Resize");
        // Fix 112d: InitProperties 负责成员数组的 ReDim/默认值 (如 ucProgressCircular
        // 的 m_PF_Colors). 运行期没有 PropertyBag, 按 VB6 语义 InitProperties 是
        // "无持久化数据时的属性初始化", 归入 init 一起调用, 否则 Draw 读未分配数组 → AV.
        const bool hasInitProps = hasProc("UserControl_InitProperties");

        c_.emitLine("// === Fix 112: UserControl 宿主描述 (供窗体宿主子窗口驱动) ===");
        if (hasInit)   c_.emitLine("static void vb6_" + ctl + "_UserControl_Initialize(vb6_cls_" + ctl + "* me);");
        if (hasPaint)  c_.emitLine("static void vb6_" + ctl + "_UserControl_Paint(vb6_cls_" + ctl + "* me);");
        if (hasShow)   c_.emitLine("static void vb6_" + ctl + "_UserControl_Show(vb6_cls_" + ctl + "* me);");
        if (hasResize) c_.emitLine("static void vb6_" + ctl + "_UserControl_Resize(vb6_cls_" + ctl + "* me);");
        if (hasInitProps) c_.emitLine("static void vb6_" + ctl + "_UserControl_InitProperties(vb6_cls_" + ctl + "* me);");
        // czUI fix: 设计器 Timer 子控件 — 前向声明其 Timer 事件处理器并生成
        // thunk, 供 RTL 设计器定时器逐实例回调 (tmrTrack 驱动 toggle 动画等)。
        for (const auto& child : frmDesc.formControl.children) {
            if (child.controlType == FrmControlType::Timer) {
                std::string cn = cIdent(child.controlName);
                c_.emitLine("static void vb6_" + ctl + "_" + cn + "_Timer(vb6_cls_" + ctl + "* me);");
                c_.emitLine("static void vb6_" + ctl + "_ucTimerThunk_" + cn + "(void* ctx) {");
                c_.emitLine("    vb6_" + ctl + "_" + cn + "_Timer((vb6_cls_" + ctl + "*)ctx);");
                c_.emitLine("}");
            }
        }
        // Fix <vbeclipse> rev22: 设计期子控件的 `<Ctrl>_Resize` 事件转发。
        // VB6 里 `Private Sub ViewArea_Resize()` 是 PictureBox 的 Resize 事件,
        // 运行时在该控件尺寸变化时自动跑。RTL 的 desc 以前没有这个槽, 于是 .ctl
        // 写在子控件事件里的布局代码永远不跑 (ucFolder.ctl:529 ViewArea_Resize
        // 是把视图窗体摆进 ViewArea 的唯一驱动)。
        // 收集判据: 模块 scope 里有 Public/Private 的 `<子控件名>_Resize` Sub,
        // 且该子控件名确实是本 .ctl 设计面上的控件 (否则跨模块重名会误收)。
        std::vector<std::string> designResizeCtrls;
        for (const auto& child : frmDesc.formControl.children) {
            if (child.controlName.empty()) continue;
            std::string procName = child.controlName + "_Resize";
            if (hasProc(procName.c_str()))
                designResizeCtrls.push_back(cIdent(child.controlName));
        }
        if (!designResizeCtrls.empty()) {
            for (const std::string& cn : designResizeCtrls)
                c_.emitLine("static void vb6_" + ctl + "_" + cn + "_Resize(vb6_cls_" + ctl + "* me);");
            c_.emitLine("static void vb6_" + ctl + "_ucHostDesignResize(void* me, const char* ctrlName) {");
            c_.emitLine("    vb6_cls_" + ctl + "* m = (vb6_cls_" + ctl + "*)me;");
            c_.emitLine("    if (!m || !ctrlName) return;");
            // ⚠ 纯 ASCII 字符串字面量: 这段代码跑在 C locale 的 stderr 打印路径上,
            //   非 ASCII 会让 fprintf 截断, 把"最后一行"伪装成崩溃点 (见 MEMORY)。
            for (const std::string& cn : designResizeCtrls) {
                c_.emitLine("    if (strcmp(ctrlName, \"" + cn + "\") == 0) { "
                            + "vb6_" + ctl + "_" + cn + "_Resize(m); return; }");
            }
            c_.emitLine("}");
        }
        c_.emitLine("static void vb6_" + ctl + "_ucHostInit(void* me) {");
        // czUI fix: 设计器子控件 (.ctl 设计面上的 TextBox 等) 属于每个实例 —
        // 逐实例创建真实子窗口 (此前 vb6_hwnd_txtEmbed 恒为 NULL, TextBox 内容
        // 与占位文本全部丢失)。必须在 Initialize/InitProperties **之前**创建:
        // VB6 语义是设计器控件先于一切生命周期代码存在; 否则后创建实例的
        // InitProperties→ConfigureForType(Case Else 隐藏 txtEmbed) 会通过全局
        // 句柄把上一个实例的 edit 隐藏掉 (czTextBox1 空白的根因)。
        // Timer 等暂不实例化 (相关 API 对 NULL 安全)。
        const bool noKids = getenv("C3_NO_DESIGNKIDS") != nullptr;
        for (const auto& child : frmDesc.formControl.children) {
            auto iprop = [&](const char* k, int defv) -> int {
                auto itc = child.properties.find(k);
                return itc != child.properties.end() ? (int)itc->second.intValue : defv;
            };
            if (!noKids && child.controlType == FrmControlType::TextBox) {
                c_.emitLine("    vb6_hwnd_" + cIdent(child.controlName) + " = vb6_UC_CreateDesignEdit("
                    + std::to_string(iprop("Left", 0)) + ", " + std::to_string(iprop("Top", 0)) + ", "
                    + std::to_string(iprop("Width", 2000)) + ", " + std::to_string(iprop("Height", 400)) + ");");
            } else if (!noKids && child.controlType == FrmControlType::Timer) {
                c_.emitLine("    vb6_hwnd_" + cIdent(child.controlName) + " = vb6_UC_CreateDesignTimer("
                    + "vb6_" + ctl + "_ucTimerThunk_" + cIdent(child.controlName) + ", me);");
                // Fix <vbeclipse>: 设计器 Timer 的 Enabled/Interval 是权威初值, 必须像
                // 上面的 TextBox 一样在 ucHostInit 落盘。vb6_GetTimerEnabled 对"从没设过"
                // 返回 True (VB6 Timer 的真实默认是 False), 于是 `Enabled = 0 'False` 的
                // tmrDrag 被当成开着的 → 每 tick 里 `Interval = 1` 自激 → Controls.Item
                // 洪泛 (~95k 次) 把消息循环饿死, 停靠面板全不刷新。
                c_.emitLine("    vb6_SetTimerInterval(vb6_hwnd_" + cIdent(child.controlName) + ", "
                    + std::to_string(iprop("Interval", 60000)) + ");");
                c_.emitLine("    vb6_SetTimerEnabled(vb6_hwnd_" + cIdent(child.controlName) + ", "
                    + std::to_string(iprop("Enabled", 0)) + ");");
            } else if (!noKids && child.controlType == FrmControlType::PictureBox) {
                // Fix <vbeclipse>: ucFolder.ViewArea 等 PictureBox 设计子控件 → STATIC 子窗,
                // 视图窗体靠 SetParent 挂进它。旧循环不建 → ViewArea 恒 NULL → 面板全空。
                c_.emitLine("    vb6_hwnd_" + cIdent(child.controlName) + " = vb6_UC_CreateDesignPicture("
                    + std::to_string(iprop("Left", 0)) + ", " + std::to_string(iprop("Top", 0)) + ", "
                    + std::to_string(iprop("Width", 2000)) + ", " + std::to_string(iprop("Height", 1500)) + ");");
            } else if (!noKids && child.controlType == FrmControlType::Label) {
                // Fix <vbeclipse> rev32: 设计子控件里的 **VB.Label** (ucTab.lblCaption)。
                // 此前这个种类不在发射循环里 ⇒ ucTab 的设计面一个子控件都没建
                // (对照: ucCaption/ucTabStrip/ucFolder 各有 CreateDesign*, ucTab = 0)
                // ⇒ `vb6_hwnd_lblCaption` 恒 NULL, 而 `ucTab.ToolTip` setter 是
                // `lblCaption.ToolTipText = NewToolTip` ⇒ 对 NULL 写属性 0xC0000005。
                c_.emitLine("    vb6_hwnd_" + cIdent(child.controlName) + " = vb6_UC_CreateDesignLabel("
                    + std::to_string(iprop("Left", 0)) + ", " + std::to_string(iprop("Top", 0)) + ", "
                    + std::to_string(iprop("Width", 2000)) + ", " + std::to_string(iprop("Height", 400)) + ");");
            } else if (!noKids && child.controlType == FrmControlType::Image) {
                // 同上: **VB.Image** (ucTab.imgIcon)。`ucTab.Icon` setter 写它。
                c_.emitLine("    vb6_hwnd_" + cIdent(child.controlName) + " = vb6_UC_CreateDesignImage("
                    + std::to_string(iprop("Left", 0)) + ", " + std::to_string(iprop("Top", 0)) + ", "
                    + std::to_string(iprop("Width", 2000)) + ", " + std::to_string(iprop("Height", 1500)) + ");");
            } else if (!noKids && child.controlType == FrmControlType::Unknown
                       && child.controlTypeName.find('.') != std::string::npos) {
                // Fix <vbeclipse>: 设计子控件里"工程内 UserControl" (parseControlType 认不出的
                // 限定名, 如 VbEclipse.ucTabStrip / VbEclipse.ucCaption) → 复用 HostCreate 逐实例
                // 建宿主子窗。非登记 UC (第三方 OCX) 时 helper 内 findDesc 落空返回 NULL, 与不建
                // 等价, 不影响别的夹具 (它们没有这类设计子控件)。
                std::string tn = child.controlTypeName;
                for (char& ch : tn) { if (ch == '"' || ch == '\\') ch = ' '; }
                c_.emitLine("    vb6_hwnd_" + cIdent(child.controlName) + " = vb6_UC_CreateDesignUserControl("
                    + "\"" + tn + "\", \"" + child.controlName + "\", "
                    + std::to_string(iprop("Left", 0)) + ", " + std::to_string(iprop("Top", 0)) + ", "
                    + std::to_string(iprop("Width", 2000)) + ", " + std::to_string(iprop("Height", 1500)) + ");");
            }
        }
        if (hasInit) c_.emitLine("    vb6_" + ctl + "_UserControl_Initialize((vb6_cls_" + ctl + "*)me);");
        if (hasInitProps) c_.emitLine("    vb6_" + ctl + "_UserControl_InitProperties((vb6_cls_" + ctl + "*)me);");
        c_.emitLine("}");
        c_.emitLine("static void vb6_" + ctl + "_ucHostPaint(void* me) {");
        if (hasPaint) c_.emitLine("    vb6_" + ctl + "_UserControl_Paint((vb6_cls_" + ctl + "*)me);");
        c_.emitLine("}");
        c_.emitLine("static void vb6_" + ctl + "_ucHostShow(void* me) {");
        if (hasShow) c_.emitLine("    vb6_" + ctl + "_UserControl_Show((vb6_cls_" + ctl + "*)me);");
        c_.emitLine("}");
        c_.emitLine("static void vb6_" + ctl + "_ucHostResize(void* me) {");
        if (hasResize) c_.emitLine("    vb6_" + ctl + "_UserControl_Resize((vb6_cls_" + ctl + "*)me);");
        c_.emitLine("}");
        c_.emitLine("static void vb6_" + ctl + "_ucHostTerminate(void* me) {");
        c_.emitLine("    vb6_cls_" + ctl + "_Destroy((vb6_cls_" + ctl + "*)me);");
        c_.emitLine("}");

        int ucScaleMode = 1;
        auto smIt = frmDesc.formControl.properties.find("ScaleMode");
        if (smIt != frmDesc.formControl.properties.end()) ucScaleMode = (int)smIt->second.intValue;

        // czUI fix: 鼠标事件封装 — 宿主 wndproc 收到鼠标消息后经由 desc 钩子
        // 调用 UserControl_MouseDown/Up/Move/DblClick (此前无转发, 控件无法交互)。
        // 注意生成的处理器的 Integer/Single 形参按指针发射。
        const bool hasMouseDown = hasProc("UserControl_MouseDown");
        const bool hasMouseUp   = hasProc("UserControl_MouseUp");
        const bool hasMouseMove = hasProc("UserControl_MouseMove");
        const bool hasDblClick  = hasProc("UserControl_DblClick");
        std::string clsShort = "vb6_cls_" + ctl;
        auto emitMouseWrapper = [&](const char* proc, const char* wrap) {
            c_.emitLine("static void vb6_" + ctl + "_" + wrap + "(void* me, int32_t button, int32_t shift, float x, float y) {");
            c_.emitLine("    int16_t b = (int16_t)button, sh = (int16_t)shift;");
            c_.emitLine("    float fx = x, fy = y;");
            c_.emitLine("    vb6_" + ctl + "_" + proc + "((" + clsShort + "*)me, &b, &sh, &fx, &fy);");
            c_.emitLine("}");
        };
        if (hasMouseDown) emitMouseWrapper("UserControl_MouseDown", "ucHostMouseDown");
        else c_.emitLine("static void vb6_" + ctl + "_ucHostMouseDown(void* me, int32_t b, int32_t sh, float x, float y) { (void)me;(void)b;(void)sh;(void)x;(void)y; }");
        if (hasMouseUp) emitMouseWrapper("UserControl_MouseUp", "ucHostMouseUp");
        else c_.emitLine("static void vb6_" + ctl + "_ucHostMouseUp(void* me, int32_t b, int32_t sh, float x, float y) { (void)me;(void)b;(void)sh;(void)x;(void)y; }");
        if (hasMouseMove) emitMouseWrapper("UserControl_MouseMove", "ucHostMouseMove");
        else c_.emitLine("static void vb6_" + ctl + "_ucHostMouseMove(void* me, int32_t b, int32_t sh, float x, float y) { (void)me;(void)b;(void)sh;(void)x;(void)y; }");
        if (hasDblClick) {
            c_.emitLine("static void vb6_" + ctl + "_ucHostDblClick(void* me) {");
            c_.emitLine("    vb6_" + ctl + "_UserControl_DblClick((" + clsShort + "*)me);");
            c_.emitLine("}");
        } else {
            c_.emitLine("static void vb6_" + ctl + "_ucHostDblClick(void* me) { (void)me; }");
        }

        // 账 #226: UserControl_Click 的封装 (与 DblClick 同形, 零参)。
        const bool hasUcClick = hasProc("UserControl_Click");
        if (hasUcClick) {
            c_.emitLine("static void vb6_" + ctl + "_ucHostClick(void* me) {");
            c_.emitLine("    vb6_" + ctl + "_UserControl_Click((" + clsShort + "*)me);");
            c_.emitLine("}");
        } else {
            c_.emitLine("static void vb6_" + ctl + "_ucHostClick(void* me) { (void)me; }");
        }

        // Fix <vbeclipse> rev18: 自有属性按名桥 (表 + thunk) —— 必须在本 desc 之前发。
        // (.ctl 不走 emitFormFramework, 所以只能落在这个函数里; 见该 .inc 头部说明.)
#include "backend/detail/module/cgen_form_uc_props.inc"

        // Fix <vbeclipse> rev21: 自有 Public Sub 按名桥 (表 + thunk) —— 同样必须在本 desc 之前.
        // 缺了它 `l_ucFolder.ShowView ViewId` 这类晚绑定方法调用会落进宿主模型的
        // "未知方法一律空实现" ⇒ 视图窗体永远 Visible=False ⇒ 停在 0x0 空白。
#include "backend/detail/module/cgen_form_uc_methods.inc"

        c_.emitLine("static const vb6_UserControlDesc vb6_" + ctl + "_ucHostDesc = {");
        c_.emitLine("    \"" + moduleName_ + "\", " + std::to_string(ucScaleMode) + ",");
        c_.emitLine("    (void* (*)(void))vb6_cls_" + ctl + "_New,");
        c_.emitLine("    vb6_" + ctl + "_ucHostInit, vb6_" + ctl + "_ucHostPaint,");
        c_.emitLine("    vb6_" + ctl + "_ucHostResize, vb6_" + ctl + "_ucHostShow, vb6_" + ctl + "_ucHostTerminate,");
        c_.emitLine("    vb6_" + ctl + "_ucHostMouseDown, vb6_" + ctl + "_ucHostMouseUp,");
        // Fix <vbeclipse> rev18: 自有属性按名桥 —— 表与 thunk 由 emitFormFramework 的尾部
        // (cgen_form_uc_props.inc) 先发; 那里把条数记进 ucHostPropCount_, 这里只引用.
        c_.emitLine("    vb6_" + ctl + "_ucHostMouseMove, vb6_" + ctl + "_ucHostDblClick,"
                    + (ucHostPropCount_ <= 0
                           ? std::string(" NULL, 0,")
                           : (" vb6_" + ctl + "_ucProps, "
                              + std::to_string(ucHostPropCount_) + ","))
                    + "  /* Fix <vbeclipse> rev18: 自有属性按名桥 */");
        // Fix <vbeclipse> rev21: 自有 Public Sub 按名桥 (紧跟 props/propCount 之后两槽)。
        c_.emitLine("    " + (ucHostMethodCount_ <= 0
                           ? std::string("NULL, 0,")
                           : ("vb6_" + ctl + "_ucMethods, "
                              + std::to_string(ucHostMethodCount_) + ","))
                    + "  /* Fix <vbeclipse> rev21: 自有方法按名桥 */");
        // Fix <vbeclipse> rev22: 设计期子控件 Resize 事件转发 (rev21 之后最后一槽)。
        c_.emitLine("    " + (designResizeCtrls.empty()
                           ? std::string("NULL")
                           : ("vb6_" + ctl + "_ucHostDesignResize"))
                    + ",  /* Fix <vbeclipse> rev22: 子控件 Resize 事件 */");
        // 账 #226: UC 自身 Click 的末槽 —— 顺序必须与 vb6_UserControlDesc 逐字一致
        // (布局式初始化, 错位一个指针宽就是运行期 AV, 见结构体 rev22 那段教训)。
        c_.emitLine("    vb6_" + ctl + "_ucHostClick  /* 账 #226: UserControl_Click */");
        c_.emitLine("};");
        c_.emitLine("void vb6_" + ctl + "_RegisterHost(void) { vb6_UC_Register(&vb6_" + ctl + "_ucHostDesc); }");
        // Fix <vbeclipse> rev14: 让每个 .ctl 在**本模块的 init 函数**里自注册宿主描述。
        // 此前只有"窗体设计面上直接摆了该 UC"的实例才会在窗体 create-controls 里
        // 调用 RegisterHost; 而运行期 `Controls.Add("VbEclipse.ucFolder", ...)` 动态
        // 创建的类型 (ucFolder/ucSplitBar/ucTab) 从未注册 → vb6_uc_findDesc 返回 NULL
        // → Add 返回 NULL。驱动在 WinMain 里逐个调用各模块 vb6_mod_<X>_init(),
        // 故把注册放进这里就保证了任何 Add 之前描述表已就绪。
        moduleInitStmts_.push_back("vb6_" + ctl + "_RegisterHost();");
        // czUI fix: 暴露 UserControl_ReadProperties — 窗体侧在设计期属性直赋后
        // 以 PropertyBag 重放一次读取, 复现 .ctl 内部"读取后同步"逻辑
        if (hasProc("UserControl_ReadProperties")) {
            c_.emitLine("static void vb6_" + ctl + "_UserControl_ReadProperties(vb6_cls_" + ctl + "* me, void** PropBag);");
            c_.emitLine("void vb6_" + ctl + "_UC_ReadProps(void* me, void** PropBag) {");
            c_.emitLine("    vb6_" + ctl + "_UserControl_ReadProperties((vb6_cls_" + ctl + "*)me, PropBag);");
            c_.emitLine("}");
        }
        c_.emitBlank();

        h_.emitLine("// Fix 112: UserControl 宿主描述注册 (窗体创建子控件前调用)");
        h_.emitLine("void vb6_" + ctl + "_RegisterHost(void);");
    }
}

// Fix 112: "Proyecto1.ucChartBar" — 判断该类是否为工程内 .ctl UserControl 模块.
// 命中返回该模块的设计器描述, 否则 nullptr.
const FrmFile* CCodeGen::findUserControlSpec(const std::string& controlTypeName) const {
    static int dbg = -1;
    if (dbg < 0) {
        const char* e = std::getenv("C3_DBG112");
        dbg = e ? 1 : 0;
        if (dbg && designerFiles_) {
            for (const auto& kv : *designerFiles_) {
                std::fprintf(stderr, "[DBG112] key='%s' form='%s'\n",
                             kv.first.c_str(), kv.second.form.formName.c_str());
            }
        }
    }
    if (dbg) std::fprintf(stderr, "[DBG112] spec query='%s' map=%s\n",
                          controlTypeName.c_str(), designerFiles_ ? "set" : "null");
    if (!designerFiles_ || designerFiles_->empty()) return nullptr;
    std::string cls = controlTypeName;
    size_t dot = cls.rfind('.');
    if (dot != std::string::npos) cls = cls.substr(dot + 1);
    if (cls.empty()) return nullptr;
    cls = Symbol::toLower(cls);   // Fix 112: VB6 大小写不敏感
    // 键是 module.moduleName (VB_Name, 保留原大小写), 故按 VB6 大小写不敏感比较
    const FrmFile* hit = nullptr;
    for (const auto& kv : *designerFiles_) {
        if (Symbol::toLower(kv.first) == cls) { hit = &kv.second; break; }
    }
    if (dbg) std::fprintf(stderr, "[DBG112]   -> %s\n", hit ? "HIT" : "miss");
    return hit;
}

} // namespace vb6c3

