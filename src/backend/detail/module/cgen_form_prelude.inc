// cgen_form_prelude.inc - emitFormFramework 函数体片段：初始化（frx 加载、菜单ID收集）与窗体属性提取
// 由 src/backend/module/cgen_form.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 54~144 行，逐行未改。

    std::string formName = frmDesc.formName;
    formName_ = formName;  // M22-Issue6: save for Form Print use

    // P24: Load .frx file if present
    bool frxLoaded = false;
    if (!frmDesc.frxFilePath.empty()) {
        frxLoaded = FrxReader::load(frmDesc.frxFilePath);
        // Fix 195: 文件在, 但读不出来 (损坏/被截断/权限) 同样会静默丢属性值。
        if (!frxLoaded) {
            diag_.warn(DiagnosticID::CodeGenFormResourceMissing, SourceLocation{},
                "无法读取二进制资源文件 " + frmDesc.frxFilePath.filename().string()
                + ": " + FrxReader::lastError()
                + "。窗体中引用该文件的属性, 其设计期取值将被忽略。");
        }
    }
    std::string clsName = "VB6_Form_" + cIdent(formName);    // Win32窗口类名
    std::string wndProc = "vb6_form_wndproc_" + cIdent(formName);  // WndProc函数名
    std::string createFn = "vb6_form_create_" + cIdent(formName);  // 控件创建函数名
    std::string showFn = "vb6_form_show_" + cIdent(formName);      // Show函数名
    int ctrlId = 0;  // P7.6: 控件ID计数器(在WM_COMMAND和CreateControls中复用)
    bool isMDIForm = (frmDesc.formControl.controlType == FrmControlType::MDIForm);  // P7.7
    bool isMDIChild = frmDesc.isMDIChild;  // P7.7

    // P20-36: 收集菜单项ID映射 (与emitMenuItem/emitMenuClickDispatch一致)
    {
        std::function<void(const FrmControl&, int&)> collectMenuIds;
        collectMenuIds = [&](const FrmControl& menuCtrl, int& mid) {
            for (const auto& child : menuCtrl.children) {
                std::string cap = child.controlName;
                auto capIt2 = child.properties.find("Caption");
                if (capIt2 != child.properties.end() && capIt2->second.type == FrmValueType::String) {
                    std::string raw = capIt2->second.rawText;
                    if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') cap = raw.substr(1, raw.size() - 2);
                }
                if (cap == "-") continue;  // separator: no ID
                if (!child.children.empty()) {
                    collectMenuIds(child, mid);  // submenu: recurse, no ID for container
                } else {
                    std::string cLower = child.controlName;
                    std::transform(cLower.begin(), cLower.end(), cLower.begin(), ::tolower);
                    // Fix <vbeclipse>: 菜单控件数组 (同名多个 Index=N 项) —— 首元素
                    // 的 id 是数组基址, 语句级 `mnuX(i).Caption = ...` 发
                    // SetMenu*(GetMenu(hwnd), 基址+i, 值) 时要用。
                    auto prevId = knownMenuIds_.find(cLower);
                    if (prevId != knownMenuIds_.end()) {
                        knownMenuArrayBaseIds_.emplace(cLower, prevId->second);
                    }
                    knownMenuIds_[cLower] = mid;
                    mid++;
                }
            }
        };
        int collId = 1000;
        knownMenuFormHwnd_ = "vb6_hwnd_" + cIdent(formName);
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (ctrl.controlType == FrmControlType::Menu) collectMenuIds(ctrl, collId);
        }
    }

    // --- 提取窗体属性 ---
    std::string caption = formName;  // 默认标题=窗体名
    int clientHeight = 3000;  // 默认客户区高度 (缇)
    int clientWidth = 4680;  // 默认客户区宽度 (缇)
    int startupPos = 3;  // 默认: Windows Default

    auto it = frmDesc.formControl.properties.find("Caption");
    if (it != frmDesc.formControl.properties.end() && it->second.type == FrmValueType::String) {
        // 去掉引号
        std::string raw = it->second.rawText;
        if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') {
            caption = raw.substr(1, raw.size() - 2);
        }
    }
    it = frmDesc.formControl.properties.find("ClientHeight");
    if (it != frmDesc.formControl.properties.end()) clientHeight = (int)it->second.intValue;
    it = frmDesc.formControl.properties.find("ClientWidth");
    if (it != frmDesc.formControl.properties.end()) clientWidth = (int)it->second.intValue;
    it = frmDesc.formControl.properties.find("StartUpPosition");
    if (it != frmDesc.formControl.properties.end()) startupPos = (int)it->second.intValue;
    // czUI fix: 窗体级 BackColor (OLE &H00BBGGRR) — 交给 RTL 做类背景刷 +
    // Ambient.BackColor, 未指定时保持 -1 (VB6 默认 BTNFACE)
    int formBackColor = -1;
    it = frmDesc.formControl.properties.find("BackColor");
    if (it != frmDesc.formControl.properties.end()) formBackColor = (int)it->second.intValue;
    // czUI fix: BorderStyle (默认 2=Sizable)。0=None 时 RTL 创建无边框窗口,
    // 否则自绘标题栏 (BorderStyle=None 的 .ctl 窗体) 会叠在系统标题栏下面。
    int formBorderStyle = 2;
    it = frmDesc.formControl.properties.find("BorderStyle");
    if (it != frmDesc.formControl.properties.end()) formBorderStyle = (int)it->second.intValue;

    // czUI fix: UC 事件接线 (文件作用域) — 对每个工程内 UserControl 实例:
    // 解析 .ctl 的 Public Event 声明, 若窗体代码定义了 <实例>_<事件> 处理器,
    // 则发射 处理器前向声明 + thunk + sink 表 (文件作用域, 函数体内不能定义函数),
    // attach 行存入 ucEventAttach_ 供 CreateControls 挂接。
    {
        std::string frmLowCode;
        if (designerFiles_) {
            for (const auto& kvFrm : *designerFiles_) {
                if (Symbol::toLower(kvFrm.first) == Symbol::toLower(formName)) {
                    frmLowCode = Symbol::toLower(kvFrm.second.codeSection);
                    break;
                }
            }
        }
        for (const auto& child : frmDesc.formControl.children) {
            const FrmFile* ucSpecEvt = findUserControlSpec(child.controlTypeName);
            if (!ucSpecEvt) continue;
            // 解析 Public Event 声明
            struct UcEvtPre { std::string name; std::string lowName; std::vector<std::string> ptypes; };
            std::vector<UcEvtPre> ucEventsPre;
            {
                std::string lowCode = Symbol::toLower(ucSpecEvt->codeSection);
                static const char* kTypeMapPre[][2] = {
                    // Fix <vbeclipse> rev17: 补 `object` —— 缺它时命中 default "int32_t",
                    // 于是 `Public Event X(ByRef Editor As Object)` 的前向声明/thunk 写成
                    // `int32_t arg0`, 而回调 typedef (emitEventSink 走 mapTypeRef) 与
                    // RaiseEvent 发送侧都是 `void*`。x86 上同宽侥幸能跑, x64 直接截断指针。
                    {"boolean", "int16_t"}, {"string", "BSTR"},
                    {"integer", "int32_t"}, {"long", "int32_t"},
                    {"single", "float"}, {"double", "double"},
                    {"object", "void*"}, {nullptr, nullptr} };
                size_t pos = 0;
                while (true) {
                    size_t e = lowCode.find("public event ", pos);
                    if (e == std::string::npos) break;
                    size_t nameStart = e + 13;
                    size_t lp = lowCode.find('(', nameStart);
                    size_t eol = lowCode.find('\n', nameStart);
                    if (lp == std::string::npos || (eol != std::string::npos && lp > eol)) { pos = nameStart; continue; }
                    UcEvtPre ev;
                    ev.name = ucSpecEvt->codeSection.substr(nameStart, lp - nameStart);
                    ev.lowName = Symbol::toLower(ev.name);
                    size_t rp = lowCode.find(')', lp);
                    if (rp != std::string::npos && (eol == std::string::npos || rp < eol)) {
                        std::string args = ucSpecEvt->codeSection.substr(lp + 1, rp - lp - 1);
                        std::string argsLow = Symbol::toLower(args);
                        size_t pcur = 0;
                        while (true) {
                            size_t asPos = argsLow.find(" as ", pcur);
                            if (asPos == std::string::npos) break;
                            size_t tStart = asPos + 4;
                            size_t tEnd = tStart;
                            while (tEnd < args.size() && args[tEnd] != ',') tEnd++;
                            std::string ty = args.substr(tStart, tEnd - tStart);
                            while (!ty.empty() && ty.front() == ' ') ty.erase(ty.begin());
                            while (!ty.empty() && (ty.back() == ' ' || ty.back() == '\r')) ty.pop_back();
                            std::string cty = "int32_t";
                            std::string tyLowPre = Symbol::toLower(ty);
                            for (int k = 0; kTypeMapPre[k][0]; k++)
                                if (tyLowPre == kTypeMapPre[k][0]) { cty = kTypeMapPre[k][1]; break; }
                            ev.ptypes.push_back(cty);
                            pcur = tEnd;
                        }
                    }
                    bool dup = false;
                    for (const auto& prev : ucEventsPre)
                        if (prev.lowName == ev.lowName) { dup = true; break; }
                    if (!dup && !ev.name.empty()) ucEventsPre.push_back(ev);
                    pos = nameStart;
                }
            }
            if (ucEventsPre.empty()) continue;
            std::string ucShort = child.controlTypeName;
            size_t dotUc = ucShort.rfind('.');
            if (dotUc != std::string::npos) ucShort = ucShort.substr(dotUc + 1);
            std::string evtPrefixType = "vb6_events_" + cIdent(ucShort);
            std::string sinkVar = "_sink_" + cIdent(child.controlName);
            bool anyHandler = false;
            std::vector<std::string> attachLines;
            for (const auto& ev : ucEventsPre) {
                std::string handlerLower = Symbol::toLower(child.controlName + "_" + ev.name);
                // Fix 160x: 精确匹配事件处理子例程 — 子串 find 会把
                // "sub vbflexgrid1_scrolltip(...)" 误判为存在 "vbflexgrid1_scroll"
                // (MainForm VBFlexGrid1 无 Scroll 处理但 c3 为其声明 thunk →
                // MainForm.c C2129 "静态函数已声明但未定义"). 匹配到 '(' 才算定义.
                if (frmLowCode.find("sub " + handlerLower + "(") == std::string::npos
                    && frmLowCode.find("sub " + handlerLower + " (") == std::string::npos) continue;
                anyHandler = true;
                // Fix <vbeclipse> rev17: 登记"这条过程是事件处理器" —— visit(SubDecl&)
                // 据此把形参按事件 ABI (ByVal) 发, 与下面 thunk 的 arg 类型、
                // emitEventSink 的 typedef、RaiseEvent 发送侧三处对齐 (见状态区注释)。
                ucEventHandlers_.insert(handlerLower);
                std::string fnOrig = cIdent(child.controlName) + "_" + cIdent(ev.name);
                std::string frmPrefix = "vb6_" + cIdent(formName) + "_";
                std::string thunkName = frmPrefix + "evtThunk_" + fnOrig;
                std::string sig = "";
                for (size_t i = 0; i < ev.ptypes.size(); i++)
                    sig += (i ? ", " : "") + ev.ptypes[i] + " arg" + std::to_string(i);
                c_.emitLine("static void " + frmPrefix + fnOrig + "(" + sig + ");");
                std::string thunkSig = "void* handler";
                for (size_t i = 0; i < ev.ptypes.size(); i++)
                    thunkSig += ", " + ev.ptypes[i] + " a" + std::to_string(i);
                std::string callArgs = "";
                for (size_t i = 0; i < ev.ptypes.size(); i++)
                    callArgs += (i ? ", " : "") + std::string("a") + std::to_string(i);
                c_.emitLine("static void " + thunkName + "(" + thunkSig + ") { " + frmPrefix + fnOrig + "(" + callArgs + "); }");
                std::string evtCbType = "vb6_evt_" + cIdent(ucShort) + "_" + cIdent(ev.name) + "_cb";
                attachLines.push_back(sinkVar + ".on" + cIdent(ev.name) + " = (" + evtCbType + ")" + thunkName + ";");
            }
            if (anyHandler) {
                c_.emitLine("static " + evtPrefixType + " " + sinkVar + ";");
                ucEventAttach_[Symbol::toLower(child.controlName)] = attachLines;
            }
        }
    }

    // P24: Form.Icon from .frx
    // Fix 161f: 数据是否可用要在 prelude 这里判一次并记入标志; 早先 create 片段
    // 只看"Icon 属性存在 + frxLoaded"就 emit vb6_frx_icon_* 引用, 遇到 .frx 里
    // 没有真图标数据的工程 (如 extlist FLVStyles.frx 只有 12 字节属性袋) ⇒
    // readPicture 返回空 ⇒ 数组不生成而引用照发 ⇒ C2065 未声明的标识符。
    bool formIconHasData_ = false;
    {
        auto iconIt = frmDesc.formControl.properties.find("Icon");
        if (iconIt != frmDesc.formControl.properties.end() && iconIt->second.type == FrmValueType::FrxReference && frxLoaded) {
            auto pic = FrxReader::readPicture(iconIt->second.frxOffset, false);
            if (!pic.data.empty()) {
                formIconHasData_ = true;
                std::string varName = "vb6_frx_icon_" + cIdent(formName);
                h_.emitLine(bytesToHexArray(pic.data.data(), pic.data.size(), varName));
                // Will be used in WM_CREATE to set window icon
                // Store a flag for later code generation
            }
        }
    }
    // P24: Form.Picture from .frx
    bool formPicHasData_ = false;
    {
        auto picIt2 = frmDesc.formControl.properties.find("Picture");
        if (picIt2 != frmDesc.formControl.properties.end() && picIt2->second.type == FrmValueType::FrxReference && frxLoaded) {
            auto pic = FrxReader::readPicture(picIt2->second.frxOffset, false);
            if (!pic.data.empty()) {
                formPicHasData_ = true;
                std::string varName = "vb6_frx_bgpic_" + cIdent(formName);
                h_.emitLine(bytesToHexArray(pic.data.data(), pic.data.size(), varName));
            }
        }
    }
