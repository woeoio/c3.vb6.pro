// cgen_expr_ident_dispatch.inc - visit(IdentifierExpr&) 函数体片段：标识符分派与 VB6 内建常量
// 由 src/backend/expr/cgen_expr_ident.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 13~233 行，逐行未改。

    // 查找符号确定类型
    std::string cName = cIdent(node.name);

    // 内置对象特殊处理
    std::string lower = node.name;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);

    if (lower == "me") {
        // Same logic as MeExpr: form→hwnd, class→me, else→error
        if (isClassModule_) {
            lastExpr_ = "me";
        } else if (isFormModule_ && !knownFormName_.empty()) {
            auto it = knownFormControlOriginalNames_.find(knownFormName_);
            if (it != knownFormControlOriginalNames_.end()) {
                lastExpr_ = "vb6_hwnd_" + cIdent(it->second);
            } else {
                lastExpr_ = "vb6_hwnd_" + moduleName_;
            }
        } else {
            lastExpr_ = "vb6_Me";  // should not happen (semantic error caught)
        }
        return;
    }

    // Fix 010c: VB6全局Err对象 — 不能用me->ERR访问
    // Err是全局内置对象, 不是类成员
    if (lower == "err") {
        // Err 的默认属性是 Number: `Err = 0` / `If Err Then` 都应读写错误号,
        // 否则会编译成恒真的 `(void*)0 == 0`(实测 BalloonTooltips 的
        // ToolTipExists 因此恒返回 True, tooltip 永不创建)。
        lastExpr_ = "vb6_ErrNumber()";  // Err.Number; 属性访问仍由 MemberAccessExpr 处理
        return;
    }

    // Fix 086: VB6全局App对象作为值 (Common4DLL: Set HostApp = App).
    // 编译产物中没有App COM对象, 用NULL对象引用; 属性/方法访问由
    // MemberAccessExpr的App分支 (vb6_App_Path/GetCurrentThreadId等) 拦截.
    if (lower == "app") {
        lastExpr_ = "(void*)0";
        return;
    }

    // Fix 161: 裸 Console 内置全局对象 (对齐 twinBASIC)。成员 (.WriteLine/
    // .Write/.ReadLine/.ReadKey/.Read) 由 MemberAccessExpr 的 Console 分支处理;
    // 裸名本身极少单独使用, 给个哨兵值避免落到未定义标识符。
    if (lower == "console") {
        lastExpr_ = "(void*)0";
        return;
    }

    // Task #39: 裸 Printer / Printers 内置全局对象。
    // cDlg.cls ShowFont: `If Not Printer Is Nothing Then ... tChooseFont.hDC =
    // Printer.hDC`; IsPrinter: `For Each iPrn In Printers` — 必须在符号解析
    // (ident_symbol 的 me-> 字段规则) 之前拦截, 否则类模块内裸名被当类字段
    // (me->Printer → C2039 "Printer 不是 vb6_cls_cDlg 成员") / 裸 Printers
    // → C2065。Printer → RTL DC 哨兵 (无打印机为 NULL, `Is Nothing` 判真);
    // Printers → 空 RTL Collection (vb6_ForEach_Init 原生支持)。
    // 成员访问 (.hDC/.Print 等) 由 MemberAccessExpr 的 Printer 分支处理。
    if (lower == "printer") {
        lastExpr_ = "vb6_Printer_Object()";
        return;
    }
    if (lower == "printers") {
        lastExpr_ = "vb6_Printers_Collection()";
        return;
    }

    // Fix <vbeclipse>: 裸 Screen 内置全局对象 — 类模块内 (ucPerspective.ctl
    // `Screen.MouseIcon = ...`) 不能被 me->Screen 类字段回退吃掉 (C2039)。
    // 成员访问 (.Width/.MouseIcon/...) 由 MemberAccessExpr 的 Screen 分支处理,
    // 这里只给全局哨兵; Screen.MouseIcon 等未登记成员给可写全局槽语义
    // (RTL vb6_Screen_* 全局, 未建模成员赋值无害)。
    if (lower == "screen") {
        lastExpr_ = "vb6_Screen_Object()";
        return;
    }

    // Fix <vbeclipse>: 裸 Parent 内建对象 (UserControl 的宿主, VB6 声明 As Object)。
    // 与 Screen/Printer/Err 同口径放在内建对象段 (必须在符号解析之前), 否则
    // ucButton.ctl:74 `Set ParentObject = Parent` 落到未解析兜底 → 裸 `Parent`
    // C2065。成员访问不在这里处理, 各走各自通道:
    //   带括号调用 Parent.EventRaise(a, b) → cgen_expr_call_callee_member.inc
    //     的后期绑定分支 (vb6_ComCall(vb6_UC_ParentObject(), ...));
    //   无括号语句 Parent.Refresh          → 同一分支的零实参形态;
    //   属性写 Parent.Active = True        → cgen_assign_host_pseudo.inc
    //     (vb6_ComSetProp(vb6_UC_ParentObject(), ...));
    //   .hWnd/.Icon.Handle/.Move 链         → cgen_base.cpp 的 fixupUserControlParentChain。
    if (lower == "parent" && isDesignerModule_ && !knownLocalVars_.count(lower)) {
        lastExpr_ = "vb6_UC_ParentObject()";
        return;
    }

    // M22: 检查是否是当前函数名 — VB6语义歧义
    // - 作为IndexOrCallExpr的callee(函数调用) → 返回函数过程名
    // - 作为普通表达式(返回值引用) → 返回返回值变量
    // P6.7: Property Get也使用返回值赋值语义
    //
    // Fix <VBFlexGridDemo/GA 36766856082>: 两边都剥类型后缀再比。
    // Task #40 (40960eb) 给 parseIdentifierOrCall 加了"使用点剥类型后缀"
    // (`Command$` → `Command`), 但过程**声明**侧走的是 parseProcedureDecl, 不在
    // parseVariableDecl 的剥后缀范围内 ⇒ `Public Function Command$()` 的符号名仍带 $。
    // 两侧不一致 ⇒ 本判定失配 (`command` vs `command$`) ⇒ 函数体内对自身函数名的
    // 引用掉过本分支, 落到末尾内置函数表命中 VB6 内置的 Command$ ⇒ 生成 vb6_Command():
    //   · `VarPtr(Command$)` → `&(void*){vb6_Command()}`          → C2102
    //   · `Command$ = LTrim$(Command$)` → `vb6_BSTR_AssignMove(&vb6_Command(), ...)`
    //   · vb6_ret_Command_ 声明了却从未被赋值 ⇒ Command$() 永远返回空 (静默语义错)
    // 剥后缀后两边同为 `command`, 回到"引用函数名=引用返回值槽"的 VB6 语义。
    auto stripSuffixForProcMatch = [](const std::string& s) -> std::string {
        if (!s.empty() && std::string("$&%!#@").find(s.back()) != std::string::npos) {
            return s.substr(0, s.size() - 1);
        }
        return s;
    };
    if (currentProc_ && stripSuffixForProcMatch(lower) == stripSuffixForProcMatch(Symbol::toLower(currentProc_->name))
        && (currentProc_->kind == SymbolKind::Function || currentProc_->kind == SymbolKind::PropertyGet)) {
        if (asCallCallee_) {
            // 在IndexOrCallExpr的callee上下文中, 返回函数名供调用
            lastExpr_ = cProcName(currentProc_->name, currentProc_->access, currentProc_->sourceModule);
        } else {
            // 在普通表达式上下文中, 返回返回值变量(VB6: 引用函数名=引用返回值)
            lastExpr_ = currentReturnVar_;
        }
        return;
    }

    // 内置常量/枚举映射
    if (lower == "vbokonly")       { lastExpr_ = "0"; return; }
    if (lower == "vbcancel")       { lastExpr_ = "2"; return; }
    if (lower == "vbabortretryignore") { lastExpr_ = "2"; return; }
    if (lower == "vbyesnocancel")  { lastExpr_ = "3"; return; }
    if (lower == "vbyesno")        { lastExpr_ = "4"; return; }
    if (lower == "vbretrycancel")  { lastExpr_ = "5"; return; }
    if (lower == "vbcritical")     { lastExpr_ = "16"; return; }
    if (lower == "vbquestion")     { lastExpr_ = "32"; return; }
    if (lower == "vbexclamation")  { lastExpr_ = "48"; return; }
    if (lower == "vbinformation")  { lastExpr_ = "64"; return; }
    if (lower == "vbcrlf")         { lastExpr_ = "vb6_BSTR_FromStr(L\"\\r\\n\")"; return; }
    if (lower == "vblf")           { lastExpr_ = "vb6_BSTR_FromStr(L\"\\n\")"; return; }
    if (lower == "vbcr")           { lastExpr_ = "vb6_BSTR_FromStr(L\"\\r\")"; return; }
    if (lower == "vbtab")          { lastExpr_ = "vb6_BSTR_FromStr(L\"\\t\")"; return; }
    // Fix 086: 补齐缺失的VB6内建常量 (ToolsJsonVba/cShadow/pvSubClass 等引用)
    if (lower == "vbback")         { lastExpr_ = "vb6_BSTR_FromStr(L\"\\b\")"; return; }
    if (lower == "vbformfeed")     { lastExpr_ = "vb6_BSTR_FromStr(L\"\\f\")"; return; }
    if (lower == "vbverticaltab")  { lastExpr_ = "vb6_BSTR_FromStr(L\"\\v\")"; return; }
    if (lower == "vbsrccopy")      { lastExpr_ = "(13369376)"; return; }  // &HCC0020
    if (lower == "vbsrcand")       { lastExpr_ = "(8913094)"; return; }   // &H8800C6
    if (lower == "vbsrcpaint")     { lastExpr_ = "(15597702)"; return; }  // &HEE0086
    if (lower == "vbsrcinvert")    { lastExpr_ = "(5588696)"; return; }   // &H555009
    if (lower == "vbblackness")    { lastExpr_ = "(66)"; return; }        // &H42
    if (lower == "vbwhiteness")    { lastExpr_ = "(16711782)"; return; }  // &HFF155A
    if (lower == "vblogeventtypeerror")     { lastExpr_ = "4"; return; }
    if (lower == "vblogeventtypewarning")   { lastExpr_ = "3"; return; }
    if (lower == "vblogeventtypeinformation") { lastExpr_ = "2"; return; }
    if (lower == "vblogeventtypeconstant")  { lastExpr_ = "1"; return; }
    if (lower == "vbnewline")      { lastExpr_ = "vb6_BSTR_FromStr(L\"\\r\\n\")"; return; }
    if (lower == "vbnullstring")   { lastExpr_ = "vb6_BSTR_Empty()"; return; }
    // vbempty is VarType constant 0, NOT vb6_VariantEmpty() — fixed in 010q
    if (lower == "vbnothing")      { lastExpr_ = "NULL"; return; }
    if (lower == "vbtrue")         { lastExpr_ = "(-1)"; return; }
    if (lower == "vbfalse")        { lastExpr_ = "0"; return; }
    // Fix 056b: vbUseDefault (-2) - 常用于字体/光标等可选参数
    if (lower == "vbusedefault")   { lastExpr_ = "(-2)"; return; }

    // Fix 010q: VarType constants
    if (lower == "vbempty")        { lastExpr_ = "0"; return; }
    if (lower == "vbnull")         { lastExpr_ = "1"; return; }
    if (lower == "vbinteger")      { lastExpr_ = "2"; return; }
    if (lower == "vblong")         { lastExpr_ = "3"; return; }
    if (lower == "vbsingle")       { lastExpr_ = "4"; return; }
    if (lower == "vbdouble")       { lastExpr_ = "5"; return; }
    if (lower == "vbcurrency")     { lastExpr_ = "6"; return; }
    if (lower == "vbdate")         { lastExpr_ = "7"; return; }
    if (lower == "vbstring")       { lastExpr_ = "8"; return; }
    if (lower == "vbobject")       { lastExpr_ = "9"; return; }
    if (lower == "vberror")        { lastExpr_ = "10"; return; }
    if (lower == "vbboolean")      { lastExpr_ = "11"; return; }
    if (lower == "vbvariant")      { lastExpr_ = "12"; return; }
    if (lower == "vbdataobject")   { lastExpr_ = "13"; return; }
    if (lower == "vbdecimal")      { lastExpr_ = "14"; return; }
    if (lower == "vbbyte")         { lastExpr_ = "17"; return; }
    if (lower == "vbarray")        { lastExpr_ = "8192"; return; }
    if (lower == "vbuserdefinedtype") { lastExpr_ = "36"; return; }

    // Fix 010q: StrConv constants
    if (lower == "vbunicode")      { lastExpr_ = "64"; return; }
    if (lower == "vbfromunicode")  { lastExpr_ = "128"; return; }
    if (lower == "vbuppercase")    { lastExpr_ = "1"; return; }
    if (lower == "vblowercase")    { lastExpr_ = "2"; return; }
    if (lower == "vbpropercase")   { lastExpr_ = "3"; return; }
    if (lower == "vbwide")         { lastExpr_ = "4"; return; }
    if (lower == "vbnarrow")       { lastExpr_ = "8"; return; }
    if (lower == "vbkatakana")    { lastExpr_ = "16"; return; }
    if (lower == "vbhiragana")     { lastExpr_ = "32"; return; }
    if (lower == "vbsimple")       { lastExpr_ = "256"; return; }

    // Fix 010q: CompareMethod constants
    if (lower == "vbbinarycompare")  { lastExpr_ = "0"; return; }
    if (lower == "vbtextcompare")    { lastExpr_ = "1"; return; }
    if (lower == "vbdatabasecompare") { lastExpr_ = "2"; return; }
    if (lower == "vbusecompare")     { lastExpr_ = "(-1)"; return; }

    // Fix 010q: Misc VB6 constants
    if (lower == "vbobjecterror")  { lastExpr_ = "(-2147221504)"; return; }  // &H800A0000
    if (lower == "vbnullchar")     { lastExpr_ = "vb6_BSTR_Empty()"; return; }  // Chr(0) ≈ empty BSTR
    if (lower == "vbobject")       { lastExpr_ = "9"; return; }  // already above but dedup-safe
    if (lower == "vbcallobject")   { lastExpr_ = "1"; return; }
    if (lower == "vbcallgetnext")  { lastExpr_ = "2"; return; }
    if (lower == "vbcalllet")      { lastExpr_ = "3"; return; }
    if (lower == "vbcallmethod")   { lastExpr_ = "13"; return; }
    if (lower == "vbcallset")      { lastExpr_ = "14"; return; }
    if (lower == "vbfirstjan1")    { lastExpr_ = "1"; return; }
    if (lower == "vbfirstfourdays") { lastExpr_ = "2"; return; }
    if (lower == "vbfirstfullweek") { lastExpr_ = "3"; return; }
    if (lower == "vbusesystem")    { lastExpr_ = "0"; return; }
    if (lower == "vbusesystemdayofweek") { lastExpr_ = "0"; return; }
    if (lower == "vbsunday")       { lastExpr_ = "1"; return; }
    if (lower == "vbmonday")       { lastExpr_ = "2"; return; }
    if (lower == "vbtuesday")      { lastExpr_ = "3"; return; }
    if (lower == "vbwednesday")    { lastExpr_ = "4"; return; }
    if (lower == "vbthursday")     { lastExpr_ = "5"; return; }
    if (lower == "vbfriday")       { lastExpr_ = "6"; return; }
    if (lower == "vbsaturday")     { lastExpr_ = "7"; return; }
    if (lower == "vbusesystemdayofweek") { lastExpr_ = "0"; return; }
    if (lower == "vbfirstdayofweek") { lastExpr_ = "1"; return; }  // vbSunday
    if (lower == "vbfirstweekofyear") { lastExpr_ = "1"; return; }  // vbFirstJan1
    if (lower == "vbgeneraldate")  { lastExpr_ = "0"; return; }
    if (lower == "vblongdate")     { lastExpr_ = "1"; return; }
    if (lower == "vbshortdate")    { lastExpr_ = "2"; return; }
    if (lower == "vblongtime")     { lastExpr_ = "3"; return; }
    if (lower == "vbshorttime")    { lastExpr_ = "4"; return; }

    // 先查符号表: 如果有用户定义的同名符号(变量/过程/常量等), 优先使用
    // 这避免了用户变量名与内置函数名冲突的问题 (如 Dim v As Long vs Val()函数)
    // 同时使用currentProc_->params来判断ByRef参数(比符号表查找更可靠)
    Symbol* foundSym = symTab_.lookup(node.name);
    Symbol* modSym = symTab_.lookupModule(node.name);

    // ByRef参数: 最优先检查 (通过currentProc_->params, 不依赖符号表作用域)
    if (currentProc_) {
        for (auto& param : currentProc_->params) {
            if (Symbol::toLower(param.name) == lower) {
                if (param.isParamArray) {
                    // P14.1.5: ParamArray is SAFEARRAY*, no dereference needed
                    lastExpr_ = cName;
                } else if (!param.isByVal) {
                    // Fix 010r-6 rev2: ByRef array parameters are vb6_SafeArray1D**,
                    // need dereference to get the actual pointer value (*name)
                    if (static_cast<uint16_t>(param.type) & static_cast<uint16_t>(Vb6Type::Array)) {
                        lastExpr_ = "(*" + cName + ")";
                    } else {
                        lastExpr_ = "(*" + cName + ")";
                    }
                } else {
                    lastExpr_ = cName;
                }
                return;
            }
        }
    }

    // Fix 010r-12c + 010r-16b: 局部变量必须遮蔽跨模块外部符号 + 同模块的 Function/Property
    // codegen期间 symTab_.current_ 处于模块作用域(语义分析已完成),
    // lookup()会找到模块级 Public/Private Sub/Function/Property. 如果局部Dim变量
    // 与这些符号同名, VB6 作用域语义规定局部变量优先 (局部 > 模块 > 跨模块).
    // 注意: 自引用(过程内引用自身函数名作为返回值)已在上方 line 181-191 处理,
    // 此处的局部变量一定不是当前过程的自引用.
    if (knownLocalVars_.count(lower)) {
        // Dim As New 自动实例化守卫 (与下方Variable块的逻辑一致)
        auto itNew = knownNewVars_.find(lower);
        if (itNew != knownNewVars_.end()) {
            std::string newExpr;
            auto itCom = knownTypedComVars_.find(lower);
            if (itCom != knownTypedComVars_.end()) {
                newExpr = comNewExprFor(itCom->second);  // Fix 177b: 遮蔽→工程类工厂+IDispatch包装; 未遮蔽→空串
                if (newExpr.empty()) {
                    const std::string& _progId = itCom->second->comProgId.empty() ? itCom->second->name : itCom->second->comProgId;
                    newExpr = "(void*)vb6_NewObject(L\"" + _progId + "\")";
                }
            } else {
                newExpr = "vb6_cls_" + itNew->second + "_New()";
            }
            // Fix 086: Variant局部变量持有对象时, 守卫用 VariantToObjectVal 判空,
            // 赋值用 _Generic vb6_VariantFromValue 包装 (避免 !VARIANT 与 C2440)
            if (knownVariantVars_.count(lower)) {
                c_.emitLine("if (vb6_VariantToObjectVal(" + cName + ") == NULL) " + cName
                            + " = vb6_VariantFromValue(" + newExpr + ");  /* Dim As New auto-instantiate (Variant) */");
            } else {
                c_.emitLine("if (!" + cName + ") " + cName + " = " + newExpr + ";  /* Dim As New auto-instantiate */");
            }
        }
        lastExpr_ = cName;
        return;
    }
