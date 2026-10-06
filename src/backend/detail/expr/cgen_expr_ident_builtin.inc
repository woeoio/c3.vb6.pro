// cgen_expr_ident_builtin.inc - visit(IdentifierExpr&) 函数体片段：内置函数表与用户过程解析
// 由 src/backend/expr/cgen_expr_ident.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 421~740 行，逐行未改。

    // 内置函数映射 (名称 → RTL函数名) - 仅当符号表中没有用户定义的函数时使用

    // Task #39: 裸 Printer / Printers 拦截已上移到 ident_dispatch (符号解析之前,
    // 类模块 me-> 字段规则会先吃掉裸名 — 见 dispatch 段 Task #39 注释)。

    // Fix 048: Strip $ type suffix before builtin lookup (Mid$ -> mid, Left$ -> left, etc.)
    std::string lookupName = lower;
    if (!lookupName.empty() && lookupName.back() == '$') {
        lookupName.pop_back();
    }
    auto it = builtinFuncs.find(lookupName);
    if (it != builtinFuncs.end()) {
        // 无参内置函数: VB6允许省略括号(如 Now, Date, Time)
        // 当IdentifierExpr引用这些函数时，必须生成调用(带括号)
        static const std::unordered_set<std::string> zeroArgBuiltinFuncs = {
            "now", "date", "time", "freefile", "command", "curdir", "timer",
            "rnd", "erl", "doevents"
        };
        if (zeroArgBuiltinFuncs.count(lower)) {
            // Fix 041: Rnd() requires 1 arg (int32_t seed). Emit vb6_Rnd(1) not vb6_Rnd().
            if (lower == "rnd") {
                lastExpr_ = "vb6_Rnd(1)";
            // Fix 094: CurDir 的 RTL 签名是 vb6_CurDir(BSTR drive)，0 参调用必须补
            // NULL (当前驱动器)，否则 C2198 参数太少。与 call 路径 fixup
            // (cgen_expr_call_builtin_fixup.inc P14.2.2) 同语义；ident 路径
            // (无括号引用 / 0 参) 不经过该 fixup，需在此处补。
            } else if (lower == "curdir") {
                lastExpr_ = "vb6_CurDir(NULL)";
            // Fix <vbeclipse>: LoadPicture() 无参 (Screen.MouseIcon = LoadPicture())
            // → vb6_LoadPictureEx(BSTR) 需 1 参, 补 NULL (=空图片). 与 call 路径
            // fixup (cgen_expr_call_builtin_fixup.inc) 同口径.
            } else if (lower == "loadpicture") {
                lastExpr_ = "vb6_LoadPictureEx(NULL)";
            } else {
                lastExpr_ = it->second + "()";
            }
        } else {
            lastExpr_ = it->second;
        }
        return;
    }

    // 用户定义的过程: 加vb6_前缀
    // 查找模块级符号
    Symbol* sym = symTab_.lookupModule(node.name);
    // Fix 110u: 裸标识符不得解析为「其它类/设计器模块的公开成员」.
    // VB6 中类模块 (含 .ctl/.pag) 的 Public 过程/属性只能经对象限定访问
    // (Obj.Member); 无 Option Explicit 时未声明的裸名是隐式 Variant 局部.
    // C3 的跨模块注入 (driver 的 globalPublicSyms → defineExternal) 会把别模块
    // 的全部 Public 成员注入本模块作用域, 于是裸名误命中 — Charts 2020
    // ucProgressCircular.ctl:949
    //   If hBrush = 0 Or Count = 0 Then Exit Function   (Count 是 lCount 的笔误)
    // 命中了 ucPieChart 的 `Public Property Get Count` →
    // vb6_ucPieChart_prop_get_Count() 缺 me 实参 → C2198.
    // 标准模块 (Module1.x / ToolsIDE.IsIDE 等) 的公开过程/变量/属性保持可见
    // (VB6 允许跨模块裸名访问), 因此仅当来源模块本身是 Class 类模块时拦截.
    bool leakedClassMember110u = false;
    if (sym && sym->isExternal && !sym->sourceModule.empty()
        && Symbol::toLower(sym->sourceModule) != Symbol::toLower(moduleName_)) {
        Symbol* srcModSym = symTab_.lookupModule(sym->sourceModule);
        if (srcModSym && srcModSym->kind == SymbolKind::Class) {
            leakedClassMember110u = true;
            sym = nullptr;
        }
    }
    if (sym && (sym->kind == SymbolKind::Sub || sym->kind == SymbolKind::Function
             || sym->kind == SymbolKind::DeclareSub || sym->kind == SymbolKind::DeclareFunc)) {
        // M22: VB6语义 — 在函数体内引用自身函数名等同于引用返回值变量
        // e.g. Function Add(): Add = a & b → vb6_ret_Add = ...; Module1.myName = Add → ...myName = vb6_ret_Add
        if (currentProc_ && !currentReturnVar_.empty() &&
            sym->kind == SymbolKind::Function && !sym->isExternal) {
            std::string procLower = currentProc_->name;
            std::transform(procLower.begin(), procLower.end(), procLower.begin(), ::tolower);
            std::string nameLower = node.name;
            std::transform(nameLower.begin(), nameLower.end(), nameLower.begin(), ::tolower);
            if (procLower == nameLower) {
                lastExpr_ = currentReturnVar_;
                return;
            }
        }
        // Declare函数: 使用VB6函数名的cIdent形式 (通过#define映射到导出名)
        if (sym->kind == SymbolKind::DeclareSub || sym->kind == SymbolKind::DeclareFunc) {
            // Fix 161b-decl-out: 有显式 Alias 时**用 Alias 名**发射。原因: VB 名
            // 可能恰是 SDK 的 A/W 宏名 (WinBase.h `#define GetUserName GetUserNameW`),
            // 此时 cgen 生成的 `#ifndef GetUserName` 守卫为假 ⇒ 调用点被 SDK 宏改写
            // 成宽版, 而我们按 VB6 语义生成的是 ANSI(char*) 编组 → 窄缓冲交给宽 API,
            // 写入 UTF-16 后按 NUL 截断只读到首字符。用 Alias 名 (GetUserNameA) 天然
            // 绕开 SDK 的 A/W 宏。
            auto itAlias = declareAliasMap_.find(lower);
            if (itAlias != declareAliasMap_.end()) {
                lastExpr_ = itAlias->second;  // 已清洗好的 C 标识符 (含 ord_ 前缀等)
            } else {
                lastExpr_ = cIdent(node.name);
            }
        } else if (sym->isExternal) {
            // 跨模块函数: 使用 vb6_<ModuleName>_<ProcName> 格式
            // (O3: 跨模块重载变体 — 语义层选定后缀在此消费, 与定义侧 _ov<fp> 同规则)
            lastExpr_ = cProcName(node.name, sym->access, sym->sourceModule)
                        + ovlCSuffixFromKey(pendingCalleeOvl_);
        } else {
            lastExpr_ = cProcName(node.name, sym->access)
                        + ovlCSuffixFromKey(pendingCalleeOvl_);  // 重载变体后缀 (O2)
        }
        pendingCalleeOvl_.clear();
        // Fix 086: VB6语义 — 无括号引用函数/过程即调用 (x = GetTickCount,
        // MySub 语句形式). callee上下文 (asCallCallee_) 由 IndexOrCallExpr 追加
        // 括号, 此处不重复; 含必选参数的过程无括号引用在 VB6 中是编译错误,
        // 保持裸名让其延续原有行为 (ParamArray 视同可选).
        // Fix 089f: 类模块内无括号引用本类成员 Function (pvSessionID =
        // GenerateSessionID) — C 签名带 me 首参, 只加 "()" 会漏 me → C2198
        // 参数太少. 与 CallStmt (cgen_stmt P6.6) / PropertyGet (上方 832-836)
        // 同一判定方式: 函数名以 "vb6_<本类>_" 前缀开头即本类成员方法.
        if (!asCallCallee_ && !std::getenv("C3_NO_BAREFN")) {
            bool hasRequired = false;
            for (const auto& p : sym->params) {
                if (!p.isOptional && !p.isParamArray) { hasRequired = true; break; }
            }
            if (!hasRequired) {
                if (isClassModule_ && currentProc_) {
                    std::string modPrefix = "vb6_" + cIdent(moduleName_) + "_";
                    if (lastExpr_.find(modPrefix) == 0) {
                        lastExpr_ += "((void*)me)";
                    } else {
                        lastExpr_ += "()";
                    }
                } else {
                    lastExpr_ += "()";
                }
            }
        }
        return;
    }

    // Fix 010: Property Get — bare reference calls the getter
    // e.g. If frMessageHWnd = 0 → vb6_cAsyncSocket_prop_get_frMessageHWnd((void*)me)
    if (sym && sym->kind == SymbolKind::PropertyGet) {
        std::string propCName = "prop_get_" + node.name;
        std::string funcName = cProcName(propCName, sym->access,
                                          sym->isExternal ? sym->sourceModule : "");
        if (asCallCallee_) {
            // Used as callee in IndexOrCallExpr — return function name, caller adds args + me
            lastExpr_ = funcName;
        } else if (isClassModule_ && currentProc_ && !sym->isExternal) {
            // Bare reference in class method — call getter with me
            // Fix 086: 外部属性 (跨模块, 如 ToolsIDE.IsIDE 标准模块属性) 不带 me —
            // 其C签名无 me 参数, 追加会 C2197 (参数太多)
            lastExpr_ = funcName + "((void*)me)";
        } else {
            // Standard module or external — no me parameter
            lastExpr_ = funcName + "()";
        }
        return;
    }

    // 外部变量/常量: 使用 vb6_<ModuleName>_<Name> 格式
    if (sym && sym->isExternal) {
        lastExpr_ = "vb6_" + cIdent(sym->sourceModule) + "_" + cName;
        return;
    }

    // 类符号引用 (直接使用类名作为标识符, 如 Dim x As MyClass)
    if (sym && sym->kind == SymbolKind::Class) {
        lastExpr_ = cName;  // 类名本身就是一个类型标识符
        return;
    }

    // M22: Cross-module variable resolution - if identifier is not found locally,
    // check if it exists as a public variable in an external module
    {
        Symbol* extSym = symTab_.lookupModule(node.name);
        if (extSym && extSym->kind == SymbolKind::Variable && extSym->isExternal) {
            // When module is #included, use the unprefixed name directly
            // (Module1.h declares "extern BSTR myName;" which is already visible)
            std::string srcLower = extSym->sourceModule;
            std::transform(srcLower.begin(), srcLower.end(), srcLower.begin(), ::tolower);
            bool isIncluded = false;
            for (const auto& extMod : externalModules_) {
                std::string extLower = extMod;
                std::transform(extLower.begin(), extLower.end(), extLower.begin(), ::tolower);
                if (extLower == srcLower) { isIncluded = true; break; }
            }
            if (isIncluded) {
                lastExpr_ = cName;
            } else {
                lastExpr_ = "vb6_" + cIdent(extSym->sourceModule) + "_" + cName;
            }
            return;
        }
    }
    // Fix 110f: UserControl/PropertyPage 模块内对宿主内建成员的非限定引用.
    // VB6 允许在控件自身的代码里直接写 `hDC`(等价于 UserControl.hDC) —— 该名字既
    // 不是过程局部变量也不在符号表中, 之前发射裸标识符 `hDC` → C2065
    // (Charts 2020: ucChartArea/ucChartBar/ucPieChart/ucTreeMaps 的 Draw 里
    // `If GdipCreateFromHDC(hDC, hGraphics) Then Exit Sub`).
    // Fix 111: UserControl 的方法在 .ctl 里也常以裸名书写 (ScaleX/ScaleY/AsyncRead/
    // PropertyChanged), 此前落到未解析兜底 → 裸符号 → LNK2019 (Charts 2020
    // LabelPlus/ucChartArea/…)。
    // 账 #159: 哪些名字能裸写、接成哪个 RTL 符号, 现在由 kHostPseudoRows 一张表回答
    // (HPF_BARE 档), 这里不再抄成员名清单。门仍是 `!sym && !knownLocalVars_`
    // (本文件上方), 所以 `Dim hDC As Long` 这类真实局部用法不受影响; 表外的名字
    // 维持原样 (仍留 C2065, 不引入新错误)。
    // Fix <vbeclipse>: `Controls` 在 .ctl/.pag 里无条件是宿主集合, 优先于任何同名符号。
    // 与本表其余条目 (hDC/ScaleWidth/…) 不同, 那些必须靠 `!sym` 保住
    // `Dim hDC As Long` 之类的真实局部用法; 而 `Controls` 是 VB6 保留的宿主成员名,
    // 用户不可能在本模块声明同名局部, driver 的跨模块 CoClass 登记却会在本模块
    // 作用域留下一个 ComClass(17) 符号 (ucTabStrip 实测) → `!sym` 被否 → 发射裸
    // `Controls` → C2065 (ucTabStrip.c For Each l_Ctrl In Controls)。
    // 成员访问形态 (Controls.Item/Add/Remove) 由 cgen_expr_member_precheck.inc 在
    // 符号解析前拦截, 不受影响。
    const bool reservedHostObj = (lower == "controls");
    if (isDesignerModule_ && !knownLocalVars_.count(lower)
        && (!sym || reservedHostObj)) {
        // Fix <vbeclipse> rev14: 裸 `Controls` 是 UserControl.Controls 集合, 必须发射
        // 真实的集合单例 vb6_UC_Controls() (RTL 已按当前 UC 实例回落 formHwnd), 而不是
        // 宿主伪对象前缀 vb6_UserControl_Controls (恒 NULL 的空桩)。ucTabStrip.ctl 的
        // `For Each l_Ctrl In Controls` 走的正是这条裸标识符通路 —— 用 NULL 会让枚举
        // 得到空集, 而 Add/Item 返回 NULL → 后续解引用崩溃。
        if (lower == "controls") {
            lastExpr_ = "vb6_UC_Controls()";
            return;
        }
        std::string rtlHp;
        if (hostPseudoBareName(isPropertyPageDesigner_ ? "PropertyPage" : "UserControl",
                               lower, rtlHp)) {
            lastExpr_ = (isPropertyPageDesigner_ ? "vb6_PropertyPage_" : "vb6_UserControl_")
                        + rtlHp;
            return;
        }
    }

    // 账 #196 第三条: 窗体模块 / 属性页模块里**裸写**的 `ScaleX(...)` / `ScaleY(...)`。
    // VB6 那一形指的是"本模块那张窗口"的换算，而换算只吃两个显式的 from/to 参数 ⇒ 与
    // 控件接收者 / `Me.` / With 块里那枚 共同一处实现 `vb6_ScaleUnitX/Y`（见 vb6rtl_com.c）。
    // 以前它落到本文件末尾的"发射裸标识符"兜底 ⇒ 产物里是一个既没声明也没定义的 `ScaleX(`
    // (隐式声明；真工程里今天只被 /OPT:REF 把整个调用者删掉才没响，见 §B31 那条读数)。
    // 刻意不走 kHostPseudoRows 那条 `vb6_<Host>_<Member>` 的命名契约 —— 那是宿主伪成员的形状，
    // 而这一对不是伪成员；.ctl 里裸写那一形仍由表里 UC 那两行负责，走到上面就 return 了。
    if ((isFormModule_ || isPropertyPageDesigner_)
        && (lower == "scalex" || lower == "scaley")) {
        lastExpr_ = (lower == "scalex") ? "vb6_ScaleUnitX" : "vb6_ScaleUnitY";
        return;
    }

    // Fix 110u (续): 隐式 Variant 局部 (VB6 无 Option Explicit 时的未声明裸名).
    // 上方已把来自其它类模块的"泄漏"符号清空; 这里发射运行期 Empty 值
    // (Charts 2020 ucProgressCircular.ctl:949 的 Count — 语义等价于 VB6 的隐式
    // Variant = Empty, 与 0 比较为真). 用 vb6_VariantEmpty() 而非裸名, 避免 C2065.
    if (leakedClassMember110u) {
        lastExpr_ = "vb6_VariantEmpty()";
        return;
    }

    lastExpr_ = cName;