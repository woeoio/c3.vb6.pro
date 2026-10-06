// cgen_expr_call_callee_withm.inc - visit(IndexOrCallExpr&) 函数体片段：Fix 060 WithMemberExpr callee、callee 发射与控件方法拦截
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1010~1167 行，逐行未改。

    // ---- Fix 158c: WithMemberExpr callee 带 **空括号** — With 块 UDT 动态/固定
    // 数组成员整体引用 .member() → 裸数组表达式 tempVar->member (不带调用括号).
    // 与下方 Fix 060 Pattern B 相同判定, 但不消费 positional (空).
    // 否则 fallback 把 .member() 当函数调用 emit 成 tempVar->member() → C2064
    // (VBFlexGrid.ctl ClearGrid: With VBFlexGridCells.Rows(iRow):
    //     .Cols() = VBFlexGridDefaultCols.Cols() — LHS/RHS 都是整体数组引用).
    if (node.callee && node.callee->kind == ASTNodeKind::WithMemberExpr
        && node.positional.empty() && node.named.empty()
        && !withObjectVars_.empty() && !withObjectInfoStack_.empty()) {
        auto& wmExpr158 = static_cast<WithMemberExpr&>(*node.callee);
        const auto& info158 = withObjectInfoStack_.back();
        if (info158.kind == WithObjKind::Unknown) {
            const std::string& tempVar158 = withObjectVars_.back();
            std::string tempLower158 = Symbol::toLower(tempVar158);
            auto it158 = knownUdtVars_.find(tempLower158);
            if (it158 != knownUdtVars_.end()) {
                const std::string prefix158c = "vb6_type_";
                const std::string& udtCType158 = it158->second;
                if (udtCType158.size() > prefix158c.size()
                    && udtCType158.compare(0, prefix158c.size(), prefix158c) == 0) {
                    std::string udtName158 = udtCType158.substr(prefix158c.size());
                    Symbol* udtSym158 = symTab_.lookupModule(udtName158);
                    if (udtSym158 && udtSym158->kind == SymbolKind::UserDefinedType) {
                        std::string memLower158 = Symbol::toLower(wmExpr158.memberName);
                        for (auto& mi158 : udtSym158->udtMembers) {
                            if (Symbol::toLower(mi158.name) == memLower158) {
                                if (mi158.isArrayDynamic || mi158.arraySize > 0) {
                                    lastExpr_ = tempVar158 + "->" + cIdent(mi158.name);
                                    return;
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }
    }

    // ---- Fix 060: WithMemberExpr callee with positional args — With block UDT array field ----
    // 处理 With 块内 .Data(index) 模式, 其中 .Data 是 UDT 动态数组字段
    // WithMemberExpr 不同于 MemberAccessExpr, 需要单独处理
    if (node.callee && node.callee->kind == ASTNodeKind::WithMemberExpr
        && !node.positional.empty() && node.named.empty()
        && !withObjectVars_.empty() && !withObjectInfoStack_.empty()) {
        auto& wmExpr = static_cast<WithMemberExpr&>(*node.callee);
        const auto& info = withObjectInfoStack_.back();
        if (info.kind == WithObjKind::Unknown) {
            // UDT With block — 查找 UDT 类型
            const std::string& tempVar = withObjectVars_.back();
            std::string tempLower = Symbol::toLower(tempVar);
            auto it = knownUdtVars_.find(tempLower);
            if (it != knownUdtVars_.end()) {
                const std::string prefix = "vb6_type_";
                const std::string& udtCType = it->second;
                if (udtCType.size() > prefix.size()
                    && udtCType.compare(0, prefix.size(), prefix) == 0) {
                    std::string udtName = udtCType.substr(prefix.size());
                    Symbol* udtSym = symTab_.lookupModule(udtName);
                    if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                        std::string memLower = Symbol::toLower(wmExpr.memberName);
                        for (auto& mi : udtSym->udtMembers) {
                            if (Symbol::toLower(mi.name) == memLower) {
                                bool isHandled = false;
                                if (mi.arraySize > 0) {
                                    // Pattern A: 固定大小数组字段 → tempVar->member[idx]
                                    // Fix 081j-2: With 块临时变量是指针，用 -> 访问成员
                                    emitExpr(*node.positional[0]);
                                    std::string idx = std::move(lastExpr_);
                                    lastExpr_ = tempVar + "->" + cIdent(mi.name) + "[" + idx + "]";
                                    isHandled = true;
                                } else if (mi.isArrayDynamic) {
                                    // Pattern B: 动态数组字段 → VB6_SA_AT(elemType, tempVar->member, idx)
                                    // Fix 081j-2: With 块临时变量是指针，用 -> 访问成员
                                    emitExpr(*node.positional[0]);
                                    std::string idx = std::move(lastExpr_);
                                    // Fix 092s: 同 3626 — UDT 元素类型成员数组
                                    std::string elemCType = (mi.type == Vb6Type::UserDefinedType
                                                             && !mi.typeRefName.empty())
                                        ? ("vb6_type_" + cIdent(mi.typeRefName))
                                        : mapSaElemCType(mi.type);
                                    lastExpr_ = "VB6_SA_AT(" + elemCType + ", "
                                              + tempVar + "->" + cIdent(mi.name) + ", " + idx + ")";
                                    isHandled = true;
                                }
                                if (isHandled) return;
                                break;
                            }
                        }
                    }
                }
            }
        } else if (info.kind == WithObjKind::FormControl
                   && (info.ctrlType == FrmControlType::ListBox
                       || info.ctrlType == FrmControlType::ComboBox)) {
            // Fix 152e: With 块内 ListBox/ComboBox 的**可索引属性读取** —
            // `.ItemData(i)` / `.List(i)` / `.Selected(i)` (callee 是
            // WithMemberExpr 且带索引实参)。不拦截时走通用路径: 先 visit
            // (WithMemberExpr) 发 `vb6_GetItemData(tempVar)` (cgen_expr_with.cpp
            // 的 FormControl 读只喂 hwnd 单参, RTL 需 (void* hwnd, int idx) 两参)
            // → C2198 "太少实参", 再把索引当函数调用拼 `(i)` → C2064.
            // 与 P13.3b (MemberAccessExpr 同位) 一致: 直接生成 fn(hwnd, idx).
            const std::string& tempVar = withObjectVars_.back();
            std::string memLower152 = Symbol::toLower(wmExpr.memberName);
            const char* fn152 = nullptr;
            if (memLower152 == "list") fn152 = "vb6_GetListItem";
            else if (memLower152 == "selected") fn152 = "vb6_GetSelected";
            else if (memLower152 == "itemdata") fn152 = "vb6_GetItemData";
            if (fn152) {
                emitExpr(*node.positional[0]);
                std::string idx152 = std::move(lastExpr_);
                lastExpr_ = std::string(fn152) + "((void*)" + tempVar + ", " + idx152
                          + ")  /* With indexed ctrl prop */";
                return;
            }
            // Fix 152f: With 块内 ListBox/ComboBox 的**方法调用** —
            // `.AddItem(item[, index])` / `.RemoveItem(index)`. 这两者既不在
            // getControlPropReadFn (非属性) 也不在通用路径里映射, 原会退化成
            // `tempVar.AddItem` → HWND__.AddItem → C2039 "不是 HWND__ 的成员"
            // (VBFlexGrid demo 里 219 处, 全部 .AddItem)。直接映射到 RTL。
            if (memLower152 == "additem") {
                emitExpr(*node.positional[0]);
                std::string item152 = std::move(lastExpr_);
                lastExpr_ = "vb6_AddItem((void*)" + tempVar + ", "
                          + wrapToBSTR(item152, *node.positional[0])
                          + ")  /* With ctrl .AddItem */";
                return;
            }
            if (memLower152 == "removeitem") {
                emitExpr(*node.positional[0]);
                std::string idx152r = std::move(lastExpr_);
                lastExpr_ = "vb6_RemoveItem((void*)" + tempVar + ", " + idx152r
                          + ")  /* With ctrl .RemoveItem */";
                return;
            }
        }
    }

    // Fix 110b: 控件数组元素默认属性读取 — VB6 里 `TxtARGB(0)` 等价于
    // `TxtARGB(0).Text` (TextBox 控件数组元素的默认属性).
    // 不处理时 callee 的 IdentifierExpr 会先展开成默认属性读取
    // `vb6_GetControlText(vb6_hwnd_TxtARGB)`, 再被拼上 `(0)` →
    // `vb6_GetControlText(vb6_hwnd_TxtARGB)(0)` (C2064 / C2106).
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr
        && !node.positional.empty()) {
        auto& caId110 = static_cast<IdentifierExpr&>(*node.callee);
        std::string caLower110 = caId110.name;
        std::transform(caLower110.begin(), caLower110.end(), caLower110.begin(), ::tolower);
        if (knownControlArrays_.count(caLower110)) {
            auto itCtrl110 = knownFormControls_.find(caLower110);
            if (itCtrl110 != knownFormControls_.end()) {
                const char* defProp110 = getDefaultPropertyName(itCtrl110->second);
                std::string readFn110 = defProp110
                    ? getControlPropReadFn(itCtrl110->second, defProp110) : std::string();
                if (!readFn110.empty()) {
                    emitExpr(*node.positional[0]);
                    std::string idxArg110 = std::move(lastExpr_);
                    std::string orig110 = knownFormControlOriginalNames_.count(caLower110)
                        ? knownFormControlOriginalNames_[caLower110] : caId110.name;
                    lastExpr_ = readFn110 + "(vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(orig110)
                              + ", " + idxArg110 + "))  /* ctrl array default prop */";
                    return;
                }
            }
        }
    }

    // P13.3b: 控件可索引属性读取 — cmbBalloonTooltip.List(0) / .Selected(0) / .ItemData(0).
    // callee 是 MemberAccessExpr(控件, 属性) 且带索引实参时, RTL 读函数需要 hwnd + idx
    // 两个实参; 若走通用路径, 先发 vb6_GetListItem(vb6_hwnd_X) 再拼 (0) →
    // vb6_GetListItem(hwnd)(0) (C2064 把索引当函数调用). 与 Fix 110b 控件数组
    // 默认属性读取同位拦截: 直接生成 vb6_GetListItem((void*)vb6_hwnd_X, idx).
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && !node.positional.empty()) {
        auto& maC131 = static_cast<MemberAccessExpr&>(*node.callee);
        if (maC131.object && maC131.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& mcId131 = static_cast<IdentifierExpr&>(*maC131.object);
            std::string mcLower131 = mcId131.name;
            std::transform(mcLower131.begin(), mcLower131.end(), mcLower131.begin(), ::tolower);
            auto itCtrl131 = knownFormControls_.find(mcLower131);
            // C29-1b: 文件系统三控件的原生窗口就是 LISTBOX / COMBOBOX, 可索引属性
            // 与 ListBox 完全同一批 RTL 读函数, 所以拦截面必须一起放宽 —— 否则
            // dirList.List(0) 落到成员读函数那条单参路径, 发成 vb6_GetListItem(hwnd)(0)。
            if (itCtrl131 != knownFormControls_.end()
                && (itCtrl131->second == FrmControlType::ListBox
                    || itCtrl131->second == FrmControlType::ComboBox
                    || itCtrl131->second == FrmControlType::DirListBox
                    || itCtrl131->second == FrmControlType::FileListBox
                    || itCtrl131->second == FrmControlType::DriveListBox)) {
                std::string prop131 = maC131.memberName;
                std::transform(prop131.begin(), prop131.end(), prop131.begin(), ::tolower);
                const char* fn131 = nullptr;
                if (prop131 == "list") fn131 = "vb6_GetListItem";
                else if (prop131 == "selected") fn131 = "vb6_GetSelected";
                else if (prop131 == "itemdata") fn131 = "vb6_GetItemData";
                if (fn131) {
                    std::string orig131 = knownFormControlOriginalNames_.count(mcLower131)
                        ? knownFormControlOriginalNames_[mcLower131] : mcId131.name;
                    emitExpr(*node.positional[0]);
                    std::string idx131 = std::move(lastExpr_);
                    lastExpr_ = std::string(fn131) + "((void*)vb6_hwnd_" + cIdent(orig131)
                              + ", " + idx131 + ")  /* indexed ctrl prop */";
                    return;
                }
            }
        }
    }

    // M22: 设置asCallCallee_标志, 让IdentifierExpr知道当前是函数调用callee上下文
    // 这确保递归调用时(如 Factorial(n-1))返回函数名而非返回值变量
    // Fix 015: 加 save/restore. 原代码硬编码 `asCallCallee_=false` 会丢失嵌套 callee
    // 上下文 (如 CallStmt→MemberAccessExpr.Fix015→IndexOrCallExpr 链中, 内层 IndexOrCallExpr
    // 反复重置成 false, 外层 Fix 015 看到的是 false 而非 CallStmt 设置的 true).
    // 同时清空 pendingChainObj_ 防止跨调用泄漏 (Fix 015 在 callee emission 时设置).
    pendingChainObj_.clear();
    bool savedAsCallCallee = asCallCallee_;
    asCallCallee_ = true;
    emitExpr(*node.callee);
    asCallCallee_ = savedAsCallCallee;
    std::string callee = std::move(lastExpr_);

    // Bug #3 fix: 清除COM Picture标志 — 每次新的IndexOrCallExpr重置
    // 后续根据被调用函数的返回类型重新设置
    lastExprIsComPicture_ = false;

    // Fix 032: args (positional + named + early-COM 早期路径) 是 值上下文, 不可能是
    // callee 上下文. 外层 CallStmt(2159-2162) 或嵌套 IndexOrCallExpr 可能把
    // asCallCallee_ 留在 true 状态. 若不强制 false, 后续所有 emitExpr(子表达式)
    // 作为参数发联时, 会被 IdentifierExpr 的自引用检查 (line ~199-213) 误判为
    // callee 上下文, 导致 Property Get / Function 返回值变量名引用错误地返回
    // 过程名 (如 vb6_cTlsReMaster_LocalHostName 而非 vb6_ret_LocalHostName).
    // RAII: 函数退出时 (任何 return 或自然走到末尾) 自动恢复为 savedAsCallCallee.
    struct CallCalleeValueScope {
        bool& ref;
        bool saved;
        CallCalleeValueScope(bool& r, bool s) : ref(r), saved(s) { ref = false; }
        ~CallCalleeValueScope() { ref = saved; }
    } _argsValueScope{asCallCallee_, savedAsCallCallee};
    (void)_argsValueScope;  // suppress unused-warning

    // --- P7.9: WebBrowser控件方法调用 ---
    // Navigate/GoBack/GoForward/Refresh via isComMarker_ flag set by MemberAccessExpr
    if (isComMarker_) {
        auto itCtrl = knownFormControls_.find(comObjExpr_);
        if (itCtrl != knownFormControls_.end() && itCtrl->second == FrmControlType::WebBrowser) {
            isComMarker_ = false;
            std::string method = std::move(comMemberName_);
            std::string ctrlName = cIdent(knownFormControlOriginalNames_.count(comObjExpr_) ? knownFormControlOriginalNames_[comObjExpr_] : comObjExpr_);
            comObjExpr_.clear();
            comMemberName_.clear();
            if (method == "navigate") {
                std::string urlArg = "0";
                if (!node.positional.empty()) {
                    emitExpr(*node.positional[0]);
                    urlArg = std::move(lastExpr_);
                }
                lastExpr_ = "(void)vb6_WebViewNavigate((void*)vb6_hwnd_" + ctrlName + ", " + urlArg + ")";
                return;
            } else if (method == "goback" || method == "goforward" || method == "refresh") {
                // Simplified: not yet implemented
                lastExpr_ = "(void)0";
                return;
            }
        }
    }

    // --- C29-Data: Data1.Recordset 的语句方法调用 (Refresh/Move*) ---
    // comObjExpr_ = vb6_Data_Self(...) (recordset 透传形态)。语句路不走 com bind,
    // 必须在这里拦, 否则落 vb6_ComCall(vb6_Data_Self(...), L"Refresh") —— 对
    // ODBC 状态当 IDispatch 用。
    if (isComMarker_ && comObjExpr_.find("vb6_Data_Self(") == 0) {
        isComMarker_ = false;
        std::string method = std::move(comMemberName_);
        comObjExpr_.clear();
        comMemberName_.clear();
        std::string dh = dataSelfHwndExpr(comObjExpr_);
        if (!dh.empty()) {
            if (method == "refresh")      { lastExpr_ = "vb6_Data_Refresh((void*)" + dh + ")"; return; }
            if (method == "movenext")     { lastExpr_ = "vb6_Data_MoveNext((void*)" + dh + ")"; return; }
            if (method == "moveprevious") { lastExpr_ = "vb6_Data_MovePrevious((void*)" + dh + ")"; return; }
            if (method == "movefirst")    { lastExpr_ = "vb6_Data_MoveFirst((void*)" + dh + ")"; return; }
            if (method == "movelast")     { lastExpr_ = "vb6_Data_MoveLast((void*)" + dh + ")"; return; }
            if (method == "fields") {
                // Fields("name") / Fields(1) —— 取值即 FieldValueStr (数字下标转宽串)
                std::string a0 = "L\"\"";
                if (!node.positional.empty()) {
                    emitExpr(*node.positional[0]);
                    a0 = lastExpr_;
                    bool allNum = !a0.empty() && a0.find_first_not_of("0123456789") == std::string::npos;
                    if (allNum) a0 = "L\"" + a0 + "\"";
                }
                lastExpr_ = "vb6_Data_FieldValueStr((void*)" + dh + ", " + a0 + ")";
                return;
            }
        }
        // 未识别: 不恢复 marker —— 恢复成脏标记会污染后续语句
        // (实测: Fields("id") 拿到 "vb6_Data_Self((" 残缺 comObjExpr_ → ComCallObject 空成员)。
        // 落默认 COM 路即可 (marker 已在上面消费清掉)。
    }

    // P13.3: ListBox/ComboBox methods (AddItem/RemoveItem/Clear/List) via isComMarker_ flag
    if (isComMarker_) {
        auto itCtrl = knownFormControls_.find(comObjExpr_);
        if (itCtrl != knownFormControls_.end() &&
            (itCtrl->second == FrmControlType::ListBox || itCtrl->second == FrmControlType::ComboBox)) {
            isComMarker_ = false;
            std::string method = std::move(comMemberName_);
            std::string ctrlName = cIdent(knownFormControlOriginalNames_.count(comObjExpr_) ? knownFormControlOriginalNames_[comObjExpr_] : comObjExpr_);
            comObjExpr_.clear();
            comMemberName_.clear();
            if (method == "additem") {
                std::string itemArg = "0";
                if (!node.positional.empty()) {
                    emitExpr(*node.positional[0]);
                    // Fix 143: 实参是 COM 属性读 (cboThemes.AddItem iTheme.Name) —
                    // 必须在此消费 COM marker, 否则 lastExpr_ 停在对象表达式上
                    // (vb6_AddItem(hwnd, iTheme) 把 IDispatch 指针当字符串塞进列表,
                    // 条目显示成 "???" 乱码), 且残留 marker 会被上层当成独立语句
                    // 发射 (vb6_ComVarFree(ComCall(iTheme,"Name"))) → 结果丢弃.
                    if (isComMarker_) {
                        // 直接生成字符串属性读 (AddItem 形参是字符串);
                        // resolveComValue 会按属性类型给出 BSTR/vb6_VARIANT 两种形态,
                        // 这里要的是 BSTR, 故显式用 vb6_ComGetStringProp.
                        isComMarker_ = false;
                        isEarlyBoundCom_ = false;
                        earlyBoundSym_ = nullptr;
                        std::string objM = std::move(comObjExpr_);
                        std::string memM = std::move(comMemberName_);
                        comObjExpr_.clear();
                        comMemberName_.clear();
                        itemArg = "vb6_ComGetStringProp(" + objM + ", L\"" + memM + "\")";
                    } else {
                        itemArg = std::move(lastExpr_);
                    }
                }
                c_.emitLine("vb6_AddItem((void*)vb6_hwnd_" + ctrlName + ", " + itemArg + ");  /* ListBox.AddItem */");
                lastExpr_ = "0";
                return;
            } else if (method == "removeitem") {
                std::string idxArg = "0";
                if (!node.positional.empty()) {
                    emitExpr(*node.positional[0]);
                    idxArg = std::move(lastExpr_);
                }
                c_.emitLine("vb6_RemoveItem((void*)vb6_hwnd_" + ctrlName + ", " + idxArg + ");  /* ListBox.RemoveItem */");
                lastExpr_ = "0";
                return;
            } else if (method == "clear") {
                c_.emitLine("vb6_ClearList((void*)vb6_hwnd_" + ctrlName + ");  /* ListBox.Clear */");
                lastExpr_ = "0";
                return;
            } else if (method == "list") {
                // List(idx) property read - List1.List(0)
                std::string idxArg = "0";
                if (!node.positional.empty()) {
                    emitExpr(*node.positional[0]);
                    idxArg = std::move(lastExpr_);
                }
                lastExpr_ = "vb6_GetListItem((void*)vb6_hwnd_" + ctrlName + ", " + idxArg + ")";
                return;
            }
        }
    }

    // D6 / C29-9: CommonDialog.Show* → 原生 comdlg32 入口（成员侧打的是小写控件名）。
    // 必须在这里消费掉标记，否则下面的通用 COM 绑定会编成运行期 no-op 的 vb6_ComCall。
    // Fix <vbeclipse>: With <CommonDialog> 块内 .ShowOpen —— 成员侧打的是 With
    // 临时变量 (_vb6_with_N，值就是 CD 属性袋 HWND)，不在 knownFormControls_ 里；
    // 直接以该临时变量为 hwnd 实参收口。
    if (isComMarker_) {
        bool withTempCd = comObjExpr_.rfind("_vb6_with_", 0) == 0;
        auto itCtrlCd = knownFormControls_.find(comObjExpr_);
        if (withTempCd
            || (itCtrlCd != knownFormControls_.end() &&
                itCtrlCd->second == FrmControlType::CommonDialog)) {
            std::string mCd = Symbol::toLower(comMemberName_);
            if (mCd == "showopen" || mCd == "showsave" || mCd == "showcolor"
                || mCd == "showfont" || mCd == "showprinter" || mCd == "showabout") {
                isComMarker_ = false;
                isEarlyBoundCom_ = false;
                earlyBoundSym_ = nullptr;
                std::string hwndCd = withTempCd
                    ? comObjExpr_
                    : cIdent(knownFormControlOriginalNames_.count(comObjExpr_)
                        ? knownFormControlOriginalNames_[comObjExpr_] : comObjExpr_);
                comObjExpr_.clear();
                comMemberName_.clear();
                std::string fnCd = "vb6_CdShow"
                    + std::string(1, (char)::toupper((unsigned char)mCd[4])) + mCd.substr(5);
                c_.emitLine("{" + fnCd + "((void*)" + hwndCd + ");}"
                            "  /* CommonDialog." + mCd + " (原生 comdlg32) */");
                lastExpr_ = "0";
                return;
            }
        }
    }

    // --- C29-WS-a/b: Winsock 的 Bind / SendData / GetData / PeekData / Close / Listen / Connect / Accept ---
    // 同一族"标记必须在这里消费掉"的形状（不消费就落到 vb6_ComCall —— 本线踩过三次的那声不响）。
    // GetData / PeekData 的**出参**是这里唯一的新形状：VB6 那两第一条实参是被写回来的变量，
    // 所以发 `&<变量名>`；实参不是光秃秃的标识符时（属性、数组元素、常量…）发 NULL 让 RTL 安全返回 0
    // —— VB6 在那里同样编不过，只是我们给不出那条编译期诊断（记进 029 §九 本格）。
    if (isComMarker_) {
        auto itWs = knownFormControls_.find(comObjExpr_);
        if (itWs != knownFormControls_.end() &&
            itWs->second == FrmControlType::Winsock) {
            std::string mWs = comMemberName_;
            static const char* kWsKnown[] = { "bind", "senddata", "getdata", "peekdata",
                                              "close", "listen", "connect", "accept" };
            bool knownWs = false;
            for (const char* k : kWsKnown) if (mWs == k) { knownWs = true; break; }
            if (knownWs) {
                std::string hwndWs = cIdent(knownFormControlOriginalNames_.count(comObjExpr_)
                    ? knownFormControlOriginalNames_[comObjExpr_] : comObjExpr_);
                isComMarker_ = false;
                isEarlyBoundCom_ = false;
                earlyBoundSym_ = nullptr;
                comObjExpr_.clear();
                comMemberName_.clear();
                std::vector<std::string> argWs;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    argWs.push_back(lastExpr_);
                }
                auto isBareIdent = [](const std::string& t) {
                    if (t.empty() || !(std::isalpha((unsigned char)t[0]) || t[0] == '_')) return false;
                    for (char ch : t)
                        if (!(std::isalnum((unsigned char)ch) || ch == '_')) return false;
                    return true;
                };
                std::string preWs = "(void*)vb6_hwnd_" + hwndWs;
                if (mWs == "close") {
                    lastExpr_ = "vb6_Ws_Close(" + preWs + ")";
                } else if (mWs == "listen") {
                    lastExpr_ = "vb6_Ws_Listen(" + preWs + ")";
                } else if (mWs == "connect") {
                    lastExpr_ = "vb6_Ws_Connect(" + preWs + ")";
                } else if (mWs == "accept") {
                    lastExpr_ = "vb6_Ws_Accept(" + preWs + ", "
                              + (argWs.size() > 0 ? ("(int32_t)" + argWs[0]) : std::string("0")) + ")";
                } else if (mWs == "bind") {
                    std::string port = argWs.size() > 0 ? ("(int32_t)" + argWs[0]) : std::string("0");
                    std::string ip = argWs.size() > 1 ? argWs[1] : std::string("L\"\"");
                    lastExpr_ = "vb6_Ws_Bind(" + preWs + ", " + port + ", " + ip + ")";
                } else if (mWs == "senddata") {
                    lastExpr_ = "vb6_Ws_SendData(" + preWs + ", "
                              + (argWs.size() > 0 ? argWs[0] : std::string("L\"\"")) + ")";
                } else {
                    std::string out = (!argWs.empty() && isBareIdent(argWs[0]))
                                    ? ("&" + argWs[0]) : std::string("((void*)0)");
                    std::string ty = argWs.size() > 1 ? ("(int32_t)" + argWs[1]) : std::string("0");
                    std::string mx = argWs.size() > 2 ? ("(int32_t)" + argWs[2]) : std::string("(-1)");
                    lastExpr_ = std::string(mWs == "getdata" ? "vb6_Ws_GetData(" : "vb6_Ws_PeekData(")
                              + preWs + ", " + out + ", " + ty + ", " + mx + ")";
                }
                return;
            }
        }
    }

    // --- C29-RT-c: RichTextBox 的 LoadFile / SaveFile / Find ---
    // 三条都是 VB6 真方法（不是判据专用），走的仍是原生消息 + RTL 出口（D6：不碰 OCX）。
    // 与 CommonDialog / Data 两族同一个坑：语句路不进 com_bind，标记必须在这里消费掉，
    // 否则落回通用兜底 = vb6_ComCall(vb6_hwnd_rt1, L"Find") —— 把 HWND 当 IDispatch 用，
    // 编得过、链接过、运行期一声不响（本线踩过三次）。
    // 可选实参按 VB6 的缺省填：fileType 缺省 0 = rtfRTF；Find 的 start/end 缺省 -1
    // （原生 CHARRANGE 里 -1 = "用当前选区起点" / "到文末"），flags 缺省 0。
    if (isComMarker_) {
        auto itRtb = knownFormControls_.find(comObjExpr_);
        if (itRtb != knownFormControls_.end() &&
            itRtb->second == FrmControlType::RichTextBox) {
            std::string mRtb = comMemberName_;
            for (auto& chR : mRtb) chR = (char)::tolower((unsigned char)chR);
            if (mRtb == "loadfile" || mRtb == "savefile" || mRtb == "find"
                || mRtb == "simnotify") {
                std::string hwndRtb = cIdent(knownFormControlOriginalNames_.count(comObjExpr_)
                    ? knownFormControlOriginalNames_[comObjExpr_] : comObjExpr_);
                // 标记必须**先清掉再发实参**：留着的话实参里任何一个字面量都会被上面那套
                // "COM 属性读"的形状捡走 —— 实测 `rt4.Find("alph", 0, -1, 1)` 的第三个实参
                // 被发成 `(-vb6_ComGetIntProp(rt4, L"find"))`（裸控制名 ⇒ C2065）。PictureBox.Print
                // 那刀的注释（Fix 143）就是同一个坑，它是在发完之后再补消费，这里直接先把状态清空。
                isComMarker_ = false;
                isEarlyBoundCom_ = false;
                earlyBoundSym_ = nullptr;
                comObjExpr_.clear();
                comMemberName_.clear();
                std::vector<std::string> argRtb;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    argRtb.push_back(lastExpr_);
                }
                std::string aTxt = argRtb.size() > 0 ? argRtb[0] : std::string("L\"\"");
                if (mRtb == "simnotify") {
                    // 判据专用的一条（RT-d，与 DT-c 的 SimChange / MV-c 的 SimDateClick 同先例）：
                    // 造一条**真**通知发给父窗，把"认来源 + 分流 + 调用"三段派发都喂一遍。
                    // 这里非 VB6 语义，不进手册的方法表。
                    c_.emitLine("{ vb6_RTB_SimNotify((void*)vb6_hwnd_" + hwndRtb + ", (int32_t)"
                                + (argRtb.size() > 0 ? argRtb[0] : std::string("1794")) + "); }"
                                "  /* RichTextBox.SimNotify (判据专用, 非 VB6) */");
                    lastExpr_ = "0";
                    return;
                }
                if (mRtb == "find") {
                    lastExpr_ = "vb6_RTB_Find((void*)vb6_hwnd_" + hwndRtb + ", " + aTxt
                              + ", " + (argRtb.size() > 1 ? argRtb[1] : std::string("-1"))
                              + ", " + (argRtb.size() > 2 ? argRtb[2] : std::string("-1"))
                              + ", " + (argRtb.size() > 3 ? argRtb[3] : std::string("0")) + ")";
                } else {
                    lastExpr_ = std::string(mRtb == "loadfile" ? "vb6_RTB_LoadFile("
                                                               : "vb6_RTB_SaveFile(")
                              + "(void*)vb6_hwnd_" + hwndRtb + ", " + aTxt + ", "
                              + (argRtb.size() > 1 ? argRtb[1] : std::string("0")) + ")";
                }
                return;
            }
        }
    }

    // Fix 185 / 账 #232②: 画布家族的**表达式码头**（带实参或带括号那一形）。
    // 接收者问 `formCtrlSlot`（裸小写名、`vb6_hwnd_X` 两种形态都认，窗体自己那枚也认），
    // 名字与出口问 `controlCanvasMethod`（cls / print / line 三档）—— 两处以前各自
    // 硬编码 `PictureBox` + `== "print" || == "cls"`，所以 `Me.Print "AB"` 没人接。
    // 成员侧 (cgen_expr_member_form_builtin.inc) 把 comObjExpr_ 存成小写控件名，
    // 而 `Me.` 那一路存的是 HWND 表达式，两种形态都由同一处折开。本片段 include 在
    // cgen_expr_call_com_bind.inc **之前**，必须在此消费掉标记，否则通用 COM 绑定会把它
    // 编成运行期 no-op 的 vb6_ComCall。
    if (isComMarker_) {
        FrmControlType cvType = FrmControlType::Unknown;
        std::string cvHwnd;
        std::string cvMem = Symbol::toLower(comMemberName_);
        std::string cvFn = (cvMem == "cls" || cvMem == "print")
            && formCtrlSlot(comObjExpr_, cvType, cvHwnd)
            ? controlCanvasMethod(cvType, cvMem) : std::string();
        if (!cvFn.empty()) {
            isComMarker_ = false;
            isEarlyBoundCom_ = false;
            earlyBoundSym_ = nullptr;
            std::string cvLabel = cvType == FrmControlType::Form ? "Form" : "PictureBox";
            comObjExpr_.clear();
            comMemberName_.clear();
            std::string textArg = "0";  // 裸 Print = 空行
            if (cvMem == "print" && !node.positional.empty()) {
                emitExpr(*node.positional[0]);
                // 同 Fix 143: 实参自身是 COM 属性读时必须就地消费，
                // 否则把对象指针当字符串画、残留标记还会被上层当独立语句发射。
                if (isComMarker_) {
                    isComMarker_ = false;
                    isEarlyBoundCom_ = false;
                    earlyBoundSym_ = nullptr;
                    std::string objM = std::move(comObjExpr_);
                    std::string memM = std::move(comMemberName_);
                    comObjExpr_.clear();
                    comMemberName_.clear();
                    textArg = "vb6_ComGetStringProp(" + objM + ", L\"" + memM + "\")";
                } else {
                    textArg = std::move(lastExpr_);
                }
            }
            c_.emitLine(cvMem == "cls"
                ? (cvFn + "((void*)" + cvHwnd + ");  /* " + cvLabel + ".Cls */")
                : (cvFn + "((void*)" + cvHwnd + ", " + textArg + ");  /* "
                   + cvLabel + ".Print */"));
            lastExpr_ = "0";
            return;
        }
    }

    // 账 #221 = C29-PL-a: PictureBox/Form 的 `Line (x1,y1)-(x2,y2)[, color][, B|C|F]` 从通用
    // COM 兜底里接出来 —— 与上面 Fix 185 的 Print/Cls 同族，两条码头共用 controlCanvasMethod
    // 这一张表（只接一头就是"编得过、跑了、什么都没画"，本线踩过三次的那一条）。
    // 缺的两格按 VB6 的默认交: color 缺席交 -1（RTL 那一头落到控件自己的 ForeColor）、
    // style 缺席交 0（线段）；坐标四格 parser 的 Fix 102 一定交满。
    // 旗标 B/C/F 不在这里出现: parser 已在同一处把它们折成位数值，到这一步是 Integer 字面量。
    if (isComMarker_) {
        FrmControlType cvTypeLine = FrmControlType::Unknown;
        std::string cvHwndLine;
        if (formCtrlSlot(comObjExpr_, cvTypeLine, cvHwndLine)
            && !controlCanvasMethod(cvTypeLine, comMemberName_).empty()) {
            std::string fnLine = controlCanvasMethod(cvTypeLine, comMemberName_);
            isComMarker_ = false;
            isEarlyBoundCom_ = false;
            earlyBoundSym_ = nullptr;
            comObjExpr_.clear();
            comMemberName_.clear();
            std::vector<std::string> la;
            for (size_t i = 0; i < node.positional.size() && i < 6; i++) {
                emitExpr(*node.positional[i]);
                // 同 Fix 143/185: 实参自身是 COM 成员读时必须就地消费标记，否则交出去的是对象
                // 指针，而残留标记还会被上层当独立语句再发一遍。
                if (isComMarker_) {
                    isComMarker_ = false;
                    isEarlyBoundCom_ = false;
                    earlyBoundSym_ = nullptr;
                    std::string objM2 = std::move(comObjExpr_);
                    std::string memM2 = std::move(comMemberName_);
                    comObjExpr_.clear();
                    comMemberName_.clear();
                    la.push_back("vb6_ComGetIntProp(" + objM2 + ", L\"" + memM2 + "\")");
                } else {
                    la.push_back(std::move(lastExpr_));
                }
            }
            while (la.size() < 4) la.push_back("0");
            if (la.size() < 5) la.push_back("-1");
            if (la.size() < 6) la.push_back("0");
            c_.emitLine(fnLine + "((void*)" + cvHwndLine + ", "
                        + "(double)" + la[0] + ", (double)" + la[1] + ", "
                        + "(double)" + la[2] + ", (double)" + la[3] + ", "
                        + "(int32_t)" + la[4] + ", (int32_t)" + la[5] + ");  /* Line */");
            // ⚠ 合并 rev38 时这段的收尾曾被丢掉 (与下面的 Form/Printer 段撞在同一
            // 位置, git 把它判成冲突)。两个 if 各自闭合, 别再省。
            lastExpr_ = "0";
            return;
        }
    }

    // 2026-10-06: Form / Printer 的绘图方法 (PSet/Line/Circle/Point/Cls)。
    // 与上面 PictureBox.Print 同构 —— 成员侧 (form_builtin / Printer 那支) 把
    // comObjExpr_ 存成 Form 的 hwnd 表达式 (vb6_hwnd_<Form>) 或 Printer 哨兵,
    // 这里按方法名翻译成 RTL 调用。
    //
    // 为什么必须有这一段: Form 的未知成员此前**一律**落进下方通用 COM 绑定,
    // 编出 `vb6_ComCall(hwnd, L"PSet", ...)` —— 窗体句柄不是 IDispatch, 运行期
    // 静默 no-op。于是 `Me.PSet / Me.Line / Me.Circle / Me.Point` 在 Form_Paint
    // 里既不画也不报错, 正是"Form 绘图方法缺失"的症状。
    //
    // ⚠ 形参要按 VB6 语义打平: Step 关键字、坐标省略、颜色省略、Box/Fill/BF
    // 都要各自成为**独立实参** (RTL 侧签名见 vb6forms_draw.h), 不能像
    // Print 那样只取第 0 个 —— `Line -(x1,y1)-(x2,y2), c, B` 的信息全在后几个。
    if (isComMarker_) {
        std::string mDraw = Symbol::toLower(comMemberName_);
        // Printer 那支的 obj 是 (void*)vb6_Printer_Object(); Form 那支是 vb6_hwnd_X
        bool looksLikePrinter = (comObjExpr_.find("Printer") != std::string::npos);
        // 账 #232②: `cls` 只在 Printer 那一档还留在这里 —— Printer.Cls 在 VB6 是"结束文档"
        // (vb6_Printer_EndDoc)，不是清画布；画布那两档 (Form / PictureBox) 的名字与出口
        // 归 controlCanvasMethod 那张表，上面两条码头先接走，这里不再答第二遍。
        bool isPrinterDraw = (mDraw == "pset" || mDraw == "line" || mDraw == "circle"
                              || mDraw == "point"
                              || (mDraw == "cls" && looksLikePrinter));
        if (isPrinterDraw && (isFormModule_ || looksLikePrinter)) {
            isComMarker_ = false;
            isEarlyBoundCom_ = false;
            earlyBoundSym_ = nullptr;
            std::string method = std::move(comMemberName_);
            std::string obj = std::move(comObjExpr_);
            comObjExpr_.clear();
            comMemberName_.clear();
            // ⚠ 必须小写化后比较: memberName_ 是**原样大小写** (用户写 `Me.PSet`
            // 就是 "PSet"), 拿它跟小写字面量比会全部落空 ⇒ 一律走 else 分支
            // (实测 PSet/Point/Line 全被发成 vb6_Form_Circle)。
            std::string m = Symbol::toLower(method);

            // 逐个求值位置实参, 顺带把 named 形参 (VB6 允许 Line x1,y1, To x2,y2)
            // 归位。命名实参不在本轮支持范围 (VB6 里 PSet/Line/Circle 的 named
            // 形参只有 color/box/fill/aspect, 已按位置覆盖)。
            std::vector<std::string> a;
            a.reserve(node.positional.size());
            for (size_t ai = 0; ai < node.positional.size(); ai++) {
                emitExpr(*node.positional[ai]);
                if (isComMarker_) {   // 实参自身是 COM 读, 就地消费
                    isComMarker_ = false;
                    isEarlyBoundCom_ = false;
                    earlyBoundSym_ = nullptr;
                    std::string o2 = std::move(comObjExpr_);
                    std::string m2 = std::move(comMemberName_);
                    comObjExpr_.clear();
                    comMemberName_.clear();
                    a.push_back("vb6_ComGetProp((void*)" + o2 + ", L\"" + m2 + "\")");
                } else {
                    a.push_back(std::move(lastExpr_));
                }
            }
            // 形状判定: Step 关键字只能在**语法位置**上出现, 编译期不可得
            // (VB6 的 Step 落到 AST 上已被 parser 吞掉)。所以按"实参个数 + 首参
            // 是不是坐标对"来分派, 这是能拿到的最强信号。
            //   PSet(x,y) / PSet(x,y,c) / PSet Step(x,y) / PSet(x,y),c
            //   Line -(x1,y1)-(x2,y2) [,c] [,B] [,F]
            //   Circle (x,y),r [,c] [,start] [,end] [,aspect]
            //   Point(x,y)
            std::string call;
            if (m == "pset") {
                // PSet(x,y) / PSet(x,y),color —— 3 参时末参是颜色。
                bool hasColor = (a.size() >= 3);
                std::string x = a.size() > 0 ? a[0] : "0";
                std::string y = a.size() > 1 ? a[1] : "0";
                std::string col = hasColor ? a[a.size() - 1] : "0";
                call = looksLikePrinter
                    ? ("vb6_Printer_PSet(0, 1, " + x + ", " + y + ", "
                       + (hasColor ? "1" : "0") + ", " + col + ")")
                    : ("vb6_Form_PSet((void*)" + obj + ", 0, 1, " + x + ", " + y
                       + ", " + (hasColor ? "1" : "0") + ", " + col + ")");
            } else if (m == "point") {
                std::string x = a.size() > 0 ? a[0] : "0";
                std::string y = a.size() > 1 ? a[1] : "0";
                lastExpr_ = looksLikePrinter
                    ? ("vb6_Printer_Point(" + x + ", " + y + ")")
                    : ("vb6_Form_Point((void*)" + obj + ", " + x + ", " + y + ")");
                return;
            } else if (m == "cls") {
                // 只有 Printer 那一形能走到这里（上面的 isPrinterDraw 已经把画布的 cls 交给了
                // controlCanvasMethod 那两条码头，账 #232②）。Printer.Cls = 结束文档。
                call = "vb6_Printer_EndDoc()";
            } else if (m == "line") {
                // Line 的两种形状: -(x1,y1)-(x2,y2) 4 参, 或 (x1,y1)-(x2,y2) 4 参
                // (VB6 里 Step 由 parser 吞, 这里统一按 4 坐标 = 2 点)。
                // 后面的可选实参: color / box / fill, 按 VB6 文档的实参顺序。
                size_t nCoord = (a.size() >= 4) ? 4 : 2;
                std::string x1 = nCoord > 0 ? a[0] : "0";
                std::string y1 = nCoord > 1 ? a[1] : "0";
                std::string x2 = nCoord > 2 ? a[2] : "0";
                std::string y2 = nCoord > 3 ? a[3] : "0";
                std::string hasColor = "0", col = "0", style = "0";
                for (size_t i = nCoord; i < a.size(); i++) {
                    // ⚠ B / BF / F **已由 parser 折成位值** (账 #220, B=1 C=2 F=4,
                    // 按字母逐个置位, BF=5) —— 见 parser_expr_postfix.cpp 的
                    // isLineCall 那一段。所以这里**不要再按名字归一**:
                    // 那会把用户自己的变量 `B` 误当成旗标 (parser 特意只在
                    // style 位置折, 正是为了避开这个)。
                    if (i == nCoord) { hasColor = "1"; col = a[i]; }
                    else if (i == nCoord + 1) { style = a[i]; }
                }
                std::string pfx = looksLikePrinter ? "vb6_Printer_Line" : "vb6_Form_Line";
                if (looksLikePrinter) {
                    call = pfx + "(0, 1, " + x1 + ", " + y1 + ", 0, 1, " + x2 + ", " + y2
                         + ", " + hasColor + ", " + col + ", " + style + ")";
                } else {
                    call = pfx + "((void*)" + obj + ", 0, 1, " + x1 + ", " + y1
                         + ", 0, 1, " + x2 + ", " + y2
                         + ", " + hasColor + ", " + col + ", " + style + ")";
                }
            } else { // circle
                // Circle (x,y),r [,color] [,start] [,end] [,aspect]
                std::string x = a.size() > 0 ? a[0] : "0";
                std::string y = a.size() > 1 ? a[1] : "0";
                std::string r = a.size() > 2 ? a[2] : "0";
                std::string hasColor = "0", col = "0";
                std::string hasStart = "0", sA = "0", hasEnd = "0", eA = "0", hasAsp = "0", asp = "0";
                if (a.size() > 3) { hasColor = "1"; col = a[3]; }
                if (a.size() > 4) { hasStart = "1"; sA = a[4]; }
                if (a.size() > 5) { hasEnd = "1"; eA = a[5]; }
                if (a.size() > 6) { hasAsp = "1"; asp = a[6]; }
                if (looksLikePrinter) {
                    call = "vb6_Printer_Circle(0, 1, " + x + ", " + y + ", " + r
                         + ", " + hasColor + ", " + col + ", " + hasStart + ", " + sA
                         + ", " + hasEnd + ", " + eA + ", " + hasAsp + ", " + asp + ")";
                } else {
                    call = "vb6_Form_Circle((void*)" + obj + ", 0, 1, " + x + ", " + y + ", " + r
                         + ", " + hasColor + ", " + col + ", " + hasStart + ", " + sA
                         + ", " + hasEnd + ", " + eA + ", " + hasAsp + ", " + asp + ")";
                }
            }
            c_.emitLine(call + ";  /* Form." + method + " */");
            lastExpr_ = "0";
            return;
        }
    }