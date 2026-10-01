// cgen_assign_stmt_special.inc - visit(AssignmentStmt&) 函数体片段：语句特殊形态（Date$/Time$ / Mid$ / LSet-RSet / 控件属性 / 跨模块变量）
// 由 src/backend/stmt/cgen_assign.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 12~405 行，逐行未改。

    if (!node.target || !node.value) return;
    // P22: Date$/Time$ statement interception (Date$ = "12-31-2025" / Time$ = "23:59:59")
    if (node.target->kind == ASTNodeKind::IdentifierExpr) {
        auto& _tgtId22 = static_cast<IdentifierExpr&>(*node.target);
        std::string _tgtLower22 = _tgtId22.name;
        std::transform(_tgtLower22.begin(), _tgtLower22.end(), _tgtLower22.begin(), ::tolower);
        if (_tgtLower22.size() > 1 && _tgtLower22.back() == '$') _tgtLower22.pop_back();
        if (_tgtLower22 == "date") {
            emitExpr(*node.value);
            c_.emitLine("vb6_DateSet(" + wrapToBSTR(lastExpr_, *node.value) + ");  /* Date$ = ... */");
            return;
        }
        if (_tgtLower22 == "time") {
            emitExpr(*node.value);
            c_.emitLine("vb6_TimeSet(" + wrapToBSTR(lastExpr_, *node.value) + ");  /* Time$ = ... */");
            return;
        }
    }
    // Fix 010r-14 (Pattern B): Mid$ statement — `Mid$(var, start[, len]) = value`
    // 当 Mid$ 被词法器吃成单一 Identifier "Mid$" 时, 解析器无法走 TokenKind::Mid 分支,
    // 会落到 parseLabelOrAssignmentOrCall(), 把 LHS 当作普通函数调用 (IndexOrCallExpr).
    // 在 codegen 层面识别此模式并改写为 vb6_MidSet(&var, start, len, value); 调用
    // (与 visit(MidStmt) 等价, 仅适用于 AssignmentStmt fallback 路径).
    if (node.target->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& call = static_cast<IndexOrCallExpr&>(*node.target);
        if (call.callee && call.callee->kind == ASTNodeKind::IdentifierExpr) {
            auto& calleeId = static_cast<IdentifierExpr&>(*call.callee);
            std::string calleeLower = calleeId.name;
            std::transform(calleeLower.begin(), calleeLower.end(), calleeLower.begin(), ::tolower);
            // 兼容 "mid$" 和 "mid" 两种写法
            if (calleeLower == "mid$" || calleeLower == "mid") {
                if (call.positional.size() >= 2 && call.positional.size() <= 3) {
                    // 1. 发出 target (字符串变量)
                    emitExpr(*call.positional[0]);
                    std::string midTarget = std::move(lastExpr_);
                    // 2. 发出 start
                    emitExpr(*call.positional[1]);
                    std::string midStart = std::move(lastExpr_);
                    // 3. 发出 length (可选, 缺省为 0 表示 "到字符串末尾")
                    std::string midLen = "0";
                    if (call.positional.size() >= 3) {
                        emitExpr(*call.positional[2]);
                        midLen = std::move(lastExpr_);
                    }
                    // 4. 发出 value (RHS)
                    emitExpr(*node.value);
                    std::string midValue = std::move(lastExpr_);
                    // Fix 086: 与 visit(MidStmt) 的 Fix 084b 一致 — Variant值
                    // (vb6_VariantArrayGet 等) 需转 BSTR, 否则 C2440
                    if (cExprIsVariant(midValue)) {
                        midValue = "vb6_VariantToString(" + midValue + ")";
                    }
                    c_.emitLine("vb6_MidSet(&" + midTarget + ", " + midStart + ", " + midLen + ", " + midValue + ");  /* Mid$ statement */");
                    return;
                }
            }
        }
    }
    // Fix 110b: 控件数组元素默认属性写入 — VB6 里 `TxtARGB(0) = v` 等价于
    // `TxtARGB(0).Text = v` (TextBox 控件数组元素的默认属性写入).
    // 不处理时 LHS 会被展开为 `vb6_GetControlText(vb6_hwnd_TxtARGB)(0)` →
    // C2064/C2106. 正确形态: vb6_SetControlText(vb6_CtrlArr_GetAt(&vb6_arr_X, i), v).
    if (node.target->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& caP = static_cast<IndexOrCallExpr&>(*node.target);
        if (caP.callee && caP.callee->kind == ASTNodeKind::IdentifierExpr
            && !caP.positional.empty()) {
            auto& caId = static_cast<IdentifierExpr&>(*caP.callee);
            std::string caLower = caId.name;
            std::transform(caLower.begin(), caLower.end(), caLower.begin(), ::tolower);
            if (knownControlArrays_.count(caLower)) {
                auto itCtrl = knownFormControls_.find(caLower);
                if (itCtrl != knownFormControls_.end()) {
                    const char* defProp = getDefaultPropertyName(itCtrl->second);
                    std::string writeFn = defProp
                        ? getControlPropWriteFn(itCtrl->second, defProp) : std::string();
                    if (!writeFn.empty()) {
                        emitExpr(*caP.positional[0]);
                        std::string idxArg = std::move(lastExpr_);
                        emitExpr(*node.value);
                        std::string valExpr = std::move(lastExpr_);
                        // Fix 194: 统一入口 (原为手写 SetControlText/SetMenuCaption 判断)
                        valExpr = controlPropValueExpr(writeFn, valExpr, *node.value);
                        std::string origName = knownFormControlOriginalNames_.count(caLower)
                            ? knownFormControlOriginalNames_[caLower] : caId.name;
                        c_.emitLine(writeFn + "(vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(origName)
                                    + ", " + idxArg + "), " + valExpr
                                    + ");  /* ctrl array default prop write */");
                        return;
                    }
                }
            }
        }
    }

    // P22: LSet/RSet statement (LSet strVar = expr / RSet strVar = expr)
    if (node.isLSet || node.isRSet) {
        if (node.target->kind == ASTNodeKind::IdentifierExpr) {
            auto& tgtId = static_cast<IdentifierExpr&>(*node.target);
            std::string tgtC = cIdent(tgtId.name);
            // P22-10: UDT LSet — memory copy between UDTs
            std::string tgtLower = tgtId.name;
            std::transform(tgtLower.begin(), tgtLower.end(), tgtLower.begin(), ::tolower);
            // Fix 160-C: VB 里"给函数名赋值"就是给返回值赋值, 其 C 存储是本函数的局部
            // `vb6_ret_<Fn>`。LSet 目标手写 `cIdent(名)` 于是把**函数名**当成了 C 变量,
            // 且因 `knownUdtVars_` 按 `vb6_ret_<fn>` 的 lower 注册 (decl/cgen_decl_func.cpp:215)
            // 而查不到 → 连类型也判错, 掉进字符串分支:
            //   Private Function GetAppVersionInfo() As VS_FIXEDFILEINFO
            //     LSet GetAppVersionInfo = Value
            //   → vb6_BSTR_Assign(&GetAppVersionInfo, vb6_LSet(vb6_CStr(...)))  C2065
            //   正确: memcpy(&vb6_ret_GetAppVersionInfo, &Value, sizeof(vb6_type_...))
            // 与 160-A 同根: 目标名应由标识符发射器决定, 不能手写。
            bool tgtIsRetSlot160 = false;
            if (currentProc_ && !currentReturnVar_.empty()
                && tgtLower == Symbol::toLower(currentProc_->name)
                && knownUdtVars_.find(tgtLower) == knownUdtVars_.end()
                && !knownLocalVars_.count(tgtLower)) {
                tgtC = currentReturnVar_;
                tgtLower = Symbol::toLower(currentReturnVar_);
                tgtIsRetSlot160 = true;
            }
            auto udtIt = knownUdtVars_.find(tgtLower);
            // Fix 160-A: LSet/RSet 的目标是**手写** `cIdent(名)` 而非走 IdentifierExpr
            // 发射器, 因此漏掉了 cgen_expr_ident_symbol.inc:109-114 的类模块字段规则 —
            // 类私有模块级 UDT (存储于 vb6_cls_<T> 结构体内) 被写成裸名:
            //   LSet VBFlexGridComboBoxRect = RC
            //     → memcpy(&VBFlexGridComboBoxRect, ...)   C2065 (应为 me->...)
            // 源侧经 emitExpr 自带前缀, 所以只有目标侧报错。规则须与标识符侧逐条
            // 一致, 否则局部 Dim UDT 会被错误前缀。
            if (isClassModule_ && currentProc_) {
                Symbol* tgtSym = symTab_.lookup(tgtId.name);
                if (tgtSym && tgtSym->kind == SymbolKind::Variable && !tgtSym->isExternal
                    && !knownLocalVars_.count(tgtLower)) {
                    Symbol* paramSym = symTab_.lookupLocal(tgtId.name);
                    if (!paramSym || paramSym->kind != SymbolKind::Parameter) {
                        tgtC = "me->" + tgtC;
                    }
                }
            }
            if (udtIt != knownUdtVars_.end()) {
                // Target is a UDT: LSet/RSet copies raw memory (truncated to target size)
                emitExpr(*node.value);
                std::string srcExpr = lastExpr_;
                std::string udtCType = udtIt->second;
                // Fix 178: 含所有权成员 (String/动态数组/Variant) 的 UDT 不能按位拷 ——
                // VB6 的 LSet 走类型描述符深拷贝, 这里退化成 memcpy 会让两侧共享成员。
                std::string deep = udtDeepCopyAssign(*node.target, tgtC, node.value.get(), srcExpr);
                if (!deep.empty()) { c_.emitLine(deep); return; }
                c_.emitLine("memcpy(&" + tgtC + ", &" + srcExpr + ", sizeof(" + udtCType + "));  /* LSet UDT */");
                return;
            }
            // String LSet/RSet
            emitExpr(*node.value);
            std::string valExpr = wrapToBSTR(lastExpr_, *node.value);
            // 定长字符串使用固定长度而非SysStringLen
            std::string strLen;
            auto fsIt = knownFixedStringLen_.find(tgtLower);
            if (fsIt != knownFixedStringLen_.end()) {
                strLen = fsIt->second;
            } else {
                strLen = "SysStringLen(" + tgtC + ")";
            }
            if (node.isLSet) {
                c_.emitLine("vb6_BSTR_Assign(&" + tgtC + ", vb6_LSet(" + valExpr + ", " + strLen + "));  /* LSet */");
            } else {
                c_.emitLine("vb6_BSTR_Assign(&" + tgtC + ", vb6_RSet(" + valExpr + ", " + strLen + "));  /* RSet */");
            }
            return;
        }
    }
    // P7.5+P7.6: 控件属性写入
    // 情况1: ctrl.Property = value (非数组)
    // 情况2: ctrlArr(idx).Property = value (数组)
    if (node.target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maExpr = static_cast<MemberAccessExpr&>(*node.target);
        // P18-C: Printer.CurrentX / Printer.CurrentY 赋值
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& objId = static_cast<IdentifierExpr&>(*maExpr.object);
            std::string objName = objId.name;
            std::transform(objName.begin(), objName.end(), objName.begin(), ::tolower);
            std::string memName = maExpr.memberName;
            std::transform(memName.begin(), memName.end(), memName.begin(), ::tolower);
            if (objName == "printer") {
                emitExpr(*node.value);
                if (memName == "currentx") { c_.emitLine("vb6_Printer_SetCurrentX((int32_t)(" + lastExpr_ + "));"); return; }
                if (memName == "currenty") { c_.emitLine("vb6_Printer_SetCurrentY((int32_t)(" + lastExpr_ + "));"); return; }
            }
            // Fix 161f: Screen.MousePointer 全局写入 (CWaitCursor 等类里用)。
            // 早先目标 emitExpr 走读路径生成 vb6_Screen_MousePointer() = ... →
            // C2106 "=" 左操作数必须为左值 (extlist CWaitCursor.cls 实测)。
            if (objName == "screen") {
                if (memName == "mousepointer") {
                    emitExpr(*node.value);
                    c_.emitLine("vb6_Screen_SetMousePointer((int32_t)(" + lastExpr_ + "));");
                    return;
                }
            }
        }

        // M22: Me.Property = expr — 窗体模块中Me的属性赋值 (如 Me.Caption = myName)
        if (maExpr.object && (maExpr.object->kind == ASTNodeKind::MeExpr ||
                              (maExpr.object->kind == ASTNodeKind::IdentifierExpr &&
                               static_cast<IdentifierExpr&>(*maExpr.object).name.size() == 2 &&
                               (static_cast<IdentifierExpr&>(*maExpr.object).name[0] == 'M' || static_cast<IdentifierExpr&>(*maExpr.object).name[0] == 'm') &&
                               (static_cast<IdentifierExpr&>(*maExpr.object).name[1] == 'e' || static_cast<IdentifierExpr&>(*maExpr.object).name[1] == 'E')))) {
            if (isFormModule_ && !knownFormName_.empty()) {
                std::string memLower = maExpr.memberName;
                std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                std::string formHwnd = "vb6_hwnd_" + moduleName_;
                auto origIt = knownFormControlOriginalNames_.find(knownFormName_);
                if (origIt != knownFormControlOriginalNames_.end()) {
                    formHwnd = "vb6_hwnd_" + cIdent(origIt->second);
                }
                // Form properties: Caption, Text, Visible, etc.
                std::string writeFn = getControlPropWriteFn(FrmControlType::Form, maExpr.memberName);
                if (!writeFn.empty()) {
                    emitExpr(*node.value);
                    std::string valExpr = std::move(lastExpr_);
                    // M22: Text/Caption property writes need BSTR value (Fix 194: 收敛到统一入口)
                    valExpr = controlPropValueExpr(writeFn, valExpr, *node.value);
                    c_.emitLine(writeFn + "(" + formHwnd + ", " + valExpr + ");  /* Me." + maExpr.memberName + " */");
                    return;
                }
            }
        }
        // P7.6: 控件数组属性写入 ctrlArr(idx).Property = value
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IndexOrCallExpr) {
            auto& idxExpr = static_cast<IndexOrCallExpr&>(*maExpr.object);
            if (idxExpr.callee && idxExpr.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& arrId = static_cast<IdentifierExpr&>(*idxExpr.callee);
                std::string arrLower = arrId.name;
                std::transform(arrLower.begin(), arrLower.end(), arrLower.begin(), ::tolower);
                if (knownControlArrays_.count(arrLower)) {
                    auto itCtrl = knownFormControls_.find(arrLower);
                    if (itCtrl != knownFormControls_.end()) {
                        std::string writeFn = getControlPropWriteFn(itCtrl->second, maExpr.memberName);
                        if (!writeFn.empty()) {
                            emitExpr(*node.value);
                            std::string valExpr = std::move(lastExpr_);
                            std::string idxArg;
                            if (!idxExpr.positional.empty()) {
                                emitExpr(*idxExpr.positional[0]);
                                idxArg = std::move(lastExpr_);
                                // Re-emit value because we just clobbered lastExpr_
                                emitExpr(*node.value);
                                // Fix 142: RHS 是 COM 属性读 (OCX 控件数组元素属性,
                                // 如 Label1(c).Caption = NewTab1(c).Theme) — 必须展开
                                // marker, 否则 lastExpr_ 停在 vb6_CtrlArr_GetAt(...) 的
                                // HWND 上, .Theme 被静默丢弃 (Caption 拿到句柄数字).
                                if (isComMarker_) resolveComValue();
                                valExpr = std::move(lastExpr_);
                            } else {
                                idxArg = "0";
                                if (isComMarker_) resolveComValue();
                                valExpr = std::move(lastExpr_);
                            }
                            // Fix 194: 字符串型属性 (Caption/Text) 的值必须转 BSTR。
                            // `Label1(0).Caption = List1.ListCount` 原先直传 int32 →
                            // 非 0 时被当指针解引用 (0xC000041D), 为 0 时静默不生效。
                            valExpr = controlPropValueExpr(writeFn, valExpr, *node.value);
                            c_.emitLine(writeFn + "(" + ctrlArrWriteCalleeArgs(writeFn, arrId.name, idxArg) + ", " + valExpr + ");  /* Control Array Property */");
                            return;
                        }
                    }
                }
            }
        }
        // P7.5: 非数组控件属性写入 ctrl.Property = value
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& objId = static_cast<IdentifierExpr&>(*maExpr.object);
            std::string objLower = objId.name;
            std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
            // P16: WithEvents控件属性写入
            {
                auto itWECtrl = knownWithEventsCtrlVars_.find(objLower);
                if (itWECtrl != knownWithEventsCtrlVars_.end()) {
                    std::string writeFn = getControlPropWriteFn(itWECtrl->second, maExpr.memberName);
                    if (!writeFn.empty()) {
                        auto itOrig = knownWithEventsCtrlOrigNames_.find(objLower);
                        std::string weVarName = (itOrig != knownWithEventsCtrlOrigNames_.end()) ? itOrig->second : objLower;
                        emitExpr(*node.value);
                        std::string valExpr = std::move(lastExpr_);
                        valExpr = controlPropValueExpr(writeFn, valExpr, *node.value);  // Fix 194
                        c_.emitLine(writeFn + "(" + weVarName + ", " + valExpr + ");  /* WithEvents ctrl prop write */");
                        return;
                    }
                }
            }
            // Fix 161f-extlist: `lv.View = 3` / `lv.Sorted = False` —— lv 是**形参/
            // 局部持有的 ListView** (`Sub RefillList(lv As ListView)`)。槽变量本身就是
            // HWND, 走控件属性写表 (与上方 knownFormControls_ 同一口径, 只是 hwnd 实参
            // 用裸槽名而非 vb6_hwnd_ 前缀)。不认就落到后续 COM 晚绑定
            // `vb6_ComSetProp(lv, L"View", …)` —— 拿 HWND 当 IDispatch, 静默无效。
            if (listViewSlotVars_.count(objLower)) {
                std::string writeFnLv = getControlPropWriteFn(FrmControlType::ListView, maExpr.memberName);
                if (!writeFnLv.empty()) {
                    emitExpr(*node.value);
                    std::string valExprLv = std::move(lastExpr_);
                    valExprLv = controlPropValueExpr(writeFnLv, valExprLv, *node.value);
                    c_.emitLine(writeFnLv + "((void*)" + objLower + ", " + valExprLv + ");  /* ListView slot prop write */");
                    return;
                }
            }
            auto itCtrl = knownFormControls_.find(objLower);
            if (itCtrl != knownFormControls_.end()) {
                std::string writeFn = getControlPropWriteFn(itCtrl->second, maExpr.memberName);
                if (!writeFn.empty()) {
                    emitExpr(*node.value);
                    // ImageList Picture fix: resolve COM marker on RHS before using value
                    // e.g. Picture2.Picture = ImageList1.ListImages(i).Picture
                    // Bug #3 fix: also detect vb6_ComIface_Picture* function return (e.g. QRCodegenBarcode)
                    bool rhsIsComPicture = false;
                    if (isComMarker_) {
                        if (comMemberName_ == "Picture" || comMemberName_ == "picture") {
                            rhsIsComPicture = true;
                            resolveComValue("Object");
                        } else {
                            resolveComValue();
                        }
                    } else if (lastExprIsComPicture_) {
                        rhsIsComPicture = true;
                        lastExprIsComPicture_ = false;  // consume the flag
                    }
                    std::string valExpr = std::move(lastExpr_);
                    // M22: Text/Caption property writes need BSTR value (Fix 194: 收敛到统一入口)
                    valExpr = controlPropValueExpr(writeFn, valExpr, *node.value);
                    // ImageList Picture fix: use SetControlPictureFromCom for COM IPictureDisp
                    if (rhsIsComPicture && writeFn.find("SetControlPicture") != std::string::npos) {
                        writeFn = "vb6_SetControlPictureFromCom";
                    }
                    c_.emitLine(writeFn + "(" + makeCtrlHwndArg(objLower, itCtrl->second) + ", " + valExpr + ");  /* Control Property */");
                    return;
                }
                // Task #43: writeFn 为空时不再告警落穿 — 控件属性 (如 SSTabEx1.Font /
                // SSTabEx1.TabHeight 等 OCX 增强/自定义属性) 由后续 P6.2 COM 晚绑定
                // 路径接管: LHS emitExpr 经 Fix 023e/089d 留 isComMarker_ (comObjExpr_
                // = 控件 hwnd) → vb6_ComSetProp(hwnd, L"Prop", packed)。
                // 前提是 P6.7 (cgen_assign_prop_write.inc) 的控件守卫已把这些形态
                // 从类 PropertySet 误接路径上拦下 (Task #43, 2026-02)。
            }
        }
    }

    // M22: 跨模块变量赋值 Module1.myName = expr → vb6_BSTR_Assign(&vb6_Module1_myName, expr)
    if (node.target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maExpr = static_cast<MemberAccessExpr&>(*node.target);
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& objIdent = static_cast<IdentifierExpr&>(*maExpr.object);
            std::string objLower = objIdent.name;
            std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
            // Check if object is NOT a known variable → assume module name
            auto* objSym = symTab_.lookup(objIdent.name);
            if (!objSym) objSym = symTab_.lookupModule(objIdent.name);
            bool isVarName = (objSym && (objSym->kind == SymbolKind::Variable || objSym->kind == SymbolKind::Parameter));
            // M22-fix: 补充检查cgen层跟踪集合，防止过程级UDT/类变量被误判为模块名
            if (!isVarName && knownUdtVars_.count(objLower)) isVarName = true;
            // Fix 086: Variant局部变量 (For Each Ctl In o) 也是合法的赋值目标对象 —
            // 成员写入走COM后期绑定 (vb6_ComSetProp), 而非 Module.X 回退 (C2065)
            if (!isVarName && knownVariantVars_.count(objLower)) isVarName = true;
            if (!isVarName && knownClassVars_.find(objLower) != knownClassVars_.end()) isVarName = true;
            if (!isVarName && knownNewVars_.count(objLower)) isVarName = true;
            if (!isVarName && knownObjectVars_.count(objLower)) isVarName = true;
            if (!isVarName && knownTypedComVars_.count(objLower)) isVarName = true;
            if (!isVarName && knownIfaceVars_.count(objLower)) isVarName = true;
            // Fix <vbeclipse>: 本模块自己的 Property Get 也是合法"对象表达式"
            // (ucSplitBar.ctl:161 `Folder.Relationship` / :176
            //   `Folder.Ratio = NewRatio` — Folder 是 Property Get 返回 Folder 类)。
            // 原状 objSym=PropertyGet 不被当变量 → "Module.var" 回退 → 接收者丢:
            //   `Ratio = NewRatio` 裸名 C2065 (ucSplitBar.c:204)。属性目标后交给
            //   prop_write/com_prop 路径按类属性写 vb6_Folder_prop_let_Ratio(...)。
            //   与 m22 的 PropertyGet 变量化 (cgen_expr_member_m22_module.inc:27-31) 对齐。
            if (!isVarName && objSym && objSym->kind == SymbolKind::PropertyGet
                && !objSym->isExternal) isVarName = true;

            // tB Inherits (ai/022 B09): `MyBase.X = v` 走基类面上的 Property Let / 数据字段。
            // 落到下面的 Module.var 回退就是 `vb6_MyBase_X = v` (C2065 未声明)。
            if (objLower == "mybase") {
                emitExpr(*node.value);
                std::string valMyBase09 = std::move(lastExpr_);
                if (tryEmitMyBaseAssign(maExpr, valMyBase09)) return;
            }

            // Fix 084o-4: 函数名引用返回 UDT 的字段赋值 ("FuncName.Field = expr").
            // VB6 函数体内可用函数名访问返回 UDT 的字段 (如 pvVfsOpen.FileName = STR_BUFFER).
            // 目标展开为 vb6_ret_FuncName.Field (与读取侧 cgen_expr.cpp:203 一致),
            // 而非误当 Module 变量展开为 vb6_FuncName_Field (未声明 + 类型错误).
            if (!isVarName && objSym && objSym->kind == SymbolKind::Function && !objSym->isExternal) {
                std::string curLower084o = currentProc_ ? currentProc_->name : std::string();
                std::transform(curLower084o.begin(), curLower084o.end(), curLower084o.begin(), ::tolower);
                // Fix 084o-5: 函数返回类实例时的字段赋值 ("FuncName.Field = expr",
                // 如 Set ConnInst.Conn = Me.Conn) → vb6_ret_FuncName->Field = expr,
                // 而非误当 Module 变量生成 vb6_FuncName_Field (C2065 未声明).
                bool isUdtRet084o = (currentReturnCType_.rfind("vb6_type_", 0) == 0);
                bool isClsRet084o = !isUdtRet084o
                    && currentReturnCType_.find("vb6_cls_") != std::string::npos;
                if (!curLower084o.empty() && objLower == curLower084o
                    && currentReturnVar_.size() > 8
                    && (isUdtRet084o || isClsRet084o)) {
                    // Fix <vbeclipse>: 函数返回**类实例**时, 其成员若是 Property Let
                    // (Perspective.cls `Public Property Let PerspectiveId`), 不是结构体
                    // 字段 → 走类属性写 vb6_<Cls>_prop_let_<M>((void*)ret, val), 否则
                    // `ret->PerspectiveId` 对 vb6_cls_Perspective 无此成员 → C2039
                    // (ucPerspective.c:1391 `p.PerspectiveId = PerspectiveId`).
                    if (isClsRet084o) {
                        std::string ct084o = currentReturnCType_;
                        while (!ct084o.empty()
                               && (ct084o.back() == '*' || ct084o.back() == ' ' || ct084o.back() == '\t'))
                            ct084o.pop_back();
                        std::string cls084o;
                        if (ct084o.size() > 8 && ct084o.compare(0, 8, "vb6_cls_") == 0)
                            cls084o = ct084o.substr(8);
                        if (!cls084o.empty()) {
                            std::string fn084o = resolveClassMemberCall(cls084o, maExpr.memberName);
                            // Fix <vbeclipse>: resolveClassMemberCall 按读方向优先返回
                            // vb6_Perspective_prop_get_PerspectiveId (Perspective.cls Get/Let
                            // 双方向), 写方向须改写为 _prop_let_ — 否则 "_prop_let_" 检查
                            // 不过 → 落到 FnRetUdt 字段写 `ret->PerspectiveId` C2039.
                            if (!fn084o.empty()
                                && fn084o.find("_prop_") != std::string::npos) {
                                std::string let084o = fn084o;
                                std::size_t pg084o = let084o.find("_prop_get_");
                                if (pg084o != std::string::npos)
                                    let084o = let084o.substr(0, pg084o) + "_prop_let_"
                                            + let084o.substr(pg084o + 10);
                                emitExpr(*node.value);
                                std::string val084o = std::move(lastExpr_);
                                Vb6Type valT084o = inferExprType(*node.value);
                                if ((valT084o == Vb6Type::Variant) || cExprIsVariant(val084o))
                                    val084o = "vb6_VariantToString(" + val084o + ")";
                                c_.emitLine(let084o + "((void*)" + currentReturnVar_ + ", "
                                            + val084o + ");  /* FnRetCls prop_let." + maExpr.memberName + " */");
                                return;
                            }
                        }
                    }
                    // 查返回 UDT 的字段类型 (与 inferUdtFieldVb6Type 相同的 udtMembers 查询)
                    std::string udtName084o = isUdtRet084o ? currentReturnCType_.substr(9) : std::string();
                    Symbol* udtSym084o = isUdtRet084o ? symTab_.lookupModule(udtName084o) : nullptr;
                    Vb6Type fldType084o = Vb6Type::Unknown;
                    if (udtSym084o && udtSym084o->kind == SymbolKind::UserDefinedType) {
                        std::string memLower084o = maExpr.memberName;
                        std::transform(memLower084o.begin(), memLower084o.end(), memLower084o.begin(), ::tolower);
                        for (auto& mi : udtSym084o->udtMembers) {
                            if (Symbol::toLower(mi.name) == memLower084o) { fldType084o = mi.type; break; }
                        }
                    }
                    std::string fieldAccess084o = currentReturnVar_
                        + (isUdtRet084o ? "." : "->") + cIdent(maExpr.memberName);
                    emitExpr(*node.value);
                    std::string valExpr084o = std::move(lastExpr_);
                    Vb6Type valType084o = inferExprType(*node.value);
                    bool valIsVariant084o = (valType084o == Vb6Type::Variant) || cExprIsVariant(valExpr084o);
                    switch (fldType084o) {
                        case Vb6Type::String:
                            if (valIsVariant084o) valExpr084o = "vb6_VariantToString(" + valExpr084o + ")";
                            c_.emitLine("vb6_BSTR_Assign(&" + fieldAccess084o + ", " + valExpr084o + ");  /* FnRetUdt." + maExpr.memberName + " */");
                            return;
                        case Vb6Type::Long: case Vb6Type::Integer: case Vb6Type::Byte:
                        case Vb6Type::Boolean: case Vb6Type::ULong: case Vb6Type::LongPtr:
                            if (valIsVariant084o) valExpr084o = "vb6_VariantToLong(" + valExpr084o + ")";
                            c_.emitLine(fieldAccess084o + " = " + valExpr084o + ";  /* FnRetUdt." + maExpr.memberName + " */");
                            return;
                        case Vb6Type::Double: case Vb6Type::Single:
                        case Vb6Type::Currency: case Vb6Type::Date:
                            if (valIsVariant084o) valExpr084o = "vb6_VariantToDouble(" + valExpr084o + ")";
                            c_.emitLine(fieldAccess084o + " = " + valExpr084o + ";  /* FnRetUdt." + maExpr.memberName + " */");
                            return;
                        case Vb6Type::Variant:
                            c_.emitLine("vb6_VariantClear(&" + fieldAccess084o + ");");
                            c_.emitLine(fieldAccess084o + " = vb6_VariantFromValue(" + valExpr084o + ");  /* FnRetUdt." + maExpr.memberName + " */");
                            return;
                        default:
                            c_.emitLine(fieldAccess084o + " = " + valExpr084o + ";  /* FnRetUdt." + maExpr.memberName + " */");
                            return;
                    }
                }
                // Fix 086: 函数名持有对象时的成员赋值 (json_ParseObject.CompareMode = 1,
                // json_ParseObject 是返回 Dictionary 的 Function). 此前误判为
                // Module.X → 生成未声明的 vb6_<fn>_<member> (C2065).
                // 此处处于 FnRetUdt 条件块内: objLower==当前过程名 且 kind 是
                // Function/PropertyGet, 仅需按返回C类型分流.
                if (currentReturnCType_.rfind("vb6_cls_", 0) == 0) {
                    emitExpr(*node.value);
                    std::string valExpr = std::move(lastExpr_);
                    c_.emitLine(currentReturnVar_ + "->" + cIdent(maExpr.memberName)
                                + " = " + valExpr + ";  /* FnRetObj." + maExpr.memberName + " */");
                    return;
                }
                // COM对象/void* 函数返回值 → COM后期绑定属性赋值
                emitExpr(*node.value);
                std::string valExpr = std::move(lastExpr_);
                std::string packFn = comPackExpr(*node.value);
                c_.emitLine("vb6_ComSetProp((void*)" + currentReturnVar_ + ", L\""
                            + maExpr.memberName + "\", " + packFn + "(" + valExpr
                            + "));  /* FnRetObj COM SetProp */");
                return;
            }

            // Fix 086: UDT返回属性的字段赋值 (CurLang.Index = 0) — VB6语义是修改
            // 属性返回的临时副本 (赋值被丢弃). 生成 属性调用.字段 的void表达式,
            // 而非 Module.X 回退 (C2065: vb6_CurLang_Index).
            if (!isVarName) {
                Symbol* propSym86 = symTab_.lookupModule(objIdent.name);
                if (propSym86 && propSym86->kind == SymbolKind::PropertyGet
                    && !propSym86->variableTypeName.empty()) {
                    Symbol* udtSym86 = lookupDotted(propSym86->variableTypeName);
                    if (udtSym86 && udtSym86->kind == SymbolKind::UserDefinedType) {
                        emitExpr(*node.value);  // RHS 副作用保留
                        std::string propFn86 = cProcName("prop_get_" + objIdent.name,
                            propSym86->access,
                            propSym86->isExternal ? propSym86->sourceModule : "");
                        std::string propCall86 = propFn86 + "("
                            + (isClassModule_ ? "(void*)me" : "") + ")";
                        c_.emitLine("(void)(" + propCall86 + "." + cIdent(maExpr.memberName)
                                    + ");  /* UDT prop copy assign (VB6 discards) */");
                        return;
                    }
                }
            }

            if (!isVarName && !knownFormControls_.count(objLower)) {
                // Module.varName = expr
                // When module is #included, use unprefixed name
                std::string modName = objIdent.name;
                std::string varName = cIdent(maExpr.memberName);
                std::string modLower = modName;
                std::transform(modLower.begin(), modLower.end(), modLower.begin(), ::tolower);
                bool isIncluded = false;
                for (const auto& extMod : externalModules_) {
                    std::string extLower = extMod;
                    std::transform(extLower.begin(), extLower.end(), extLower.begin(), ::tolower);
                    if (extLower == modLower) { isIncluded = true; break; }
                }
                std::string memberAccess = isIncluded ? varName : ("vb6_" + cIdent(modName) + "_" + varName);
                emitExpr(*node.value);
                std::string valExpr = std::move(lastExpr_);
                std::string memLower = maExpr.memberName;
                std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                // M22: Check if cross-module variable is BSTR type
                // knownBstrVars_ only has current module's BSTR vars, so also check symbol table
                bool isBstrVar = knownBstrVars_.count(memLower) > 0;
                if (!isBstrVar) {
                    Symbol* memSym = symTab_.lookupModule(maExpr.memberName);
                    if (!memSym) memSym = symTab_.lookup(maExpr.memberName);
                    if (memSym && memSym->kind == SymbolKind::Variable && memSym->type == Vb6Type::String) {
                        isBstrVar = true;
                    }
                }
                if (isBstrVar) {
                    c_.emitLine("vb6_BSTR_Assign(&" + memberAccess + ", " + valExpr + ");  /* Module." + maExpr.memberName + " */");
                } else {
                    c_.emitLine(memberAccess + " = " + valExpr + ";  /* Module." + maExpr.memberName + " */");
                }
                return;
            }
        }
    }
