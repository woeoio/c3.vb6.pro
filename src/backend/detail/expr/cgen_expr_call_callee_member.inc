// cgen_expr_call_callee_member.inc - visit(IndexOrCallExpr&) 函数体片段：Fix 037 MemberAccessExpr callee（Pattern A / B / C / D）
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 806~1009 行，逐行未改。

    // ---- Fix 037: MemberAccessExpr callee with positional args — array field/COM dispatch ----
    // 处理 obj.field(idx) 模式, field 可能是:
    //   - UDT 数组字段 (固定/动态)     — Pattern A/B → obj.member[idx] / VB6_SA_AT(...)
    //   - 类 Variant 字段持有 SafeArray — Pattern C → vb6_VariantArrayGet(&obj->member, idx)
    //   - 类 void* 字段持有 COM 对象    — Pattern D → vb6_ComCall(obj->member, L"Item", args, argc)
    // 若 member 是该类的方法/属性, resolveClassMemberCall 返回非空 → 不拦截 (正常函数调用路径).
    // 不拦截则 fallback path (line ~2158) 把 field 误当函数调用 → C2064.

    // ---- Fix 158a: MemberAccessExpr callee 带 **空括号** — UDT 动态/固定数组成员
    // 整体引用 obj.member() → 裸数组表达式 obj.member (不带调用括号).
    // 此前 fallback 把它当函数调用 emit 成 obj.member() → C2064 (项不会计算为接受
    // N 参数的函数), 或作为 ByRef 数组实参被包裹成 &(obj.member()) (VBFlexGrid.ctl
    // SortGrid: Call MergeSortRec(..., VBFlexGridCells.Rows(), Sort) — ByRef
    // Data() As TCOLS, C 形参 vb6_SafeArray1D**). 与 Pattern A/B 相同判定
    // (inferUdtTypeOfExpr + udtMembers + isArrayDynamic / arraySize>0), 但空括号
    // 不消费 positional. 固定数组成员 (Pattern A) 空括号引用同样走这里:
    // C 侧固定数组字段是静态数组 obj.member[], 整体引用即裸名 obj.member.
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && node.positional.empty() && node.named.empty()) {
        auto& maExpr158 = static_cast<MemberAccessExpr&>(*node.callee);
        std::string memLower158 = Symbol::toLower(maExpr158.memberName);
        if (maExpr158.object) {
            std::string udtCType158 = inferUdtTypeOfExpr(*maExpr158.object);
            const std::string prefix158 = "vb6_type_";
            if (udtCType158.size() > prefix158.size()
                && udtCType158.compare(0, prefix158.size(), prefix158) == 0) {
                std::string udtName158 = udtCType158.substr(prefix158.size());
                Symbol* udtSym158 = symTab_.lookupModule(udtName158);
                if (udtSym158 && udtSym158->kind == SymbolKind::UserDefinedType) {
                    for (auto& mi158 : udtSym158->udtMembers) {
                        if (Symbol::toLower(mi158.name) == memLower158) {
                            if (mi158.isArrayDynamic || mi158.arraySize > 0) {
                                emitExpr(*node.callee);
                                return;
                            }
                            break;
                        }
                    }
                }
            }
        }
    }

    // ============================================================
    // P20-44: OLE 拖放的方法调用 —— 两种形状 (独立于下面的 positional 大门,
    // 因为 X.OLEDrag 通常**无实参**, `!node.positional.empty()` 挡不住它):
    //   ① DataObject 形参的方法: Data.SetData(v) / Data.Files(i) / Data.Clear()
    //   ② 控件的拖动方法:        X.OLEDrag
    // 都不能掉 COM 派发 (本机 hwnd 不是 IDispatch)。
    // ============================================================
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maDD = static_cast<MemberAccessExpr&>(*node.callee);
        if (maDD.object && maDD.object->kind == ASTNodeKind::IdentifierExpr
            && node.named.empty() && node.positional.size() <= 1) {
            auto& objIdDD = static_cast<IdentifierExpr&>(*maDD.object);
            std::string objLowerDD = Symbol::toLower(objIdDD.name);
            std::string memLowerDD = Symbol::toLower(maDD.memberName);
            // ① DataObject 形参方法
            if (dataObjectParams_.count(objLowerDD)) {
                std::string ddArg = "(void*)(*" + objIdDD.name + ")";
                if (memLowerDD == "setdata" && node.positional.size() == 1) {
                    emitExpr(*node.positional[0]);
                    lastExpr_ = "{ vb6_oleDD_SetText(" + ddArg + ", "
                              + std::move(lastExpr_) + "); (void)0; }";
                    return;
                }
                if (memLowerDD == "clear" && node.positional.empty()) {
                    lastExpr_ = "{ vb6_oleDD_Clear(" + ddArg + "); (void)0; }";
                    return;
                }
                if (memLowerDD == "files" && node.positional.size() == 1) {
                    emitExpr(*node.positional[0]);
                    lastExpr_ = "vb6_oleDD_GetFileBstr(" + ddArg + ", "
                              + std::move(lastExpr_) + ")  /* DataObject.Files(i) */";
                    return;
                }
            }
            // ② 控件的 OLEDrag 方法 → 造 DataObject(控件文本) + 源事件注册 + DoDragDrop
            auto itCtlDD = knownFormControls_.find(objLowerDD);
            if (itCtlDD != knownFormControls_.end() && memLowerDD == "oledrag") {
                std::string hwndDD = makeCtrlHwndArg(objLowerDD, itCtlDD->second);
                static const char* ddSrcEv[4] = { "_OLEStartDrag", "_OLESetData",
                                                  "_OLEGiveFeedback", "_OLECompleteDrag" };
                std::string regDD;
                for (int ddK = 0; ddK < 4; ddK++) {
                    auto* ddSym = symTab_.lookup(objIdDD.name + ddSrcEv[ddK]);
                    if (!ddSym) continue;
                    std::string ddFn = cProcName(ddSym->name, AccessLevel::Private);
                    regDD += "  { extern void " + ddFn + "(); vb6_OLEDrop_SetHandler("
                           + hwndDD + ", " + std::to_string(ddK + 2) + ", (void*)" + ddFn + "); }" + "\n";
                }
                std::string allowedDD = "3";   // 默认 vbCFCopy|vbCFMove
                if (node.positional.size() == 1) {
                    emitExpr(*node.positional[0]);
                    allowedDD = std::move(lastExpr_);
                }
                lastExpr_ = (std::string)"{ void* ddObj44 = (void*)vb6_oleDD_MakeDataObject("
                          "vb6_GetControlText(" + hwndDD + "), NULL, 0);" + "\n"
                          + regDD
                          + "  vb6_OLEDrag_Start((void*)" + hwndDD + ", ddObj44, "
                          + allowedDD + "); }";
                return;
            }
        }
    }

    // ============================================================
    // Fix 162a: X.Move l[, t[, w[, h]]] —— VB6 语言级方法 (Form/控件通用)。
    // 之前落 vb6_ComGetObjectProp(hwnd, L"Move") → 在 HWND 上取属性必失败,
    // Form_Load/Form_Resize 的布局全部静默失效 (extlist 实测: ListView 停在
    // 设计宽 5835 twips 只 5 列可见; 基准 VB6 里被 Move 拉到全宽 10 列全见)。
    // 生成 vb6_ControlMove(hwnd, L,T,W,H, mask); mask 位 1=L 2=T 4=W 8=H,
    // RTL 侧缺失位保持当前值, 单位 twips 与 Left/Top/Width/Height 一致。
    // ============================================================
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maMV = static_cast<MemberAccessExpr&>(*node.callee);
        if (Symbol::toLower(maMV.memberName) == "move" && node.named.empty()
            && node.positional.size() <= 4 && maMV.object) {
            std::string hwndMV;
            if (maMV.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objIdMV = static_cast<IdentifierExpr&>(*maMV.object);
                std::string objLowerMV = Symbol::toLower(objIdMV.name);
                auto itCtlMV = knownFormControls_.find(objLowerMV);
                if (itCtlMV != knownFormControls_.end()) {
                    hwndMV = makeCtrlHwndArg(objLowerMV, itCtlMV->second);
                } else if (listViewSlotVars_.count(objLowerMV)) {
                    hwndMV = objLowerMV;   // Fix 161f: 槽变量本身就是 HWND
                } else if (objLowerMV == "me") {   // 保险: Me 被归为 IdentifierExpr 的解析器
                    auto itOrigMV = knownFormControlOriginalNames_.find(knownFormName_);
                    hwndMV = "vb6_hwnd_" + (itOrigMV != knownFormControlOriginalNames_.end()
                                            ? cIdent(itOrigMV->second) : moduleName_);
                }
            } else if (maMV.object->kind == ASTNodeKind::MeExpr
                       && isFormModule_ && !knownFormName_.empty()) {
                auto itOrigMV = knownFormControlOriginalNames_.find(knownFormName_);
                hwndMV = "vb6_hwnd_" + (itOrigMV != knownFormControlOriginalNames_.end()
                                        ? cIdent(itOrigMV->second) : moduleName_);
            }
            if (!hwndMV.empty()) {
                std::string aMV[4] = {"0", "0", "0", "0"};
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    aMV[i] = std::move(lastExpr_);
                }
                int maskMV = (1 << (int)node.positional.size()) - 1;
                lastExpr_ = "{ vb6_ControlMove(" + hwndMV + ", " + aMV[0] + ", "
                          + aMV[1] + ", " + aMV[2] + ", " + aMV[3] + ", "
                          + std::to_string(maskMV) + "); (void)0; }";
                return;
            }
        }
    }

    // Fix <vbeclipse>: UserControl 的隐式 `Parent` 伪对象 (VB6 声明 As Object) 上的
    // 成员调用必须走**后期绑定**, 不能按"模块名.成员"做全工程按名查找 ——
    //   ucFolder.ctl:338  Parent.EventRaise "RemoveFolder", FolderId, ViewId
    //   ucTab.ctl:298    Parent.EventRaise "DblClick", Me.ViewId
    //   ucSplitBar.ctl:138  Parent.Refresh          (无括号语句式)
    //   ucTabStrip.ctl:199  Parent.Active = True    (属性写 → cgen_assign_host_pseudo.inc)
    // `Parent` 不是已知变量 → 落到 class_module.inc 优先级3, 用 lookupModule 命中
    // 工程内**别的**类的同名方法 (ucTabStrip.EventRaise), 发出裸函数名
    // vb6_ucTabStrip_EventRaise: ① 丢接收者 (原型首参 me) ② 形参表按被误选的 2 参
    // 签名裁剪, 3 实参被 Fix 041c 截成 2 个 → C2198 ×18; 无括号形态则发
    // vb6_Parent_Refresh (C2065, ucSplitBar.c:205).
    // VB6 里 Parent As Object 的成员调用本就是 IDispatch 后期绑定, 这里同构发
    // vb6_ComCall(vb6_UC_ParentObject(), L"<成员>", 实参包, argc).
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && node.named.empty() && isDesignerModule_) {
        auto& maPar = static_cast<MemberAccessExpr&>(*node.callee);
        if (maPar.object && maPar.object->kind == ASTNodeKind::IdentifierExpr
            && Symbol::toLower(static_cast<IdentifierExpr&>(*maPar.object).name) == "parent") {
            std::vector<std::string> packedPar;
            for (size_t i = 0; i < node.positional.size(); i++) {
                std::string packFnPar = comPackExpr(*node.positional[i]);
                emitExpr(*node.positional[i]);
                { std::string resolvedPar = resolveComMarkerForPack(packFnPar);
                  if (!resolvedPar.empty()) lastExpr_ = resolvedPar; }
                packedPar.push_back(packFnPar + "(" + lastExpr_ + ")");
            }
            int32_t argcPar = (int32_t)packedPar.size();
            std::string argsArrayPar = "NULL";
            if (argcPar > 0) {
                argsArrayPar = "(void*[]){";
                for (int i = 0; i < argcPar; i++) {
                    if (i > 0) argsArrayPar += ", ";
                    argsArrayPar += packedPar[i];
                }
                argsArrayPar += "}";
            }
            lastExpr_ = "vb6_ComCall(vb6_UC_ParentObject(), L\""
                      + escapeWideCString(maPar.memberName) + "\", "
                      + argsArrayPar + ", " + std::to_string(argcPar) + ")";
            return;
        }
    }

    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && !node.positional.empty() && node.named.empty()) {
        auto& maExpr = static_cast<MemberAccessExpr&>(*node.callee);
        std::string memLower = Symbol::toLower(maExpr.memberName);
        std::string memLowerM = "m_" + memLower;
        bool handled = false;

        // ---- P20-42: SSTab 的带下标索引属性 TabCaption(i) / TabVisible(i) ----
        // 与 StatusBar 的 `Panels(i).Text` 不是一回事: 那个是"集合对象再取成员属性",
        // 这里 TabCaption 是**控件自身的参数化属性**, VB6 写作 SSTab1.TabCaption(0)。
        // 不拦就会掉到下面的 COM 派发, 生成 vb6_ComCall(vb6_hwnd_SSTab1, L"TabCaption", …)
        // —— 可本机 hwnd 不是 IDispatch, 运行期静默答错或直接崩。
        // 写完这一侧还要配 cgen_util_comwrite.cpp 的对应写分支 (Get→Set)。
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr
            && node.positional.size() == 1) {
            auto& objIdTs = static_cast<IdentifierExpr&>(*maExpr.object);
            std::string objLowerTs = Symbol::toLower(objIdTs.name);
            auto itCtlTs = knownFormControls_.find(objLowerTs);
            if (itCtlTs != knownFormControls_.end()
                && itCtlTs->second == FrmControlType::SSTab) {
                std::string fnTs;
                if (memLower == "tabcaption")      fnTs = "vb6_SSTab_GetTabCaption";
                else if (memLower == "tabvisible") fnTs = "vb6_SSTab_GetTabVisible";
                else if (memLower == "tabtooltiptext") fnTs = "vb6_SSTab_GetTabToolTipText";
                if (!fnTs.empty()) {
                    emitExpr(*node.positional[0]);
                    std::string idxTs = std::move(lastExpr_);
                    lastExpr_ = fnTs + "(" + makeCtrlHwndArg(objLowerTs, itCtlTs->second)
                              + ", " + idxTs + ")  /* SSTab indexed prop */";
                    return;
                }
            }
        }

        // ---- Pattern A/B: UDT 数组字段 ----
        // 仅当 maExpr.object 的 UDT 类型可推断 (IdentifierExpr / WithMemberExpr / 嵌套
        // MemberAccessExpr). inferUdtTypeOfExpr 找到 UDT C 类型 → 在 udtMembers 中查 member.
        if (maExpr.object) {
            std::string udtCType = inferUdtTypeOfExpr(*maExpr.object);
            if (!udtCType.empty()) {
                const std::string prefix = "vb6_type_";
                if (udtCType.size() > prefix.size()
                    && udtCType.compare(0, prefix.size(), prefix) == 0) {
                    std::string udtName = udtCType.substr(prefix.size());
                    Symbol* udtSym = symTab_.lookupModule(udtName);
                    if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                        for (auto& mi : udtSym->udtMembers) {
                            if (Symbol::toLower(mi.name) == memLower) {
                                if (mi.arraySize > 0) {
                                    // Pattern A: 固定大小数组成员 → obj.member[idx]
                                    emitExpr(*node.callee);  // emits "obj.member"
                                    std::string fieldExpr = std::move(lastExpr_);
                                    emitExpr(*node.positional[0]);
                                    std::string idx = std::move(lastExpr_);
                                    lastExpr_ = fieldExpr + "[" + idx + "]";
                                    handled = true;
                                } else if (mi.isArrayDynamic) {
                                    // Pattern B: 动态数组成员 → VB6_SA_AT(elemType, obj.member, idx)
                                    emitExpr(*node.callee);
                                    std::string fieldExpr = std::move(lastExpr_);
                                    emitExpr(*node.positional[0]);
                                    std::string idx = std::move(lastExpr_);
                                    // Fix 092s: UDT 元素类型成员数组须用 vb6_type_<UDT>
                                    // (mapSaElemCType 对 UserDefinedType 退化为 vb6_VARIANT
                                    // → 元素大小/取址全错, cTlsSocket 4086/4093)
                                    std::string elemCType = (mi.type == Vb6Type::UserDefinedType
                                                             && !mi.typeRefName.empty())
                                        ? ("vb6_type_" + cIdent(mi.typeRefName))
                                        : mapSaElemCType(mi.type);
                                    lastExpr_ = "VB6_SA_AT(" + elemCType + ", "
                                              + fieldExpr + ", " + idx + ")";
                                    handled = true;
                                } else if (mi.type == Vb6Type::Variant) {
                                    // Pattern J: UDT Variant 字段 (持有 SafeArray) →
                                    // vb6_VariantArrayGet(&obj.field, idx)
                                    emitExpr(*node.callee);  // emits "obj.field"
                                    std::string fieldExpr = std::move(lastExpr_);
                                    emitExpr(*node.positional[0]);
                                    // Fix 084o: 索引为 Variant 时转 Long
                                    std::string idx = toLongIfVariant(std::move(lastExpr_), node.positional[0].get());
                                    lastExpr_ = "vb6_VariantArrayGet(&" + fieldExpr + ", " + idx + ")";
                                    handled = true;
                                } else if (mi.type == Vb6Type::Object) {
                                    // Pattern H: UDT void* (Object) 字段 (持有 COM 对象) →
                                    // vb6_ComCall(obj.field, L"Item", args, argc)
                                    emitExpr(*node.callee);
                                    std::string objExpr = std::move(lastExpr_);
                                    std::vector<std::string> packedArgs;
                                    for (size_t i = 0; i < node.positional.size(); i++) {
                                        std::string packFn = comPackExpr(*node.positional[i]);
                                        emitExpr(*node.positional[i]);
                                        { std::string resolved = resolveComMarkerForPack(packFn);
                                          if (!resolved.empty()) lastExpr_ = resolved; }
                                        packedArgs.push_back(packFn + "(" + lastExpr_ + ")");
                                    }
                                    int32_t argc = (int32_t)packedArgs.size();
                                    std::string argsArray = "(void*[]){";
                                    for (int i = 0; i < argc; i++) {
                                        if (i > 0) argsArray += ", ";
                                        argsArray += packedArgs[i];
                                    }
                                    argsArray += "}";
                                    lastExpr_ = "vb6_ComCall(" + objExpr + ", L\"Item\", "
                                              + argsArray + ", " + std::to_string(argc) + ")";
                                    handled = true;
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }

        // ---- Pattern C/D: 类实例字段访问 ----
        if (!handled && maExpr.object) {
            std::string className = inferClassTypeOfExpr(*maExpr.object);
            if (!className.empty()) {
                // resolveClassMemberCall 验证 member 是否为方法/属性:
                // 是 → 跳过 (正常函数调用 fallback 处理); 否 → 数据字段 → 检查 Variant/void*
                std::string resolvedFn = resolveClassMemberCall(className, maExpr.memberName);
                if (resolvedFn.empty()) {
                    // Pattern C: Variant 字段 → vb6_VariantArrayGet(&obj->member, idx)
                    if (classVariantMembers_.count(memLower)
                        || classVariantMembers_.count(memLowerM)) {
                        emitExpr(*node.callee);  // emits "obj->member  /* class var .X field */"
                        std::string fieldExpr = std::move(lastExpr_);
                        emitExpr(*node.positional[0]);
                        // Fix 084o: 索引为 Variant 时转 Long
                        std::string idx = toLongIfVariant(std::move(lastExpr_), node.positional[0].get());
                        lastExpr_ = "vb6_VariantArrayGet(&" + fieldExpr + ", " + idx + ")";
                        handled = true;
                    } else {
                        // Pattern D: void* COM 字段 → vb6_ComCall(obj->member, L"Item", args, argc)
                        bool isVoidPtr = false;
                        if (classVoidFieldMap_) {
                            auto itV = classVoidFieldMap_->find(className);
                            if (itV != classVoidFieldMap_->end()) {
                                if (itV->second.count(memLower)
                                    || itV->second.count(memLowerM)) {
                                    isVoidPtr = true;
                                }
                            }
                        }
                        if (isVoidPtr) {
                            // 模仿 line 1880-1938 COM 默认属性调用 (comPackExpr/resolveComMarkerForPack)
                            emitExpr(*node.callee);  // emits "obj->member  /* ... voidptr */"
                            std::string objExpr = std::move(lastExpr_);
                            std::vector<std::string> packedArgs;
                            for (size_t i = 0; i < node.positional.size(); i++) {
                                std::string packFn = comPackExpr(*node.positional[i]);
                                emitExpr(*node.positional[i]);
                                { std::string resolved = resolveComMarkerForPack(packFn);
                                  if (!resolved.empty()) lastExpr_ = resolved; }
                                packedArgs.push_back(packFn + "(" + lastExpr_ + ")");
                            }
                            int32_t argc = (int32_t)packedArgs.size();
                            std::string argsArray = "(void*[]){";
                            for (int i = 0; i < argc; i++) {
                                if (i > 0) argsArray += ", ";
                                argsArray += packedArgs[i];
                            }
                            argsArray += "}";
                            // 默认成员名 "Item" — VB6 Collection / ADODB.Recordset 等大多数
                            // void* 字段都是用 .Item(idx) 索引的 (DISPID_VALUE 默认成员).
                            lastExpr_ = "vb6_ComCall(" + objExpr + ", L\"Item\", "
                                      + argsArray + ", " + std::to_string(argc) + ")";
                            handled = true;
                        } else if (classTypedFieldMap_) {
                            // Fix 037b: typed (非 void*) 对象字段 — 项目类或 COM 接口
                            auto itT = classTypedFieldMap_->find(className);
                            if (itT != classTypedFieldMap_->end()) {
                                auto itF = itT->second.find(memLower);
                                if (itF == itT->second.end()) itF = itT->second.find(memLowerM);
                                if (itF != itT->second.end()) {
                                    const std::string& fieldType = itF->second;
                                    if (fieldType.compare(0, 4, "COM:") == 0) {
                                        // Pattern K: COM 接口字段 (如 Header As Dictionary)
                                        // → vb6_ComCall((void*)obj->member, L"Item", args, argc)
                                        emitExpr(*node.callee);
                                        std::string objExpr = std::move(lastExpr_);
                                        std::vector<std::string> packedArgs;
                                        for (size_t i = 0; i < node.positional.size(); i++) {
                                            std::string packFn = comPackExpr(*node.positional[i]);
                                            emitExpr(*node.positional[i]);  // Fix: Pattern K 实参丢失 (漏 emit 致 vb6_ComPackXxx() 零实参), 与 Pattern D 对齐

                                            { std::string resolved = resolveComMarkerForPack(packFn);
                                              if (!resolved.empty()) lastExpr_ = resolved; }
                                            packedArgs.push_back(packFn + "(" + lastExpr_ + ")");
                                        }
                                        int32_t argc = (int32_t)packedArgs.size();
                                        std::string argsArray = "(void*[]){";
                                        for (int i = 0; i < argc; i++) {
                                            if (i > 0) argsArray += ", ";
                                            argsArray += packedArgs[i];
                                        }
                                        argsArray += "}";
                                        lastExpr_ = "vb6_ComCall((void*)" + objExpr + ", L\"Item\", "
                                                  + argsArray + ", " + std::to_string(argc) + ")";
                                        handled = true;
                                    } else if (node.positional.size() == 1) {
                                        // Pattern L: 项目类字段 (如 Rows As cCollection)
                                        // → vb6_<Type>_prop_get_Item(obj->member, vb6_VariantFromValue(arg))
                                        std::string itemFn = resolveClassMemberCall(fieldType, "Item");
                                        if (!itemFn.empty()) {
                                            emitExpr(*node.callee);
                                            std::string objExpr = std::move(lastExpr_);
                                            emitExpr(*node.positional[0]);
                                            std::string arg = std::move(lastExpr_);
                                        // 类虚表 (tB, ai/022 B08e-5 第⑦站): 默认属性调用 `m_up(4)` 此前静默发
                                        // `vb6_<Base>_prop_get_Item(接收者, ...)`, 接收者声明成基类而实际持派生实例时会切片回
                                        // 基类的 Item; 显式写 `m_up.Item(4)` 走的是 ①/B08d 那条早已派发的路, 两种写法必须一致。
                                        // 口径同 ①⑤⑪: 派发串与实参里的接收者是同一个 objExpr, 不纯时由 mustDispatch 出诊断,
                                        // 左值求值 (suppressVirtDispatch_) 不改派。
                                        if (!suppressVirtDispatch_) {
                                            const std::string dispatchFnL = virtDispatchCallee(
                                                fieldType, "Item", objExpr, /*mustDispatch=*/true, node.loc);
                                            if (!dispatchFnL.empty()) itemFn = dispatchFnL;
                                        }
                                            lastExpr_ = itemFn + "(" + objExpr + ", vb6_VariantFromValue(" + arg + "))";
                                            lastExprNeedsObjectUnpack_ = true;  // Set 语句需转 void*
                                            handled = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if (handled) return;
    }
