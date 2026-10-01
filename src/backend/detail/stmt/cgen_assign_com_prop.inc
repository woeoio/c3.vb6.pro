// cgen_assign_com_prop.inc - visit(AssignmentStmt&) 函数体片段：参数化属性与 COM 写（P25/P25b / 类内 Property Let / COM SetProp / 返回值解封）
// 由 src/backend/stmt/cgen_assign.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 829~1055 行，逐行未改。

    // ---- P25b: 链式 COM 默认属性索引赋值 ---- (实现提取为
    // tryEmitChainedComWrite, 090ae; AssignmentStmt/SetStmt 共用)
    if (tryEmitChainedComWrite(node.target.get(), node.value.get())) {
        return;
    }

    // P25: 参数化COM属性赋值检测: dic.Item(key) = value → vb6_ComSetPropArg
    if (node.target->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& callTarget = static_cast<IndexOrCallExpr&>(*node.target);
        if (callTarget.callee && callTarget.callee->kind == ASTNodeKind::MemberAccessExpr) {
            auto& maTarget = static_cast<MemberAccessExpr&>(*callTarget.callee);
            // 检查对象是否是COM变量
            if (maTarget.object && maTarget.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objId = static_cast<IdentifierExpr&>(*maTarget.object);
                std::string objLower = objId.name;
                std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
                bool isComVar = knownObjectVars_.count(objLower) || knownVariantVars_.count(objLower);
                // 也检查前期绑定COM变量
                if (!isComVar) {
                    isComVar = knownTypedComVars_.count(objLower) > 0;
                }
                if (isComVar && !callTarget.positional.empty()) {
                    // 生成: vb6_ComSetPropArg(obj, L"Prop", {pack(arg1),...}, argc, pack(value))
                    emitExpr(*maTarget.object);
                    std::string objExpr = lastExpr_;
                    std::string memberName = maTarget.memberName;
                    // 打包索引参数
                    std::vector<std::string> packedIdxArgs;
                    for (size_t i = 0; i < callTarget.positional.size(); i++) {
                        std::string packFn = comPackExpr(*callTarget.positional[i]);
                        emitExpr(*callTarget.positional[i]);
                        { std::string resolved = resolveComMarkerForPack(packFn); if (!resolved.empty()) lastExpr_ = resolved; }
                        packedIdxArgs.push_back(packFn + "(" + lastExpr_ + ")");
                    }
                    int32_t idxArgc = (int32_t)packedIdxArgs.size();
                    std::string idxArgsArray = "(void*[]){";
                    for (int i = 0; i < idxArgc; i++) {
                        if (i > 0) idxArgsArray += ", ";
                        idxArgsArray += packedIdxArgs[i];
                    }
                    idxArgsArray += "}";
                    // 打包值参数并生成调用
                    emitExpr(*node.value);
                    // Fix 160w: RHS 为 WithMemberExpr/COM 成员时解封 (裸对象实参
                    // 送 vb6_ComPack* → C2440) 且避免 marker 泄漏.
                    //
                    // Fix <vbeclipse>: 解封类型必须**由 packer 反推**, 不能恒用
                    // "Variant"。此前固定 resolveComValue("Variant") 得到
                    // vb6_VariantFromComResult(vb6_ComCall(...)) (vb6_VARIANT 结构体),
                    // 紧接着 comPackExpr 又按 inferExprType 选出 BSTR/Double — 载荷与
                    // packer 期望类型不匹配, frmViewSnapshot.frm:106-107
                    //   li.SubItems(1) = .RefId / li.SubItems(2) = .Ratio
                    // → vb6_ComPackBSTR(vb6_VariantFromComResult(...)) C2440
                    //   (无法从 vb6_VARIANT 转换为 const wchar_t* / double, 4 处).
                    // 口径与下方 SetProp 分支的 Fix 110i 一致: packer 决定解封类型。
                    // comPackExpr 必须在 emitExpr 之前求 (它会消费/依赖部分状态),
                    // 故先算 packer 再解封。
                    const std::string packFnVal = comPackExpr(*node.value);
                    if (isComMarker_) {
                        std::string hintVal = "BSTR";
                        if (packFnVal == "vb6_ComPackDouble") hintVal = "Double";
                        else if (packFnVal == "vb6_ComPackInt") hintVal = "Long";
                        else if (packFnVal == "vb6_ComPackObject") hintVal = "Object";
                        else if (packFnVal == "vb6_ComPackValue") hintVal = "Variant";
                        resolveComValue(hintVal);
                    }
                    std::string valExpr = std::move(lastExpr_);
                    c_.emitLine("vb6_ComSetPropArg(" + objExpr + ", L\"" + memberName + "\", " +
                                 idxArgsArray + ", " + std::to_string(idxArgc) + ", " +
                                 packFnVal + "(" + valExpr + "));  /* COM SetPropArg */");
                    return;
                }
            }
        }
    }

    // Fix 152e: With 块内 ListBox/ComboBox 的**可索引属性写** —
    //   .ItemData(i) = v / .Selected(i) = v / .List(i) = s
    // target 是 IndexOrCallExpr{callee=WithMemberExpr(.Prop), args}. 通用路径会
    // 走 prop_get 读 (vb6_GetItemData(tempVar) 单参) 再拼 " = value" → C2198 +
    // C2106 (非左值). 这里按 RTL 写函数 (hwnd, idx, value) 直接生成.
    if (node.target->kind == ASTNodeKind::IndexOrCallExpr
        && !withObjectVars_.empty() && !withObjectInfoStack_.empty()) {
        auto& wcCall = static_cast<IndexOrCallExpr&>(*node.target);
        if (wcCall.callee && wcCall.callee->kind == ASTNodeKind::WithMemberExpr
            && !wcCall.positional.empty() && wcCall.named.empty()) {
            const auto& wcInfo = withObjectInfoStack_.back();
            if (wcInfo.kind == WithObjKind::FormControl
                && (wcInfo.ctrlType == FrmControlType::ListBox
                    || wcInfo.ctrlType == FrmControlType::ComboBox)) {
                auto& wcMem = static_cast<WithMemberExpr&>(*wcCall.callee);
                const std::string wcTempVar = withObjectVars_.back();
                std::string wcProp = Symbol::toLower(wcMem.memberName);
                if (wcProp == "itemdata" || wcProp == "selected" || wcProp == "list") {
                    emitExpr(*wcCall.positional[0]);
                    std::string wcIdx = std::move(lastExpr_);
                    emitExpr(*node.value);
                    // Fix 160w: RHS 为 WithMemberExpr/COM 成员时解封
                    // (.List(i)= → BSTR, .ItemData/.Selected(i)= → Long).
                    if (isComMarker_) resolveComValue(wcProp == "list" ? "BSTR" : "Long");
                    std::string wcVal = std::move(lastExpr_);
                    if (wcProp == "itemdata") {
                        c_.emitLine("vb6_SetItemData((void*)" + wcTempVar + ", " + wcIdx
                                    + ", (int32_t)(" + wcVal + "));  /* With .ItemData(i)= */");
                    } else if (wcProp == "selected") {
                        c_.emitLine("vb6_SetSelected((void*)" + wcTempVar + ", " + wcIdx
                                    + ", " + wcVal + ");  /* With .Selected(i)= */");
                    } else {
                        c_.emitLine("vb6_SetListItem((void*)" + wcTempVar + ", " + wcIdx
                                    + ", " + wrapToBSTR(wcVal, *node.value)
                                    + ");  /* With .List(i)= */");
                    }
                    return;
                }
            }
        }
    }

    // Fix 092u: 当前类内**参数化** Property Let/Set 赋值 —
    //   Extend(Array(A, C)) = Split(...)     (cToolsArray.cls 97, Sub test 内隐式 me)
    // 此前落到通用路径: emitExpr(target) 生成裸名 + 补的 value 占位
    // "Extend(&_arr_1, &(vb6_VARIANT){0})", 再拼 " = <RHS>" → C2106 ("=" 左侧必须
    // 是左值, ToolsArray.c 130). 正确形态是按**写方向**形参表生成
    //   vb6_cToolsArray_prop_let_Extend(me, &Vars, &Value);
    if (node.target->kind == ASTNodeKind::IndexOrCallExpr && isClassModule_) {
        auto& plCall092u = static_cast<IndexOrCallExpr&>(*node.target);
        if (plCall092u.callee && plCall092u.callee->kind == ASTNodeKind::IdentifierExpr
            && !plCall092u.positional.empty() && plCall092u.named.empty()) {
            auto& plId092u = static_cast<IdentifierExpr&>(*plCall092u.callee);
            const std::string plLower092u = Symbol::toLower(plId092u.name);
            // 局部变量 / 当前过程名同名 → 不是属性 (同 642 段规则)
            bool plSkip092u = knownLocalVars_.count(plLower092u) > 0;
            if (!plSkip092u && currentProc_) {
                plSkip092u = (Symbol::toLower(currentProc_->name) == plLower092u);
            }
            if (!plSkip092u) {
                std::vector<ParameterInfo> wp092u;
                bool isSet092u = false;
                bool found092u = findClassMemberWriteParams(moduleName_, plId092u.name,
                                                            false, wp092u);
                Symbol* plSym092u = nullptr;
                if (found092u) {
                    plSym092u = symTab_.lookupModuleByKind(plId092u.name,
                                                           SymbolKind::PropertyLet);
                } else {
                    found092u = findClassMemberWriteParams(moduleName_, plId092u.name,
                                                           true, wp092u);
                    if (found092u) {
                        isSet092u = true;
                        plSym092u = symTab_.lookupModuleByKind(plId092u.name,
                                                               SymbolKind::PropertySet);
                    }
                }
                // 形参数须恰好比括号实参多 1 (末参是 value); 且**不含 Optional 形参** —
                // Optional 在 C 侧另有存在标志参数 (cAsyncSocket 2611:
                // Property Let ThunkPrivateData(pThunk, Optional ByVal Index, ByVal lValue)
                // 的 C 签名是 (me, IUnknown**, int32_t, int32_t, int)), 本路径只按符号表
                // 拼参会 C2198 → 这类仍交给原有 prop_get_ 重写路径 (Pattern C/D2) 处理.
                bool hasOptional092u = false;
                for (const auto& pOpt092u : wp092u) {
                    if (pOpt092u.isOptional) { hasOptional092u = true; break; }
                }
                if (found092u && plSym092u && !hasOptional092u
                    && wp092u.size() == plCall092u.positional.size() + 1) {
                    const std::string tmp092u = "_pv" + std::to_string(tempCounter_++) + "_";
                    std::string args092u;
                    int slot092u = 0;
                    auto pushArg092u = [&](Expr* argNode, const ParameterInfo& pi) {
                        emitExpr(*argNode);
                        std::string argExpr = std::move(lastExpr_);
                        std::string one092u;
                        if (pi.type == Vb6Type::Variant) {
                            if (pi.isByVal) {
                                // 装箱走 boxToVariant 这一权威 (按 VB 类型选档)
                                one092u = boxToVariant(argNode, argExpr);
                            } else {
                                // ByRef Variant 需可寻址: 先落到栈变量再取址
                                // (复合字面量 {Variant值} 触发 C2440)
                                std::string tv = tmp092u + std::to_string(slot092u++);
                                c_.emitLine("vb6_VARIANT " + tv + " = vb6_VariantFromValue("
                                            + argExpr + ");");
                                one092u = "&" + tv;
                            }
                        } else if (pi.type == Vb6Type::String && cExprIsVariant(argExpr)) {
                            one092u = wrapToBSTR(argExpr, *argNode);
                        } else if (!pi.isByVal) {
                            // ByRef 非 Variant 形参: 简单标识符直接取址; 函数结果 /
                            // 字段链等非左值先落栈变量再取址 — 否则 &<非左值> C2102
                            // (cAsyncSocket 2395: &vb6_BSTR_Concat(...)).
                            bool simple092u = !argExpr.empty()
                                && ((argExpr[0] >= 'a' && argExpr[0] <= 'z')
                                    || (argExpr[0] >= 'A' && argExpr[0] <= 'Z')
                                    || argExpr[0] == '_')
                                && argExpr.find('(') == std::string::npos
                                && argExpr.find("->") == std::string::npos
                                && argExpr.find('.') == std::string::npos
                                && argExpr.find(' ') == std::string::npos;
                            if (simple092u) {
                                one092u = "&" + argExpr;
                            } else {
                                std::string tv = tmp092u + std::to_string(slot092u++);
                                c_.emitLine(mapType(pi.type) + " " + tv + " = " + argExpr + ";");
                                one092u = "&" + tv;
                            }
                        } else {
                            one092u = argExpr;
                        }
                        if (!args092u.empty()) args092u += ", ";
                        args092u += one092u;
                    };
                    for (size_t ai = 0; ai < plCall092u.positional.size(); ai++) {
                        pushArg092u(plCall092u.positional[ai].get(), wp092u[ai]);
                    }
                    pushArg092u(node.value.get(), wp092u.back());
                    c_.emitLine(cProcName((isSet092u ? "prop_set_" : "prop_let_")
                                              + plId092u.name,
                                          plSym092u->access, moduleName_)
                                + "((void*)me, " + args092u
                                + ");  /* Property Let/Set (param, own class) */");
                    return;
                }
            }
        }
    }

    // P6.4+: 类模块对象 (默认实例或类变量) 的 **参数化** Property Let/Set 赋值:
    //   cTT.EnableTooltips(cmdButtonInPic) = Not cTT.EnableTooltips(cmdButtonInPic)
    // target 是 IndexOrCallExpr{callee=MemberAccess(obj, Prop), args}. 通用路径会
    // 先按 prop_get 生成读调用再拼 " = value" → C2106 (非左值). 这里按**写方向**
    // 形参表直接生成 vb6_cX_prop_let_Prop(<obj>, <形参...>, <value>, <has标志...>).
    if (node.target->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& plCallDef = static_cast<IndexOrCallExpr&>(*node.target);
        if (plCallDef.callee && plCallDef.callee->kind == ASTNodeKind::MemberAccessExpr
            && !plCallDef.positional.empty()) {
            auto& plMaDef = static_cast<MemberAccessExpr&>(*plCallDef.callee);
            if (plMaDef.object && plMaDef.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& plObjDef = static_cast<IdentifierExpr&>(*plMaDef.object);
                const std::string plObjLowerDef = Symbol::toLower(plObjDef.name);
                // 类解析: 默认实例 (裸类名) → 注册 + 访问器; 已知类变量 → 类名
                std::string plClsDef = registerDefaultInstanceClass(plObjDef.name);
                if (plClsDef.empty()) {
                    auto itClsDef = knownClassVars_.find(plObjLowerDef);
                    if (itClsDef != knownClassVars_.end()) plClsDef = itClsDef->second;
                }
                if (!plClsDef.empty()) {
                    // 写方向形参表 (Let 优先, Set 兜底)
                    std::vector<ParameterInfo> wpDef;
                    bool isSetDef = false;
                    Symbol* plSymDef = nullptr;
                    if (findClassMemberWriteParams(plClsDef, plMaDef.memberName, false, wpDef)) {
                        plSymDef = symTab_.lookupModuleByKind(plMaDef.memberName, SymbolKind::PropertyLet);
                    } else if (findClassMemberWriteParams(plClsDef, plMaDef.memberName, true, wpDef)) {
                        isSetDef = true;
                        plSymDef = symTab_.lookupModuleByKind(plMaDef.memberName, SymbolKind::PropertySet);
                    }
                    bool plOwnsDef = false;
                    if (plSymDef) {
                        std::string propModDef = plSymDef->isExternal ? plSymDef->sourceModule : plClsDef;
                        std::string propModLowerDef = Symbol::toLower(propModDef);
                        std::string clsLowerDef = Symbol::toLower(plClsDef);
                        plOwnsDef = (propModLowerDef == clsLowerDef);
                    }
                    if (plOwnsDef && wpDef.size() == plCallDef.positional.size() + 1) {
                        // 函数名 (按类内声明拼写规范化)
                        std::string srcModDef = plSymDef->isExternal ? plSymDef->sourceModule : plClsDef;
                        std::string fnDef = cProcName(
                            (isSetDef ? "prop_set_" : "prop_let_")
                                + canonicalClassMemberName(srcModDef, plMaDef.memberName),
                            plSymDef->access, srcModDef);
                        std::string objDef;
                        // Fix 152c: 放置于窗体的 UserControl 实例 (如 VBFlexGrid1)
                        // 的**参数化属性写** (VBFlexGrid1.TextMatrix(i,11,..) = "...").
                        // 裸 emitExpr(*plMaDef.object) 只输出控件名 (C 里无声明) →
                        // C2065。this 应取宿主窗口实例, 与 Fix 133y / 读取路径一致。
                        auto itUCdef = knownUserControlCtrlVars_.find(plObjLowerDef);
                        if (itUCdef != knownUserControlCtrlVars_.end()) {
                            auto itCtrlDef = knownFormControls_.find(plObjLowerDef);
                            std::string hwndArgDef = (itCtrlDef != knownFormControls_.end())
                                ? makeCtrlHwndArg(plObjLowerDef, itCtrlDef->second)
                                : ("vb6_hwnd_" + cIdent(plObjDef.name));
                            objDef = "(vb6_cls_" + cIdent(itUCdef->second)
                                   + "*)vb6_UC_InstanceOf(" + hwndArgDef + ")";
                        } else if (!defaultInstanceClassName(plObjDef.name).empty()) {
                            objDef = "vb6_cls_" + cIdent(plClsDef) + "_Default()";
                        } else {
                            emitExpr(*plMaDef.object);
                            objDef = std::move(lastExpr_);
                        }
                        // 逐形参打包 (含末参 value), 与 Fix 092u pushArg092u 规则一致
                        const std::string tmpDef = "_pv" + std::to_string(tempCounter_++) + "_";
                        std::string argsDef = objDef;
                        int slotDef = 0;
                        auto pushArgDef = [&](Expr* argNode, const ParameterInfo& pi) {
                            // Fix 143: 形参是对象类型 (Object/Control) 且实参是窗体控件
                            // 标识符时, 抑制默认属性解析, 传控件句柄引用而非默认属性值.
                            // 否则 ByRef Object 实参 (如 cTT.EnableTooltips(cmdButtonInPic))
                            // 会取到 Caption BSTR 指针, 被调方 *objControl 读到字符串指针
                            // 当对象 → 访问 hWnd 崩溃 (0xC000041D). 与 arg_emit.inc 的
                            // Fix 090q 同源.
                            bool savedSuppressDef143 = suppressDefaultProp_;
                            if (static_cast<Vb6Type>(static_cast<uint16_t>(pi.type)
                                    & ~static_cast<uint16_t>(Vb6Type::Array)) == Vb6Type::Object
                                && argNode && argNode->kind == ASTNodeKind::IdentifierExpr) {
                                std::string idLower143 = Symbol::toLower(
                                    static_cast<IdentifierExpr&>(*argNode).name);
                                if (knownFormControls_.count(idLower143)
                                    || knownWithEventsCtrlVars_.count(idLower143)) {
                                    suppressDefaultProp_ = true;
                                }
                            }
                            emitExpr(*argNode);
                            suppressDefaultProp_ = savedSuppressDef143;
                            std::string argExpr = std::move(lastExpr_);
                            std::string oneDef;
                            if (pi.type == Vb6Type::Variant) {
                                if (pi.isByVal) {
                                    // 同上: 装箱走 boxToVariant 权威
                                    oneDef = boxToVariant(argNode, argExpr);
                                } else {
                                    std::string tv = tmpDef + std::to_string(slotDef++);
                                    c_.emitLine("vb6_VARIANT " + tv + " = "
                                                + boxToVariant(argNode, argExpr) + ";");
                                    oneDef = "&" + tv;
                                }
                            } else if (pi.type == Vb6Type::String && cExprIsVariant(argExpr)) {
                                oneDef = wrapToBSTR(argExpr, *argNode);
                            } else if (pi.type == Vb6Type::String && pi.isByVal && argNode
                                       && isScalarImplicitStr159(inferExprType(*argNode))) {
                                // Fix 159-B: ByVal String 形参收到**日期/数值**实参 —
                                // VB6 按 CStr 隐式转换, 此前只在实参是 vb6_VARIANT 时才
                                // wrapToBSTR, 标量/日期实参裸拼 → C2440 "double→BSTR".
                                //   UserEditingForm.frm: `TextMatrix(i,3) = DateAdd("d",i,Now())`
                                //   MainForm.frm:        `EditText = ComboCalendarValue`
                                // 仅收标量族 (Date/Double/Single/Long/Integer/Byte/Boolean/
                                // Currency), Object/Variant/String 一律不介入 (走原路径).
                                oneDef = wrapToBSTR(argExpr, *argNode);
                            } else if (!pi.isByVal) {
                                bool simpleDef = !argExpr.empty()
                                    && ((argExpr[0] >= 'a' && argExpr[0] <= 'z')
                                        || (argExpr[0] >= 'A' && argExpr[0] <= 'Z')
                                        || argExpr[0] == '_')
                                    && argExpr.find('(') == std::string::npos
                                    && argExpr.find("->") == std::string::npos
                                    && argExpr.find('.') == std::string::npos
                                    && argExpr.find(' ') == std::string::npos;
                                if (simpleDef) {
                                    oneDef = "&" + argExpr;
                                } else {
                                    std::string tv = tmpDef + std::to_string(slotDef++);
                                    c_.emitLine(mapType(pi.type) + " " + tv + " = " + argExpr + ";");
                                    oneDef = "&" + tv;
                                }
                            } else {
                                oneDef = argExpr;
                            }
                            argsDef += ", " + oneDef;
                        };
                        for (size_t ai = 0; ai < plCallDef.positional.size(); ai++) {
                            pushArgDef(plCallDef.positional[ai].get(), wpDef[ai]);
                        }
                        pushArgDef(node.value.get(), wpDef.back());
                        // _has_ 标志: 可选形参逐一追加 (出现在括号实参中 → 1, 否则 0)
                        for (size_t oi = 0; oi < wpDef.size(); oi++) {
                            const auto& poDef = wpDef[oi];
                            if (!poDef.isOptional) continue;
                            int hasFlagDef = (oi < plCallDef.positional.size()) ? 1 : 0;
                            if (oi + 1 == wpDef.size()) hasFlagDef = 1;  // value 形参必提供
                            argsDef += ", " + std::to_string(hasFlagDef);
                        }
                        c_.emitLine(fnDef + "(" + argsDef + ");  /* class prop_let_ (default-instance/class obj) */");
                        return;
                    }
                }
            }
        }
    }

    // Fix 170: 目标是 VB6 **整体数组** (`A() = expr`, 空括号) —— emit 期间置
    // wholeArrayLvalue_, 让 IndexOrCallExpr 发数组载体 (vb6_SafeArray1D*) 而不是
    // VB6_SA_AT(elem, A, 0)。值语义侧用 tgtIsWholeArray170 决定右侧是否需要深拷贝。
    bool tgtIsWholeArray170 = isWholeArrayRef(node.target.get());
    wholeArrayLvalue_ = tgtIsWholeArray170;
    // 左值求值期间关掉虚表派发 (ai/022 B08e-4 站点⑭: Pattern C/D2 要对这段文本做
    // `prop_get_` -> `prop_let_` 的字符串级改写, 派发表达式会被换成表里不存在的槽字段)
    const bool savedVirtLhs = suppressVirtDispatch_;
    suppressVirtDispatch_ = true;
    emitExpr(*node.target);
    suppressVirtDispatch_ = savedVirtLhs;
    std::string target = std::move(lastExpr_);
    wholeArrayLvalue_ = false;

    // COM属性赋值检测 (P6.2): obj.Property = value → vb6_ComSetProp(obj, L"Property", pack(value))
    if (isComMarker_) {
        isComMarker_ = false;
        std::string objExpr = std::move(comObjExpr_);
        std::string memberName = std::move(comMemberName_);
        // P20-40: StatusBar 原生复刻的**链式写** `StatusBar1.Panels(i).Prop = v`。
        // 必须在 vb6_ComSetProp 之前截住 —— 走 COM 就是拿 HWND 当 IDispatch 去
        // PutProp, 链接都过、运行静默无效 (读那边早已改道, 写这边漏了就只错一半)。
        {
            std::string sbMem = memberName;
            std::transform(sbMem.begin(), sbMem.end(), sbMem.begin(), ::tolower);
            std::string sbHwnd, sbIdx;
            std::string sbFn = statusBarPanelWriteFn(objExpr, sbMem, &sbHwnd, &sbIdx);
            if (!sbFn.empty()) {
                emitExpr(*node.value);
                std::string sbVal = std::move(lastExpr_);
                c_.emitLine(sbFn + "((void*)" + sbHwnd + ", " + sbIdx + ", " + sbVal + ");"
                            "  /* StatusBar Panels write */");
                return;
            }
        }
        emitExpr(*node.value);
        std::string valExpr = std::move(lastExpr_);
        std::string packFn = comPackExpr(*node.value);
        // Fix 110i: 右侧本身也是 COM 属性读 (obj.Prop = other.Prop) 时, emitExpr 只
        // 留下对象表达式, 成员名还挂在 comMemberName_ 上 — 不消费就整体丢失
        // (Value = oPC.Min → vb6_ComPackDouble(me->oPC)). 按 packer 反推解封类型后
        // 调用 resolveComValue 生成 vb6_ComGetXxxProp (Charts 2020
        // ppProgressCircular.Timer1_Timer: oPC.Value = oPC.Max / oPC.Min).
        if (isComMarker_) {
            std::string hint = "BSTR";
            if (packFn == "vb6_ComPackDouble") hint = "Double";
            else if (packFn == "vb6_ComPackInt") hint = "Long";
            else if (packFn == "vb6_ComPackBSTR") hint = "BSTR";
            else if (packFn == "vb6_ComPackObject") hint = "Object";
            else if (packFn == "vb6_ComPackValue") hint = "Variant";
            resolveComValue(hint);
            valExpr = std::move(lastExpr_);
        }
        c_.emitLine("vb6_ComSetProp(" + objExpr + ", L\"" + memberName + "\", " +
                     packFn + "(" + valExpr + "));  /* COM SetProp */");
        return;
    }

    emitExpr(*node.value);
    // COM属性值: 如果右侧是COM属性, 根据目标变量类型解析为适当C类型
    if (isComMarker_) {
        // 推断目标变量类型, 用于COM值解封
        std::string unpackHint;
        if (node.target->kind == ASTNodeKind::IdentifierExpr) {
            auto& idExpr = static_cast<IdentifierExpr&>(*node.target);
            std::string lower = idExpr.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            // 从已知变量集合推断类型
            if (knownLongPtrVars_.count(lower)) unpackHint = "LongPtr";
            else if (knownLongVars_.count(lower)) unpackHint = "Long";
            else if (knownDoubleVars_.count(lower)) unpackHint = "Double";
            else if (knownObjectVars_.count(lower)) unpackHint = "Object";
            else if (knownBstrVars_.count(lower)) unpackHint = "BSTR";
            else if (knownVariantVars_.count(lower)) unpackHint = "Variant";
        }
        // Fix 110d: 目标是数组/UDT 字段元素 (VB6_SA_AT(<C类型>, ...)) 时,
        // 从生成的目标表达式文本推断元素 C 类型 — 否则裸 vb6_ComCall (void*)
        // 会直接赋给 float/int32_t/BSTR 元素 → C2440.
        // 例: Values(i - 1) = cValues.Item(i) (ucTreeMaps.ctl)
        //   → VB6_SA_AT(float, _vb6_with_0->Values, ...) = vb6_ComCall(...) C2440
        // Fix 110r: 前缀长度修正 — "VB6_SA_AT(" 是 10 字符 (此前误写 11),
        // 导致本块实际从未命中 (out-v13 ucTreeMaps.c 326/336 残留 C2440).
        if (unpackHint.empty() && target.compare(0, 10, "VB6_SA_AT(") == 0) {
            const size_t comma = target.find(',');
            if (comma != std::string::npos && comma > 10) {
                std::string elemC = target.substr(10, comma - 10);
                while (!elemC.empty() && elemC.back() == ' ') elemC.pop_back();
                if (elemC == "float" || elemC == "double") unpackHint = "Double";
                else if (elemC == "BSTR") unpackHint = "BSTR";
                else if (elemC == "int16_t" || elemC == "int32_t" || elemC == "uint8_t"
                         || elemC == "LONG") unpackHint = "Long";
            }
        }
        resolveComValue(unpackHint);
    }
    std::string value = std::move(lastExpr_);

    // Fix 092w: 目标是否为已知 Byte 数组变量 (vb6_SafeArray1D*) — 赋值右侧为
    // StrConv(...)/字符串构造字节数组内容时需改写 helper.
    bool tgtIsByteArray092w = false;
    {
        std::string tname = target;
        if (tname.compare(0, 4, "me->") == 0) tname = tname.substr(4);
        if (tname.size() > 4 && tname[0] == '(' && tname[1] == '*'
            && tname.back() == ')') {
            tname = tname.substr(2, tname.size() - 3);
        }
        std::string tlower = tname;
        std::transform(tlower.begin(), tlower.end(), tlower.begin(), ::tolower);
        // Fix 140: 类成员 Byte() 数组 (me->m_Caption) 亦须走 Byte 数组改写路径 —
        // knownByteArrayVars_ 已在过程入口 copy classByteArrayMembers_, 此处再显式
        // 兜底判定一次 (防止类字段在非类过程上下文被 clear 后漏判).
        if (knownByteArrayVars_.count(tlower) || classByteArrayMembers_.count(tlower)) {
            tgtIsByteArray092w = true;
        }
    }

    // P6.12: COM方法调用返回值类型化解封
    // vb6_ComCall 返回 VARIANT* (void*), 赋值给typed变量时需用对应解封函数
    // vb6_ComCall(obj, L"Method", args, argc) → vb6_ComCallInt/Double/BSTR/Object(...)
    if (value.find("vb6_ComCall(") == 0) {
        std::string callArgs = value.substr(strlen("vb6_ComCall"));
        // 推断目标变量类型
        std::string comCallType;
        if (node.target->kind == ASTNodeKind::IdentifierExpr) {
            auto& idExpr = static_cast<IdentifierExpr&>(*node.target);
            std::string lower = idExpr.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (knownLongVars_.count(lower)) comCallType = "Int";
            else if (knownDoubleVars_.count(lower)) comCallType = "Double";
            else if (knownBstrVars_.count(lower)) comCallType = "BSTR";
            else if (knownObjectVars_.count(lower)) comCallType = "Object";
        }
        // Fix 110h: 目标是数组/UDT 字段元素 (VB6_SA_AT(<C类型>, ...)) 时, 从目标
        // 表达式文本推断元素 C 类型 —— 否则裸 vb6_ComCall (void*) 直接赋给
        // float/BSTR 元素 → C2440 (ucTreeMaps.ctl 415-418:
        //   .Values(i - 1) = cValues(i) → VB6_SA_AT(float, ...) = vb6_ComCall(...)).
        if (comCallType.empty() && target.compare(0, 10, "VB6_SA_AT(") == 0) {
            const size_t comma = target.find(',');
            if (comma != std::string::npos && comma > 10) {
                std::string elemC = target.substr(10, comma - 10);
                while (!elemC.empty() && elemC.back() == ' ') elemC.pop_back();
                if (elemC == "float" || elemC == "double") comCallType = "Double";
                else if (elemC == "BSTR") comCallType = "BSTR";
                else if (elemC == "int16_t" || elemC == "int32_t" || elemC == "uint8_t"
                         || elemC == "LONG" || elemC == "int64_t") comCallType = "Int";
            }
        }
        // Fix 140: 目标为 Byte() 数组 (类成员 m_Caption / me->m_Xxx) — COM 结果
        // 需解封为字节数组: BSTR → vb6_StringToByteArray, SAFEARRAY → parray.
        // 场景: LabelPlus.ctl ReadProperties 里
        //   m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        // PropertyBag 存的是字符串, 旧路径把裸 void* (VARIANT*) 直接赋给
        // vb6_SafeArray1D* 字段 → UBound/StrPtr 读不到 → 右侧 KPI 卡片空白.
        if (comCallType.empty() && tgtIsByteArray092w) {
            comCallType = "ByteArray";
        }
        if (!comCallType.empty()) {
            value = "vb6_ComCall" + comCallType + callArgs;
        }
    }

    // Fix 010r-15+010r-16 (Pattern A/C/D2/F): LHS 是非左值 COM 调用或 Property Get,
    // 改写为对应的 COM SetProp/SetPropArg 或 prop_let_/prop_set_ 调用.
    // 覆盖模式:
    //   A: vb6_ComCall(obj, L"Item", args, n) = value → vb6_ComSetPropArg(...)
    //   F: vb6_ComGetStringProp(obj, L"Prop") = value → vb6_ComSetProp(...)
    //   C/D2: vb6_X_prop_get_Y(args) = value → vb6_X_prop_let_Y(args, value)
    if (tryRewriteCOMLvalue(target, value, node.value.get(), /*isSet=*/false)) {
        return;
    }
