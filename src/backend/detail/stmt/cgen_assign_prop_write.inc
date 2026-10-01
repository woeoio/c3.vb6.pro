// cgen_assign_prop_write.inc - visit(AssignmentStmt&) 函数体片段：属性写入分派（类 Property Let / 默认属性 / With 块 / 链式 COM）
// 由 src/backend/stmt/cgen_assign.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 406~828 行，逐行未改。

    // P6.7: 检测类Property Let赋值: obj.Prop = value → vb6_prop_let_Prop(obj, value)
    // 在emitExpr左侧前, 先检查target是否为MemberAccessExpr且成员是Property
    if (node.target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maExpr = static_cast<MemberAccessExpr&>(*node.target);
        // 查找Property Let符号
        auto* propLetSym = symTab_.lookupModuleByKind(maExpr.memberName, SymbolKind::PropertyLet);
        if (!propLetSym) {
            // 也尝试Property Set
            propLetSym = symTab_.lookupModuleByKind(maExpr.memberName, SymbolKind::PropertySet);
        }
        if (propLetSym) {
            // 检查对象是否是类实例变量
            if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objId = static_cast<IdentifierExpr&>(*maExpr.object);
                std::string objLower = objId.name;
                std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
                std::string objClass;
                auto itClassVar = knownClassVars_.find(objLower);
                if (itClassVar != knownClassVars_.end()) {
                    objClass = itClassVar->second;
                } else {
                    // Fix 084g: 局部变量/参数类对象 (如 Dim Response As cHttpServerResponse)
                    // 不在 knownClassVars_ 中, 通过符号表推断类名
                    objClass = inferClassTypeOfExpr(*maExpr.object);
                }
                // Fix 084g-2: 外部注入的跨类属性符号 (Friend/Public Property) —
                // lookupModuleByKind 已确认该成员是某类的属性, 即使对象类型无法
                // 从符号表推断 (cgen阶段访问不到过程作用域的局部变量类型), 也
                // 生成属性调用; sourceModule 由外部符号提供.
                // Fix 084j: UDT 变量的成员是结构体字段而非类属性 — lookupModuleByKind
                // 会全局命中同名类属性 (如 dcb As MODBUS_DCB 的 .BaudRate 命中
                // cModbusSlave.BaudRate; m_uCtx 的 .LastError 命中 cZipArchive.LastError),
                // 误生成 vb6_cX_prop_let_Y(udtVar, ...) → C2440. 此类对象必须跳过,
                // 交给 visit(MemberAccessExpr) 的 Fix 031 UDT 字段路径生成 obj.field.
                // Fix 084y: 对象类与属性所属类不一致时, 该属性是其他类的同名成员,
                // 不属于当前对象 — lookupModuleByKind 全局命中, 需校验类归属.
                // 例: Set oCallback.Socket = New cTlsReMaster 中 oCallback 是
                // cClientCallback, Socket 是数据字段, 但全局命中 cTlsSocket.Socket
                // PropertySet → 原生成 vb6_cTlsSocket_prop_set_Socket(oCallback,..)
                // (C2065 未声明 + C2440 参数类型错误). 必须跳过属性路径, 交给
                // visit(MemberAccessExpr) 生成 oCallback->Socket = ... 字段赋值.
                bool objClassMatchesProp = true;
                // Fix 093a: 属性符号与对象类必须同类. 非 external 符号 (属性定义于
                // 本模块) 的"所属类"即当前模块名 — 此前只在 isExternal 时校验, 导致
                // `HttpSvr.MaxCacheFileSize = v` (MaxCacheFileSize 在本类 cHttpServer
                // 是 Property, 而 HttpSvr 是 cHttpServerSvr, 该名在后者是公有字段)
                // 误发 vb6_cHttpServerSvr_prop_let_MaxCacheFileSize → LNK2019
                // (cSSEClient.RequestTimeOut / cHttpServerSvr.CacheTTLSeconds 同类).
                if (!objClass.empty()) {
                    std::string propModLower = propLetSym->isExternal
                        ? propLetSym->sourceModule
                        : (isClassModule_ ? moduleName_ : std::string());
                    std::transform(propModLower.begin(), propModLower.end(), propModLower.begin(), ::tolower);
                    std::string objClassLower = objClass;
                    std::transform(objClassLower.begin(), objClassLower.end(), objClassLower.begin(), ::tolower);
                    objClassMatchesProp = (propModLower == objClassLower);
                }
                // Fix 092b: objClass 无法推断且对象是 Object/COM 后期绑定变量时,
                // 该属性写应走 COM 路径 (vb6_ComSetProp / tryEmitChainedComWrite) —
                // 否则 lookupModuleByKind 全局命中他类同名 Property Let, 误生成
                // 类属性调用 (cToolsList 25: Rs As Object 的 COM 属性写
                // "Rs.Filter = ..." → vb6_cDialog_prop_let_Filter(Rs, ...),
                // C2440 "vb6_VARIANT→BSTR" 并丢失 COM 后期绑定语义).
                // Fix 084g-2 的"objClass 空 + isExternal 即生成"过宽, 此处收窄.
                bool objIsComVar092b = false;
                if (objClass.empty()) {
                    objIsComVar092b = knownObjectVars_.count(objLower) > 0
                                      || knownTypedComVars_.count(objLower) > 0;
                }
                // Task #43: 对象是窗体控件时, member 是 OCX/控件属性而非跨类
                // PropertySet — lookupModuleByKind 全局查找会误命中他类同名属性:
                //   SSTabEx1.Font = iDlg.Font
                //     → vb6_cDlg_prop_set_Font(SSTabEx1, ...)  (接收者 C2065 裸名 +
                //        属性语义整体错位)
                // 控件自有 UserControl 类的属性 (VBFlexGrid1.CellPicture, 符号
                // sourceModule == knownUserControlCtrlVars_ 类名) 仍走类路径;
                // 其余控件属性交给下方 P6.2 COM SetProp — LHS 经 Fix 023e/089d
                // 留 isComMarker_ (comObjExpr_ = 控件 hwnd) → vb6_ComSetProp。
                bool formCtrlNotOwned43 = false;
                if (knownFormControls_.count(objLower)) {
                    formCtrlNotOwned43 = true;
                    auto itUC43 = knownUserControlCtrlVars_.find(objLower);
                    if (itUC43 != knownUserControlCtrlVars_.end() && propLetSym->isExternal
                        && Symbol::toLower(propLetSym->sourceModule)
                           == Symbol::toLower(itUC43->second)) {
                        formCtrlNotOwned43 = false;  // 控件自有 UserControl 类属性
                    }
                }
                if (!formCtrlNotOwned43 && !objIsComVar092b && !knownUdtVars_.count(objLower) && objClassMatchesProp &&
                    (!objClass.empty() || propLetSym->isExternal)) {
                    // 生成Property Let调用: vb6_prop_let_Name(obj, value)
                    std::string prefix = (propLetSym->kind == SymbolKind::PropertySet) ? "prop_set_" : "prop_let_";
                    // Fix 010r-10: 使用map中的类名作为sourceModule
                    std::string sourceModule = propLetSym->isExternal ? propLetSym->sourceModule : objClass;
                    // Fix 093a: 成员名按类内声明拼写规范化 (VB6 大小写不敏感),
                    // 避免调用点拼写与类定义不一致 → LNK2019.
                    std::string funcName = cProcName(
                        prefix + canonicalClassMemberName(sourceModule, maExpr.memberName),
                        propLetSym->access, sourceModule);
                    emitExpr(*maExpr.object);
                    std::string objExpr = std::move(lastExpr_);
                    // Fix 133y: 目标是工程内 UserControl 子控件 (czUI:czLabel3.Caption,
                    // ucChartBar1.ChartStyle 等)。对象在 knownUserControlCtrlVars_
                    // (Fix 112 已登记), 裸 emitExpr 输出控件名标识符 (无声明) →
                    // C2065。this 应取宿主窗口的 UserControl 实例 (与读取路径
                    // cgen_expr_member_form_builtin.inc 一致)。
                    auto itUC133y = knownUserControlCtrlVars_.find(objLower);
                    if (itUC133y != knownUserControlCtrlVars_.end()) {
                        auto itCtrl133y = knownFormControls_.find(objLower);
                        std::string hwndArg133y = (itCtrl133y != knownFormControls_.end())
                            ? makeCtrlHwndArg(objLower, itCtrl133y->second)
                            : ("vb6_hwnd_" + cIdent(objId.name));
                        objExpr = "(vb6_cls_" + cIdent(itUC133y->second)
                                + "*)vb6_UC_InstanceOf(" + hwndArg133y + ")";
                    }
                    emitExpr(*node.value);
                    // Fix 160w: RHS 是 WithMemberExpr 或早期绑定 COM 成员访问
                    // (`VBFlexGrid1.CellToolTipText = .Result`, `VBFlexGrid1.CellFontSize
                    // = PropCellFont.Size`) 时, With/成员表达式路径只把**对象表达式**
                    // 留在 lastExpr_ (bare `_vb6_with_N` / `PropCellFont`), 成员名挂在
                    // comObjExpr_/comMemberName_ 上且 isComMarker_=true. 此处必须按
                    // 目标 PropertyLet/Set 末形参类型立即解封, 否则:
                    //  (a) 裸对象作为值实参 → C2440 (vb6_ComIface_Font*→float,
                    //       MainForm.c:346 CellFontSize = PropCellFont.Size);
                    //  (b) 标记残留, 下一过程 BinaryExpr 误消费 → C2065
                    //       (_vb6_with_14/_vb6_with_15). 同 assign_com_prop Fix 110i.
                    if (isComMarker_) {
                        std::string hint160w = "BSTR";  // String 属性默认解 BSTR
                        if (propLetSym && !propLetSym->params.empty()) {
                            const Vb6Type pt160w = propLetSym->params.back().type;
                            if (pt160w == Vb6Type::Double || pt160w == Vb6Type::Single
                                || pt160w == Vb6Type::Date || pt160w == Vb6Type::Currency) {
                                hint160w = "Double";
                            } else if (pt160w == Vb6Type::Long || pt160w == Vb6Type::Integer
                                       || pt160w == Vb6Type::Boolean || pt160w == Vb6Type::Byte
                                       || pt160w == Vb6Type::Error) {
                                hint160w = "Long";
                            } else if (pt160w == Vb6Type::LongPtr) {
                                hint160w = "LongPtr";
                            } else if (pt160w == Vb6Type::Object) {
                                hint160w = "Object";
                            } else if (pt160w == Vb6Type::Variant || pt160w == Vb6Type::Empty
                                       || pt160w == Vb6Type::Null) {
                                hint160w = "Variant";
                            }
                        }
                        resolveComValue(hint160w);
                    }
                    std::string valExpr = std::move(lastExpr_);
                    // Fix 091n: 按 Let/Set **写方向**参数表适配值实参 —
                    //  · Variant 形参 ← 具体类型值 (Byte 数组/SafeArray1D* 等) 需
                    //    packLetValueArg 打包 (cWinsock: oClient.UserData = baBuffer,
                    //    实参 vb6_SafeArray1D* → C2440 "→ vb6_VARIANT");
                    //  · BSTR 形参 ← Variant 值 (COM 属性结果) 需提取
                    //    (cToolsList: Rs.Filter = <ComGetProp 结果>, C2440 "→ BSTR").
                    // 注意不能用裸 propLetSym->params: lookupModuleByKind 会全局命中
                    // 他类同名属性 (其 Variant 参数表) → 误打包 → 大规模 C2440
                    // (cHttpServer/cDataBase/cJson 的 Rs.Status = 200 等).
                    std::string padArgs158h; // Fix 158h 入口参数 padding 结果
                    {
                        std::vector<ParameterInfo> wp091n;
                        bool isSet091n = (propLetSym->kind == SymbolKind::PropertySet);
                        if (!sourceModule.empty()
                            && findClassMemberWriteParams(sourceModule, maExpr.memberName,
                                                          isSet091n, wp091n)
                            && !wp091n.empty()) {
                            const ParameterInfo& lastP091n = wp091n.back();
                            if (lastP091n.type == Vb6Type::Variant) {
                                // 仅打包"数组载体"值: 直接给具体类型标量 (如
                                // Dictionary.CompareMode = 1, C 形参 int32_t)
                                // 打包会 C2440 — 符号表 Variant 判定对 COM/后期
                                // 绑定属性不可靠.
                                bool arrCarrier091n = valExpr.find("vb6_SafeArray1D") != std::string::npos
                                    || valExpr.rfind("_arr_", 0) == 0
                                    || valExpr.find("vb6_Split(") == 0
                                    || valExpr.find("vb6_ArrayCreate") == 0
                                    || valExpr.find("vb6_VariantToSafeArray1D") == 0;
                                if (!arrCarrier091n && node.value
                                    && node.value->kind == ASTNodeKind::IdentifierExpr) {
                                    std::string vLower091n = Symbol::toLower(
                                        static_cast<IdentifierExpr&>(*node.value).name);
                                    arrCarrier091n = knownArrays_.count(vLower091n) > 0
                                                     || knownByteArrayVars_.count(vLower091n) > 0;
                                }
                                if (arrCarrier091n) {
                                    valExpr = packLetValueArg(lastP091n, node.value.get(), valExpr);
                                }
                            } else if (lastP091n.type == Vb6Type::String) {
                                // Fix 091n 原条件 (String 形参 ← Variant 值) 只覆盖
                                // COM 属性结果; UserEditingForm.c:588 `EditText =
                                // ComboCalendarValue` (EditText Property Let As String,
                                // ComboCalendarValue 解析为 double → C 实参 double → 形参
                                // BSTR) 生成 vb6_VBFlexGrid_prop_let_EditText(..., double)
                                // → C2440. VB6 这里把 Date/数值隐式转字符串 (CStr 规则).
                                // 对数值/日期类 RHS 同样 wrapToBSTR (wrapToBSTR 按类型
                                // 用 vb6_CStrDate/CStrDbl/CStrLong 等).
                                if (cExprIsVariant(valExpr)) {
                                    valExpr = wrapToBSTR(valExpr, *node.value);
                                } else if (node.value) {
                                    const Vb6Type vt160w = inferExprType(*node.value);
                                    if (vt160w == Vb6Type::Double || vt160w == Vb6Type::Single
                                        || vt160w == Vb6Type::Date || vt160w == Vb6Type::Currency
                                        || vt160w == Vb6Type::Long || vt160w == Vb6Type::Integer
                                        || vt160w == Vb6Type::Boolean || vt160w == Vb6Type::Byte) {
                                        valExpr = wrapToBSTR(valExpr, *node.value);
                                    }
                                }
                            }
                        }
                        // Fix 158h: 属性 Let/Set 写方向的**入口参数** (Optional ByRef
                        // After 等前置参数) 缺失 → 生成的 prop_let_ 调用 C2198 参数太少
                        // (VBFlexGrid.ctl:9212 "Public Property Let InsertMark(Optional
                        // ByRef After As Boolean, ByVal Value As Long)", C 声明
                        // (me, int16_t* After, int32_t Value, int _has_After), 而调用点
                        // MainForm.frm:746/751 ".InsertMark = .FixedRows" / ".InsertMark
                        // = -1" 只传 (obj, value)). 按写方向形参表补每个前置入口参数的
                        // 默认值 + Optional 的 _has_ 标志 (与 m22_module Fix 090am 同模板).
                        if (wp091n.size() > 1) {
                            padArgs158h = objExpr;
                            size_t np158h = wp091n.size() - 1;
                            for (size_t j = 0; j < np158h; j++) {
                                const ParameterInfo& pp158h = wp091n[j];
                                std::string defP158h =
                                    (pp158h.hasDefaultValue && !pp158h.defaultValueExpr.empty())
                                    ? pp158h.defaultValueExpr : defaultValue(pp158h.type);
                                padArgs158h += ", ";
                                if (pp158h.isByVal) {
                                    padArgs158h += defP158h;
                                } else {
                                    std::string ct158h = mapType(pp158h.type);
                                    if (pp158h.type == Vb6Type::Variant
                                        || pp158h.type == Vb6Type::Empty
                                        || pp158h.type == Vb6Type::Null
                                        || pp158h.type == Vb6Type::Object) {
                                        padArgs158h += "&(" + ct158h + "){0}";
                                    } else {
                                        padArgs158h += "&(" + ct158h + "){" + defP158h + "}";
                                    }
                                }
                            }
                            padArgs158h += ", " + valExpr;
                            for (size_t j = 0; j < np158h; j++) {
                                if (wp091n[j].isOptional) padArgs158h += ", 0";
                            }
                        }
                    }
                    if (!padArgs158h.empty()) {
                        c_.emitLine(funcName + "(" + padArgs158h
                                    + ");  /* Property Let (Fix 158h) */");
                    } else {
                        c_.emitLine(funcName + "(" + objExpr + ", " + valExpr
                                    + ");  /* Property Let */");
                    }
                    return;
                }
            }
        }
    }

    // P11.7: 默认属性写入 — 如果赋值目标是已知窗体控件标识符 (如 Label1 = value)
    // 自动转换为默认属性写入 (如 Label1.Caption = value → vb6_SetControlText(...))
    if (node.target->kind == ASTNodeKind::IdentifierExpr) {
        auto& tgtId = static_cast<IdentifierExpr&>(*node.target);
        std::string tgtLower = tgtId.name;
        std::transform(tgtLower.begin(), tgtLower.end(), tgtLower.begin(), ::tolower);
        // Fix 056b: 裸标识符目标可能是类属性 (pvState = sckOpen → vb6_<Class>_prop_let_pvState(me, value))
        // 之前漏到此路径生成裸名赋值 → C2065 "未声明的标识符"
        // Fix 083a: 当前过程的参数/局部变量优先于属性符号 — 例如 SyncReceiveText 的
        // Optional ByVal TimeOut 参数会被全局符号表误匹配为 cHttpServerSession.TimeOut 属性,
        // 生成 vb6_cHttpServerSession_prop_let_TimeOut(me, ...) → 错误调用
        {
            bool isLocalTarget = knownLocalVars_.count(tgtLower);
            if (!isLocalTarget && currentProc_) {
                for (const auto& p : currentProc_->params) {
                    std::string pLower = p.name;
                    std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);
                    if (pLower == tgtLower) { isLocalTarget = true; break; }
                }
            }
            Symbol* propSym = nullptr;
            if (!isLocalTarget) propSym = symTab_.lookupModuleByKind(tgtId.name, SymbolKind::PropertyLet);
            if (!propSym && !isLocalTarget) propSym = symTab_.lookupModuleByKind(tgtId.name, SymbolKind::PropertySet);
            if (propSym) {
                // 属性过程内部 PropertyName = value 是设置返回值, 不是PropertyLet调用
                // (如 Property Get pvThunkGlobalData 内: pvThunkGlobalData = Val(...) → vb6_ret_xxx = ...)
                // Fix 090f: 普通 Function 内 "FuncName = expr" 也是返回赋值 (VB6 允许,
                // 如 cCollection.Items(): Items = Array()). 此前仅识别 Property 过程,
                // 当全局属性符号与函数同名时 (cJson.Items Property Let/Set 存在,
                // lookupModuleByKind 命中), 被误拦截生成 vb6_cJson_prop_let_Items(me,..)
                // → C2198 参数太少 (Items() 无参却按 2 参带参属性拼调用).
                bool isAssigningReturnValue = false;
                if (currentProc_) {
                    std::string curName = currentProc_->name;
                    std::transform(curName.begin(), curName.end(), curName.begin(), ::tolower);
                    std::string tgtName = tgtId.name;
                    std::transform(tgtName.begin(), tgtName.end(), tgtName.begin(), ::tolower);
                    if (curName == tgtName &&
                        (currentProc_->kind == SymbolKind::Function ||
                         currentProc_->kind == SymbolKind::PropertyGet ||
                         currentProc_->kind == SymbolKind::PropertyLet ||
                         currentProc_->kind == SymbolKind::PropertySet)) {
                        isAssigningReturnValue = true;
                    }
                }
                // Fix 093a: 本类成员字段优先 — `recvBuffer = data` (cClientCallback.cls,
                // recvBuffer 是本类字段) 被全局符号表误命中 cWinsock 的 Property Let
                // RecvBuffer → vb6_cWinsock_prop_let_recvBuffer((void*)me, ...) → LNK2019.
                // 命中本类字段时不走属性路径, 交给下方通用赋值生成 me->recvBuffer = ...;
                if (!isAssigningReturnValue && !isOwnClassField(tgtId.name)) {
                    std::string prefix = (propSym->kind == SymbolKind::PropertySet) ? "prop_set_" : "prop_let_";
                    std::string funcName = cProcName(prefix + tgtId.name, propSym->access,
                        propSym->isExternal ? propSym->sourceModule : (isClassModule_ ? moduleName_ : ""));
                    emitExpr(*node.value);
                    // Fix 160w: bare 属性写 RHS 同为主路径场景 (WithMemberExpr/COM 成员)
                    // — 消费标记以避免裸对象实参 C2440 与跨过程泄漏.
                    if (isComMarker_) resolveComValue("BSTR");
                    std::string valExpr = std::move(lastExpr_);
                    // Fix 186: 裸属性写入 (类内 `FontName = sTmp`, 无 obj 前缀) 的 RHS
                    // 此前原样传值, 但 Property Let 值形参 VB6 默认 ByRef → C 形参是
                    // BSTR*/int32_t* 等 — BSTR 值被当地址解引用, 读字符串头 4 字节当
                    // 指针 → 0xC0000005 (cDlg.cls `FontName = sByteArrayToString(...)`,
                    // av read target=0x53004D 即 UTF-16 "MS"). 与 With 路径 Fix 185
                    // 同款 packLetValueArg 打包; 限定本类属性 (isClassModule_ &&
                    // !isExternal, 与下方 funcName 生成的分支对齐, 保证打包所用参数表
                    // 与被调函数名同源) — lookupModuleByKind 可能全局命中他类同名属性
                    // (Fix 158h 教训), external 符号维持原发射不介入.
                    if (isClassModule_ && !propSym->isExternal
                        && propSym->kind == SymbolKind::PropertyLet
                        && !propSym->params.empty()
                        && (propSym->params.back().type == Vb6Type::Variant
                            || !propSym->params.back().isByVal)) {
                        valExpr = packLetValueArg(propSym->params.back(),
                                                  node.value.get(), valExpr);
                    }
                    c_.emitLine(funcName + (isClassModule_ ? "((void*)me, " : "(") + valExpr + ");  /* Property Let/Set (bare) */");
                    return;
                }
            }
        }
        auto itCtrl = knownFormControls_.find(tgtLower);
        if (itCtrl != knownFormControls_.end()) {
            const char* defaultProp = getDefaultPropertyName(itCtrl->second);
            if (defaultProp) {
                std::string writeFn = getControlPropWriteFn(itCtrl->second, defaultProp);
                if (!writeFn.empty()) {
                    emitExpr(*node.value);
                    std::string valExpr = std::move(lastExpr_);
                    // M22: Text/Caption default prop writes need BSTR value (Fix 194: 收敛到统一入口)
                    valExpr = controlPropValueExpr(writeFn, valExpr, *node.value);
                    c_.emitLine(writeFn + "(" + makeCtrlHwndArg(tgtLower, itCtrl->second) + ", " + valExpr + ");  /* default prop: ." + std::string(defaultProp) + " */");
                    return;
                }
            }
        }
        // P20-31: WithEvents控件变量默认属性写入
        if (node.target->kind == ASTNodeKind::IdentifierExpr) {
            auto& tgtId2 = static_cast<IdentifierExpr&>(*node.target);
            std::string tgtLower2 = tgtId2.name;
            std::transform(tgtLower2.begin(), tgtLower2.end(), tgtLower2.begin(), ::tolower);
            auto itWECtrl = knownWithEventsCtrlVars_.find(tgtLower2);
            if (itWECtrl != knownWithEventsCtrlVars_.end()) {
                const char* defaultProp2 = getDefaultPropertyName(itWECtrl->second);
                if (defaultProp2) {
                    std::string writeFn2 = getControlPropWriteFn(itWECtrl->second, defaultProp2);
                    if (!writeFn2.empty()) {
                        auto itOrig2 = knownWithEventsCtrlOrigNames_.find(tgtLower2);
                        std::string weVarName2 = (itOrig2 != knownWithEventsCtrlOrigNames_.end()) ? itOrig2->second : cIdent(tgtId2.name);
                        emitExpr(*node.value);
                        std::string valExpr2 = std::move(lastExpr_);
                        // Fix 194: 原缺 BSTR 转换 (数值 RHS → 指针解引用 / 静默无效)
                        valExpr2 = controlPropValueExpr(writeFn2, valExpr2, *node.value);
                        c_.emitLine(writeFn2 + "(" + weVarName2 + ", " + valExpr2 + ");  /* WithEvents ctrl default prop: ." + std::string(defaultProp2) + " */");
                        return;
                    }
                }
            }
        }
    }
    // P17.1: WithMemberExpr作为赋值目标 (With块内 .Property = value)
    if (node.target->kind == ASTNodeKind::WithMemberExpr && !withObjectVars_.empty() && !withObjectInfoStack_.empty()) {
        auto& wmExpr = static_cast<WithMemberExpr&>(*node.target);
        const auto& info = withObjectInfoStack_.back();
        const std::string& tempVar = withObjectVars_.back();

        switch (info.kind) {
        case WithObjKind::FormControl: {
            std::string writeFn = getControlPropWriteFn(info.ctrlType, wmExpr.memberName);
            if (!writeFn.empty()) {
                emitExpr(*node.value);
                if (info.ctrlType == FrmControlType::Menu) {  // P20-36
                    std::string mnuLower = info.ctrlOrigName;
                    std::transform(mnuLower.begin(), mnuLower.end(), mnuLower.begin(), ::tolower);
                    c_.emitLine(writeFn + "(" + makeCtrlHwndArg(mnuLower, info.ctrlType) + ", " + lastExpr_ + ");  /* With menu prop write */");
                } else {
                    // M22: Text/Caption prop writes need BSTR (Fix 194: 收敛到统一入口)
                    lastExpr_ = controlPropValueExpr(writeFn, lastExpr_, *node.value);
                    c_.emitLine(writeFn + "(" + tempVar + ", " + lastExpr_ + ");");
                }
                return;
            }
            break;
        }
        case WithObjKind::WithEventsCtrl: {
            std::string writeFn = getControlPropWriteFn(info.ctrlType, wmExpr.memberName);
            if (!writeFn.empty()) {
                emitExpr(*node.value);
                // Fix 194: 原缺 BSTR 转换 (数值 RHS → 指针解引用 / 静默无效)
                std::string valExpr = controlPropValueExpr(writeFn, lastExpr_, *node.value);
                c_.emitLine(writeFn + "(" + info.ctrlOrigName + ", " + valExpr + ");");
                return;
            }
            break;
        }
        case WithObjKind::COMObject: {
            // Fix 161f-extlist: With 目标是 ListView 槽 (形参/局部) —— 走控件属性写表,
            // 不发 vb6_ComSetProp (那是拿 HWND 当 IDispatch 用, 静默无效)。
            if (info.ctrlType == FrmControlType::ListView) {
                std::string writeFnLv = getControlPropWriteFn(FrmControlType::ListView, wmExpr.memberName);
                if (!writeFnLv.empty()) {
                    emitExpr(*node.value);
                    std::string valExprLv = controlPropValueExpr(writeFnLv, lastExpr_, *node.value);
                    c_.emitLine(writeFnLv + "((void*)" + tempVar + ", " + valExprLv + ");  /* With ListView slot prop write */");
                    return;
                }
            }
            emitExpr(*node.value);
            std::string valExpr = std::move(lastExpr_);
            std::string packFn = comPackExpr(*node.value);
            c_.emitLine("vb6_ComSetProp(" + tempVar + ", L\"" + wmExpr.memberName + "\", " +
                         packFn + "(" + valExpr + "));  /* With COM SetProp */");
            return;
        }
        case WithObjKind::ClassInstance: {
            // Fix 011r-1: 优先用 info.className 精确查找该类的 Property Let/Set
            // (原 symTab_.lookupModule 会捡错模块, 导致 Pattern A2 在 With 块赋值时出错)
            // Fix 011r-1b: VB6 case-insensitive — 类名匹配与函数名输出均用规范类名 (来自符号表)
            if (!info.className.empty()) {
                std::string letFn, setFn;
                const Symbol* letPropSym = nullptr;  // Fix 090x: 记录 PropertyLet 符号以取末参方向
                const Symbol* setPropSym = nullptr;
                std::string canonicalClassName;  // 与struct定义一致的大小写
                std::string classNameLower = Symbol::toLower(info.className);
                if (symTab_.moduleScope()) {
                    std::string memberLower = Symbol::toLower(wmExpr.memberName);
                    for (const auto& [key, sym] : symTab_.moduleScope()->symbols()) {
                        if (sym->lowerName != memberLower) continue;
                        bool matches = false;
                        if (sym->isExternal) {
                            if (Symbol::toLower(sym->sourceModule) == classNameLower) {
                                matches = true;
                                if (canonicalClassName.empty()) canonicalClassName = sym->sourceModule;
                            }
                        } else if (isClassModule_ && Symbol::toLower(moduleName_) == classNameLower) {  // Fix 013: moduleName_ = VB_Name
                            matches = true;
                            if (canonicalClassName.empty()) canonicalClassName = moduleName_;
                        }
                        if (!matches) continue;
                        if (sym->kind == SymbolKind::PropertyLet) {
                            letFn = "vb6_" + cIdent(canonicalClassName) + "_prop_let_" + cIdent(wmExpr.memberName);
                            letPropSym = sym.get();
                        } else if (sym->kind == SymbolKind::PropertySet) {
                            setFn = "vb6_" + cIdent(canonicalClassName) + "_prop_set_" + cIdent(wmExpr.memberName);
                            setPropSym = sym.get();
                        }
                    }
                }
                if (!letFn.empty()) {
                    emitExpr(*node.value);
                    // Fix 090x: With 块 .Expires = vb6_DateAdd(...) — 属性
                    // Property Let Expires(v As Variant) ByRef 值参是 vb6_VARIANT*,
                    // double/BSTR/int 实参裸拼 → C2440 (cHttpServer.c 913). 与
                    // Fix 090w (Pattern C/D2) 同款值参打包, 复用 packLetValueArg.
                    // Fix 185: 除 Variant 外, ByRef 值类型形参 (String/Boolean/
                    // Integer/Long/...) 的 C 形参同样是 <T>* (VB6 默认 ByRef),
                    // 实参裸传值会被当地址解引用 → 非确定 AV/堆破坏.
                    if (letPropSym && !letPropSym->params.empty()
                        && (letPropSym->params.back().type == Vb6Type::Variant
                            || !letPropSym->params.back().isByVal)) {
                        lastExpr_ = packLetValueArg(letPropSym->params.back(),
                                                    node.value.get(), lastExpr_);
                    }
                    // Fix <VBFlexGridDemo>: With 块这条 prop_let_ 与上面 Fix 158h 那条
                    // (非 With 的 MemberAccess 写) 同病 —— 属性带**前置入口参数**
                    // (Optional ByRef After) 时只发 (obj, value):
                    //   VBFlexGrid.ctl:9212 Property Let InsertMark(Optional ByRef After
                    //   As Boolean, ByVal Value As Long) ⇒ C 声明
                    //   (me, int16_t* After, int32_t Value, int _has_After)
                    //   MainForm.frm:746/751 `.InsertMark = .FixedRows` / `.InsertMark = -1`
                    //   ⇒ prop_let_InsertMark(_vb6_with_22, …) 只有 2 实参 ⇒ C2198×2。
                    // 复用 Fix 158h 的 padding 模板 (默认值 + Optional 的 _has_ 0)。
                    std::string padArgsWithLet;
                    {
                        std::vector<ParameterInfo> wpWith;
                        if (!canonicalClassName.empty()
                            && findClassMemberWriteParams(canonicalClassName, wmExpr.memberName,
                                                          false, wpWith)
                            && wpWith.size() > 1) {
                            padArgsWithLet = tempVar;
                            size_t npW = wpWith.size() - 1;
                            for (size_t j = 0; j < npW; j++) {
                                const ParameterInfo& ppW = wpWith[j];
                                std::string defW =
                                    (ppW.hasDefaultValue && !ppW.defaultValueExpr.empty())
                                    ? ppW.defaultValueExpr : defaultValue(ppW.type);
                                padArgsWithLet += ", ";
                                if (ppW.isByVal) {
                                    padArgsWithLet += defW;
                                } else {
                                    std::string ctW = mapType(ppW.type);
                                    if (ppW.type == Vb6Type::Variant
                                        || ppW.type == Vb6Type::Empty
                                        || ppW.type == Vb6Type::Null
                                        || ppW.type == Vb6Type::Object) {
                                        padArgsWithLet += "&(" + ctW + "){0}";
                                    } else {
                                        padArgsWithLet += "&(" + ctW + "){" + defW + "}";
                                    }
                                }
                            }
                            padArgsWithLet += ", " + lastExpr_;
                            for (size_t j = 0; j < npW; j++) {
                                if (wpWith[j].isOptional) padArgsWithLet += ", 0";
                            }
                        }
                    }
                    if (!padArgsWithLet.empty()) {
                        c_.emitLine(letFn + "(" + padArgsWithLet
                                    + ");  /* With class prop_let_ (Fix <VBFlexGridDemo>) */");
                    } else {
                        c_.emitLine(letFn + "(" + tempVar + ", " + lastExpr_ + ");  /* With class prop_let_ */");
                    }
                    return;
                }
                if (!setFn.empty()) {
                    emitExpr(*node.value);
                    c_.emitLine(setFn + "(" + tempVar + ", " + lastExpr_ + ");  /* With class prop_set_ (fallback) */");
                    return;
                }
                // Fix 090h: 跨模块 storageKey 冲突 (如 CookieAttr.Value 的 value$pl
                // 被先注入的其它类同名属性覆盖, 消费模块作用域只留一个外部符号)
                // 使上面 moduleScope 扫描漏掉 PropertyLet/Set → 误走字段写 →
                // C2039 (CookieAttr 无 Value 数据字段, With 块 .Value= 生成
                // _vb6_with_9->Value, cHttpServer.c 908). findClassMemberCallParams
                // Phase B 从 Class 符号自身的 memberParams/memberProcKinds 确认
                // 该类确实有该属性成员 → 生成 prop_let_ 调用.
                if (letFn.empty() && setFn.empty() && !info.className.empty()) {
                    std::vector<ParameterInfo> clsParams090h;
                    bool clsBuiltin090h = false;
                    std::string canonCls090h = info.className;
                    if (symTab_.moduleScope()) {
                        auto clsIt090h = symTab_.moduleScope()->symbols().find(
                            Symbol::toLower(info.className));
                        if (clsIt090h != symTab_.moduleScope()->symbols().end()
                            && clsIt090h->second->kind == SymbolKind::Class) {
                            canonCls090h = clsIt090h->second->name;
                        }
                    }
                    if (findClassMemberCallParams(canonCls090h, wmExpr.memberName,
                                                  clsParams090h, clsBuiltin090h)) {
                        emitExpr(*node.value);
                        // Fix 185: 与主路径一致 — ByRef 值类型形参需打包成 <T>*.
                        std::string val090h = lastExpr_;
                        if (!clsParams090h.empty()
                            && (clsParams090h.back().type == Vb6Type::Variant
                                || !clsParams090h.back().isByVal)) {
                            val090h = packLetValueArg(clsParams090h.back(),
                                                      node.value.get(), val090h);
                        }
                        c_.emitLine("vb6_" + cIdent(canonCls090h) + "_prop_let_"
                                    + cIdent(wmExpr.memberName) + "(" + tempVar + ", "
                                    + val090h + ");  /* With class prop_let_ (Fix 090h) */");
                        return;
                    }
                    // Fix <vbeclipse>: findClassMemberCallParams 也会漏 (SplitBar.cls
                    // 的 `Property Let Orientation` 生成 vb6_SplitBar_prop_let_Orientation,
                    // 但语义层 Class 符号的 memberParams/memberProcKinds 未记录该属性 →
                    // 090h 走不到 → 字段写 `_vb6_with_1->Orientation` → C2039
                    // "Orientation 不是 vb6_cls_SplitBar 的成员" (ucSplitBar.ctl:148
                    // `With SplitBar: .Orientation = m_Orientation`). 用
                    // resolveClassMemberCall 兜底 (它查运行期类成员解析, 与属性函数名
                    // 生成一源, 能找到 prop_let_).
                    {
                        std::string rc090h = resolveClassMemberCall(canonCls090h,
                                                                    wmExpr.memberName);
                        if (!rc090h.empty()
                            && rc090h.find("_prop_let_") != std::string::npos) {
                            emitExpr(*node.value);
                            std::vector<ParameterInfo> p090h;
                            bool b090h = false;
                            if (findClassMemberCallParams(canonCls090h, wmExpr.memberName,
                                                          p090h, b090h)
                                && !p090h.empty()
                                && (p090h.back().type == Vb6Type::Variant
                                    || !p090h.back().isByVal)) {
                                lastExpr_ = packLetValueArg(p090h.back(),
                                                            node.value.get(), lastExpr_);
                            }
                            c_.emitLine(rc090h + "(" + tempVar + ", " + lastExpr_
                                        + ");  /* With class prop_let_ (vbeclipse) */");
                            return;
                        }
                    }
                }
                // Fix <vbeclipse>: SplitBar.cls 类符号被 ucSplitBar.ctl 同名控件
                // 字段遮蔽 (ucSplitBar.h:140 `void* SplitBar;`) → resolveClassMemberCall/
                // findClassMemberCallParams 均失效 → 上面 090h/FnRet 兜底落空 → 字段写
                // `_vb6_with_1->Orientation` C2039. 成员不是该类的数据字段
                // (classTypedFieldMap_ 无此项) → 按 090h 同一命名约定直发
                // vb6_<Cls>_prop_let_<Member>(With, value) (SplitBar.h 已生成
                // vb6_SplitBar_prop_let_Orientation(vb6_cls_SplitBar*, int32_t)).
                {
                    bool isDataField090c = false;
                    if (classTypedFieldMap_) {
                        auto itCf090c = classTypedFieldMap_->find(info.className);
                        if (itCf090c != classTypedFieldMap_->end()
                            && itCf090c->second.count(
                                Symbol::toLower(wmExpr.memberName)))
                            isDataField090c = true;
                    }
                    if (!isDataField090c && !info.className.empty()) {
                        std::string canonCls090c = info.className;
                        if (symTab_.moduleScope()) {
                            auto clsIt090c = symTab_.moduleScope()->symbols().find(
                                Symbol::toLower(info.className));
                            if (clsIt090c != symTab_.moduleScope()->symbols().end()
                                && clsIt090c->second->kind == SymbolKind::Class)
                                canonCls090c = clsIt090c->second->name;
                        }
                        emitExpr(*node.value);
                        c_.emitLine("vb6_" + cIdent(canonCls090c) + "_prop_let_"
                                    + cIdent(wmExpr.memberName) + "(" + tempVar + ", "
                                    + lastExpr_ + ");  /* With class prop_let_ direct (vbeclipse) */");
                        return;
                    }
                }
                // 既无Let也无Set → 视为数据字段写: tempVar->member = value
                emitExpr(*node.value);
                // Fix 092m: RHS 是 COM/typed-COM 成员读时必须先消费 COM 标记 —
                // 否则 lastExpr_ 只含对象本身 (cHttpServer 372/373:
                //   .IP/.Port = me->m_oServer, 属性名整体丢失 = 编译通过但语义错),
                // 且标记泄漏到下一条语句 (374: .ConnectAt = ... L"RemotePort" C2440).
                // 无 hint 的 resolveComValue 在早期绑定签名可用时按 returnType 取
                // ComGetStringProp/IntProp/DoubleProp, 否则回退 BSTR.
                if (isComMarker_) {
                    // Fix 092m: 按目标字段类型解包 (String→ComGetStringProp /
                    // Long→ComGetIntProp / Date→ComGetDoubleProp). 目标类字段类型表
                    // 由语义分析填充; 表缺失时 helper 回退 "BSTR".
                    resolveComValue(classFieldComUnpackHint(info.className,
                                                            wmExpr.memberName));
                }
                c_.emitLine(tempVar + "->"
                            + cIdent(canonicalClassFieldName(info.className,
                                                             wmExpr.memberName))
                            + " = " + lastExpr_ + ";  /* With class field write */");
                return;
            }

            // className未知 — 使用原symTab查找(可能捡错模块, 但无法避免)
            Symbol* memSym = symTab_.lookupModule(wmExpr.memberName);
            if (!memSym) memSym = symTab_.lookup(wmExpr.memberName);
            if (memSym && memSym->kind == SymbolKind::PropertyLet) {
                std::string propFn = "prop_let_" + wmExpr.memberName;
                std::string funcName = cProcName(propFn, memSym->access,
                    memSym->isExternal ? memSym->sourceModule : (isClassModule_ ? moduleName_ : ""));
                emitExpr(*node.value);
                c_.emitLine(funcName + "(" + tempVar + ", " + lastExpr_ + ");");
                return;
            }
            if (memSym && memSym->kind == SymbolKind::PropertySet) {
                std::string propFn = "prop_set_" + wmExpr.memberName;
                std::string funcName = cProcName(propFn, memSym->access,
                    memSym->isExternal ? memSym->sourceModule : (isClassModule_ ? moduleName_ : ""));
                emitExpr(*node.value);
                c_.emitLine(funcName + "(" + tempVar + ", " + lastExpr_ + ");  /* With class PropertySet */");
                return;
            }
            // Fix 010n: 未知成员赋值 → COM后期绑定
            // .DataMember = value → vb6_ComSetProp(obj, L"Member", packedValue)
            emitExpr(*node.value);
            std::string valExpr = std::move(lastExpr_);
            std::string packFn = comPackExpr(*node.value);
            c_.emitLine("vb6_ComSetProp(" + tempVar + ", L\"" + wmExpr.memberName + "\", " +
                         packFn + "(" + valExpr + "));  /* With class .unknown COM SetProp */");
            return;
        }
        case WithObjKind::BuiltinObject: {
            // Fix 010l: .Property = value on builtin object (Err/App/etc.)
            // Err properties are read-only in C RTL — emit as comment to avoid MSVC error
            emitExpr(*node.value);
            c_.emitLine("/* With " + info.ctrlOrigName + "." + wmExpr.memberName +
                        " = <value> — builtin object property write (no-op) */");
            return;
        }
        default:
            break;
        }
    }
