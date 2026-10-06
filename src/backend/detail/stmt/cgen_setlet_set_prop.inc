// cgen_setlet_set_prop.inc - visit(SetStmt&) 函数体片段：目标侧分派（Nothing / 属性 / COM / 接口引用）
// 由 src/backend/stmt/cgen_setlet.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 12~321 行，逐行未改。

    if (!node.target || !node.value) return;

    // ai/022 B09c: `Set MyBase.<属性> = obj` —— 目标侧走基类面上的 Property Set。
    // 实测改码前发的是 `vb6_T9Base_prop_set_Peer(MyBase, me);  /* Set Property */` —— 函数名
    // 那条通路本来就查对了, 错的只是**接收者**: `MyBase` 被当裸标识符发出去 → C2065 未声明。
    // 这里在 Nothing/COM/链式那些分支之前就接管, 与 Let 侧的 tryEmitMyBaseAssign 同一条口径。
    if (node.target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma09c = static_cast<MemberAccessExpr&>(*node.target);
        if (ma09c.object && ma09c.object->kind == ASTNodeKind::IdentifierExpr
            && Symbol::toLower(static_cast<IdentifierExpr&>(*ma09c.object).name) == "mybase") {
            emitExpr(*node.value);
            std::string rhs09c = std::move(lastExpr_);
            if (tryEmitMyBaseAssign(ma09c, rhs09c, /*forSet=*/true)) return;
        }
    }

    // 检测 Set obj = Nothing → vb6_ReleaseObject(&obj)
    if (node.value->kind == ASTNodeKind::LiteralExpr) {
        auto& lit = static_cast<LiteralExpr&>(*node.value);
        if (lit.literalKind == LiteralKind::Nothing) {
            // 左值求值期间关掉虚表派发 (ai/022 B08e-4 站点⑭: Pattern C/D2 要对这段文本做
                // `prop_get_` -> `prop_let_` 的字符串级改写, 派发表达式会被换成表里不存在的槽字段)
            const bool savedVirtLhs = suppressVirtDispatch_;
            suppressVirtDispatch_ = true;
            emitExpr(*node.target);
            suppressVirtDispatch_ = savedVirtLhs;
            std::string target = std::move(lastExpr_);

            // COM属性SetRef Nothing: Set obj.Property = Nothing → vb6_ComSetRef(obj, L"Property", NULL)
            if (isComMarker_) {
                isComMarker_ = false;
                c_.emitLine("vb6_ComSetRef(" + comObjExpr_ + ", L\"" + comMemberName_ + "\", NULL);  /* COM SetRef Nothing */");
                comObjExpr_.clear();
                comMemberName_.clear();
                return;
            }

            // Fix 056b: Set obj.Prop = Nothing → 属性setter调用传 NULL
            // target 形如 vb6_cXxx_prop_set_fClient(obj) (来自 emitExpr 属性访问路径),
            // 不能做 &(prop_set_...) (void* 左值) → C2198/C2440.
            // Fix 083b: 匹配条件放宽 — 类名前缀使函数名形如 vb6_cXxx_prop_set_fClient(,
            // 原先的 "prop_set_(" 子串匹配不到 (prop_set_ 后是函数名不是左括号)
            if ((target.find("prop_set_") != std::string::npos && target.find("prop_set_") < target.find('(')) ||
                (target.find("prop_let_") != std::string::npos && target.find("prop_let_") < target.find('('))) {
                if (!target.empty() && target.back() == ')') {
                    target = target.substr(0, target.size() - 1) + ", NULL)";
                }
                c_.emitLine(target + ";  /* Set Nothing (property setter) */");
                return;
            }

            // Fix 086: 目标表达式解析为空对象 stub ((void*)0) — 无可释放引用,
            // 跳过整条语句 (避免 &(void*)0 → C2101; 例: Set Response.fClient = Nothing
            // 中 Response.fClient 链被解析为 COM NULL stub).
            if (target == "(void*)0" || target == "NULL" || target == "0") {
                return;
            }

            // vbeclipse: Variant 容器左值 `Set x = Nothing` → 清空容器, 不释放对象.
            // Set m_Items(Idx) = Nothing (m_Items() As Variant) 的 target 文本是
            // `VB6_SA_AT(vb6_VARIANT, me->m_Items, (i-1))`, 下方 isCallTarget 判定
            // 成立 → 包装成 `vb6_ReleaseObject((void**)(&(void*){...}))` →
            // C2440 (`(void*){vb6_VARIANT}` 复合字面量不能用结构体初始化, List.c
            // 36/50). 语义上清空 Variant 即 VB6 的 Set Nothing, 与同循环里
            // 非对象元素的清空口径一致 (List.c 32/46: vb6_VariantClear(&VB6_SA_AT(...))).
            // 裸 Variant 局部 / me->Variant 字段同样走这里, 不再对 VARIANT 取址
            // 当 IDispatch* 释放 (此前 vb6_ReleaseObject(&v) 编译得过但运行期
            // 会把 Variant 首 8 字节当 vtable).
            {
                std::string vcLower = target;
                if (vcLower.rfind("me->", 0) == 0) vcLower = vcLower.substr(4);
                if (vcLower.size() > 4 && vcLower[0] == '(' && vcLower[1] == '*'
                    && vcLower.back() == ')') {
                    vcLower = vcLower.substr(2, vcLower.size() - 3);
                }
                std::transform(vcLower.begin(), vcLower.end(), vcLower.begin(), ::tolower);
                const bool vcArr = (target.find("VB6_SA_AT(vb6_VARIANT,") == 0);
                const bool vcVar = (!vcArr && knownVariantVars_.count(vcLower) > 0);
                const bool vcFld = (!vcArr && !vcVar && target.rfind("me->", 0) == 0
                    && classVariantFields_.count(vcLower) > 0);
                if (vcArr || vcVar || vcFld) {
                    c_.emitLine("vb6_VariantClear(&" + target + ");  /* Set Nothing (Variant 容器) */");
                    return;
                }
            }

            // P6.3: 早期绑定COM变量 → vb6_ComReleaseTyped
            std::string targetLower = target;
            std::transform(targetLower.begin(), targetLower.end(), targetLower.begin(), ::tolower);
            // tB Interface B05: Set <接口变量> = Nothing → 先经槽 Release 再清空.
            // 不拦下来会落到下面的 vb6_ReleaseObject —— 那会把薄指针当 IDispatch* 用.
            if (knownIvrefVars_.count(targetLower) &&
                target.find_first_not_of(
                    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_")
                    == std::string::npos) {
                c_.emitLine("if (" + target + ") " + target + "->vt->Release(" + target + ");");
                c_.emitLine(target + " = NULL;  /* Set Nothing (tB Interface 值) */");
                return;
            }
            // Fix 084aa: Set dict(key) = Nothing 的目标是 vb6_ComCall(...) 函数调用,
            // 对调用结果取址 &vb6_ComCall(...) → C2102. 检测"以标识符开头+含("的
            // 函数调用形态, 改用 (&(void*){...}) 复合字面量包装 (仅释放引用).
            // me->Field / Client->Context->SSE 等可寻址形态不含 '(' 保持 &target.
            bool isCallTarget = !target.empty()
                && (std::isalpha(static_cast<unsigned char>(target[0])) || target[0] == '_');
            if (isCallTarget) {
                size_t paren = target.find('(');
                if (paren != std::string::npos) {
                    for (size_t j = 0; j < paren; j++) {
                        char ch = target[j];
                        if (!std::isalnum(static_cast<unsigned char>(ch)) && ch != '_') {
                            isCallTarget = false;
                            break;
                        }
                    }
                } else {
                    isCallTarget = false;
                }
            }
            std::string relObjTarget = isCallTarget ? "(&(void*){" + target + "})" : ("&" + target);
            // Fix 101: 项目类实例判定补上「类类型推断」— 原 knownClassVars_ 只按
            // 裸变量名登记, 类模块内的目标形如 me-><字段> (如 cHttpServer.SSE
            // As cSSE, 已在 classTypedFieldMap_ 登记为项目类), 永不命中 →
            // 落到 vb6_ReleaseObject, 把纯 C 结构体当 IUnknown* 解引用其首成员
            // __comObj (恒 NULL) → 释放实例时 0xC0000005
            // (cHttpServer / cRedisClient 的 Class_Terminate 首条语句).
            // inferClassTypeOfExpr 递归覆盖 MeExpr / me->字段 / With 成员 /
            // 方法返回类, 且对 COM 字段 (classTypedFieldMap_ 的 "COM:" 标记) 与
            // COM 类型化类 (SymbolKind::ComClass) 一律返回空串, 不误伤本应走
            // vb6_ComReleaseTyped / vb6_ReleaseObject 的目标.
            // 函数调用形态 (isCallTarget) 排除: 其求值非左值, 生成 "<call> = NULL"
            // 会 C2106/C2440.
            std::string targetCls101 = isCallTarget ? std::string()
                                                    : inferClassTypeOfExpr(*node.target);
            if (!targetCls101.empty() || knownClassVars_.count(targetLower)) {
                // 项目类实例 (Dim c As CCircle / Dim WithEvents s As EventSource) 不是 COM
                // 对象, 没有引用计数: 调用 vb6_ReleaseObject 会对纯 C 结构体走 vtable
                // Release → 运行期 0xC0000005. Set Nothing 只需断开引用.
                c_.emitLine(target + " = NULL;  /* Set Nothing (项目类实例) */");
            } else if (knownTypedComVars_.count(targetLower)) {
                c_.emitLine("vb6_ComReleaseTyped((void**)" + relObjTarget + ");  /* Set Nothing (early bound) */");
            } else {
                c_.emitLine("vb6_ReleaseObject((void**)" + relObjTarget + ");  /* Set Nothing */");
            }
            return;
        }
    }

    // P25: Set ctrl.Property = COM_value (e.g. Set Picture2.Picture = ImageList1.ListImages(i).Picture)
    // SetStmt的emitExpr左侧走属性读取路径, 但Picture赋值需要写函数
    if (node.target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& _setMa = static_cast<MemberAccessExpr&>(*node.target);
        if (_setMa.object && _setMa.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& _setId = static_cast<IdentifierExpr&>(*_setMa.object);
            std::string _setObjLower = _setId.name;
            std::transform(_setObjLower.begin(), _setObjLower.end(), _setObjLower.begin(), ::tolower);
            auto _setItCtrl = knownFormControls_.find(_setObjLower);
            if (_setItCtrl != knownFormControls_.end()) {
                std::string _setWriteFn = getControlPropWriteFn(_setItCtrl->second, _setMa.memberName);
                std::string _writeFnRaw = _setWriteFn;
                if (!_setWriteFn.empty()) {
                    // Fix 110z8: Set ctrl.Picture = <StdPicture 变量> — RHS 是局部
                    // StdPicture/IPictureDisp 变量时也应走 FromCom (此前仅识别函数调用
                    // 返回值, 局部变量走 vb6_SetControlPicture 把 COM 指针截断成句柄值,
                    // 导致 Image1.Picture 赋值时把无效句柄当 HBITMAP → 0xC0000005).
                    bool _rhsRhsIsComPicture = false;
                    std::string _rhsRawName;
                    if (node.value->kind == ASTNodeKind::IdentifierExpr) {
                        auto& _rhsId = static_cast<IdentifierExpr&>(*node.value);
                        _rhsRawName = _rhsId.name;
                        const Symbol* _rhsSym = symTab_.lookup(_rhsId.name);
                        if (_rhsSym && !_rhsSym->variableTypeName.empty()) {
                            std::string _rhsT = _rhsSym->variableTypeName;
                            std::transform(_rhsT.begin(), _rhsT.end(), _rhsT.begin(), ::tolower);
                            if (_rhsT == "stdpicture" || _rhsT == "ipicturedisp" || _rhsT == "picture")
                                _rhsRhsIsComPicture = true;
                        }
                    }
                    // Fix 110z9: 局部 StdPicture 变量 (符号表作用域不可达, rhsSym=null)
                    // 且 SetControlPicture write 函数时, 直接按 COM 指针变量传给 FromCom,
                    // 避免 emitExpr 把 StdPicture 变量误当 .Width/默认属性 (生成
                    // vb6_ComGetIntProp(pic, L"Width") 截断句柄 → 崩溃).
                    if (_setWriteFn.find("SetControlPicture") != std::string::npos
                        && !_rhsRawName.empty()) {
                        _rhsRhsIsComPicture = true;
                    }
                    emitExpr(*node.value);
                    // Resolve COM marker on RHS
                    // Bug #3 fix: also detect vb6_ComIface_Picture* function return
                    bool _rhsIsComPicture = _rhsRhsIsComPicture;
                    if (isComMarker_) {
                        if (comMemberName_ == "Picture" || comMemberName_ == "picture") {
                            _rhsIsComPicture = true;
                            resolveComValue("Object");
                        } else {
                            resolveComValue("Object");
                        }
                    } else if (lastExprIsComPicture_) {
                        _rhsIsComPicture = true;
                        lastExprIsComPicture_ = false;  // consume the flag
                    }
                    std::string _setValExpr = std::move(lastExpr_);
                    // Fix 110z9: RHS 是标识符变量时直接用变量名作 COM 指针,
                    // 避免 emitExpr 把 StdPicture 变量误当 .Width/默认属性
                    // (生成 vb6_ComGetIntProp(pic, L"Width") 截断句柄 → 崩溃).
                    if (!_rhsRawName.empty() && _writeFnRaw.find("SetControlPicture") != std::string::npos) {
                        _setValExpr = _rhsRawName;
                    }
                    // Use SetControlPictureFromCom for COM IPictureDisp
                    if (_rhsIsComPicture && _setWriteFn.find("SetControlPicture") != std::string::npos) {
                        _setWriteFn = "vb6_SetControlPictureFromCom";
                    }
                    // Fix 194: 字符串型控件属性的值必须转 BSTR (数值 RHS → 指针解引用 / 静默无效)
                    _setValExpr = controlPropValueExpr(_setWriteFn, _setValExpr, *node.value);
                    // Fix <vbeclipse> 2026-10-06: Picture 属性的 RHS 是 Variant 载体
                    // (LoadResPicture 等, controlPropValueExpr 已包 vb6_VariantToObjectVal)
                    // 时是 COM 对象, 必须走 FromCom (QI IPicture) 而不是裸句柄 setter,
                    // 否则 IPicture 指针被当 HBITMAP 挂上去 → 画面错乱。
                    if (_setWriteFn == "vb6_SetControlPicture"
                        && _setValExpr.compare(0, 23, "vb6_VariantToObjectVal(") == 0) {
                        _setWriteFn = "vb6_SetControlPictureFromCom";
                    }
                    c_.emitLine(_setWriteFn + "(" + makeCtrlHwndArg(_setObjLower, _setItCtrl->second) + ", " + _setValExpr + ");  /* Set Control Property */");
                    return;
                }
            }
        }
    }

    // P6.8: Set obj.Prop = value — 跨类 Friend/Public Property Set/Let 赋值
    // SetStmt 之前没有属性分支: target 被 emitExpr 当表达式生成 prop_set_ 单参数,
    // 后续 tryRewriteCOMLvalue 只认 prop_get_ 前缀, 导致 "prop_set_(obj) = value"
    // 左值错误 (C2440/C2198). 此处用符号信息直接生成完整属性调用.
    if (node.target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& _ma = static_cast<MemberAccessExpr&>(*node.target);
        Symbol* _psSym = symTab_.lookupModuleByKind(_ma.memberName, SymbolKind::PropertySet);
        if (!_psSym) {
            _psSym = symTab_.lookupModuleByKind(_ma.memberName, SymbolKind::PropertyLet);
        }
        // Task #43 守卫: object 是窗体控件且属性不属控件自有 UserControl 类时,
        // member 是 OCX/控件属性 (Set SSTabEx1.Font = ...) — lookupModuleByKind
        // 全局查找会误命中他类同名 PropertySet (cDlg.cls 的 Font), 收者/语义双错。
        // 置空符号交给下方 COM SetRef 分支 (vb6_ComSetRef)。
        if (_psSym && _ma.object && _ma.object->kind == ASTNodeKind::IdentifierExpr) {
            std::string _oidG = Symbol::toLower(
                static_cast<IdentifierExpr&>(*_ma.object).name);
            if (knownFormControls_.count(_oidG)) {
                auto _itUCG = knownUserControlCtrlVars_.find(_oidG);
                bool _ucOwned = (_itUCG != knownUserControlCtrlVars_.end())
                    && _psSym->isExternal
                    && (Symbol::toLower(_psSym->sourceModule)
                        == Symbol::toLower(_itUCG->second));
                if (!_ucOwned) _psSym = nullptr;
            }
        }
        if (_psSym) {
            std::string _objClass = inferClassTypeOfExpr(*_ma.object);
            // Fix 084g-2: 外部注入的跨类属性符号 — 对象类型无法从符号表推断
            // (cgen阶段访问不到过程作用域局部变量) 时, 由外部符号提供 sourceModule.
            // Fix 084y-2: 同 P6.7 — 对象类与属性所属类不一致时, 属性是其他类的
            // 同名成员, 不属于当前对象 (Set oCallback.Socket 命中 cTlsSocket.Socket
            // PropertySet, 而 oCallback 是 cClientCallback, Socket 是数据字段) →
            // 跳过属性路径, 交给 visit(MemberAccessExpr) 生成 obj->field 字段赋值.
            bool _objClassMatchesProp = true;
            if (!_objClass.empty() && _psSym->isExternal) {
                std::string _psModLower = Symbol::toLower(_psSym->sourceModule);
                std::string _objClassLower = Symbol::toLower(_objClass);
                _objClassMatchesProp = (_psModLower == _objClassLower);
            }
            if (_objClassMatchesProp && (!_objClass.empty() || _psSym->isExternal)) {
                std::string _verb = (_psSym->kind == SymbolKind::PropertySet) ? "prop_set_" : "prop_let_";
                std::string _src = _psSym->isExternal ? _psSym->sourceModule : _objClass;
                std::string _fn = cProcName(_verb + _ma.memberName, _psSym->access, _src);
                emitExpr(*_ma.object);
                std::string _objE = std::move(lastExpr_);
                // Fix 160-E: 与 cgen_assign_com_prop.inc:249-266 (Fix 152c) 同一条规则,
                // 但 Set/Let 属性写分支此前**没有**套用。放在窗体上的 UserControl 实例
                // (VBFlexGrid1) 裸 emitExpr 只发出控件名, C 里无此声明 → C2065:
                //   VBFlexGrid1.CellPicture = Picture1.Picture
                //     → vb6_VBFlexGrid_prop_set_CellPicture(VBFlexGrid1, …)
                //   同函数下一行的 Let 却是 (vb6_cls_VBFlexGrid*)vb6_UC_InstanceOf(…)
                // this 是宿主窗口实例, 必须经 InstanceOf 反查:
                if (_ma.object && _ma.object->kind == ASTNodeKind::IdentifierExpr) {
                    auto& _oid160 = static_cast<IdentifierExpr&>(*_ma.object);
                    std::string _oidLower160 = Symbol::toLower(_oid160.name);
                    auto itUC160 = knownUserControlCtrlVars_.find(_oidLower160);
                    if (itUC160 != knownUserControlCtrlVars_.end()) {
                        auto itCtrl160 = knownFormControls_.find(_oidLower160);
                        std::string hwnd160 = (itCtrl160 != knownFormControls_.end())
                            ? makeCtrlHwndArg(_oidLower160, itCtrl160->second)
                            : ("vb6_hwnd_" + cIdent(_oid160.name));
                        _objE = "(vb6_cls_" + cIdent(itUC160->second)
                              + "*)vb6_UC_InstanceOf(" + hwnd160 + ")";
                    }
                }
                emitExpr(*node.value);
                std::string _valE = std::move(lastExpr_);
                // Task #43: RHS 是窗体控件属性读 (Set iDlg.Font = SSTabEx1.Font)
                // 时, Fix 023e/089d 兜底只留下裸 hwnd + isComMarker_ — 不消费则
                // &(void*){hwnd} 被当对象指针存进属性, 后续解引用即崩。
                // 与 cgen_assign_com_prop.inc Fix 110i 同一律, 对象属性按 Object 解。
                if (isComMarker_) {
                    _valE = resolveComValue("Object");
                }
                // Object 形参 (void** ByRef 槽): 对象指针实参包装为 &(void*){...}
                // 复合字面量, 与 cgen_expr 的 ByRef 实参规则一致.
                std::string _arg2 = _valE;
                // Fix 159-E: 实参已是 vb6_VARIANT 时按 value 形参类型解包/直传,
                // 不再无条件套 &(void*){...} 复合字面量。
                //   VBFlexGrid.ctl `Select Case VarType(Value)` 里的
                //   prop_set_CellPicture/CellTag(me, Value) 曾被发成 &(void*){Value}
                //   → VBFlexGrid.c 22824/22827 两条 C2440 (void**→VARIANT 与
                //   (void*){vb6_VARIANT})。
                // 实参**不是** Variant 时维持原发射 (不新增包装): _psSym 可能因跨模块
                // storageKey 命中到别的同名符号, 其 params.back().type 不可信 —
                // prop_set_DataSource(me, void* Value) 曾被误包成
                // vb6_VariantFromValue(Value) → C2172 实参不是指针 ×3。
                bool _valIsVar159 = cExprIsVariant(_valE)
                    || (node.value && node.value->kind == ASTNodeKind::IdentifierExpr
                        && knownVariantVars_.count(Symbol::toLower(
                               static_cast<IdentifierExpr&>(*node.value).name)));
                bool _psValVariant159 = !_psSym->params.empty()
                    && _psSym->params.back().type == Vb6Type::Variant;
                // Fix 166: 被调方形参是 **ByVal** 对象时, 其 C 形参就是单个指针
                // (VBFlexGrid.ctl:3258 `Property Set Font(ByVal NewFont As StdFont)` →
                // `prop_set_Font(vb6_cls_VBFlexGrid*, vb6_ComIface_Font* NewFont)`),
                // 此时再套 &(void*){...} 等于传"指针的指针"。
                // 只有 ByRef 对象槽 (C 形参 void**) 才需要那层复合字面量。
                // 现场后果 (demo 启动即 0xC0000005, 全靠 /MAP 才查出来):
                //   VBFlexGrid.c:2082 prop_let_Font → prop_set_Font(me, &(void*){NewFont})
                //   → me->PropFont = 复合字面量地址(栈) → 2099 CreateGDIFontFromOLEFont
                //   → Common.c:1068 vb6_ComGetObjectProp(Font, L"hFont") 的 disp 是栈地址、
                //   其首 8 字节恰为 &g_vb6_UserControl_FontObj → 当成 vtable 调用 →
                //   RIP 落进 .data。**C 里 void** 传成 vb6_ComIface_Font* 只是指针不匹配
                //   警告, 所以编译期全绿、错误数为 0 也照样崩。**别再加回无条件包装。**
                bool _psValByVal166 = !_psSym->params.empty()
                    && _psSym->params.back().isByVal;
                if (_valIsVar159 && _psValVariant159) {
                    _arg2 = _valE;
                } else if (_valIsVar159) {
                    _arg2 = "vb6_VariantToObjectVal(" + _valE + ")";
                } else if (_arg2 != "NULL" && _arg2 != "0" && _arg2[0] != '&'
                           && _arg2.find("vb6_ComPack") == std::string::npos
                           && !_psValByVal166) {
                    _arg2 = "&(void*){" + _valE + "}";
                }
                c_.emitLine(_fn + "(" + _objE + ", " + _arg2 + ");  /* Set Property */");
                return;
            }
        }
    }

    // Fix 090ae: 链式 COM 默认属性索引写 (P25b helper) — SetStmt 原先缺失该分支,
    // Set Data(Line)(Col) = Dat 的 LHS 被 emitExpr 按链式读生成
    // vb6_VariantFromComResult(vb6_ComCall(...)) = value (非左值) → C2440.
    // 此处分流为 vb6_ComSetPropArg(ComCallObject(...), ...) 链写.
    if (tryEmitChainedComWrite(node.target.get(), node.value.get())) {
        return;
    }

    // 左值求值期间关掉虚表派发 (ai/022 B08e-4 站点⑭: Pattern C/D2 要对这段文本做
    // `prop_get_` -> `prop_let_` 的字符串级改写, 派发表达式会被换成表里不存在的槽字段)
    const bool savedVirtLhs = suppressVirtDispatch_;
    suppressVirtDispatch_ = true;
    emitExpr(*node.target);
    suppressVirtDispatch_ = savedVirtLhs;
    std::string target = std::move(lastExpr_);

    // COM属性SetRef: Set obj.Property = objRef → vb6_ComSetRef(obj, L"Property", objRef)
    if (isComMarker_) {
        isComMarker_ = false;
        std::string objExpr = std::move(comObjExpr_);
        std::string memberName = std::move(comMemberName_);
    emitExpr(*node.value);
    // COM属性值: 如果右侧是COM属性, 解析为值
    if (isComMarker_) resolveComValue();
    std::string value = std::move(lastExpr_);
    // Fix <vbeclipse>: SetRef 的第3参是 void* (对象指针), 右侧若是 Variant 表达式
    // (如 .Item(i) — vb6_List_Item 返回 vb6_VARIANT 结构体, SubClass.cls:56
    //   `Set l_View = .Item(i)` → `vb6_ComSetRef(l_View, L"ViewId", vb6_List_Item(...))`
    //   → C2172 "实参不是指针"), 需先提取其中的 IDispatch (VB6: 对象 As Object
    //   由 Variant 槽持有). 裸 Variant 变量取 & 后经 VariantToObject.
    {
        std::string vTrim = value;
        while (vTrim.size() >= 2 && vTrim.front() == '(' && vTrim.back() == ')')
            vTrim = vTrim.substr(1, vTrim.size() - 2);
        if (knownVariantVars_.count(Symbol::toLower(vTrim))) {
            value = "vb6_VariantToObject(&" + value + ")";
        } else if (cExprIsVariant(value)) {
            value = "vb6_VariantToObjectVal(" + value + ")";
        }
    }
        c_.emitLine("vb6_ComSetRef(" + objExpr + ", L\"" + memberName + "\", " + value + ");  /* COM SetRef */");
        return;
    }

    emitExpr(*node.value);
    // Set语句: 如果右侧是COM调用返回的VARIANT*, 需要解封为对象
    if (isComMarker_) {
        // COM属性值作为对象引用: vb6_ComUnpackObject(vb6_ComGetProp(...))
        resolveComValue("Object");
    }
    std::string value = std::move(lastExpr_);

    // P16: Set cmd = Command1 → value应为vb6_hwnd_Command1而非默认属性值
    {
        std::string targetLower = target;
        std::transform(targetLower.begin(), targetLower.end(), targetLower.begin(), ::tolower);
        if (knownWithEventsCtrlVars_.count(targetLower)) {
            // value可能是vb6_GetControlXxx(vb6_hwnd_Name)形式, 需提取为vb6_hwnd_Name
            size_t hwndPos = value.find("vb6_hwnd_");
            if (hwndPos != std::string::npos) {
                size_t endPos = value.find(")", hwndPos);
                if (endPos != std::string::npos) {
                    value = value.substr(hwndPos, endPos - hwndPos);
                }
            }
        }
    }

    // 如果ComCall/ComGetProp返回VARIANT*含对象, 需要UnpackObject
    // 使用一体化函数: vb6_ComCallObject 内部完成 UnpackObject+VarFree
    if (value.find("vb6_ComCall(") == 0) {
        // vb6_ComCall(obj, L"Method", args, argc) → vb6_ComCallObject(obj, L"Method", args, argc)
        value = "vb6_ComCallObject" + value.substr(strlen("vb6_ComCall"));
    }

    // Fix 037b: 项目类 typed 字段的 Item 属性调用返回 VARIANT, Set 语句需要提取 void* 对象引用.
    // 检测: value 以 "vb6_" 开头且第一个 '(' 恰在 "_prop_get_Item" 之后 → 顶层 Item 调用.
    // 嵌套场景 (如 SomeFunc(obj.Rows(1))) 的第一个 '(' 在 SomeFunc 之后, 不会被误匹配.
    if (value.find("vb6_") == 0
        && value.find("vb6_VariantToObjectVal") == std::string::npos) {
        size_t itemPos = value.find("_prop_get_Item(");
        if (itemPos != std::string::npos) {
            size_t firstParen = value.find('(');
            if (firstParen == itemPos + 15) {  // 15 = strlen("_prop_get_Item")
                value = "vb6_VariantToObjectVal(" + value + ")";
            }
        }
    }
    lastExprNeedsObjectUnpack_ = false;  // 清除标记 (字符串检测已覆盖)

    // P6.3: 早期绑定COM变量赋值: Set fso = CreateObject("X") → fso = (Type*)vb6_ComCreateTyped(L"X", "{IID}")
    // 检查target是否是早期绑定COM变量, 且value是vb6_CreateObject
    if (value.find("vb6_CreateObject(") == 0) {
        // 提取target变量名
        std::string targetLower = target;
        std::transform(targetLower.begin(), targetLower.end(), targetLower.begin(), ::tolower);
        auto it = knownTypedComVars_.find(targetLower);
        if (it != knownTypedComVars_.end()) {
            const Symbol* comSym = it->second;
            // 生成类型转换: (vb6_ComIface_<Iface>*)vb6_ComCreateTyped(progId, iidStr)
            std::string ifaceName = comSym->name;
            if (comSym->kind == SymbolKind::ComClass && !comSym->comDefaultIfaceName.empty()) {
                ifaceName = comSym->comDefaultIfaceName;
            }
            std::string ifaceType = "vb6_ComIface_" + cIdent(ifaceName);
            // 从vb6_CreateObject(progId)中提取progId参数
            size_t start = value.find('(');
            size_t end = value.rfind(')');
            if (start != std::string::npos && end != std::string::npos && end > start) {
                std::string progIdArg = value.substr(start + 1, end - start - 1);
                std::string iidStr = comSym->comIidStr.empty() ? "" : "\"" + comSym->comIidStr + "\"";
                if (!iidStr.empty()) {
                    value = "(" + ifaceType + "*)vb6_ComCreateTyped(" + progIdArg + ", " + iidStr + ")";
                }
                // 如果没有IID, 降级为后期绑定 (保持vb6_CreateObject)
            }
        }
    }

    // P6.3: 早期绑定COM的vtable调用返回对象 → 自动类型转换
    // vb6_ComVtableGetObject(...) → (Type*)vb6_ComVtableGetObject(...)
    if (value.find("vb6_ComVtableGetObject(") == 0) {
        std::string targetLower = target;
        std::transform(targetLower.begin(), targetLower.end(), targetLower.begin(), ::tolower);
        auto it = knownTypedComVars_.find(targetLower);
        if (it != knownTypedComVars_.end()) {
            const Symbol* comSym = it->second;
            std::string ifaceName = comSym->name;
            if (comSym->kind == SymbolKind::ComClass && !comSym->comDefaultIfaceName.empty()) {
                ifaceName = comSym->comDefaultIfaceName;
            }
            std::string ifaceType = "vb6_ComIface_" + cIdent(ifaceName);
            value = "(" + ifaceType + "*)" + value;
        }
    }

    // P6.3: 早期绑定COM变量Set Nothing → vb6_ComReleaseTyped
    // (已在前面的Nothing分支处理, 但那里用的是vb6_ReleaseObject)
    // 这里检查target是否是早期绑定变量, 将vb6_ReleaseObject改为vb6_ComReleaseTyped

    // tB Interface 契约 (B04): Set ivref = obj → ivref = &(obj)->__iv_<I>
    // 薄指针 = 对象内该接口槽字段的地址 (D3); Nothing 已是 NULL, 这里不动.
    {
        std::string ivTargetLower = target;
        std::transform(ivTargetLower.begin(), ivTargetLower.end(), ivTargetLower.begin(), ::tolower);
        auto itIv = knownIvrefVars_.find(ivTargetLower);
        if (itIv != knownIvrefVars_.end() && value != "NULL") {
            if (const IfaceView* v = ivLookupIface(itIv->second)) {
                std::string ivValLower = value;
                std::transform(ivValLower.begin(), ivValLower.end(), ivValLower.begin(), ::tolower);
                const bool isClassVar = knownClassVars_.count(ivValLower) != 0;
                const bool isNewExpr = value.find("_New(") != std::string::npos;
                // B06a: RHS 也是接口变量 → 真 QueryInterface（VB6 跨接口 Set 的语义）。
                // QI 成功即已 AddRef（B06a 的 QI 实现负责），失败得到 NULL = Nothing。
                // 目标是裸标识符 (= 过程级接口变量) 才接计数; me->字段 等形式
                // 沿用 B04 的纯赋值改写, 生命周期留给后续批次.
                const bool plainIvTarget = !target.empty() &&
                    target.find_first_not_of(
                        "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_")
                        == std::string::npos;
                const bool isIvrefVar = knownIvrefVars_.count(ivValLower) != 0;
                if (plainIvTarget && isIvrefVar) {
                    const std::string ivCType = "vb6_ivref_" + cIdent(v->name) + "*";
                    const std::string ivOld = "__ivold" + std::to_string(ivrefSetSeq_);
                    const std::string ivGot = "__ivgot" + std::to_string(ivrefSetSeq_);
                    ivrefSetSeq_++;
                    c_.emitLine("{  /* tB Interface B06a: Set " + target + " = QueryInterface */");
                    c_.indent();
                    c_.emitLine(ivCType + " " + ivOld + " = " + target + ";");
                    c_.emitLine("void* " + ivGot + " = NULL;");
                    c_.emitLine("if (" + value + ") " + value + "->vt->QueryInterface(" + value +
                                ", vb6_iv_iid_" + cIdent(v->name) + ", &" + ivGot + ");");
                    c_.emitLine(target + " = (" + ivCType + ")" + ivGot + ";");
                    c_.emitLine("if (" + ivOld + ") " + ivOld + "->vt->Release(" + ivOld + ");");
                    c_.dedent();
                    c_.emitLine("}");
                    return;
                }
                if (isClassVar || isNewExpr) {
                    // D22-10: 类变量可能是 Nothing —— &NULL->__iv_<I> 算出来是个**非空**野地址
                    // (NULL + offsetof)，紧接着的 AddRef 就会解引用它。New 的返回值恒非空，
                    // 而且对它做三目会让 _New() 求值两次（凭空多创建一个实例），所以守卫只加在
                    // "裸类变量"这条路上。
                    const std::string ivField = "__iv_" + cIdent(v->name);
                    const std::string ivPtr = isClassVar
                        ? "((" + value + ") ? &(" + value + ")->" + ivField + " : NULL)"
                        : "&(" + value + ")->" + ivField;
                    if (plainIvTarget) {
                        // COM 口径: 存旧值 → 覆盖 → AddRef 新 → Release 旧.
                        // RHS 是 New 时那 1 次初始引用直接移交, 故不 AddRef.
                        const std::string ivCType = "vb6_ivref_" + cIdent(v->name) + "*";
                        const std::string ivOld = "__ivold" + std::to_string(ivrefSetSeq_++);
                        c_.emitLine("{  /* tB Interface B05: Set " + target + " (refcounted) */");
                        c_.indent();
                        c_.emitLine(ivCType + " " + ivOld + " = " + target + ";");
                        c_.emitLine(target + " = " + ivPtr + ";");
                        if (!isNewExpr) {
                            c_.emitLine("if (" + target + ") " + target + "->vt->AddRef(" +
                                        target + ");");
                        }
                        c_.emitLine("if (" + ivOld + ") " + ivOld + "->vt->Release(" + ivOld + ");");
                        c_.dedent();
                        c_.emitLine("}");
                        return;
                    }
                    value = ivPtr;
                } else {
                    diag_.error(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                        "Interface binding: Set " + target + " = <expr> needs a class instance,"
                        " New, or another interface value on the right side"
                        " (tB Interface " + v->name + ")");
                }
            }
        }
    }

    // tB Interface B06b: Set <工程类变量> = <接口变量> → 下行转换。
    // 用按类导出的 vb6_iv_from_iv_<C> 验明正身（薄指针所属实例的槽表是 C 那个 TU 的
    // static const 对象，使用点比不了地址 → 绝不能写裸强转）；命中才 AddRef ——
    // 类变量那一遍引用按 D21-5 的不变式永不释放，不加码的话接口侧 Release 到 0
    // 会把对象从这个类变量脚下抽走。
    {
        std::string dcTgt = target, dcRhs = value;
        std::transform(dcTgt.begin(), dcTgt.end(), dcTgt.begin(), ::tolower);
        std::transform(dcRhs.begin(), dcRhs.end(), dcRhs.begin(), ::tolower);
        auto itDcCls = knownClassVars_.find(dcTgt);
        const bool dcPlain =
            !target.empty() && target.find_first_not_of(
                "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_")
                == std::string::npos &&
            !value.empty() && value.find_first_not_of(
                "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_")
                == std::string::npos;
        if (itDcCls != knownClassVars_.end() && knownIvrefVars_.count(dcRhs) && dcPlain) {
            const std::string dcClsC = cIdent(itDcCls->second);
            c_.emitLine("{  /* tB Interface B06b: Set " + target + " = downcast of " + value + " */");
            c_.indent();
            c_.emitLine("void* __ivdown = vb6_iv_from_iv_" + dcClsC + "(" + value + ");");
            c_.emitLine(target + " = (vb6_cls_" + dcClsC + "*)__ivdown;");
            c_.emitLine("if (" + target + ") " + value + "->vt->AddRef(" + value + ");");
            c_.dedent();
            c_.emitLine("}");
            return;
        }
    }

    // P6.4: 接口引用赋值: Set ifaceRef = obj → ifaceRef = vb6_iface_IFoo_wrap(obj)
    {
        std::string targetLower = target;
        std::transform(targetLower.begin(), targetLower.end(), targetLower.begin(), ::tolower);
        auto itIface = knownIfaceVars_.find(targetLower);
        if (itIface != knownIfaceVars_.end()) {
            std::string ifaceName = itIface->second;
            std::string ifaceType = "vb6_iface_" + cIdent(ifaceName);
            // 如果右侧值包含_New()或是指向类实例的变量, 包装为接口引用
            if (value.find("_New()") != std::string::npos) {
                value = ifaceType + "_wrap(" + value + ")";
            } else {
                // 简单变量引用: value可能是类实例指针变量名
                std::string valLower = value;
                std::transform(valLower.begin(), valLower.end(), valLower.begin(), ::tolower);
                if (knownClassVars_.find(valLower) != knownClassVars_.end()) {
                    value = ifaceType + "_wrap(" + value + ")";
                } else {
                    // P6.4+: RHS 是 Variant (ByRef 参数/表达式), 只含对象引用.
                    // 需唯一实现类才能安全 cast:
                    //   vb6_iface_IFoo_wrap((vb6_cls_Impl*)vb6_VariantToObjectVal(<var>))
                    // 生成的串已含 "vb6_VariantToObjectVal", Fix 038b-6 下方守卫会跳过,
                    // 不会二次转换. 无唯一实现类时无法推断 → 报编译错误 (绝不静默误编).
                    bool rhsIsVariant = cExprIsVariant(value);
                    // ByRef Variant 参数解引用形式 "(*vName)": cExprIsVariant 剥掉前导
                    // '(' 检测不到 ("*vName)" 无前缀命中). 用 AST 标识符兜底判定.
                    if (!rhsIsVariant && node.value
                        && node.value->kind == ASTNodeKind::IdentifierExpr) {
                        auto& idv = static_cast<IdentifierExpr&>(*node.value);
                        std::string vl = idv.name;
                        std::transform(vl.begin(), vl.end(), vl.begin(), ::tolower);
                        if (knownVariantVars_.count(vl)) rhsIsVariant = true;
                    }
                    if (rhsIsVariant) {
                        const std::string impl = singleImplementationClass(ifaceName);
                        if (impl.empty()) {
                            diag_.error(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                                "P6.4: Set <interface> = Variant 需要接口 '" + ifaceName +
                                "' 在全项目有唯一实现类, 当前无法推断转换类型; "
                                "可将 RHS 改为 objptr/类型化对象变量后重试");
                        } else {
                            value = ifaceType + "_wrap((vb6_cls_" + cIdent(impl) +
                                    "*)vb6_VariantToObjectVal(" + value + "))";
                        }
                    }
                }
            }
        }
    }

    // Fix 010r-16: Set 语句中, 当 LHS 是非左值的 COM 调用或 Property Get
    // (常见于 With-block 跨模块成员: Set .Request = value, Set dict.Item(k) = v)
    // 重写为 vb6_ComSetProp / vb6_ComSetPropArg / prop_set_ 调用.
    if (tryRewriteCOMLvalue(target, value, node.value.get(), /*isSet=*/true)) {
        return;
    }
