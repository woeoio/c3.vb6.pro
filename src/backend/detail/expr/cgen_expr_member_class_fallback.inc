// cgen_expr_member_class_fallback.inc - 函数体片段：类实例成员 fallback 与函数尾
// 由 src/backend/expr/cgen_expr_member.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1255~1493 行，逐行未改。

    // Fix <vbeclipse> rev8: 局部判据 —— `node` 是"取某个类成员的属性", 且该成员
    // 的返回类型是 `As Object` (C 侧 void*)。定义在此处是因为要用到 node.object
    // 的 MemberAccessExpr 形态拆解, 而调用点在下方 class_fallback 的通用兜底。
    // `View.View.hWnd` 的 node.object = MemberAccessExpr(View, `View`) ⇒ 宿主类
    // 走 inferClassTypeOfExpr(它的 object) = "View", 成员名 = "View"。
    // ⚠ node.object 是 ExprPtr (智能指针), 不是裸指针 —— lambda 形参用 ExprPtr。
    auto objIsAsObjectMemberV8 = [this](const ExprPtr& objAst) -> bool {
        if (!objAst) return false;
        if (objAst->kind == ASTNodeKind::MemberAccessExpr) {
            const auto& maV8 = static_cast<const MemberAccessExpr&>(*objAst);
            if (!maV8.object) return false;
            return classMemberReturnsAsObject(inferClassTypeOfExpr(*maV8.object),
                                             maV8.memberName);
        }
        // Fix 246c: **With 块内的 `.X` 是 WithMemberExpr, 不是 MemberAccessExpr** ——
        //   上面那条一个 kind 判断就把整族挡在门外, 于是 `.Picture.Handle` 在
        //   `With Button.ButtonMenus(i)` 里永远走不到 rev8 的 `vb6_ComGetProp` 分支,
        //   掉到最后的 `obj + "." + member` ⇒ C2224 "左侧必须具有结构/联合类型"。
        //   实测 ComCtlsDemo ToolBar.c 7147/7151/7153/7198/7202/7204 共 6 条
        //   (TbrButtonMenu.cls:137 `Property Get Picture() As IPictureDisp`)。
        //   宿主类改从 With 栈顶取 —— 与 inferClassTypeOfExpr 的 WithMemberExpr
        //   分支同一个来源, 不另立一套。
        if (objAst->kind == ASTNodeKind::WithMemberExpr) {
            if (withObjectInfoStack_.empty()) return false;
            const std::string& hostCls = withObjectInfoStack_.back().className;
            if (hostCls.empty()) return false;
            const auto& wm = static_cast<const WithMemberExpr&>(*objAst);
            return classMemberReturnsAsObject(hostCls, wm.memberName);
        }
        // Fix 262: **窗体的控件属性读**里也有返回对象的那几枚 —— 最典型是
        // `Picture` (`vb6_GetControlPicture(hwnd)` → void*, 即 IPictureDisp 的 C 面)。
        //   `ImagePicture.Picture.Handle` 的 node.object 是
        //   MemberAccessExpr(ImagePicture, Picture): 宿主不是类, 上面各分支都捡不到
        //   ⇒ 落 `obj + "." + member` → `vb6_GetControlPicture(hwnd).Handle`
        //   ⇒ C2224 "左侧必须具有结构/联合类型"。
        //   返回 true 后下游改发 `vb6_ComGetProp(<getter>, L"Handle")`
        //   (RTL 的 vb6_ComGetProp 对 HWND/宿主对象有专档, 见 vb6com_invoke.c:409)。
        //   实测 ComCtlsDemo PPStatusBarPanels.pag:1182/1323
        //   (`If Not ImagePicture.Picture Is Nothing Then If ImagePicture.Picture.Handle <> 0`),
        //   PPStatusBarPanels.c 1362/1548。
        if (objAst->kind == ASTNodeKind::MemberAccessExpr) {
            const auto& maC262 = static_cast<const MemberAccessExpr&>(*objAst);
            if (maC262.object && maC262.object->kind == ASTNodeKind::IdentifierExpr) {
                const std::string ctrlName262 = Symbol::toLower(
                    static_cast<const IdentifierExpr&>(*maC262.object).name);
                auto itC262 = knownFormControls_.find(ctrlName262);
                if (itC262 != knownFormControls_.end()
                    && getControlPropReadFn(itC262->second, maC262.memberName)
                           == "vb6_GetControlPicture") {
                    return true;
                }
            }
        }
        return false;
    };

    // Fix 010r-10: 类实例变量的成员访问 fallback
    // 需要区分: 数据字段(obj->member) vs 方法/属性(vb6_ClassName_MethodName(obj))
    // Fix 011r-1: 用 resolveClassMemberCall 精确解析 (避免原 vb6_cls_ 前缀bug
    // 同时处理Pattern D1 — 跨模块类Public数据字段访问)
    // Fix 014: 链式类成员访问 — 如 me.m_oSocket.Create(...)
    // emitExpr(node.object) 后 obj 可能是 "me->m_oSocket" 这样的 C 表达式,
    // 而 knownClassVars_ 中的 key 注册的只是简单名 "m_osocket" (不带 me-> 前缀),
    // 因此完整 objLower 在 map 中找不到. 此时提取 obj 尾部标识符
    // (即最后一个 "->" 或 "." 之后的标识符) 再做一次 fallback 查找,
    // 即可解析出 obj 的真实类类型, 让方法分发正确工作.
    {
        // Fix 026: 剥掉 (*name) 解引用外层 (ByRef class/UDT 参数 emit 形式为 (*name)).
        // ByRef class 在 C 中是 cls**, (*name) 得 cls*, 应走类成员路径 (obj->member).
        // ByRef UDT 在 C 中是 struct_t*, (*name) 得 struct_t, struct.field 仍正确.
        // 仅剥完整 (*X) 外层; 其他形式 (me->X / a.b / 直接 X) 不动.
        std::string objBase = obj;
        if (objBase.size() > 4 && objBase[0] == '(' && objBase[1] == '*'
            && objBase.back() == ')') {
            objBase = objBase.substr(2, objBase.size() - 3);
        }
        std::string objLower = objBase;
        std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
        auto itClassVar = knownClassVars_.find(objLower);

        // Fix 014: 链式成员访问 — 若完整 objLower 找不到, 尝试提取尾部标识符
        // 例: obj = "me->m_oSocket" → trailingLower = "m_osocket"
        //     obj = "vb6_x()->y"     → trailingLower = "y"
        //     obj = "(*p).field"     → trailingLower = "field"
        // Fix <vbeclipse>: 对象表达式是**函数调用结果** (vb6_List_Item(me->m_Tabs, ...))
        // → 禁止尾部标识符回退 (把 m_tabs 误当类实例 → .ViewId 发成 `->ViewId`
        // 且调用实参被丢 → C2039 + C2198). 调用结果整体交给 Variant/COM 通道
        // (P24-04/voidptr_com) 按 rvalue 处理. 仅 `.field` 形态 (结尾是字段名)
        // 才允许 Fix 014 启发式.
        bool objIsCallExpr090 = (!objLower.empty() && objLower.back() == ')');
        std::string trailingLower;
        if (!objIsCallExpr090 && itClassVar == knownClassVars_.end()) {
            size_t lastArrow = objLower.rfind("->");
            size_t lastDot = objLower.rfind('.');
            size_t start = std::string::npos;
            if (lastArrow != std::string::npos && lastDot != std::string::npos) {
                // 取后出现的边界: 对于 "->", '-' 在 pos, 跳过2字符到标识符; 对于 ".", '.' 在 pos, 跳过1字符
                start = (lastDot > lastArrow) ? lastDot + 1 : lastArrow + 2;
            } else if (lastArrow != std::string::npos) {
                start = lastArrow + 2;
            } else if (lastDot != std::string::npos) {
                start = lastDot + 1;
            }
            if (start != std::string::npos && start < objLower.size()) {
                // 跳过非标识符字符 ('>', '*', '(', ')', 空白 等)
                while (start < objLower.size()
                       && !isalnum((unsigned char)objLower[start])
                       && objLower[start] != '_') {
                    start++;
                }
                size_t endPos = start;
                while (endPos < objLower.size()
                       && (isalnum((unsigned char)objLower[endPos])
                           || objLower[endPos] == '_')) {
                    endPos++;
                }
                trailingLower = objLower.substr(start, endPos - start);
            }
            if (!trailingLower.empty()) {
                auto itTrailing = knownClassVars_.find(trailingLower);
                if (itTrailing != knownClassVars_.end()) {
                    itClassVar = itTrailing;
                }
            }
        }

        if (itClassVar != knownClassVars_.end()) {
            // Fix 090a: obj 若是 As New 自动实例化项目类变量 (knownNewVars_, C 存储 void*),
            // 继续访问其成员前需 cast 回具体类指针 ((vb6_cls_X*)obj), 否则 void* 上
            // 直接 -> 触发 C2223 / COM 误 dispatch.
            // 例: cIni 内 Dim FileStream As New cToolsStream → me->FileStream->UseLine
            //     (C2223). 仅当 As New 目标是工程内 Class (kind==Class) 时 cast;
            //     As New 外部 COM (ADODB.Stream, kind==ComClass) 走 COM dispatch 不 cast.
            std::string obj090 = obj;
            if (!trailingLower.empty()) {
                auto itNew090 = knownNewVars_.find(trailingLower);
                if (itNew090 != knownNewVars_.end()) {
                    auto* newClsSym090 = symTab_.lookupModule(itNew090->second);
                    if (newClsSym090 && newClsSym090->kind == SymbolKind::Class) {
                        obj090 = "((vb6_cls_" + cIdent(itNew090->second) + "*)" + obj + ")";
                    }
                }
            }
            // 对象是类实例变量 — 用resolveClassMemberCall查证成员身份
            std::string resolvedFn = resolveClassMemberCall(itClassVar->second, node.memberName);
            // 类虚表 (tB B08d): `Me.<槽>` 的发码落在这里 —— 上面"优先级2"那一支整块在
            // `node.object` 是 IdentifierExpr 的分支里, `Me` 是 MeExpr → 从来进不到那一支, 是本支
            // 按 emit 出来的对象文本 ("me") 认出类实例的。所以两处都得改写, 且这一处必须派发:
            // 派发不了就由 virtDispatchCallee 出诊断, 不留静默直调。
            const std::string directFnFb = resolvedFn;
            const std::string dispatchFnFb = virtDispatchCallee(
                itClassVar->second, node.memberName, obj090, /*mustDispatch=*/true, node.loc);
            if (!dispatchFnFb.empty()) resolvedFn = dispatchFnFb;
            if (!resolvedFn.empty()) {
                // 方法/属性调用: vb6_<className>_<prefix><memberName>(obj)
                // Fix 089h: 值上下文无括号方法引用 (x = obj.ReadLine 或参数内引用)
                // 只发 this → C2198 参数太少 (ReadLine 声明3参). 补默认参数
                // (同 Fix 015 链式值上下文 / With 类 Fix 044a). asCallCallee_
                // (带括号调用) 由外层 IndexOrCallExpr 经 split 补参, 此处不 pad.
                if (!asCallCallee_
                    && directFnFb.find("_prop_") == std::string::npos) {
                    std::vector<ParameterInfo> params89h;
                    bool isBuiltin89h = false;
                    if (findClassMemberCallParams(itClassVar->second, node.memberName,
                                                   params89h, isBuiltin89h)
                        && !params89h.empty() && !isBuiltin89h) {
                        std::string argList89h = obj090;
                        for (size_t i = 0; i < params89h.size(); i++) {
                            const auto& param = params89h[i];
                            argList89h += ", ";
                            std::string defVal89h;
                            if (param.hasDefaultValue && !param.defaultValueExpr.empty()) {
                                defVal89h = param.defaultValueExpr;
                            } else {
                                defVal89h = defaultValueForParam(param);
                            }
                            if (param.isByVal) {
                                argList89h += defVal89h;
                            } else {
                                std::string cType89h = mapType(param.type);
                                if (param.type == Vb6Type::Variant
                                    || param.type == Vb6Type::Empty
                                    || param.type == Vb6Type::Null
                                    || param.type == Vb6Type::Object) {
                                    argList89h += "&(" + cType89h + "){0}";
                                } else {
                                    argList89h += "&(" + cType89h + "){" + defVal89h + "}";
                                }
                            }
                        }
                        for (size_t i = 0; i < params89h.size(); i++) {
                            const auto& param = params89h[i];
                            if (param.isOptional && !param.isParamArray) {
                                argList89h += ", 0";
                            }
                        }
                        lastExpr_ = resolvedFn + "(" + argList89h + ")";
                    } else {
                        lastExpr_ = resolvedFn + "(" + obj090 + ")";
                    }
                } else {
                    lastExpr_ = resolvedFn + "(" + obj090 + ")";
                }
            } else {
                // 非方法/属性 → 数据字段访问: obj->member
                // (此时obj类型为 vb6_cls_<className>*, ->访问可正确编译)
                // Fix 023: 检查该字段是否是 void* (外部 COM 指针). 若是, 在注释中
                // 加入 "voidptr" 标记, 由外层 MemberAccessExpr 的链式 COM 检测2
                // (~line 1477 之后) 识别并切换为 COM dispatch.
                // 例: Db.Rs (Rs As New ADODB.Recordset → C 结构体中 void* 字段)
                //     内层 emit: "Db->Rs  /* class var .Rs field voidptr */"
                //     外层 .EOF 识别 voidptr → vb6_ComGetIntProp(Db->Rs..., L"EOF")
                bool isVoidPtr = false;
                if (classVoidFieldMap_) {
                    auto itV = classVoidFieldMap_->find(itClassVar->second);
                    if (itV != classVoidFieldMap_->end()) {
                        std::string memLower = node.memberName;
                        std::transform(memLower.begin(), memLower.end(),
                                       memLower.begin(),
                                       [](unsigned char c) { return (char)std::tolower(c); });
                        if (itV->second.count(memLower) ||
                            itV->second.count("m_" + memLower)) {
                            isVoidPtr = true;
                        }
                    }
                }
                // Fix 092x: COM 接口字段 (如 cHttpServerRequest.Header As Dictionary)
                // 由 driver 预扫描记入 classTypedFieldMap_, 值为 "COM:<Interface>" 前缀,
                // 而 classVoidFieldMap_ 不含 (isVoidFieldType 为真时被 continue 跳过,
                // 但内置 COM 类走 COM: 前缀分支). 此处一并视为 COM 指针字段并追加
                // "voidptr" 标记, 供外层 MemberAccessExpr (Fix 023, line 2437) 切换为
                // COM dispatch. 例: Request.Header.Exists("Cookie") — 此前无标记 →
                // 外层 Fix 088b inferClassTypeOfExpr 遇 COM: 前缀返回空串 → 落入
                // obj->member 结构体访问 → C2037 (cHttpServer.c 821).
                if (!isVoidPtr && classTypedFieldMap_) {
                    auto itT = classTypedFieldMap_->find(itClassVar->second);
                    if (itT != classTypedFieldMap_->end()) {
                        std::string memLower = node.memberName;
                        std::transform(memLower.begin(), memLower.end(),
                                       memLower.begin(),
                                       [](unsigned char c) { return (char)std::tolower(c); });
                        auto itF = itT->second.find(memLower);
                        if (itF == itT->second.end()) itF = itT->second.find("m_" + memLower);
                        if (itF != itT->second.end()
                            && itF->second.compare(0, 4, "COM:") == 0) {
                            isVoidPtr = true;
                        }
                    }
                }
                // 账 #305 §B172 (第六十三刀): 这一整段「拼字段访问 + 尾注」收进唯一出口
                //   classFieldReadExpr —— 尾注里的档位 (` variant` 那一个词, §B171) 与
                //   字段名规范化 (Fix 092p) 从此只有一个作者, 站点只递"是不是 voidptr"。
                lastExpr_ = classFieldReadExpr(obj090, itClassVar->second, node.memberName,
                                               isVoidPtr, "class var");
            }
        // Fix 089b: trailing 变量名匹配在 obj 含 '.' 时跳过 — obj 是 UDT/属性字段链
        // (如 uFile.Data) 时, 尾段 "data" 是字段名而非变量名. 若外部 COM 变量恰好
        // 与 UDT 字段同名 (如 cSSE/cCsv 的 Public Data As Dictionary), trailing 匹配
        // 会误把 UDT 值字段当 COM 指针 → obj->member (C2232: uFile.Data->X).
        // 含 '.' 的链交由 Fix 088c / Fix 085(inferUdtTypeOfExpr) 兜底, 生成正确 . 访问.
        // 形如 me->field / 裸变量名的 trailing (Fix 014 场景) 不受影响.
        } else if (knownTypedComVars_.count(objLower)
                   || (obj.find('.') == std::string::npos && !trailingLower.empty()
                       && knownTypedComVars_.count(trailingLower))) {
            lastExpr_ = obj + "->" + cIdent(node.memberName);
        } else if (node.object) {
            // Fix 088c: AST 层兜底推断 — obj 是 knownClassVars_ 未覆盖的类实例
            // 表达式文本 (prop_get 调用文本 vb6_cY_prop_get_Z(a)、类方法链返回、
            // WithMemberExpr 等). inferClassTypeOfExpr 走 AST 递归
            // (knownClassVars_/classTypedFieldMap_/方法返回类型表), 成功后按类成员
            // 分发, 避免生成 obj'.'member/obj'->'member
            // → C2039 (SendData 不是 vb6_cls_cWinsock 的成员等).
            std::string clsFromAst = inferClassTypeOfExpr(*node.object);
            if (!clsFromAst.empty()) {
                std::string resolvedFn =
                    resolveClassMemberCall(clsFromAst, node.memberName);
                if (!resolvedFn.empty()) {
                    // 类虚表 (ai/022 B08e-6 站点⑩): 接收者是"任意已发码头"。纯读
                    // (`me->m_x` 这类) 与 ⑤ 同法接派发; 含调用的 (链式/属性返回对象) 由
                    // mustDispatch=true 出 VB3027 —— 这一支此前一律静默直调, 正是
                    // D27-13 那类"编得过、虚调用消失"的形状。
                    std::string thisArg088c = "(void*)" + obj;
                    if (!suppressVirtDispatch_
                        && resolvedFn.find("_prop_let_") == std::string::npos
                        && resolvedFn.find("_prop_set_") == std::string::npos) {
                        const std::string dispatchFn088c = virtDispatchCallee(
                            clsFromAst, node.memberName, thisArg088c,
                            /*mustDispatch=*/true, node.loc);
                        if (!dispatchFn088c.empty()) resolvedFn = dispatchFn088c;
                    }
                    if (asCallCallee_) {
                        // 外层 IndexOrCallExpr/CallStmt 前置 this 并补参数
                        pendingChainObj_ = thisArg088c;
                        lastExpr_ = resolvedFn;
                    } else {
                        lastExpr_ = resolvedFn + "(" + thisArg088c + ")";
                    }
                    return;
                }
                // 类中无此成员 → 数据字段
                //
                // Fix 219: 但接收者是 **Variant 载体** (`vb6_VariantFromComResult(...)`
                // 这类, 静态类型是 vb6_VARIANT 值, 不是 vb6_cls_X*) 时, 它压根没有
                // 结构字段可读 —— `obj->member` 必是 C2039 "不是 vb6_VARIANT 的成员"。
                // VB6 语义是在该 Variant 所持 IDispatch 上做后期绑定的属性读。
                // 实测 ComCtlsDemo: CoolBar.ctl:874 `Me.Bands(i).Caption` — 内层
                // `Bands(i)` 已按 Fix 086 走了晚绑定的 Item 调用 (冷知识: 那正是
                // "项目类默认成员" 的既有出口), 于是外层就不能再当结构体字段,
                // 否则同一表达式里一半是 COM 一半是结构体 (整族 135+53 处)。
                if (isVariantCarrierExpr(obj)) {
                    lastExpr_ = lateBoundReadOnVariantCarrier(obj, node.memberName);
                    return;
                }
                // 账 #305 §B172 (第六十三刀): **这一站就是那条被漏掉的成员读**。接收者是
                //   "文本没进 knownClassVars_ 的类实例" —— 实测类形参 `Private Sub UseAll(h As
                //   K63Holder)` 的标识符发成 `(*h)`, 于是走到这里, 而它以前只发裸的
                //   `(*h)->PV` (无尾注) ⇒ §B171 那两半判据都看不见它 ⇒ InStr(h.PV, …) 把整只
                //   vb6_VARIANT 递给 BSTR 形参 = C2440 (两台实测)。现在与上面那一站同出口、
                //   同尾注、同 variant 词。voidptr 那一问刻意仍交 false: 本刀只补"认槽",
                //   不顺手改 COM 指针字段的链式路由 (那是另一问, 台账 §B172 记着)。
                lastExpr_ = classFieldReadExpr(obj, clsFromAst, node.memberName,
                                               /*isVoidPtr=*/false, "class var");
                return;
            }
            // Fix 085: 更深嵌套的 UDT 链末端字段访问 (Fix 031 只覆盖直接 UDT 变量).
            // object 仍是 UDT 表达式时, 若 member 为对象字段则追加标记,
            // 由外层 MemberAccessExpr (Fix 085 消费点) 转类方法/COM 路径.
            std::string udtChainType = inferUdtTypeOfExpr(*node.object);
            if (!udtChainType.empty()) {
                lastExpr_ = appendUdtObjFieldMarker(obj, udtChainType, node.memberName);
            // Fix <vbeclipse> rev8: 接收者是**类成员且其返回类型是 As Object** ——
            // `View.View.hWnd` / `l_View.View.Caption` 里中段是
            // `Property Get View() As Object` (View.cls:66), C 侧 vb6_View_prop_get_View()
            // 返回 void*。此前列在这里走 `obj + "." + member` → C2224
            // "左侧必须具有结构/联合类型" ×63 (ucPerspective.c 36 .hWnd /
            // ucFolder.c 11 .Caption / 6 .Icon / modSubClass.c 6 .hWnd)。
            // 语义 = IDispatch 上的属性读, 必须晚绑定。
            } else if (objIsAsObjectMemberV8(node.object)) {
                lastExpr_ = "vb6_ComGetProp(" + obj + ", L\"" + node.memberName + "\")";
            } else if (isVariantCarrierExpr(obj)) {
                // Fix <vbeclipse>: Variant 表达式 (vb6_List_Item(...) 返回 vb6_VARIANT)
                // 上的 COM 属性读取 → 取回对象后用 ComGetProp 晚绑定, 避免
                // `.ViewId`/`->ViewId` → C2039 (ucTabStrip.c:260 / ucPerspective.c:331).
                // List.* 属函数返回对象型Variant, rfind 兜底 variantPrefixes 表缺失项.
                // Fix 219: 表达式形状挪进 lateBoundReadOnVariantCarrier —— 上面
                // clsFromAst 分支的同一个形状要与这里逐字一致, 不能再各写一遍。
                lastExpr_ = lateBoundReadOnVariantCarrier(obj, node.memberName);
            } else {
                lastExpr_ = obj + "." + cIdent(node.memberName);
            }
        } else {
            lastExpr_ = obj + "." + cIdent(node.memberName);
        }
    }