// cgen_expr_member_voidptr_com.inc - 函数体片段：void* 字段的 COM 化与类方法链式调用
// 由 src/backend/expr/cgen_expr_member.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1064~1254 行，逐行未改。

    // Fix 023: void* struct 字段访问结果作为 COM 对象使用.
    // 内层 MemberAccessExpr 的 knownClassVars_ fallback 路径 (见本函数尾部) 在检测到
    // 当前访问的类字段是 void* 时, 在 emit 的表达式中加入 "voidptr" 注释标记:
    //   例: Db.Rs  →  "Db->Rs  /* class var .Rs field voidptr */"
    // 此处 outer MemberAccessExpr 检测该标记, 将其视为 void* IDispatch* 指针,
    // 设置 COM marker 让下游 (IndexOrCallExpr / BinaryExpr / resolveComValue) 通过
    // vb6_ComCall/vb6_ComGet*Prop 走 COM dispatch 通道:
    //   Db.Rs.EOF         → vb6_ComGetXXXProp(Db->Rs..., L"EOF")
    //   Db.Rs.FileExists()→ vb6_ComCallXXX(Db->Rs..., L"FileExists", args, argc)
    // 注: class_voidptr 的注释会被 MSVC 视为空白, 不影响 C 代码语义.
    if (obj.find("/* class var .") != std::string::npos) {
        if (obj.find(" voidptr */") != std::string::npos) {
            comObjExpr_ = obj;
            comMemberName_ = node.memberName;
            isComMarker_ = true;
            lastExpr_ = obj;  // 保持 void* 对象表达式供外层使用
            return;
        }
        // Fix 088b: typed 类对象字段链 — obj 是 "objExpr  /* class var .X field */"
        // (无 voidptr 注释), 字段是项目类实例 (vb6_cls_Y*). 外层 .Method/Property 应
        // 解析为 canonical 调用, 与 udt objfield 的 vb6_cls_ 分支同机制. 此前只有
        // void* (COM) 字段被消费, typed 字段 (如 oCallback.Socket As cTlsReMaster)
        // 落入主 fallback 生成 obj'.'member → C2039 (Accept 不是 vb6_cls_cTlsReMaster
        // 的成员). 例: oCallback.Socket.Accept requestId
        //   → vb6_cTlsReMaster_Accept((void*)oCallback->Socket, &requestId)
        std::string objExpr = obj;
        size_t mkPos = obj.find("  /* class var .");
        if (mkPos != std::string::npos) objExpr = obj.substr(0, mkPos);
        std::string fieldCls =
            node.object ? inferClassTypeOfExpr(*node.object) : "";
        if (!fieldCls.empty()) {
            std::string resolvedFn =
                resolveClassMemberCall(fieldCls, node.memberName);
            if (!resolvedFn.empty()) {
                std::string thisArg = "(void*)" + objExpr;
                // 类虚表 (tB, ai/022 B08e 第⑤站): `Me.m_h.Pick()` 与裸 `m_h.Pick()` 是同一个调用,
                // 后者早就按 __cvtbl 派发, 这一支却静默发 `vb6_<Base>_Pick((void*)me->m_h)` —— 基类型
                // 字段持有派生实例时切片回基类实现。接收者 (void*)<字段链> 是纯读, 派发串与 this 实参
                // 取**同一个** thisArg (拼两遍就是双副作用)。属性写方向不改写: 3.4b 无 Let/Set 槽。
                if (!suppressVirtDispatch_
                    && resolvedFn.find("_prop_let_") == std::string::npos
                    && resolvedFn.find("_prop_set_") == std::string::npos) {
                    const std::string dispatchFn88b = virtDispatchCallee(
                        fieldCls, node.memberName, thisArg, /*mustDispatch=*/true, node.loc);
                    if (!dispatchFn88b.empty()) resolvedFn = dispatchFn88b;
                }
                if (asCallCallee_) {
                    // 由外层 IndexOrCallExpr/CallStmt 把 thisArg 前置到参数首 (同 Fix 015)
                    pendingChainObj_ = thisArg;
                    lastExpr_ = resolvedFn;
                } else {
                    // Fix 090ak: 值上下文无括号裸调用 (If Db.Sql(s).Param(p).QueryParam Then) —
                    // 对象链尾段方法经 Fix 088b typed 字段链分支 (此分支先于 Fix 015/089h
                    // 命中) 到达此处, 只发 this 会 C2198 参数太少 (QueryParam 声明 5 参传 1 参).
                    // 按形参表补 Optional 默认值 (同 Fix 089h 2404 分支模板).
                    std::vector<ParameterInfo> params090;
                    bool isB090 = false;
                    if (resolvedFn.find("_prop_") == std::string::npos
                        && findClassMemberCallParams(fieldCls, node.memberName,
                                                     params090, isB090)
                        && !params090.empty() && !isB090) {
                        std::string argList090 = thisArg;
                        for (size_t i = 0; i < params090.size(); i++) {
                            const auto& param = params090[i];
                            argList090 += ", ";
                            std::string defVal090;
                            if (param.hasDefaultValue && !param.defaultValueExpr.empty()) {
                                defVal090 = param.defaultValueExpr;
                            } else {
                                defVal090 = defaultValueForParam(param);
                            }
                            if (param.isByVal) {
                                argList090 += defVal090;
                            } else {
                                std::string cType090 = mapType(param.type);
                                if (param.type == Vb6Type::Variant
                                    || param.type == Vb6Type::Empty
                                    || param.type == Vb6Type::Null
                                    || param.type == Vb6Type::Object) {
                                    argList090 += "&(" + cType090 + "){0}";
                                } else {
                                    argList090 += "&(" + cType090 + "){" + defVal090 + "}";
                                }
                            }
                        }
                        for (size_t i = 0; i < params090.size(); i++) {
                            const auto& param = params090[i];
                            if (param.isOptional && !param.isParamArray) {
                                argList090 += ", 0";
                            }
                        }
                        lastExpr_ = resolvedFn + "(" + argList090 + ")";
                    } else {
                        lastExpr_ = resolvedFn + "(" + thisArg + ")";
                    }
                }
                return;
            }
            // 类中无此方法/属性 → 类数据字段 (Public Field): objExpr->field
            // Fix 229: 接收者是 **Variant 载体** (VARIANT 值) 时没有结构体字段可走 ——
            // `objExpr->field` 必是 C2039 "field 不是 vb6_VARIANT 的成员"。
            // 与 Fix 219 (class_fallback / 本文件 221 行) 同口径: 改发后期绑定属性读。
            if (isVariantCarrierExpr(objExpr)) {
                lastExpr_ = lateBoundReadOnVariantCarrier(objExpr, node.memberName);
                return;
            }
            lastExpr_ = objExpr + "->" + cIdent(node.memberName);
            return;
        }
        // 字段类型推断失败 → 按结构体数据字段访问
        // Fix 229: 同上 —— 这一档比上面更"兜底", 更可能撞上 VARIANT 值接收者。
        if (isVariantCarrierExpr(objExpr)) {
            lastExpr_ = lateBoundReadOnVariantCarrier(objExpr, node.memberName);
            return;
        }
        lastExpr_ = objExpr + "->" + cIdent(node.memberName);
        return;
    }

    // Fix 015: 类方法链式调用 — db.Sql(s).Exec(...) 模式
    // node.object 是 IndexOrCallExpr, 其 callee 是 MemberAccessExpr.
    // 通过 inferClassTypeOfExpr 递归推断 node.object 求值后的类类型:
    //   - 从最内层 IdentifierExpr 起在 knownClassVars_ 中拿到 base 类名
    //   - 逐层用 getClassMethodReturnType 查方法的返回类型名 (Fix 015 在
    //     semantic_analyzer 给 Function/PropertyGet 补了 variableTypeName)
    // 若 node.object 确实返回类实例, 用 resolveClassMemberCall 分发外层成员.
    // 类方法返回的是 vb6_cls_X* 指针, 链上中间结果可直接作为下一段的 this
    // 指针透传 (嵌套函数调用), 无需复合字面量:
    //   db.Sql(s).Exec(args) → vb6_cDataBase_Exec((void*)vb6_cDataBase_Sql(db, s), args)
    // 注: 链上的每一段 (Sql→Param→Exec 等) 都会经此分支处理, 嵌套调用本身合法.
    if (node.object && node.object->kind == ASTNodeKind::IndexOrCallExpr) {
        std::string retClassName = inferClassTypeOfExpr(*node.object);
        if (!retClassName.empty()) {
            // 内层调用返回类实例 (vb6_cls_X* 指针) → 直接作为 this 参数透传.
            std::string wrappedObj = obj;

            std::string resolvedFn =
                resolveClassMemberCall(retClassName, node.memberName);

            // 类虚表 (ai/022 B08e-6 站点⑫): 接收者是内层调用的发射文本 (`obj.Sql(s)`),
            // 按构造含调用 → 求值两次就是双副作用, 派发不了。mustDispatch=true 把这里的
            // 静默直调换成一条 VB3027: 与其"编得过、虚调用消失"(D27-13 判例), 不如判死。
            if (!resolvedFn.empty() && !suppressVirtDispatch_
                && resolvedFn.find("_prop_let_") == std::string::npos
                && resolvedFn.find("_prop_set_") == std::string::npos) {
                const std::string dispatchFn015 = virtDispatchCallee(
                    retClassName, node.memberName, wrappedObj,
                    /*mustDispatch=*/true, node.loc);
                if (!dispatchFn015.empty()) resolvedFn = dispatchFn015;
            }

            // 关键决策点: 外层是否要把本节点当作 callee 调用?
            // - asCallCallee_=true: visit(IndexOrCallExpr)/visit(CallStmt) 会随后附加
            //   用户参数 + Optional 默认值填充. 此时只 emit 裸函数名, wrappedObj 走
            //   pendingChainObj_ 通道在 IndexOrCallExpr / CallStmt 里前置. 这样可
            //   正确生成 vb6_cDataBase_Exec(wrappedObj, def1, def2, ...) (5个参数).
            // - asCallCallee_=false: 上下文是值引用 (bX = obj.Sql(s).Prop), 没有外层
            //   call 来填补默认值. 此时只能退化为 func(wrappedObj) 形式 (只有 this 指针,
            //   默认参数缺失). 与非链式 obj.Method (无括号) 的 Priority 2 路径行为一致,
            //   同样需要后续统一改进.
            if (asCallCallee_) {
                if (!resolvedFn.empty()) {
                    pendingChainObj_ = wrappedObj;
                    lastExpr_ = resolvedFn;  // 裸函数名, 由外层补全
                } else if (isVariantCarrierExpr(wrappedObj)) {
                    // Fix 230: 接收者是 **Variant 载体** (VARIANT 值) 且本成员不是已知类方法时,
                    // `wrappedObj->member` 是对 VARIANT **结构体**做成员访问 ⇒ C2039
                    // "member 不是 vb6_VARIANT 的成员"; 而外层 IndexOrCallExpr 还要把实参
                    // 追加在这个"callee 文本"上, 实参也随之错位 (伴生 C2198)。
                    //   ComCtlsDemo ToolBar.ctl:1131
                    //     PropBagButtonMenus.WriteProperty …, StrToVar(Me.Buttons(i).ButtonMenus(ii).Key), …
                    //   Me.Buttons(i) 是集合 Item ⇒ VARIANT 值; .ButtonMenus(ii) 是它上面的
                    //   后期绑定成员访问 (**带下标**) ⇒ 以前发成
                    //     vb6_VariantFromComResult(vb6_ComCall(…, L"Item", …))->ButtonMenus()
                    //   ToolBar.c:1707 等, C2039×32 + C2198×32 (ToolBar.c 共 101 条)。
                    // VB6 语义就是 IDispatch 后期绑定 —— 交给 com_bind 那条
                    // `vb6_ComCall(obj, L"member", args, n)` 通路: com_bind.inc:6-7 消费
                    // isComMarker_ + comObjExpr_/comMemberName_ (与 generic_access.inc:228-234
                    // 的 void* 字段那一档同款)。对象先用 vb6_VariantToObjectVal 从 VARIANT 取指针。
                    // 判据收窄到"必然编译报错"的那一格 (Variant 载体 ∧ 无已知类方法 ∧ 作 callee),
                    // 故不改变任何当前能编过的组合。
                    pendingChainObj_.clear();
                    comObjExpr_ = "vb6_VariantToObjectVal(" + wrappedObj + ")";
                    comMemberName_ = node.memberName;
                    isComMarker_ = true;
                    lastExpr_ = wrappedObj;
                } else {
                    // 数据成员但被当作 callee 调用 — 罕见, 保持 wrappedObj->member 形式
                    pendingChainObj_.clear();
                    lastExpr_ = wrappedObj + "->" + cIdent(node.memberName);
                }
            } else {
                if (!resolvedFn.empty()) {
                    // 值上下文 (If Db.Sql(s).Param(p).QueryParam Then 等无括号裸引用):
                    // emit 完整 func(wrappedObj) 形式 — 但要补默认参数. 仅发 this 会
                    // C2198 参数太少 (QueryParam 声明5参传1参). 参数表经
                    // findClassMemberCallParams 精确取得 (Fix 033 优先级与
                    // resolveClassMemberCall 一致); 无参属性 Get 读 (params 空) 不受影响.
                    // Fix 089h: Optional/必选参数按默认值补全 (同 With 类 Fix 044a),
                    // 必选参数在 VB 无括号引用时非法, 但补默认值可保持可编译.
                    std::vector<ParameterInfo> params89h;
                    bool isBuiltin89h = false;
                    if (resolvedFn.find("_prop_") == std::string::npos
                        && findClassMemberCallParams(retClassName, node.memberName,
                                                   params89h, isBuiltin89h)
                        && !params89h.empty() && !isBuiltin89h) {
                        std::string argList89h = wrappedObj;
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
                        lastExpr_ = resolvedFn + "(" + wrappedObj + ")";
                    }
                } else {
                    // Fix 219: 类中无此成员 → 数据字段; 但 Variant 载体没有结构字段,
                    // `obj->member` 必是 C2039 —— 改走晚绑定属性读。
                    // 实测: CoolBar.c:1226/1232 CoolBar.ctl:874 `Me.Bands(i).Caption`
                    // → `vb6_VariantFromComResult(vb6_ComCall(…, L"Item", …))->Caption`。
                    // (上面 resolvedFn 非空的两支不动: 那里成员确实解析成了规范函数,
                    //  接收者若是载体另有其因, 不在这里一并改。)
                    if (isVariantCarrierExpr(wrappedObj)) {
                        lastExpr_ = lateBoundReadOnVariantCarrier(wrappedObj, node.memberName);
                    } else {
                        lastExpr_ = wrappedObj + "->" + cIdent(node.memberName);
                    }
                }
            }
            return;
        }
    }
