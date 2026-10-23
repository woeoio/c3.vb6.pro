// cgen_setlet_set_rhs.inc - visit(SetStmt&) 函数体片段：值侧语义（Variant 容器与包装 / WithEvents 事件连接）
// 由 src/backend/stmt/cgen_setlet.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 322~565 行，逐行未改。

    // Fix 038b-6 + 051 共用: 判定 LHS 是否是 Variant 容器
    // (Variant 变量 / Variant 数组元素 / UDT Variant 字段). Set 到 Variant 容器
    // 语义 = 把对象引用存进 Variant (拷贝/FromValue), 而非提取对象到 typed 指针.
    std::string checkName = target;
    if (checkName.substr(0, 4) == "me->") checkName = checkName.substr(4);
    if (checkName.size() > 4 && checkName[0] == '(' && checkName[1] == '*'
        && checkName.back() == ')') {
        checkName = checkName.substr(2, checkName.size() - 3);
    }
    std::string targetLower = checkName;
    std::transform(targetLower.begin(), targetLower.end(), targetLower.begin(), ::tolower);
    bool targetIsVariant = knownVariantVars_.count(targetLower) > 0;
    // Fix 090af: Variant 数组元素 (VB6_SA_AT(vb6_VARIANT, arr, i)) 也是 Variant 容器
    if (!targetIsVariant && checkName.find("VB6_SA_AT(vb6_VARIANT,") == 0) {
        targetIsVariant = true;
    }
    // Fix 091p: 类字段 (me->field) 的 Variant 判定 — knownVariantVars_ 每过程
    // clear(), 类模块字段不在其中 (091m 回灌只覆盖标准模块级变量) → Set
    // m_vUserData = Value 被误判为 typed 目标 → VariantToObjectVal 提取 →
    // C2440 (cWinsock 172: void* → vb6_VARIANT). 字段访问带 me-> 前缀,
    // 用字段集合单独判定, 不污染裸名集合.
    if (!targetIsVariant && target.rfind("me->", 0) == 0
        && classVariantFields_.count(targetLower) > 0) {
        targetIsVariant = true;
    }
    // Fix 084n: UDT 的 Variant 字段 (如 ZipFileInfo.SourceFile As Variant)
    // Set .SourceFile = obj → 也需 vb6_VariantFromValue 包装对象指针 (void*→vb6_VARIANT C2440)
    // ai/022 B08f-1 (D37): 但 `As <项目类>` 的 UDT 字段在语义层**也**记成 Variant，而它在 C 里是
    // `vb6_cls_X* h;` —— 再包一层就是把 vb6_VARIANT 结构体灌进指针字段 (C2440, `Set u.h = obj`)。
    // 所以按字段的真实 C 类型否决：只有不是 `vb6_cls_*` 时才当它是 Variant 容器
    // (`void*` 的 COM/Collection 字段维持 Fix 084n 原样，那条路才是它当初的动机)。
    if (!targetIsVariant) {
        const std::string udtFldC = udtFieldCTypeOfTarget(node.target.get());
        if (udtFldC.rfind("vb6_cls_", 0) != 0) {
            targetIsVariant = (inferUdtFieldVb6Type(node.target.get()) == Vb6Type::Variant);
            // <vbeclipse> <VBFlexGridDemo>: 判据看 mapTypeRef 那份**具体 C 类型**,
            // 与 arg_emit / arg_variant 三条 wrap guard 同一份权威 —— VTableHandle.bas
            // `Set VTableIPAOData.OriginalIOleIPAO = This` 一条, 字段声明
            // `As OLEGuids.IOleInPlaceActiveObject`, TLB 命中让 inferUdtFieldVb6Type
            // 兜底成 Variant → 装箱 → 字段 typed ptr 收 VARIANT → C2440。
            // 但"As Variant"字段本身 resolveTypeName 也给 vb6_VARIANT, 装箱是必须的,
            // 不能一并挡掉。所以判据不能只看 typeRefName 是否非空, 要问那份 C 类型。
            if (targetIsVariant) {
                const std::string fldTRN = udtFieldTypeRefNameOfTarget(node.target.get());
                if (!fldTRN.empty()
                    && cTypeForDeclaredTypeName(fldTRN) != "vb6_VARIANT") {
                    targetIsVariant = false;
                }
            }
        }
    }

    // Fix 038b-6: Set 语句中 Variant 值 → 对象引用提取
    // 当 RHS 是 Variant (如 vb6_VariantFromStackVARIANT, vb6_VariantArrayGet,
    // Variant 变量等) 而 LHS 是 typed 对象指针时, 用 vb6_VariantToObjectVal 提取.
    // 使用 cExprIsVariant (C 字符串级) + knownVariantVars_ 检测.
    // Fix 090af: LHS 是 Variant 容器时跳过本步 — Set d(0) = d(0).Root 的 RHS 是
    // VariantFromComResult(ComGetProp(...)) (Variant 值), 应直接拷贝进容器
    // (051 的 vb6_VariantFromValue _Generic Identity 安全处理), 而非先 ToObjectVal
    // 提取 void* 再重包 (多余且二次语义).
    // Fix <vbeclipse> 2026-10-07 (P24-10b 的漏网): `Set o = <晚绑定默认成员>`
    //   (`Set o = d("x")`, d 是晚绑定 Object) 的 RHS 发出来是 vb6_ComCallByDispid(...)
    //   —— 它返回 **VARIANT*** (void*), 不是 IDispatch*。原样塞进对象槽
    //   (`void* o = ...`) 后, 后续任何 .成员 / CallByName 都会把那只 VARIANT 结构体
    //   当 IDispatch 的 lpVtbl 解引用 → 0xC0000005。
    //   实测: tests_github/t2_cases	est_exe_com_bridge —— 打印到 "EXEB:3 got object"
    //   之后 `CallByName o, "Ping", VbMethod` 段错误 (退出码 139; 夜间 GitHub Tests
    //   T0+T1+T2 run 37543812934 的 t2 shard 2/2 就是这一条)。
    //   与 cgen_expr_member_generic_access.inc 的 ByDispid 成员访问**同口径**:
    //   先 vb6_VariantFromComResult 把 VARIANT* 解成 by-value VARIANT, 再交给下面
    //   既有的 vb6_VariantToObjectVal 取 IDispatch* (工程类槽还会再剥一层)。
    //   只动 ByDispid 这一族 —— 具名晚绑定那条路 (`d.Item("x")`) 早在发码时就按目标
    //   类型挑了 vb6_ComCallObject, 实测已发成 `o = vb6_ComCallObject(d, L"Item", ...)`。
    if (!targetIsVariant && value.compare(0, 20, "vb6_ComCallByDispid(") == 0) {
        value = "vb6_VariantFromComResult(" + value + ")";
    }

    {
        bool valueIsVariant = cExprIsVariant(value);
        if (!valueIsVariant && node.value && node.value->kind == ASTNodeKind::IdentifierExpr) {
            auto& id = static_cast<IdentifierExpr&>(*node.value);
            std::string idLower = id.name;
            std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
            if (knownVariantVars_.count(idLower)) valueIsVariant = true;
        }
        if (valueIsVariant && !targetIsVariant
            && value.find("vb6_VariantToObjectVal") == std::string::npos) {
            // Fix <vbeclipse> rev5: LHS 是**工程类**变量 (vb6_cls_X*) 时, 光取
            // pdispVal 不够 —— Variant 里装的是 vb6_ComObject 包装器, 必须再经
            // vb6_ComObject_GetInstance 剥出裸实例, 否则把包装器当类指针用。
            // 实测 play78.exe: `Set l_Perspective = m_Perspectives.Item(PerspectiveId)`
            // (ucPerspective.ctl ShowPerspective) 生成 `l_Perspective =
            // vb6_VariantToObjectVal(vb6_List_Item(...))`, 随后
            // `l_Perspective.Folders.Contains(...)` 崩在 vb6_List_Contains+0xB6,
            // vb6_UBound 打出 `sig=0x640045 et=7602281` —— 那是 BSTR "Ed..." 的
            // 字节, 因为 m_Folders 槽其实读到的是前一个成员 (m_PerspectiveId) 的 BSTR。
            // 判据 = "LHS 这个槽收不收工程类实例指针", 两个来源:
            //   ① knownClassVars_ —— 局部/模块级/形参的工程类变量 (含 `Dim x As C`)
            //   ② currentReturnCType_ —— `Property Get Foo() As C` 里的 `Set Foo = ...`,
            //      LHS 是 vb6_ret_Foo, **不在**任何变量表里 (实测 ucPerspective.ctl
            //      `Property Get ActivePerspective() As Perspective` 的
            //      `Set ActivePerspective = Perspectives.Item(...)` 就是这一型:
            //      未剥壳 → ActivePerspective 返回包装器 → ShowPerspective 里
            //      `l_Perspective.Folders` 读垃圾 → "No active perspective!" 模态框)
            // 与 cgen_base_type.cpp 的 deliverToClassSlot 同一口径 (那边已经在剥)。
            // void* (Object) 目标维持原样: 那是真 IDispatch 语义, 不该剥。
            std::string setLhs = target;
            if (setLhs.rfind("me->", 0) == 0) setLhs = setLhs.substr(4);
            std::string setLhsLower = setLhs;
            std::transform(setLhsLower.begin(), setLhsLower.end(), setLhsLower.begin(), ::tolower);
            auto itCls = knownClassVars_.find(setLhsLower);
            if (itCls != knownClassVars_.end() && !itCls->second.empty()) {
                value = "(vb6_cls_" + cIdent(itCls->second) + "*)vb6_ComObject_GetInstance("
                        "vb6_VariantToObjectVal(" + value + "))";
            } else if (!currentReturnVar_.empty() && target == currentReturnVar_
                       && currentReturnCType_.compare(0, 8, "vb6_cls_") == 0) {
                value = "(" + currentReturnCType_ + ")vb6_ComObject_GetInstance("
                        "vb6_VariantToObjectVal(" + value + "))";
            } else {
                value = "vb6_VariantToObjectVal(" + value + ")";
            }
        }
    }

    // 账 #305 §B216 格② (#404): `Set <Variant 槽> = <工程内实例>` 这一站以前**连表都不问** ——
    // 裸实例指针直接交给下面的 _Generic, 指针臂把它投成对象档。第一百一十八刀的真跑读数
    // (E03/E11/E15, .build/k820out/): vt=9 看着像 VT_DISPATCH 而 pdispVal 是一枚裸 `vb6_cls_X*`
    // ⇒ TypeName 答 "Object"、晚绑定 stderr `vb6_ComCall: method "Twice" not found` + 空答。
    // 现在与泛对象槽问**同一位主人** (唯一发码处 boxIntoLateSlot): 装箱成功交回的是包装器 (void*),
    // 下面那一站的 _Generic 正好走 vb6_VariantObject = 真 IDispatch; 交回原样的形态 (认不出工程类 /
    // 值自己就是槽 / 宿主句柄 / 已剥壳) 一字不变。
    if (targetIsVariant && node.value) {
        value = boxIntoLateSlot(value, node.value.get());
    }

    // Fix 051: Set 语句中 void* → vb6_VARIANT 包装
    // 当 LHS 是 Variant 变量 (如 vb6_ret_Xxx, Dim x As Variant) 而 RHS 是
    // void* (对象指针, 如 vb6_ComCallObject, 类字段访问, vb6_VariantToObjectVal 等) 时,
    // 用 vb6_VariantFromValue 包装. vb6_VariantFromValue 是 _Generic 宏:
    //   - void* → vb6_VariantObject (包装对象指针)
    //   - vb6_VARIANT → vb6_VariantIdentity (no-op, 安全)
    // 因此始终包装是安全的, 只需避免对已包装的表达式双重包装.
    // Fix 090af: targetIsVariant 判定上移共用 (含 Variant 数组元素 LHS).
    if (targetIsVariant
        && value.find("vb6_VariantFromValue(") != 0
        && value.find("vb6_VariantFromComResult(") != 0
        && value.find("vb6_VariantBool(") != 0) {
        // 装箱走 boxToVariant 这一权威 (按 VB 类型选档; 未知才让 _Generic 猜)
        value = boxToVariant(node.value.get(), value);
    }

    // Fix 098: Set 左值是函数调用表达式 → 非左值 (C2106), 且多半缺参 (C2198).
    // VB6 语义: Set FnName = obj 在 FnName 不是当前函数自身时, 是对**别的函数
    // 返回值变量**的误写, VB6 编译器接受但赋值不改变本函数返回值. 等效生成:
    // 求值 RHS 后丢弃 (语义 = 该 Set 不影响任何可见状态).
    // 来源: ToolsJs.NewObj 内 "Set NewArr = MSSC.Eval(\"{};\")" (vbman 源码
    // 疑似笔误, NewArr 是 ParamArray 函数, 生成为 vb6_ToolsJs_NewArr() = ...)
    // → C2198+C2106 (ToolsJs.c:45).
    // 注: 左值 = 当前函数名时 M22 判定已在 emitExpr(IdentifierExpr) 转为
    // currentReturnVar_ (vb6_ret_NewObj), 不会落入本分支.
    {
        bool targetIsCallExpr = !target.empty() && target.back() == ')'
            && target.find('(') != std::string::npos
            && target.compare(0, 4, "me->") != 0
            && target.find("->") == std::string::npos
            && target.find(".") == std::string::npos
            && target.find("[") == std::string::npos
            && target.find("VB6_SA_AT(") != 0
            && !isVariantSlotElemRead(target)
            && target.find("vb6_PA_Get") != 0
            && target.find("prop_get_") == std::string::npos
            && target.find("prop_let_") == std::string::npos
            && target.find("prop_set_") == std::string::npos;
        if (targetIsCallExpr) {
            // 左括号前必须全是标识符字符 (排除 wb6_ComCall(...) 等包装调用已
            // 由前面的 tryRewriteCOMLvalue/链式写分支处理的情况兜底)
            size_t lp = target.find('(');
            bool isPlainFn = true;
            for (size_t ci = 0; ci < lp; ci++) {
                char ch = target[ci];
                if (!std::isalnum(static_cast<unsigned char>(ch)) && ch != '_') {
                    isPlainFn = false;
                    break;
                }
            }
            if (isPlainFn && lp > 0 && target.compare(0, 4, "vb6_") == 0) {
                c_.emitLine("(void)(" + value + ");  /* Set to non-lvalue function ref (Fix 098, VB6 no-op) */");
                return;
            }
        }
    }

    // Fix 179a: Set <Object(void*)目标> = New <工程类> → 创建点包成 IDispatch.
    // 177b 的非遮蔽扩展 (裸结构体指针进 void* 后 ComCall 把首字段当 vtable → AV,
    // 2026-09-21 CI VBMAN cTlsSocket 实证). void* 目标的表示是 IDispatch*
    // (vb6_ComCall 解 lpVtbl / vb6_ReleaseObject Release), 不是 VARIANT*;
    // 不能用 vb6_ComPack_<X>() (返回 VARIANT*, 仅作 COM 实参).
    // ⚠ 账 #305 §B216 格①: 这一句**发码**已经搬进唯一发码处 boxIntoLateSlot (cgen_base_type.cpp),
    // 那一形认得 New 就交拥有型; typed 目标照旧保持裸指针, Variant 目标走上面格② 那一站。
    bool fix179bUnwrapped = false;
    {
        // 「这枚目标槽是不是后期绑定的泛对象槽 (void* 里该装 IDispatch*)」——
        // 唯一发码处与登记表那一问**同一问**，所以收成一条 lambda 问一次:
        //   me-> 字段 → classVoidFieldMap_ (driver 预扫描, 含 As Object / As New)
        //   其余      → knownObjectVars_   (本过程登记过的晚绑定变量/参数)
        auto lateBoundSlotTarget = [&]() -> bool {
            if (target.rfind("me->", 0) == 0) {
                if (!classVoidFieldMap_) return false;
                auto itV = classVoidFieldMap_->find(moduleName_);
                return itV != classVoidFieldMap_->end()
                    && itV->second.count(targetLower) > 0;
            }
            return knownObjectVars_.count(targetLower) > 0;
        };

        // 进站时的值 (下面几条臂会改写 value, 登记表那一问要吃**原样**那一枚)
        const std::string rawSrcForReg = value;

        bool wrappedBy179a = false;   // 行内 New 那一形交回**拥有型**包装器
        bool boxedHere = false;       // 装箱腿: 借用型包装器
        // 账 #305 §B216 格①: 两臂归一 —— Fix 179a 原来在这一处手抄的
        // `(vb6_cls_<X>_New())` 文本判据与那句拥有型发码一起搬进了唯一发码处 boxIntoLateSlot
        // (它按同一形认出新实例、交拥有型, 其余交借用型), 站点这一处只剩**一问**:
        // 这枚目标槽是不是后期绑定的泛对象槽。表外那一档 (工程类不占运行期 coclass 表那行)
        // 两臂现在也答同一个: 交**原样** —— 台账 #443 原先写的「归一成交 Nothing」已由本机
        // A/B 推翻, 两处实物与理由记在 cgen_base_type.cpp 那一格 (尺 .build/k842_ab.py)。
        if (!targetIsVariant && node.value && lateBoundSlotTarget()) {
            value = boxIntoLateSlot(value, node.value.get());
            // 下面那张登记表要问「这一站刚交的是不是包装器」: 唯一发码处交出的只有这两枚形状
            // (前缀判据用 rfind(…, 0), 不写长度字面量 —— 那一族的坑是恒假也不报错)。
            boxedHere = value.rfind("(void*)vb6_ComObject_FromBorrowedInstance(", 0) == 0;
            wrappedBy179a = value.rfind("(void*)vb6_ComObject_FromInstance(", 0) == 0;
        }

        // 登记表的口径: wrappedComVars_ 记的是「这个名字**最后一次** Set 进去的是不是
        // 一枚包装器」—— 只有那一次交的是包装器才许登记，交了别的东西 (裸实例 / 外部
        // IDispatch / Nothing) 一律注销。收不拢就会留一枚过期答案给三处读者:
        //   Set o = New Foo          → 登记 (179a 交的是包装器)
        //   Set o = CreateObject(…)  → 这一站既不装箱也不注销 ⇒ 表里还挂着"o 装着包装器"
        //   If o Is Nothing          → 拿着旧答案去 GetInstance 一枚活着的外部 IDispatch
        //                              ⇒ 认不出 vtable ⇒ 交回 NULL ⇒ **真对象被判成 Nothing**
        // 原来只有 `Set x = Nothing` 一条注销, 那本来就是这条规则的一个特例。
        // 复制 (`Set v2 = v1`, v1 装着包装器) 也算"最后一次交的是包装器" —— 槽里换的
        // 就是同一枚指针，不登记会让读者② / ③ 在 v2 上漏一次剥壳。
        if (!targetIsVariant) {
            bool holdsWrapper = wrappedBy179a || boxedHere
                || (lateBoundSlotTarget() && varHoldsComWrapper(rawSrcForReg));
            if (holdsWrapper) wrappedComVars_.insert(targetLower);
            else              wrappedComVars_.erase(targetLower);
        } else {
            // Variant 槽里的值是 vb6_VARIANT 那份存储，不是裸包装器指针 —— 名字不许进
            // 登记表 (否则读者② / ③ 会把它当指针剥壳)。`Set v = Nothing` 那半直译自这里。
            wrappedComVars_.erase(targetLower);
        }

        // 账 #305 §B187 装箱腿的读者①: Set <typed 工程类**局部/形参**目标> = <装箱过的
        // 泛对象变量> → 剥回裸实例。Fix 179b 只接了 me-> 字段那一格 (且限类模块), 局部
        // 那一格漏着: 装箱腿把这枚变量的内容从裸实例换成了包装器, 原样赋给 typed 目标
        // 就是把包装器当实例用 (读字段整排错位)。判据问 varHoldsComWrapper (唯一那张
        // 登记表), 目标类名问 knownClassVars_ (与 Fix 084g / deliverToClassSlot 同一条)。
        {
            std::string srcBoxed = value;
            if (srcBoxed.size() >= 2 && srcBoxed.front() == '(' && srcBoxed.back() == ')') {
                srcBoxed = srcBoxed.substr(1, srcBoxed.size() - 2);
            }
            if (!targetIsVariant && varHoldsComWrapper(srcBoxed)
                && target.rfind("me->", 0) != 0) {
                auto itTgtCls = knownClassVars_.find(targetLower);
                if (itTgtCls != knownClassVars_.end() && !itTgtCls->second.empty()) {
                    value = "(vb6_cls_" + cIdent(itTgtCls->second) + "*)"
                          + unwrapComVarExpr(srcBoxed);
                }
            }
        }
        // Fix 179b: Set <typed工程类目标> = <179a包装的Object变量> → 解包回裸实例
        // (仅 wrappedComVars_ 登记的变量; 限类模块, .bas 不 include vb6comserver.h)
        std::string src179 = value;
        if (src179.size() >= 2 && src179.front() == '(' && src179.back() == ')') {
            src179 = src179.substr(1, src179.size() - 2);
        }
        if (isClassModule_ && target.rfind("me->", 0) == 0
            && !src179.empty()
            && src179.find_first_of(" ()+*/,") == std::string::npos
            && !targetIsVariant) {
            // 账 #305 §B187: 「这枚源变量此刻装着包装器吗」收成一条问句
            // (varHoldsComWrapper) —— 别名复制、类槽实参交付、`Is` 同一性那三站问的
            // 就是同一句话，这里不再手抄「lowercase + 剥 me-> + 查表」那一段。
            if (varHoldsComWrapper(src179)) {
                std::string tgtClass179;
                if (classTypedFieldMap_) {
                    auto itT179 = classTypedFieldMap_->find(moduleName_);
                    if (itT179 != classTypedFieldMap_->end()) {
                        auto itF179 = itT179->second.find(targetLower);
                        if (itF179 != itT179->second.end()
                            && itF179->second.rfind("COM:", 0) != 0) {
                            tgtClass179 = itF179->second;
                        }
                    }
                }
                if (!tgtClass179.empty()) {
                    c_.emitLine(target + " = (vb6_cls_" + cIdent(tgtClass179)
                        + "*)vb6_ComObject_GetInstance((void*)" + src179
                        + ");  /* Set (Fix 179b: wrapped Object -> typed instance) */");
                    fix179bUnwrapped = true;
                }
            }
        }
        // (原来这条 "Set Nothing 清除跟踪" 已并入上面那条"最后一次 Set 交的是不是包装器"
        //  —— Nothing 那一形本来就落在它的 else 分支上，两条并着就是一处问。)
    }

    if (!fix179bUnwrapped) {
        // Fix <vbeclipse> rev16: `Set <Variant 数组元素> = x` 走"接管"语义。
        // 原先只发裸 `slot = value` —— 既不释放槽里旧值 (漏), 又是 identity 浅拷贝
        // (与借用方共用同一只 BSTR / 同一个 Dispatch, 借用方一 Clear 槽即悬垂)。
        // 对照同一函数的非 Set 路径: 它至少发了 vb6_VariantClear + 浅拷贝;
        // 本条把两条路径并到 vb6_VariantAssign (Clear + 深拷贝 / AddRef)。
        // 实测 List.cls: `Set m_Items(i) = m_Items(i - 1)` 之后
        // `Set m_Items(i - 1) = Nothing` 会把两只槽共用的那只 BSTR 放掉。
        /* Fix 296 / 3b 最后一格: Set 语句的**右值**那一问同样要过实例入口。
           `Property Get Container` 的返回槽那句**也是 Set 语句** —— 行尾那个标记就是证据
           (目标侧是 vb6_ret_Container, 所以上面那道目标闸正确地没碰它, 但右值
           `Extender.Container` 还是那枚裸全局)。ComCtlsDemo 实测 66 处。
           换道放在发码**之前**, 下面 targetIsVariantArraySlot 那一支的 VariantAssign
           包装、以及写侧那一支的 (void*)(intptr_t) 都照旧套在外面 —— 与 4a 同款。 */
        value = extenderMemberReadExpr(value);
        bool targetIsVariantArraySlot =
            (target.compare(0, 7, "VB6_SA_") == 0
             && target.find("(vb6_VARIANT,") != std::string::npos);
        if (targetIsVariantArraySlot) {
            c_.emitLine("vb6_VariantAssign(&" + target + ", " + value + ");  /* Set */");
        } else if (target.rfind("vb6_Extender_", 0) == 0) {
            /* Fix 296 / 3a 的**镜像**, 但走的是**另一条**发码路径: Property Set 存取器体
               (`Public Property Set Container(ByVal Value As Object)`) 的那句赋值由
               visit(SetStmt&) 发出, 不是 visit(AssignmentStmt&) —— 所以 3a 那道闸
               (cgen_assign_host_pseudo.inc) 压根够不着, 66 处
               `vb6_Extender_Container = Value;` (line-tail Set marker not pasted: it would close this comment early) 一路活到今天。
               成员名直接取发射符号的后缀 = 表里那一行的 rtl 列, ExtenderField 按
               _wcsicmp 匹配, 所以 `vb6_Default` 那行 name/rtl 不同拼法也对得上。 */
            c_.emitLine("vb6_UC_SetExtenderMember(L\""
                        + escapeWideCString(target.substr(std::string("vb6_Extender_").size()))
                        + "\", (void*)(intptr_t)" + value + ");"
                        "  /* Extender.<prop> = v (per-instance, Set stmt) */");
        } else {
            // 账 #305 §B218 (甲) 腿②: 目标是一格 typed 工程类槽 ⇒ 走那位主人。
            // 类名两问按 #408 定的次序排: 先问 AST (作用域感知), 落空才问按名那张表。
            // 拿不到类名时维持今天那条裸拷贝。**下标形态从第一百四十刀起也进主人** (走地址
            // 那一臂, 见下面 elemTy140)，仍留在外面的只有**真调用形态** (`dict(key)`)：那种
            // target 不是左值，取不了址，而「下标算术里的副作用算几次」已由第一百三十九刀
            // 量成一句有答案的口径 (两栏都答一次) ⇒ 地址临时量那一臂就是它的答案。
            std::string clsOfSlot = node.target ? inferClassTypeOfExpr(*node.target)
                                                : std::string();
            if (clsOfSlot.empty()) {
                auto itSlot = knownClassVars_.find(targetLower);
                if (itSlot != knownClassVars_.end()) clsOfSlot = itSlot->second;
            }
            // 账 #305 §B218 第一百六十五刀: 槽装的是 legacy 接口视图 ⇒ 本主人不接手
            // (它只会写 `vb6_cls_<名>* 旧 = 目标`，而那一型从没为接口造过 = cl C2440)。
            // 视图那一档今天落回下面那条裸拷贝 —— 与**局部**声明那一路发的一字相同 (实测)。
            if (!clsOfSlot.empty() && slotIsLegacyIfaceView(targetLower)) clsOfSlot.clear();
            // 账 #305 §B218 (甲) 腿②': 目标是**出参槽** `(*p)` (ByRef 形参那座桥的写回站)。
            // 写进去的那一枚由本过程一位局部持有, 而腿④ 会在过程出口放手它 ⇒ 这一格不自己取走
            // 一枚引用, 调用者还看着它时它就已经拆了 (门 #763 实测 = b187 第五臂
            // probe exit=-1073740940, 两台同形)。类名问**右侧**: 出参槽没有类型档 (void*),
            // 只有 RHS 答得出。旧值仍不退 = 出参写回那一格泄漏, 与下标槽同在册。
            bool outSlotClaim = false;
            std::string clsOfRhs;
            if (target.rfind("(*", 0) == 0 && value.find('(') == std::string::npos) {
                clsOfRhs = node.value ? inferClassTypeOfExpr(*node.value) : std::string();
                outSlotClaim = !clsOfRhs.empty();
            }
            // 账 #305 §B218 (甲) 形A (第一百四十四刀): 出参槽那一格**由调用方持有**那一枚引用
            // (调用侧现在是把持有者自己的地址交出来的)，所以覆写 = 取走新的 + **当场退旧**，
            // 与第一百三十九刀⑤栏那句代理读数「落回调用者并当场退旧」同一条口径。
            // 类名问**形参声明档** (classOfProjectRefOutSlot)，不再问右侧那枚文本的形状 ——
            // 下面那条 outSlotClaim 是按「RHS 发码文本有没有括号」分臂的旧判据，同一条语句
            // 换种写法就换一套答案 (New / Nothing / (*q) 三形落裸写、只有裸标识符落腿②′)，
            // 它现在只剩 §B187 那一档 (出参槽是泛对象 / Variant) 还会走到，一字未动。
            const std::string outCls144 = classOfProjectRefOutSlot(target);
            // 账 #305 §B218 (甲) 刀3: 「不 claim」那一臂有两个成因 —— `New <类>`, 以及
            // 工程内过程交回来的那一枚 (自带对价)。问句只有一处: rhsTransfersProjectRef。
            // 第一百三十四刀给它补了第二档 (forViewSlot): 本站的槽装实例 ⇒ 报 false。
            bool rhsTransferred = node.value && node.value->kind == ASTNodeKind::NewExpr;
            if (!rhsTransferred && node.value) {
                rhsTransferred = rhsTransfersProjectRef(*node.value, false);
            }
            // 账 #305 §B218 第一百四十刀: 「带括号就不进主人」那条老闸对**数组元素**不成立 ——
            // VB6_SA_AT(…) 是产出左值的函数式宏 (Fix 113f)，第一百三十九刀量到它今天裸写
            // ⇒ 覆写不退旧值。改走主人的地址那一臂：元素档非空 = 这一形有可寻址的地址临时量，
            // 下标算术仍只算一次；真调用形态 (`dict(key)`) 元素档为空串 ⇒ 维持裸拷贝。
            const std::string elemTy140 = classArraySlotElemType(target);
            if (!clsOfSlot.empty() && !procClassFormals_.count(targetLower)
                && (target.find('(') == std::string::npos || !elemTy140.empty())) {
                emitClassSlotRefSet(target, clsOfSlot, value, rhsTransferred, false, elemTy140);
            } else if (!outCls144.empty()) {
                emitClassSlotRefSet(target, outCls144, value, rhsTransferred, false);
            } else if (outSlotClaim) {
                emitClassSlotRefSet(target, clsOfRhs, value, rhsTransferred, true);
            } else {
                c_.emitLine(target + " = " + value + ";  /* Set */");
            }
        }
    }

    // P6.5: WithEvents变量事件连接
    // Set obj = newInst -> 如果obj是WithEvents变量, 设置事件接收器
    {
        std::string targetLower = target;
        std::transform(targetLower.begin(), targetLower.end(), targetLower.begin(), ::tolower);
        auto itWE = knownWithEventsVars_.find(targetLower);
        if (itWE != knownWithEventsVars_.end()) {
            std::string sourceClass = itWE->second;
            std::string sinkName = "vb6_events_" + cIdent(sourceClass);
            // 查找源类符号，区分内部类与外部COM类
            auto* srcClsSym = symTab_.lookup(sourceClass);
            if (srcClsSym && srcClsSym->kind == SymbolKind::Class) {
                // 内部类 WithEvents: if (target) { static sink = {...}; target->events = &sink; }
                // 需要查找当前模块中是否有 obj_EventName 形式的事件处理器
                c_.emitLine("if (" + target + ") {");
                c_.indent();
                // 生成静态事件接收器实例
                c_.emitLine("static " + sinkName + " " + targetLower + "_sink = {");
                c_.indent();
                // handler: 指向当前对象(me)
                std::string handlerExpr = isClassModule_ ? "(void*)me" : "NULL";
                c_.emitLine(".handler = " + handlerExpr + ",");
                // 为源类的每个事件生成回调指针
                bool firstEvent = true;
                for (auto& evtName : srcClsSym->eventNames) {
                    std::string handlerName = target + "_" + evtName;  // VB6: obj_Click
                    std::string handlerLower = handlerName;
                    std::transform(handlerLower.begin(), handlerLower.end(), handlerLower.begin(), ::tolower);
                    // 查找处理器函数符号 (用lookup跨模块查找)
                    auto* handlerSym = symTab_.lookup(handlerName);
                    std::string cbField = ".on" + cIdent(evtName) + " = ";
                    if (handlerSym) {
                        // 生成包装函数名: vb6_evt_<EventName>_wrap_<varName>
                        std::string wrapperName = evtWrapperName(targetLower, evtName);
                        cbField += wrapperName;
                    } else {
                        cbField += "NULL";
                    }
                    if (!firstEvent || srcClsSym->eventNames.size() > 1) {
                        // 多事件时加逗号
                    }
                    c_.emitLine(cbField + ",");
                    firstEvent = false;
                }
                c_.dedent();
                c_.emitLine("};");
                // WithEvents 变量在 C 侧声明为 void* (类实例统一用对象指针表示),
                // 直接生成 "var->events" 会触发 C2223 ("->" 的左侧必须指向结构/联合)。
                // 必须先 cast 回具体类结构体指针, 与 cgen_expr.cpp 中 void* 成员访问同机制。
                c_.emitLine("((vb6_cls_" + cIdent(sourceClass) + "*)" + target + ")->events = &"
                            + targetLower + "_sink;");
                c_.dedent();
                c_.emitLine("}");
            } else if (srcClsSym && srcClsSym->kind == SymbolKind::ComClass && srcClsSym->comHasSourceIface) {
                // P13.23: External COM WithEvents
                std::string iidStr = srcClsSym->comSourceIfaceIid;
                std::string iidInit = emitGuidInitializer(iidStr);
                // Static declarations at function scope (before the if block)
                c_.emitLine("static int " + targetLower + "_evt_cookie = 0;");
                c_.emitLine("static const char* " + targetLower + "_evt_iid = \"" + iidStr + "\";");
                if (!iidInit.empty()) {
                    c_.emitLine("static const IID " + targetLower + "_evt_iid_struct = " + iidInit + ";");
                }
                c_.emitLine("if (" + target + ") {");
                c_.indent();
                c_.emitLine("if (" + targetLower + "_evt_cookie != 0) {");
                c_.indent();
                c_.emitLine("vb6_ComUnadvise((IUnknown*)" + target + ", " + targetLower + "_evt_iid, " + targetLower + "_evt_cookie);");
                c_.emitLine(targetLower + "_evt_cookie = 0;");
                c_.dedent();
                c_.emitLine("}");
                if (srcClsSym->comSourceIfaceIsDispOnly) {
                    // dispinterface: use existing IDispatch sink
                    std::vector<std::string> dispids;
                    std::vector<std::string> callbacks;
                    for (auto& evtName : srcClsSym->eventNames) {
                        std::string handlerName = target + "_" + evtName;
                        auto* handlerSym = symTab_.lookup(handlerName);
                        if (handlerSym) {
                            std::string evtLower = Symbol::toLower(evtName);
                            auto itDispId = srcClsSym->comEventDispids.find(evtLower);
                            int dispid = (itDispId != srcClsSym->comEventDispids.end()) ? itDispId->second : 0;
                            dispids.push_back(std::to_string(dispid));
                            callbacks.push_back(comEvtWrapperName(targetLower, evtName));
                        }
                    }
                    if (!dispids.empty()) {
                        std::string dispidsVar = targetLower + "_evt_dispids";
                        std::string dispidsInit;
                        for (size_t di = 0; di < dispids.size(); di++) {
                            if (di > 0) dispidsInit += ", ";
                            dispidsInit += dispids[di];
                        }
                        c_.emitLine("static int " + dispidsVar + "[] = {" + dispidsInit + "};");
                        std::string callbacksVar = targetLower + "_evt_cbs";
                        std::string callbacksInit;
                        for (size_t ci = 0; ci < callbacks.size(); ci++) {
                            if (ci > 0) callbacksInit += ", ";
                            callbacksInit += "(void(*)(void*,VARIANT*,int,VARIANT*))" + callbacks[ci];
                        }
                        c_.emitLine("static void (*" + callbacksVar + "[])(void*,VARIANT*,int,VARIANT*) = {" + callbacksInit + "};");
                        std::string sinkVar = targetLower + "_comsink";
                        std::string iidStructRef = iidInit.empty() ? "NULL" : "&" + targetLower + "_evt_iid_struct";
                        // P13.23: handler = 当前类实例(me), 传入包装函数供事件处理器调用
                        std::string handlerExpr = isClassModule_ ? "(void*)me" : "NULL";
                        c_.emitLine("void* " + sinkVar + " = vb6_CreateEventSink(" +
                            dispidsVar + ", (void**)" + callbacksVar + ", " + std::to_string(dispids.size()) +
                            ", " + iidStructRef + ", " + handlerExpr + ");");
                        c_.emitLine("vb6_ComAdvise((IUnknown*)" + target + ", " + targetLower + "_evt_iid, " + sinkVar + ", &" + targetLower + "_evt_cookie);");
                    }
                } else {
                    // vtable source interface: use custom vtable sink
                    std::string sinkVar = targetLower + "_comsink";
                    std::string sinkHandlerExpr = isClassModule_ ? "(void*)me" : "NULL";
                    c_.emitLine("void* " + sinkVar + " = vb6_vsink_" + targetLower + "_create(" + sinkHandlerExpr + ");");
                    c_.emitLine("vb6_ComAdvise((IUnknown*)" + target + ", " + targetLower + "_evt_iid, " + sinkVar + ", &" + targetLower + "_evt_cookie);");
                }
                c_.dedent();
                c_.emitLine("}");
            }
        }
    }

    // P16: WithEvents控件变量赋值 - 无需额外操作
    // Set cmd = Command1 → cmd = ctrl_Command1 (HWND拷贝已在赋值行完成)
    // 事件通过WndProc的HWND匹配分发，不需要COM Sink/Advise
    {
        std::string targetLower = target;
        std::transform(targetLower.begin(), targetLower.end(), targetLower.begin(), ::tolower);
        auto itCtrl = knownWithEventsCtrlVars_.find(targetLower);
        if (itCtrl != knownWithEventsCtrlVars_.end()) {
            // 控件WithEvents变量已在赋值行 target = value; 完成HWND拷贝
        }
    }