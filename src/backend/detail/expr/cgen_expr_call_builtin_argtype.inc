// cgen_expr_call_builtin_argtype.inc - visit(IndexOrCallExpr&) 函数体片段：内建函数实参类型适配（Fix 108c）：裸 COM 结果解包 / ObjPtr 接口取值
// 由 src/backend/detail/expr/cgen_expr_call_pad_conv.inc 拆出（2026-09-19），
// 在函数体内被 #include，非独立编译单元。内容 = 拆分前该文件第 437~546 行，逐行未改。
// 账 #305 §B174 (第六十五刀) 起: 原 (b)「BSTR 形参数值实参」那一格不住这里了 (见下面 (b))。

    // Fix 108c: 内建函数实参的类型适配 (走到此处说明没有专门分支处理).
    //  (a) 数值型内建函数收到"无类型 COM 调用结果" (vb6_ComCall → void*, 指向
    //      VARIANT) 时先解包为数值, 否则 void* 传入 double 形参报 C2440
    //      (Charts 2020 ucChartBar SumSerieValues: vb6_Abs(vb6_ComCall(...))).
    //  (b) 形参为 String 的内建函数收到数值实参时按 VB6 隐式转换转 BSTR —— 这一格原先靠
    //      本站的一张 13 枚名单做**文本替换**, 第六十五刀起整段撤掉: 同一条规则早就住在
    //      cgen_expr_call_arg_emit.inc 的 Fix 175 (按形参表问位、形状走 wrapToBSTR), 两位主人
    //      合一之后名单本身就没剩什么可管的了 —— 那 13 枚**全部带类型化形参表** (实测),
    //      本站拿到实参时它们已被 Fix 175 折过, 名单自己那条「已含 vb6_CStr 就跳过」的护栏
    //      把它们全部放行。留下的缺口是**没有形参表**那一族 (InStr) ⇒ 补在 Fix 175 那一站。
    //      搬的另一层理由: 本站拿到的是**已拼成一整串的 argList**, 找回第 N 个实参只能靠
    //      "从上次命中往后数子串"这种对齐 (replaceInArgList108c 现存那两格也共用它) ——
    //      实参文本是别处生成的, 一旦重复出现或内含同款子串, 对上的就不是那一位。
    //      这一条本刀的 A/B 里有一枚实物 (b171_slotarg / b172_paramslot 的**单文件** --emit-c 形状,
    //      x64+x86 共四份捕获): 上游那条臂已把槽解包成 `vb6_VariantToString(...)` (BSTR), 本站的
    //      护栏只认「文本里含 vb6_CStr」⇒ 认不出它已经是串, 于是**又折一次** ——
    //      `vb6_Left(vb6_CStrDbl((double)(vb6_VariantToString(vb6_VariantFromComResult(
    //                  vb6_ComGetProp(h, L"PV"))))), 4)`
    //      即把 BSTR 指针 cast 成 double 再格式化, 打出来的是指针数值。修后这一形不再二次折算
    //      (Fix 175 在解包之前就问过档位, 槽文本被 cExprIsVariant 挡在门外)。
    //      ⚠ 那一形在**整工程** (.vbp, 带着 .cls) 的发射里没有 —— 单文件发射时 h 认不出宿主,
    //      走的是晚绑定那一路; 门里钉的 b171/b172 读数因此一字不动 (实测四份 .vbp 捕获逐行相同)。
    {
        static const std::unordered_set<std::string> numericBuiltins108c = {
            "vb6_Abs", "vb6_Int", "vb6_Fix", "vb6_Sgn", "vb6_Sqr", "vb6_Sin",
            "vb6_Cos", "vb6_Tan", "vb6_Atn", "vb6_Log", "vb6_Exp", "vb6_Round",
            "vb6_CDbl", "vb6_CSng", "vb6_CInt", "vb6_CLng", "vb6_CByte",
            "vb6_CBool", "vb6_Hex", "vb6_Oct",
        };
        bool argsFixed108c = false;

        // argList 在更早处已按 args 拼好, 其后还有"补默认参数"分支向 argList 追加
        // (如 vb6_Replace 的 ", 1, -1, 0"). 因此不能整体重建 argList —— 只能就地把
        // 被改写的实参文本替换掉, 否则会丢掉默认参数 (曾导致 vb6_Replace 只剩 3 参).
        size_t searchFrom108c = 0;
        auto replaceInArgList108c = [&](const std::string& oldVal,
                                        const std::string& newVal) {
            size_t p = argList.find(oldVal, searchFrom108c);
            if (p == std::string::npos) return;
            argList.replace(p, oldVal.size(), newVal);
            searchFrom108c = p + newVal.size();
        };
        if (numericBuiltins108c.count(callee)) {
            for (auto& a : args) {
                if (a.compare(0, 12, "vb6_ComCall(") == 0
                            || a.compare(0, 20, "vb6_ComCallByDispid(") == 0) {
                    std::string wrapped =
                        "vb6_VariantToDouble(vb6_VariantFromComResult(" + a + "))";
                    replaceInArgList108c(a, wrapped);
                    a = wrapped;
                    argsFixed108c = true;
                }
            }
        }
        // Fix 108c-2: vb6_Val 形参为 BSTR, 收到"无类型 COM 调用结果"
        // (vb6_ComCall → void*, 指向 VARIANT) 时解包为字符串, 而非数值 —
        // 若走 numericBuiltins108c 的 vb6_VariantToDouble 包裹, double 传入
        // BSTR 形参报 C2440 (BalloonTooltips cTT.cls: Val(colTooltips(sKey))).
        if (callee == "vb6_Val") {
            for (auto& a : args) {
                bool isComResult = (a.compare(0, 12, "vb6_ComCall(") == 0
                                    || a.compare(0, 20, "vb6_ComCallByDispid(") == 0
                                    || a.compare(0, 12, "vb6_ComGetProp(") == 0);
                if (isComResult) {
                    std::string wrapped =
                        "vb6_VariantToString(vb6_VariantFromComResult(" + a + "))";
                    replaceInArgList108c(a, wrapped);
                    a = wrapped;
                    argsFixed108c = true;
                }
            }
        }
        // 账 #305 §B174 (第六十五刀): 原此处的「数值实参 → BSTR 形参」13 枚名单整段撤掉 ——
        // 同一条规则的唯一站点在 cgen_expr_call_arg_emit.inc 的 Fix 175, 位号/档位/形状三个问句
        // 分别收在形参表(+ builtinNoParamTableStringSlot) / typeTakesImplicitCStr / wrapToBSTR。
        // 本站留下的三条都是**另一种**问句: 「这条文本是不是裸 COM 结果」「这个接口的值在哪」,
        // 与形参类型无关, 所以不跟着搬。
        // Fix 108c-3: vb6_ObjPtr(ifaceVar) — 接口引用是值结构体, 不能直接作为
        // uintptr_t 实参 (C2172). 对象指针位于 .obj 字段: 改写为 (<var>).obj.
        // mdlSubclass.bas: vb6_ObjPtr(Subclass) → vb6_ObjPtr((Subclass).obj).
        // 仅精确匹配裸接口变量名, 不触碰其它形式 (LongPtr 变量等走别处转换).
        if (callee == "vb6_ObjPtr" && !args.empty() && !node.positional.empty()
            && node.positional[0] && node.positional[0]->kind == ASTNodeKind::IdentifierExpr) {
            std::string aLower = args[0];
            std::transform(aLower.begin(), aLower.end(), aLower.begin(), ::tolower);
            if (knownIfaceVars_.count(aLower)) {
                // Fix 241: 上面的 `.obj` 改写只对 **legacy 胖值结构体**
                // (vb6_iface_<I>, 字段 obj + vtbl) 成立。新式 ivref 表示
                // (vb6_ivref_<I>*, 结构体里**只有** vt, 见 ai/022 B04) 没有 .obj
                // ⇒ 生成 `(This).obj` 直接 C2039。ivref 的薄指针本来就是接口引用
                // 的地址，与 Fix 215 的 `PtrToObj` 往返同口径 —— 存进去什么、
                // 取回来就是什么，所以只需把裸指针当 void* 交出去。
                // 实测 ComCtlsDemo Builds/ListView/ListView.ctl:5226
                // `ByVal ObjPtr(This)` (This As ISubclass 形参，C 侧 vb6_ivref_ISubclass*)
                // 与 Builds/ComCtlsBase.bas:786 `ObjPtr(This)`：
                // 修复前 ListView.c:7552 / ComCtlsBase.c:783 发 `vb6_ObjPtr((This).obj)`
                // ⇒ C2039×5 + C2198×5。
                const std::string ifaceNm241 = knownIfaceVars_[aLower];
                if (!ivrefCType(ifaceNm241).empty()) {
                    std::string newVal = "(void*)(" + args[0] + ")";
                    replaceInArgList108c(args[0], newVal);
                    args[0] = newVal;
                    argsFixed108c = true;
                } else {
                    std::string newVal = "(" + args[0] + ").obj";
                    replaceInArgList108c(args[0], newVal);
                    args[0] = newVal;
                    argsFixed108c = true;
                }
            }
        }
        // Fix (vbNullChar 一族): VB6 `String$(n, charcode)` 的 charcode **也可以是字符串** ——
        // 取它**首字符**的码 (手册: String$(3, "A") = "AAA")。RTL 原型是
        // vb6_String(int32_t n, int32_t charCode), 把 BSTR 直接递进去是**指针→int**:
        // 低 32 位被当字符码 ⇒ `String$(260, vbNullChar)` 造出 260 个**垃圾宽字符**
        // (屏上是整条 CJK 乱码)。实测 Builds/ImageCombo/ImageCombo.ctl:1400
        // `Buffer = String$(CBEMAXSTRLEN, vbNullChar)` 是 ComCtls 控件分配缓冲的惯用写法。
        // 改用 vb6_AscW(BSTR) 取首字符 (空串/NULL 交 0, 无 Asc 的错误副作用)。
        if (callee == "vb6_String" && args.size() >= 2
            && node.positional.size() >= 2 && node.positional[1]
            && inferExprType(*node.positional[1]) == Vb6Type::String) {
            std::string wrappedStr = "vb6_AscW(" + args[1] + ")";
            replaceInArgList108c(args[1], wrappedStr);
            args[1] = wrappedStr;
            argsFixed108c = true;
        }

        (void)argsFixed108c;
    }

    // Fix <vbeclipse> 2026-10-06: Declare 的 Err.LastDllError 快照**不在这里**做 ——
    // 调用点注入语句会把 `While Api() <> 0` / `IIf(Api(), ...)` 这类表达式位置的
    // 调用撕成 `(Api(); vb6_ErrSetLastDllError(...) != 0)` 这种非法 C。快照改由
    // cgen_decl_api.cpp 为每个 Declare 生成的 static __inline 包装函数承担
    // (调用点经 #define / declareAliasMap_ 已被重定向到该包装), 这里保持原样。
    lastExpr_ = callee + "(" + argList + ")";
