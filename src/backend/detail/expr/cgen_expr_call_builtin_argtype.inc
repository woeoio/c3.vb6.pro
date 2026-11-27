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

        // 账 #305 §11.106 (第一百二十四刀): 九枚**只读谓词**的实参按借用档装箱。
        // 它们形参是按值 `vb6_VARIANT`, 而装箱站发的是 `vb6_VariantFromValue(...)` 这枚
        // **表达式里的临时** —— 谁也不会 VivariantClear 它。对象档那一臂 (vb6_VariantObject)
        // 按 Fix 024 必须 AddRef ⇒ 每读一次漏一枚引用 (第一百二十三刀 FProbe6 的 D2 term=0
        // 就是这个: `Set x.A = Nothing` 只能把槽退到 1)。
        // 判据: callee 是这九枚之一 ⇒ 把实参**最外层那几层**装箱名前缀换成借用档; 里面
        // 的嵌套子表达式一律不动 (它们不在实参的顶层, 自有一套所有权)。
        {
            static const std::unordered_set<std::string> readonlyVariantPredicates124 = {
                "vb6_TypeName", "vb6_VarType", "vb6_IsObject", "vb6_IsNull",
                "vb6_IsEmpty", "vb6_IsArray", "vb6_IsDate", "vb6_IsError",
                "vb6_IsNumeric",
            };
            if (readonlyVariantPredicates124.count(callee)) {
                static const std::string kFrom124 = "vb6_VariantFromValue(";
                static const std::string kTo124 = "vb6_VariantBorrowFromValue(";
                size_t cursor124 = 0;
                for (auto& a124 : args) {
                    std::string b124 = a124;
                    size_t at124 = 0;
                    while (b124.compare(at124, kFrom124.size(), kFrom124) == 0) {
                        b124.replace(at124, kFrom124.size(), kTo124);
                        at124 += kTo124.size();
                    }
                    if (b124 == a124) continue;
                    size_t p124 = argList.find(a124, cursor124);
                    if (p124 == std::string::npos) continue;
                    argList.replace(p124, a124.size(), b124);
                    cursor124 = p124 + b124.size();
                    a124 = b124;
                }
            }
        }

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
        // §11.209 (第二百二十六刀): Hex/Oct 的实参收窄 = CLng 那一档。
        // 真机 .temp/k225p/P305：Hex(44259.62)=ACE4 / Hex(44259.5)=ACE4 / Hex(1.5)=2 /
        // Hex(3.5)=4 / Hex(-2.5)=FFFFFFFE / Oct 同族，即**先四入后进制**（半到偶），
        // 越界报 6（Hex(2147483647.6) 真机直接 Error 6）。改前这里让 C 隐式截断 ⇒
        // ACE3 / 1 / 3，且越界静默饱和。
        // 判据 = 实参的**静态档位**（inferExprType），不是文本：只有 C 隐式转换会
        // 走样那几档（Single/Double/Date/Currency 是浮点形，String 是 BSTR 指针）
        // 才加一层收窄；Integer/Long/Byte/Boolean 的 C 隐式加宽与 VB6 同值，不动。
        // 与 vb6rtl_com.c 的 `~` 那一站同口径（同一枚 vb6_CLng 主人）。
        if ((callee == "vb6_Hex" || callee == "vb6_Oct")
            && args.size() == 1 && !node.positional.empty() && node.positional[0]) {
            const Vb6Type t209 = inferExprType(*node.positional[0]);
            std::string narrowed209;
            if (t209 == Vb6Type::String) {
                narrowed209 = "vb6_CLng(vb6_NumVal(" + args[0] + "))";
            } else if (t209 == Vb6Type::Single || t209 == Vb6Type::Double
                       || t209 == Vb6Type::Date || t209 == Vb6Type::Currency) {
                narrowed209 = "vb6_CLng(" + args[0] + ")";
            }
            if (!narrowed209.empty()) {
                replaceInArgList108c(args[0], narrowed209);
                args[0] = narrowed209;
                argsFixed108c = true;
            }
            // §11.210 (第二百二十七刀): 答案宽度。真机 .temp/k225p/P308：
            // Hex(CByte(255))="FF" / Hex(CInt(-1))="FFFF" / Hex(True)="FFFF" /
            // Hex(CLng(-1))="FFFFFFFF"。正值不补零，所以宽度只体现在负值的二补码上。
            // 宽度表只在 RTL 一处（vb6_HexMaskOf）——这里只选名字，不另立一张表。
            //
            // （1）整只变体槽那一形：字符串上下文里 `Hex(v)` 已经落成
            //     `vb6_Hex(vb6_VariantToLong(v))`（装箱站还在后面，改不到它），
            //     而宽度全在槽的 vt 上 ⇒ 整串换成只读一次槽的 vb6_HexSlot(v)。
            //     判据是“这一串就是一层 VariantToLong”，不是猜来源。
            {
                static const std::string preV210 = "vb6_VariantToLong(";
                const bool slotShaped = args[0].compare(0, preV210.size(), preV210) == 0
                    && args[0].size() > preV210.size() + 1
                    && args[0].back() == ')';
                if (slotShaped) {
                    // 括号配平：确认整串就是那一次提取（尾括号必须是它自己的）
                    const size_t open210 = preV210.size() - 1;
                    int depth210 = 0;
                    size_t close210 = std::string::npos;
                    for (size_t k210 = open210; k210 < args[0].size(); ++k210) {
                        if (args[0][k210] == '(') depth210++;
                        else if (args[0][k210] == ')') {
                            if (--depth210 == 0) { close210 = k210; break; }
                        }
                    }
                    if (close210 == args[0].size() - 1) {
                        const std::string slot210 =
                            args[0].substr(preV210.size(), args[0].size() - preV210.size() - 1);
                        replaceInArgList108c(args[0], slot210);
                        args[0] = slot210;
                        callee = (callee == "vb6_Hex" ? "vb6_HexSlot" : "vb6_OctSlot");
                        argsFixed108c = true;
                    }
                }
            }
            // §11.212 (第二百二十九刀): 整枚数组进 Hex/Oct 的实参位 ⇒ 13。
        // 真机 .temp/k225p/P320/P321：`Hex(a)`/`Oct(a)`/`Hex(a())`/`Hex(b)`（String 元素）/
        // `Hex(c)`（未 ReDim）全部是 [keep]/13（目标一字不动）；`Hex(a(1))` 仍是 FFFFFFFB。
        // 改前本站把数组名当数值交给 vb6_Hex(double)：x64 上截的是描述符指针
        // （P312 的 A3 打出 EA921860），错误号也没有。
        // 判据只问 isWholeArrayRef（裸名与 `A()` 两形都算，见 cgen_util_type.cpp），
        // 报错本体在 RTL 那一位（vb6_ArrayArgMismatch）—— 本站只换实参，不拼 ErrRaise。
        // 必须放在宽度那一块之前：数组没有「宽度」，增宽只会把元素档误折进去。
        if (args.size() == 1 && node.positional.size() == 1 && node.positional[0]
            && (callee == "vb6_Hex" || callee == "vb6_Oct")
            && isWholeArrayRef(node.positional[0].get())) {
            const std::string oldA212 = args[0];
            args[0] = "vb6_ArrayArgMismatch()";
            replaceInArgList108c(oldA212, args[0]);
            argsFixed108c = true;
        }
        // （2）静态档：Byte → *Byte、Integer/Boolean → *Int。
            //     Integer 变量在 inferExprType 里答 Long（C 侧 int32 口径），
            //     所以那一档必须问 isNarrowIntExpr 这枚既有谓词（§C30-10 第三十四/四十五刀）。
            if ((callee == "vb6_Hex" || callee == "vb6_Oct")
                && args[0] != "vb6_ArrayArgMismatch()") {
                if (t209 == Vb6Type::Byte) callee += "Byte";
                else if (t209 == Vb6Type::Integer || t209 == Vb6Type::Boolean
                         || isNarrowIntExpr(node.positional[0].get())) callee += "Int";
            }
        }
        // §11.213 (第二百三十刀): 这一族在**非变体消费面**上都要严格。
        // 真机 .temp/k225p/P323（VB6Mini）：
        //   S1/S2 `s = CStr(Abs(v))`（v=""/"x"）= [keep]/13；S3 "2.5" = [2.5]/0；
        //   S4 `CStr(Abs(Null))` = [keep]/94；S5 Empty = [0]/0；
        //   T1/T2 `CStr(Int(""))/CStr(Fix("x"))` = 13；T3 `CStr(Int(Null))` = 94；
        //   U1..U3 `s = Abs(v)`（直接字符串赋值）同形；V1..V3 `s = "n=" & Abs(v)` 同形；
        //   W1 `w = Abs(Null)`（Variant 目标）= **Null**（IsNull=True）—— 只有变体那一面保 Null。
        // 改前这些格子全发 `vb6_Abs(vb6_VariantToDouble(v))`：Null 与文本都落 0。
        // 处置与 §11.211 的 Sgn 同一形（整只槽 + VariantToDouble ⇒ vb6_DblStrict），
        // 只是 Abs/Int/Fix 有 *V 出口，所以两条装箱表必须同步认新前缀
        // （cgen_util_ctrl.cpp 的 mathNullFamilyV / boxToVariant），变体消费面才能仍拿到
        // vb6_AbsV/IntV/FixV（否则 W1 会从 Null 变成 94）。
        if ((callee == "vb6_Abs" || callee == "vb6_Sgn" || callee == "vb6_Int"
             || callee == "vb6_Fix") && args.size() == 1 && !args[0].empty()) {
            static const std::string preS211 = "vb6_VariantToDouble(";
            if (args[0].compare(0, preS211.size(), preS211) == 0 && args[0].back() == ')') {
                int depthS211 = 0;
                size_t closeS211 = std::string::npos;
                for (size_t kS211 = preS211.size() - 1; kS211 < args[0].size(); ++kS211) {
                    if (args[0][kS211] == '(') depthS211++;
                    else if (args[0][kS211] == ')') {
                        if (--depthS211 == 0) { closeS211 = kS211; break; }
                    }
                }
                if (closeS211 == args[0].size() - 1) {
                    const std::string slotS211 = args[0].substr(
                        preS211.size(), args[0].size() - preS211.size() - 1);
                    const std::string strictS211 = "vb6_DblStrict(" + slotS211 + ")";
                    args[0] = strictS211;
                    replaceInArgList108c(std::string("vb6_VariantToDouble(") + slotS211 + ")",
                                         strictS211);
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
    // 账 #464 §B164 回收刀: 实参位上「整枚就是一个 mint 调用」的那一格 ⇒ 提成具名临时量,
    // 语句末按档收回。
    // ⚠ 这一格必须**放在最后**: 到这一步实参表才是**最终**形状（前面的 fixup 会补缺省参数、
    // 换 callee 名、改实参文本）⇒ 「消费者第几格」取的是**发码里那一格的位号**, 与
    // scripts/owned_temp_census_164.py 那份清点（同样按发码文本算）**同源**。
    // 判据全部来自 src/common/owned_temp_keeps.hpp（表由 RTL 推导 ⇒ 哨兵双向对账）。
    hoistOwnedTempArgs(argList, callee);

    // 账 §39 (C:/D: 少两行, 第 2 层): VB6 的 0 参 Function/PropertyGet 返回 Variant 时,
    // `F(i)` 的 (i) 是**对返回数组的下标**, 不是调用实参 (真 VB6 实测 /make 通过:
    // ucTreeView ClsTreeExplorer.cls:317 `GetDrive(i)`)。上面 Fix 041c 会把多出来的
    // 实参按「超过形参个数」截掉 ⇒ 拿到整个数组 (探针实测 VarType=8200、GetDrive(0/1) 空)。
    // 这里改发 vb6_VariantArrayGetVal(F(), i): F 收 0 个实参, 结果再取下标。
    if (calleeZeroArgVariantSubscript && !node.positional.empty()) {
        std::string selfCall = callee + "(" + argList + ")";
        std::string idxExpr = args.empty() ? "0" : args[0];
        lastExpr_ = "vb6_VariantArrayGetVal(" + selfCall + ", "
                    + toLongIfVariant(idxExpr, node.positional[0].get()) + ")";
        return;
    }

    // C3-SENTINEL(g4-value-member-empty-paren): 值成员被写成"带空括号"的收口点。
    // §B94: `cmdOK.FontSize()`, `cboUpdates.UBound()`, `PDImages.GetActiveImage.Width()` ——
    // 成员读取本身已经产出**完整** C 表达式 (尾部带"读法尾注"), 到这里再追加 `()` 就是
    // 调用一个非函数 ⇒ cl C2064 (PhotoDemon 实测 9 条: options_Updates 4 / SelectionFilters 4 /
    // pdCommandBarMini 1)。判据只认**产出尾注**, 不猜符号表:
    //   · 有实参的调用不进这一格 (argList 非空);
    //   · 真方法调用 (resolveClassMemberCall 有解) 在 callee 阶段就被折成函数名, 没有尾注;
    //   · 尾注是这三条读法自己写的 (`/* ctrl prop read */`、`/* ctrl array U|LBound */`、
    //     类字段的字段尾注) ⇒ 产出的是值访问, 空括号只是 VB6 的噪声写法。
    if (argList.empty() && callee.size() > 3 && callee.compare(callee.size() - 2, 2, "*/") == 0) {
        static const std::string kValueReadTails[] = {
            "  /* ctrl prop read */",
            "  /* ctrl array UBound */",
            "  /* ctrl array LBound */",
        };
        bool valueReadNote = false;
        for (const auto& tail : kValueReadTails) {
            if (callee.size() >= tail.size()
                && callee.compare(callee.size() - tail.size(), tail.size(), tail) == 0) {
                valueReadNote = true;
                break;
            }
        }
        // 类字段读: 尾注里带字段名 —— 认这一问只许问 §B172 那个出口的识别半
        // (cExprIsClassFieldRead), 不在这里认尾注文字。
        if (!valueReadNote && cExprIsClassFieldRead(callee)) {
            valueReadNote = true;
        }
        if (valueReadNote) {
            lastExpr_ = callee;
            return;
        }
    }

    lastExpr_ = callee + "(" + argList + ")";
