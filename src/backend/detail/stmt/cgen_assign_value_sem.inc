// cgen_assign_value_sem.inc - visit(AssignmentStmt&) 函数体片段：值语义（返回值变量 / BSTR / Variant 赋值）
// 由 src/backend/stmt/cgen_assign.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1056~1362 行，逐行未改。

    // 如果赋值目标是当前Function/PropertyGet名 (VB6语义: 设置返回值), 替换为返回值变量
    if (currentProc_ && (currentProc_->kind == SymbolKind::Function ||
                         currentProc_->kind == SymbolKind::PropertyGet)) {
        std::string procCName = cProcName(currentProc_->name, currentProc_->access, currentProc_->sourceModule);
        if (target == procCName) {
            target = currentReturnVar_;
        }
    }

    /* Fix 296 / 4a: 目标就是本 Property Get / Function 的返回槽, 而右值是发射出去的
       `vb6_Extender_*` 全局 ⇒ 换实例感知读入口 (写侧 3a 的镜像)。**只改右边**:
       vb6_ret_<Prop> 这个名字来自 currentReturnVar_ (decl/cgen_decl_func.cpp:325), 不动;
       外层的 vb6_ChkInt / vb6_BSTR_Assign 也照旧套在上面 —— 数值与字符串两种右值都成立。
       ComCtlsDemo 实测 333 处走的是这枚 push/pop 副本, 而权威值在 vb6_UCRec.ext*。
       4a 时这一问还带着 `target == currentReturnVar_` 的**窄**条件 (只接"填返回值槽"),
       4b 放开到**任何目标**: 自由表达式 `Align = Extender.Align` 的目标是个普通局部量,
       窄条件下永远不触发 (ComCtlsDemo ToolBar/CoolBar/StatusBar 共 4 处)。放开是安全的 ——
       闸门只在整条右值**就是**一枚 vb6_Extender_<已知成员> 时才动手, 而表里 type=Unknown 的
       Container/DragIcon 原样返回, 所以对象成员不会被塞进 int32 入口。 */
    value = extenderMemberReadExpr(value);

    // BSTR赋值检测: 如果目标是BSTR变量, 使用vb6_BSTR_Assign防止悬垂指针和双重释放
    bool targetIsBstr = false;
    // 账 #118: "这个目标是不是定长串, 定长多少" —— 在下面 targetIsBstr 块内算一次,
    // 发码处(本文件 240 行那一支)还要用它把值收口到 N, 故提到块外。
    std::string fixedLenTarget;
    {
        // 检查目标是否是已知BSTR变量 (me->field 或 模块级变量)
        std::string checkName = target;
        if (checkName.substr(0, 4) == "me->") checkName = checkName.substr(4);  // 去掉me->前缀
        // Fix 025: (*name) 解引用形式 — ByRef BSTR/Variant 参数写穿透, 剥掉 (* ... ) 取内部名
        if (checkName.size() > 4 && checkName[0] == '(' && checkName[1] == '*'
            && checkName.back() == ')') {
            checkName = checkName.substr(2, checkName.size() - 3);
        }
        std::string lower = checkName;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownBstrVars_.count(lower)) targetIsBstr = true;
        // 也检查不含前缀的原始名 (Property Let参数等)
        if (node.target->kind == ASTNodeKind::IdentifierExpr) {
            auto& id = static_cast<IdentifierExpr&>(*node.target);
            std::string idLower = id.name;
            std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
            if (knownBstrVars_.count(idLower)) targetIsBstr = true;
        }
        // Fix 177: UDT 动态数组元素上的 String 字段 (`...Cols(j).Text = TextIn`)。
        // 裸名集合捡不到这种目标, 原先落到最底下的 `target = value;` 裸赋值 →
        // 存进单元格的是**调用方临时 BSTR 的地址**, 过程返回即悬垂: VBFlexGridDemo
        // 里 148/148 行读到同一个 "149"/列头, 写 "ZZZ" 后读回是指针数值。
        // 交给下方 targetIsBstr 分支 → vb6_BSTR_Assign (释放旧值 + 深拷贝)。
        if (!targetIsBstr && udtFieldIsBstrInCTarget(target)) targetIsBstr = true;
        // Fix <vbeclipse> 2026-10-06: **裸 UDT 变量**的 String 字段 (`r.name = ...` /
        // With 块 `.name = ...`, 含嵌套链 `a.b.name`)。账 #118 当时只修了定长串字段
        // (fixedStrLenOfExpr), 动态 String 字段刻意留下 (注释"那是另一笔账") ——
        // udtFieldIsBstrInCTarget 只认 SA_AT 宏形状, 裸 UDT 变量捡不到 ⇒ 落最底下
        // 裸赋值 `r.name = r.num;`: 整型 RHS 静默把数字当 BSTR 指针存进字段
        // (第一次读/释放就崩), 浮点 RHS 才是 C2440。udtFieldCTypeOfTarget 对
        // MemberAccessExpr/WithMemberExpr 返回成员 C 类型, == "BSTR" 即 String
        // 字段 (数组成员被 == 判等天然排除), 交上方 targetIsBstr 分支走
        // vb6_BSTR_Assign + 账 #122 的隐式 CStr。
        if (!targetIsBstr) {
            std::string fieldCType206 = udtFieldCTypeOfTarget(node.target.get());
            if (fieldCType206 == "BSTR") targetIsBstr = true;
        }
        // Fix <vbeclipse> rev11: **String 数组元素赋值** (`m_Keys(Position) = Key`,
        // `arr(i) = s`) 的目标 C 串是 `VB6_SA_AT(BSTR, me->m_Keys, Position)` ——
        // 一个**表达式**, 上面四道按名字/符号表的判据全都捡不到它, 于是落到最底下的
        // 裸赋值 `VB6_SA_AT(BSTR, ...) = Key;`。
        // 裸赋值对 BSTR 就是**指针别名**, 而 VB6 语义是**字符串值拷贝**。实测真工程
        // (VbEclipse 的 List.cls, `Dim m_Keys() As String` + `m_Keys(Position) = Key`):
        //   Add "Editor"(6 字符) 存进去, 紧接着同一 List 的 Contains 读 m_Keys(0)
        //   读出 **长度 1**; 后面几个 key 读出 5/12/12/23 —— 全是别人用过的堆块。
        // 机制: `Key` 来自 `Public Const PERSPECTIVE_EDITOR As String = "Editor"`
        // 生成的 `#define PERSPECTIVE_EDITOR (vb6_BSTR_FromStr(L"Editor"))`, 是**每次
        // 求值新分配、从不释放**的临时块; `List_Add` 里的 ReDim Preserve 又只 memcpy
        // 指针。真正致命的一环是 List.cls 自己在 Add 结尾 `Set m_Items(Position) = Item`
        // 走 Variant 容器 → vb6_VariantClear 会 free 掉与 BSTR 元素共享的堆块。
        // 判据落在**宏的第一个实参(元素 C 类型) == BSTR**, 与 Fix 177 同源同口径:
        // 元素是 BSTR ⇒ 该槽必须走 vb6_BSTR_Assign (释放旧值 + 深拷贝)。
        // 排除 UDT/数值族 —— 它们的元素类型不是 BSTR, 本就不该深拷贝。
        // 用 sizeof()-1 而不是手数字面量长度 (这串前面就栽过: "VB6_SA_AT(" 写成 11)。
        static const std::string kSaBstrAt   = "VB6_SA_AT(BSTR,";
        static const std::string kSaNdBstrAt = "VB6_SA_ND_AT2(BSTR,";
        if (!targetIsBstr
            && (target.compare(0, kSaBstrAt.size(), kSaBstrAt) == 0
                || target.compare(0, kSaNdBstrAt.size(), kSaNdBstrAt) == 0)) {
            targetIsBstr = true;
        }
        // 账 #113: 模块级 String 变量不进 knownBstrVars_ —— 那张表在每过程入口 clear
        // (cgen_decl_func.cpp:43 / cgen_decl_proc.cpp:51), 只有局部声明与形参进表。
        // 于是 `gS = s`(String 形参) 一路落到最底下的裸赋值 `gS = s;` —— 存下来的是
        // **调用方那只临时串的地址**, 过程一返回 RTL 就释放 ⇒ 再读是别人用过的堆块
        // (WS-c 实测读出过 "WS38=" / "Y")。用符号表那把尺对**裸标识符目标**补问一次
        // (inferExprType 的 IdentifierExpr 分支认得模块级变量)。
        // 只认 IdentifierExpr: 成员/控件属性目标另有各条专门通道 (cgen_assign_stmt_special.inc
        // 里那些), 放进来会把它们的发码面抢掉。
        if (!targetIsBstr && node.target->kind == ASTNodeKind::IdentifierExpr
            && inferExprType(*node.target) == Vb6Type::String) {
            targetIsBstr = true;
        }
        // 账 #118 ①: 定长串 UDT 字段 (`Type Rec: f As String * 8` 的 `r.f = ...`)。
        // 上面 udtFieldIsBstrInCTarget 那条只认 **SA_AT 宏**形状 (它要求对象部分以 ')' 收尾),
        // 对象是裸 UDT 变量时捡不到 ⇒ 此前连 BSTR 分支都进不来, 落最底下的裸赋值
        // (`r.f = vb6_BSTR_FromStr(L"ef")`), 既不收口到 8, 重赋也不释放旧值。
        // 判据刻意**只认定长**字段 (宽度可静态得知): 动态 String 字段维持既有口径,
        // 免得顺手改了 UDT 字段赋值的深拷贝/借用语义 (那是另一笔账)。
        // 账 #305 §B169 (第六十刀): `<对象>.<类的公共 String 字段>` 与类内的 `Me.<同一字段>`
        // —— 上面五道判据问的全是「**本模块**认不认识这个名字」(裸名 knownBstrVars_ /
        // UDT 字段类型 / SA_AT 宏文本 / inferExprType 对裸标识符的答案), 而这条字段的 C 类型
        // 由**那个类自己的声明**决定, 于是这一形整个漏掉, 发的是一句**裸指针赋值** (实测
        // MSNMacro CDLG.cls `Me.SelectedPrinterName = ...`):
        //   h.PS = s          → `h->PS = s;`  ⇒ 存进去的是**别人那只 BSTR 的地址**(别名),
        //                        源头一重赋值/一释放, 这枚字段就悬空;
        //   h.PS = "x" 再 h.PS = "y" → 每赋一次漏一块 (旧值没人 free, 而类析构会 free 字段);
        //   h.PS = <COM 属性读> → `h->PS = vb6_ComGetStringProp(o, L"DeviceName");` 交的是
        //                        **借来**的那只块, 析构 `vb6_BSTR_Free(me->PS)` 于是 free 第二遍。
        // 这一问收成一处: 档问工程级那张表 (classFieldKind, §B168 那一趟预扫描扩的档),
        // 命中即交 emitBstrSlotStore (释放旧值 + 深拷贝/接管) —— 与 §B168 的 Variant 那一支
        // 同一个走法, 也**不碰槽位语义** (自有还是借来由那一份权威答, 站点不抄)。
        // 否决与 §B168 同一条: member 若是该类的**方法/属性** (resolveClassMemberCall 非空),
        // 这一句归属性写那一路。
        if (!targetIsBstr && node.target
            && node.target->kind == ASTNodeKind::MemberAccessExpr) {
            auto& maB = static_cast<const MemberAccessExpr&>(*node.target);
            const std::string clsB
                = maB.object ? inferClassTypeOfExpr(*maB.object) : std::string();
            if (!clsB.empty()
                && resolveClassMemberCall(clsB, maB.memberName).empty()
                && classFieldIsBstrSlot(clsB, maB.memberName)) {
                targetIsBstr = true;
            }
        }
        fixedLenTarget = fixedStrLenOfExpr(*node.target);
        if (!targetIsBstr && !fixedLenTarget.empty()) targetIsBstr = true;
    }
    // P8.4: Variant赋值检测 - 如果目标是Variant变量, 用Variant构造函数包装值
    bool targetIsVariant = false;
    // Fix 232: 标记"目标 = UDT 的 Variant 字段", 供下方选**深拷贝**槽位语义
    bool targetIsUdtVariantField232 = false;
    {
        std::string checkName = target;
        if (checkName.substr(0, 4) == "me->") checkName = checkName.substr(4);
        // Fix 025: (*name) 解引用形式 — ByRef Variant 参数 (Dim X As Variant 走 ByRef),
        //   codegen emit 写穿透为 (*X) = ...; 必须剥掉 (* ... ) 才能在 knownVariantVars_ 查到 X.
        if (checkName.size() > 4 && checkName[0] == '(' && checkName[1] == '*'
            && checkName.back() == ')') {
            checkName = checkName.substr(2, checkName.size() - 3);
        }
        std::string lower = checkName;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownVariantVars_.count(lower)) targetIsVariant = true;
        // Fix 054: 也检查 classVariantMembers_ — 类成员 Variant 字段 (如 me->m_P1)
        if (classVariantMembers_.count(lower)) targetIsVariant = true;
        if (node.target->kind == ASTNodeKind::IdentifierExpr) {
            auto& id = static_cast<IdentifierExpr&>(*node.target);
            std::string idLower = id.name;
            std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
            if (knownVariantVars_.count(idLower)) targetIsVariant = true;
        }
        // Fix 154: VARIANT 数组元素赋值 — 目标 emit 后是 `VB6_SA_AT(vb6_VARIANT, ...)`
        // / `VB6_SA_ND_AT2(vb6_VARIANT, ...)` 等下标宏 (返回 vb6_VARIANT 左值引用).
        // 裸名 knownVariantVars_ 检测捡不到它, RHS 若是裸 BSTR/int32/double → 直接
        // `elem = Buffer` → C2440 "无法从 BSTR 转换为 vb6_VARIANT" (VBFlexGrid 二维
        // Variant 表数组写入). 按 VARIANT 元素处理, 交给下方 wrapVariantValue 包装;
        // vb6_VariantClear(&elem) 与 elem = wrapped 对宏返回的引用均合法.
        if (!targetIsVariant && target.compare(0, 7, "VB6_SA_") == 0
            && target.find("(vb6_VARIANT,") != std::string::npos) {
            targetIsVariant = true;
        }
        // Fix 232: UDT 的 **Variant 字段** (`.Image` / `With` 块内的 `.Image`) —— 与
        // 上方 Fix 154 同源同口径, 只是判据换成字段的**声明类型**。目标发出来是
        // `VB6_SA_AT(vb6_type_<Udt>, arr, i).Image`, 裸名集合与 Fix 154 的
        // `(vb6_VARIANT,` 子串两道判据都捡不到它, RHS 又是裸 BSTR/int32_t →
        // 直接 `elem.Image = ...` → C2440 "无法从 BSTR/int32_t 转换为 vb6_VARIANT"。
        // 实测 ComCtlsDemo CoolBar.ctl / PPCoolBarBands.pag / StatusBar.ctl 的
        // `InitBands(i).Image = <BSTR|Long>` 一族。
        // 字段的真 C 类型是 `vb6_cls_*` 时否决 —— 那是 typed 对象指针槽, 不是
        // Variant 容器 (与 Set 侧 cgen_setlet_set_rhs.inc Fix 084n/B08f-1 同一条
        // 否决规则); `As Variant` 字段自身 resolveTypeName 就给 `vb6_VARIANT`,
        // 装箱是必须的, 不能一并挡掉。字段类型问 inferUdtFieldVb6Type —— 那是
        // UDT 字段类型的既有唯一出口, 不另抄一份成员类型表。
        if (!targetIsVariant) {
            const std::string udtFldC232 = udtFieldCTypeOfTarget(node.target.get());
            if (udtFldC232.rfind("vb6_cls_", 0) != 0
                && inferUdtFieldVb6Type(node.target.get(), nullptr) == Vb6Type::Variant) {
                const std::string fldTRN232 = udtFieldTypeRefNameOfTarget(node.target.get());
                if (fldTRN232.empty() || cTypeForDeclaredTypeName(fldTRN232) == "vb6_VARIANT") {
                    targetIsVariant = true;
                    targetIsUdtVariantField232 = true;
                }
            }
        }
        // 账 #305 §B168 (第五十九刀): `<对象>.<类的公共 Variant 字段>` 与类内的
        // `Me.<同一字段>` —— 上面四条判据问的全是「**本模块**认不认识这个名字」(裸名
        // knownVariantVars_ / 本模块 classVariantMembers_ / 发出来的文本长得像 VB6_SA_*
        // / UDT 字段类型表), 而这条字段的 C 类型由**那个类自己的声明**决定, 于是这一形
        // 整个漏掉, 发的是一句**完全不装箱**的裸赋值 (实测 use3.bas + holder2.cls):
        //   h.PV = 7      → `h->PV = 7;`                        → cl C2440 int32_t→vb6_VARIANT
        //   h.PV = "s"    → `h->PV = vb6_BSTR_FromStr(L"s");`    → 同形 C2440
        //   Me.PV = x     → `me->PV = (*x);`                    → 编得过, 但浅存 + 不回收旧值
        // C2440 是硬错误 ⇒ 整棵工程编不过 (语料 94 个写点 / 28 份文件)。
        // 这一问收成一处: 类名问既有出口 inferClassTypeOfExpr (IdentifierExpr / MeExpr /
        // 链式成员都在它覆盖里), 字段档问 driver 注入的工程级「类 → vb6_VARIANT 字段」表
        // (classFieldIsVariantSlot, 与 classVoidFieldMap_/classTypedFieldMap_ 同格)。
        // **不**认发出来的文本尾注 (`/* class var .X field */`) —— 那是拼出来的注释,
        // 不是答案; 也**不**再抄一份字段类型表。
        // 命中后交下方站④: wrapVariantValue 装箱 + variantRhsOwnsValue 定所有权
        // (§B164 第五十五刀 / §B166 第五十六刀 / §B167 第五十七、五十八刀同一口径),
        // 所以这里只把「是不是槽」答对, 不碰槽位语义 —— 借用值 (h.PV = 活的 String 变量)
        // 由那份权威给出 vb6_VariantAssign 深拷贝, 与 UDT 的 232 那一支同一个答案。
        // 否决: member 若是该类的**方法/属性** (resolveClassMemberCall 非空), 这一句归
        // 属性写那一路, 不许在这里抢 —— 判据与读侧 cgen_expr_call_callee_member.inc
        // Pattern C/D 那条「是方法→跳过, 否→数据字段」完全一致。
        if (!targetIsVariant && node.target
            && node.target->kind == ASTNodeKind::MemberAccessExpr) {
            auto& maT = static_cast<const MemberAccessExpr&>(*node.target);
            const std::string clsT
                = maT.object ? inferClassTypeOfExpr(*maT.object) : std::string();
            if (!clsT.empty()
                && resolveClassMemberCall(clsT, maT.memberName).empty()
                && classFieldIsVariantSlot(clsT, maT.memberName)) {
                targetIsVariant = true;
            }
        }
    }
    if (targetIsBstr) {
        // 账 #305 §B169: 这一问只住 emitBstrSlotStore 一处 (With 块那一站共用同一枚)。
        emitBstrSlotStore(target, node.value.get(), value, fixedLenTarget);
    } else if (targetIsVariant) {
        // 账 #305 §B169: 装箱 + 所有权 + 槽位语义只住 emitVariantSlotAssign 一处。
        emitVariantSlotAssign(target, node.value.get(), value, targetIsUdtVariantField232);
    } else if (target.compare(0, 10, "vb6_PA_Get") == 0) {
        // Fix 091c: ParamArray 元素赋值 OutVars(i) = value — vb6_PA_GetXxx 是取值
        // 函数 (右值), 直接赋值 → C2106 "左操作数必须为左值" (cToolsArray.c 130);
        // 且 Variant RHS 传给具体类型 Set 形参需提取 → C2440 (cToolsArray.c 158/167).
        // 改写为 vb6_PA_SetXxx(args, <按类型提取后的 value>).
        size_t lp091c = target.find('(');
        size_t rp091c = target.rfind(')');
        if (lp091c != std::string::npos && rp091c != std::string::npos && rp091c > lp091c) {
            std::string args091c = target.substr(lp091c + 1, rp091c - lp091c - 1);
            std::string kind091c = target.substr(10, lp091c - 10);  // "Long"/"Double"/"BSTR"/...
            bool valIsVar091c = cExprIsVariant(value);
            if (!valIsVar091c && node.value && node.value->kind == ASTNodeKind::IdentifierExpr) {
                auto& id091c = static_cast<IdentifierExpr&>(*node.value);
                std::string idLower091c = id091c.name;
                std::transform(idLower091c.begin(), idLower091c.end(), idLower091c.begin(), ::tolower);
                if (knownVariantVars_.count(idLower091c)) valIsVar091c = true;
            }
            std::string conv091c = value;
            // Fix <vbeclipse> rev37: rev37 起 ParamArray 的 Variant 元素按
            //   `vb6_PA_GetVariant(...)` 取 (按声明类型分派), 它在 LHS 上对应
            //   `vb6_PA_SetVariant(args, VARIANT*)` —— **要指针**, 不能把值直接塞进去。
            //   走一个临时 VARIANT 承接 (发射后不释放: SafeArrayPutElement 已深拷贝,
            //   而栈上 VARIANT 若含堆成员由 VariantClear 负责, 这里跟既有的
            //   ComPackMissing 同款用法 —— 临时块交给 vb6_PA_SetVariant 消费)。
            if (kind091c == "Variant") {
                // 栈上 VARIANT 承接 —— vb6_PA_SetVariant 的第二参要 VARIANT*,
                // 而 boxToVariant 给的是**值**。SafeArrayPutElement 会深拷贝
                // (另配一只 BSTR/对象), 所以这里借用栈地址、发射后不额外释放,
                // 也不调 VariantClear (那会把刚交给 SAFEARRAY 的副本再清一次)。
                std::string tmp091c = "_pav_" + std::to_string(tempCounter_++);
                c_.emitLine("{ VARIANT " + tmp091c + " = " + boxToVariant(node.value.get(), value) + ";");
                c_.emitLine("  vb6_PA_SetVariant(" + args091c + ", (VARIANT*)&" + tmp091c + "); }");
                return;
            }
            if (valIsVar091c) {
                if (kind091c == "Long" || kind091c == "LongPtr") {
                    conv091c = "vb6_VariantToLong(" + value + ")";
                } else if (kind091c == "Double") {
                    conv091c = "vb6_VariantToDouble(" + value + ")";
                } else if (kind091c == "BSTR") {
                    conv091c = "vb6_VariantToString(" + value + ")";
                }
            }
            c_.emitLine("vb6_PA_Set" + kind091c + "(" + args091c + ", " + conv091c + ");");
        } else {
            c_.emitLine(target + " = " + value + ";");
        }
    } else if (isVariantSlotElemRead(target)) {
        // Fix 038 Group 2 + 账 #305 §B150: 槽元素读改写成那一**对**写出口 —— 一维 Get->Set,
        //   多维 GetND->SetND, 装箱后的值追加在参数表末尾。参数拼法只在 variantSlotAssignLine 一处。
        std::string slotLine;
        if (variantSlotAssignLine(target, boxToVariant(node.value.get(), value), slotLine)) {
            c_.emitLine(slotLine + ";");
        } else {
            c_.emitLine(target + " = " + value + ";");
        }
    } else if (target.find("VB6_SA_AT(vb6_VARIANT,") != std::string::npos) {
        // Fix 038 Group 1: VB6_SA_AT(vb6_VARIANT, arr, idx) = value
        //   Variant 数组元素赋值: value 不是 VARIANT 时需要用 VariantFromValue 包装.
        //   vb6_VariantFromValue 对已存在的 VARIANT 是 identity (no-op), 安全.
        c_.emitLine(target + " = " + boxToVariant(node.value.get(), value) + ";");
    } else {
        // Fix 178: 两端是同一「含所有权成员 UDT」(如 `Rows(i) = VBFlexGridDefaultCols`)
        // → C 结构体赋值/memcpy 都是浅拷贝, 成员会被两侧共享; 换生成的深拷贝函数。
        {
            std::string deep = udtDeepCopyAssign(*node.target, target, node.value.get(), value);
            if (!deep.empty()) { c_.emitLine(deep); return; }
        }
        // Fix 038b-6/045: 具体类型目标 + Variant 值 → 自动提取
        // 覆盖: TimeOut = obj.Method() (Method 返回 Variant, TimeOut 是 Long)
        //       VB6_SA_AT(BSTR, arr, i) = func() (func 返回 Variant)
        //       int32_t_var = vb6_VariantArrayGet(...)
        // 使用 cExprIsVariant (C 字符串级) + knownVariantVars_ 检测.
        bool valueIsVariant = cExprIsVariant(value);
        if (!valueIsVariant && node.value && node.value->kind == ASTNodeKind::IdentifierExpr) {
            auto& id = static_cast<IdentifierExpr&>(*node.value);
            std::string idLower = id.name;
            std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
            if (knownVariantVars_.count(idLower)) valueIsVariant = true;
        }
        // Fix 089j: RHS 为类 Variant 字段 (me->mParentsColKey 等) 也视为
        // Variant 表达式 — cExprIsVariant 只查函数前缀, me-> 成员不命中,
        // 导致 PropertyGet 里 `ret = me->VarField` 不转 → C2440.
        if (!valueIsVariant && value.compare(0, 4, "me->") == 0) {
            std::string rhsMember = value.substr(4);
            std::transform(rhsMember.begin(), rhsMember.end(), rhsMember.begin(), ::tolower);
            if (classVariantMembers_.count(rhsMember)) valueIsVariant = true;
        }
        // Fix 092c: RHS 是当前过程的返回变量 (vb6_ret_<Proc>) 且过程返回类型为
        // Variant (未声明返回类型 / As Variant) — cExprIsVariant 只认函数前缀,
        // 裸返回变量名不命中 → 目标为具体类型时不做提取 → C2440:
        //   cDialog 393: Function ShowOpen(...) As Variant 内
        //     If VarType(ShowOpen) = vbString Then mData.mstrFileName = ShowOpen
        //     → me->mData.mstrFileName = vb6_ret_ShowOpen (vb6_VARIANT→BSTR).
        if (!valueIsVariant && !currentReturnVar_.empty()
            && value == currentReturnVar_
            && currentReturnCType_ == "vb6_VARIANT") {
            valueIsVariant = true;
        }
        // Fix 124: RHS 是裸 COM 调用结果 (vb6_ComCall* / ComGet*Prop 返回 void*,
        // 承载 VARIANT*) 且目标是具体类型标量 — 直接赋值 → void*→double/int32_t
        // C2440 (cZipArchive: pvVfsSeek = uFile.Stream.VfsSetFilePointer(...) →
        // vb6_ret_pvVfsSeek = vb6_ComCall(...) 2630). 按目标 C 类型解封:
        // VariantFromComResult 把 void* 包回 vb6_VARIANT, 再 VariantToXxx 提取.
        bool valueIsBareComResult = (value.compare(0, 12, "vb6_ComCall(") == 0
            || value.compare(0, 20, "vb6_ComCallByDispid(") == 0
            || value.compare(0, 17, "vb6_ComCallObject(") == 0
            || value.compare(0, 15, "vb6_ComCallBSTR(") == 0
            || value.compare(0, 14, "vb6_ComCallInt(") == 0
            || value.compare(0, 17, "vb6_ComCallDouble(") == 0
            || value.compare(0, 14, "vb6_ComGetProp(") == 0
            || value.compare(0, 19, "vb6_ComGetObjectProp(") == 0
            || value.compare(0, 18, "vb6_ComGetStringProp(") == 0
            || value.compare(0, 15, "vb6_ComGetIntProp(") == 0
            || value.compare(0, 18, "vb6_ComGetDoubleProp(") == 0);
        if (valueIsBareComResult) {
            std::string unwrapped = "vb6_VariantFromComResult(" + value + ")";
            std::string tgtCt124;
            if (!currentReturnVar_.empty() && target == currentReturnVar_
                && !currentReturnCType_.empty()) {
                tgtCt124 = currentReturnCType_;      // 函数名赋值: 按返回类型
            } else {
                // 按目标名查已知变量类型 (与 Fix 084n 已知变量集合同口径)
                std::string chk124 = target;
                if (chk124.compare(0, 4, "me->") == 0) chk124 = chk124.substr(4);
                if (chk124.size() > 4 && chk124[0] == '(' && chk124[1] == '*'
                    && chk124.back() == ')') {
                    chk124 = chk124.substr(2, chk124.size() - 3);
                }
                std::string lo124 = chk124;
                std::transform(lo124.begin(), lo124.end(), lo124.begin(), ::tolower);
                if (knownDoubleVars_.count(lo124)) tgtCt124 = "double";
                else if (knownLongVars_.count(lo124)) tgtCt124 = "int32_t";
                else if (knownBstrVars_.count(lo124)) tgtCt124 = "BSTR";
                // Fix 161f: 上述集合在**每个过程入口被 clear** (cgen_decl_func.cpp:48
                // 起), 模块级变量 Ax As Long 进过程后就没了集合条目 → 走不到按 CLong
                // 提取的分支, 直接 `x = vb6_VariantFromComResult(...)` → C2440
                // "无法从 vb6_VARIANT 转换为 int32_t" (extlist MListViewEx
                // `m_lvHWnd = lv.hWnd` 实测)。模块级变量仍在符号表里, 回查一次:
                // 只在目标是裸标识符且命中 Variant/未知时按声明类型补判。
                if (tgtCt124.empty() && node.target
                    && node.target->kind == ASTNodeKind::IdentifierExpr) {
                    switch (inferExprType(*node.target)) {
                    case Vb6Type::Long: case Vb6Type::ULong: case Vb6Type::Integer:
                    case Vb6Type::Byte: case Vb6Type::Boolean:
                        tgtCt124 = "int32_t"; break;
                    case Vb6Type::Double: case Vb6Type::Single:
                    case Vb6Type::Currency: case Vb6Type::Date:
                        tgtCt124 = "double"; break;
                    case Vb6Type::String:
                        tgtCt124 = "BSTR"; break;
                    default: break;
                    }
                }
            }
            // Fix P24-10b(2): 目标是标量**数组元素**时 knownXxxVars_ 与
            // IdentifierExpr 兜底都捡不到 —— 目标串是 `VB6_SA_AT(float, arr, i)`
            // / `VB6_SA_ND_AT2(int32_t, ...)` 下标宏, 裸 COM 结果(RHS 为
            // vb6_ComCallByDispid 等返回值 void* 承载 VARIANT*)原样赋给元素
            // → C2440 "无法从 void* 转换为 float" (Charts2020 ucTreeMaps
            // `Values(i) = cValues(i)` 实测, Values As Single())。
            // 元素 C 类型按宏第一实参解析, 与下方 valueIsVariant 分支
            // (账 038b-6/086 的 VB6_SA_AT(...) 判定) 同口径。
            if (tgtCt124.empty()) {
                static const std::string kSaAt   = "VB6_SA_AT(";
                static const std::string kSaNdAt = "VB6_SA_ND_AT2(";
                size_t elemPos = std::string::npos;
                if (target.compare(0, kSaAt.size(), kSaAt) == 0) {
                    elemPos = kSaAt.size();
                } else if (target.compare(0, kSaNdAt.size(), kSaNdAt) == 0) {
                    elemPos = kSaNdAt.size();
                }
                if (elemPos != std::string::npos) {
                    size_t comma = target.find(',', elemPos);
                    if (comma != std::string::npos) {
                        std::string elemCt = target.substr(elemPos, comma - elemPos);
                        size_t tb = elemCt.find_first_not_of(" \t");
                        size_t te = elemCt.find_last_not_of(" \t");
                        if (tb != std::string::npos)
                            elemCt = elemCt.substr(tb, te - tb + 1);
                        if (elemCt == "float" || elemCt == "double")
                            tgtCt124 = "double";
                        else if (elemCt == "BSTR")
                            tgtCt124 = "BSTR";
                        else if (elemCt == "int32_t" || elemCt == "int16_t"
                                 || elemCt == "uint8_t" || elemCt == "LONG")
                            tgtCt124 = "int32_t";
                    }
                }
            }
            // Fix P24-10b(2b): UDT 成员目标 (`me->mData.fValue = ...` 折算后
            // 目标串带 `->` 点链) 同样不在 known 集合/裸标识符/SA_AT 三类里——
            // 按 Fix 084n 同款 inferUdtFieldVb6Type 补问一次声明类型, 否则
            // 裸 COM 结果原样赋给 float/double 字段 → C2440。
            if (tgtCt124.empty()) {
                Vb6Type udtT124 = inferUdtFieldVb6Type(node.target.get());
                if (udtT124 == Vb6Type::String) tgtCt124 = "BSTR";
                else if (udtT124 == Vb6Type::Long || udtT124 == Vb6Type::Integer
                         || udtT124 == Vb6Type::Byte || udtT124 == Vb6Type::Boolean
                         || udtT124 == Vb6Type::ULong || udtT124 == Vb6Type::LongPtr)
                    tgtCt124 = "int32_t";
                else if (udtT124 == Vb6Type::Double || udtT124 == Vb6Type::Single
                         || udtT124 == Vb6Type::Currency || udtT124 == Vb6Type::Date)
                    tgtCt124 = "double";
            }
            if (tgtCt124 == "double" || tgtCt124 == "float") {
                c_.emitLine(target + " = vb6_VariantToDouble(" + unwrapped + ");");
                return;
            }
            if (tgtCt124 == "int32_t" || tgtCt124 == "int16_t" || tgtCt124 == "uint8_t"
                || tgtCt124 == "int64_t" || tgtCt124 == "LONG") {
                c_.emitLine(target + " = vb6_VariantToLong(" + unwrapped + ");");
                return;
            }
            if (tgtCt124 == "BSTR") {
                c_.emitLine(target + " = vb6_VariantToString(" + unwrapped + ");");
                return;
            }
            // 目标类型未知 (void*/对象等) → 不动, 维持原赋值路径.
        }
        if (valueIsVariant) {
            // 检查目标类型
            std::string convertedValue = value;
            if (target.find("VB6_SA_AT(BSTR,") != std::string::npos
                || target.find("VB6_SA_ND_AT2(BSTR,") != std::string::npos) {
                convertedValue = "vb6_VariantToString(" + value + ")";
            } else if (target.find("VB6_SA_AT(int32_t,") != std::string::npos
                       || target.find("VB6_SA_AT(int16_t,") != std::string::npos
                       || target.find("VB6_SA_AT(uint8_t,") != std::string::npos
                       || target.find("VB6_SA_AT(LONG,") != std::string::npos
                       || target.find("VB6_SA_ND_AT2(int32_t,") != std::string::npos
                       || target.find("VB6_SA_ND_AT2(int16_t,") != std::string::npos
                       || target.find("VB6_SA_ND_AT2(uint8_t,") != std::string::npos
                       || target.find("VB6_SA_ND_AT2(LONG,") != std::string::npos) {
                convertedValue = "vb6_VariantToLong(" + value + ")";
            } else if (target.find("VB6_SA_AT(double,") != std::string::npos
                       || target.find("VB6_SA_AT(float,") != std::string::npos
                       || target.find("VB6_SA_ND_AT2(double,") != std::string::npos
                       || target.find("VB6_SA_ND_AT2(float,") != std::string::npos) {
                convertedValue = "vb6_VariantToDouble(" + value + ")";
            } else if (!currentReturnVar_.empty() && target == currentReturnVar_
                       && !currentReturnCType_.empty()) {
                // Fix 092f: 赋值目标是当前过程的返回变量 → 按返回 C 类型提取 Variant
                // 值. 返回变量不在 UDT 字段/数组元素判定内 (inferUdtFieldVb6Type
                // 对 vb6_ret_X 返回 Unknown) → 之前直接赋值触发 C2440:
                //   cHttpClient 399: Function ReturnBody() As Byte() 内
                //     ReturnBody = Inst.ResponseBody
                //     → vb6_ret_ReturnBody = vb6_VariantFromComResult(...)
                //       (vb6_VARIANT → vb6_SafeArray1D*).
                std::string retCt092f = currentReturnCType_;
                retCt092f.erase(std::remove(retCt092f.begin(), retCt092f.end(), ' '),
                                retCt092f.end());
                if (retCt092f == "vb6_SafeArray1D*") {
                    convertedValue = "vb6_VariantToSafeArray1D(" + value + ")";
                } else if (retCt092f == "BSTR") {
                    convertedValue = "vb6_VariantToString(" + value + ")";
                } else if (retCt092f == "void*") {
                    convertedValue = "vb6_VariantToObjectVal(" + value + ")";
                } else if (retCt092f == "double" || retCt092f == "float") {
                    convertedValue = "vb6_VariantToDouble(" + value + ")";
                } else if (retCt092f == "intptr_t") {
                    convertedValue = "vb6_VariantToLongPtr(" + value + ")";
                } else if (retCt092f == "int32_t" || retCt092f == "int16_t"
                           || retCt092f == "uint8_t" || retCt092f == "LONG"
                           || retCt092f == "int64_t") {
                    convertedValue = "vb6_VariantToLong(" + value + ")";
                }
            } else {
                // Fix 084n: UDT 字段目标 — 按字段 Vb6Type 转换 Variant RHS
                // (如 cZipArchive 的 .FileName As String ← vb6_VariantArrayGet(...)
                //  → vb6_VariantToString; uBuf.MaxMatch As Long ← At() → vb6_VariantToLong)
                Vb6Type udtFieldT = inferUdtFieldVb6Type(node.target.get());
                if (udtFieldT == Vb6Type::String) {
                    convertedValue = "vb6_VariantToString(" + value + ")";
                } else if (udtFieldT == Vb6Type::LongPtr) {
                    convertedValue = "vb6_VariantToLongPtr(" + value + ")";
                } else if (udtFieldT == Vb6Type::Long || udtFieldT == Vb6Type::Integer
                           || udtFieldT == Vb6Type::Byte || udtFieldT == Vb6Type::Boolean
                           || udtFieldT == Vb6Type::ULong) {
                    convertedValue = "vb6_VariantToLong(" + value + ")";
                } else if (udtFieldT == Vb6Type::Double || udtFieldT == Vb6Type::Single
                           || udtFieldT == Vb6Type::Currency || udtFieldT == Vb6Type::Date) {
                    convertedValue = "vb6_VariantToDouble(" + value + ")";
                } else if (udtFieldT == Vb6Type::Object) {
                    convertedValue = "vb6_VariantToObjectVal(" + value + ")";
                } else if ((static_cast<uint16_t>(udtFieldT) & static_cast<uint16_t>(Vb6Type::Array))
                           == static_cast<uint16_t>(Vb6Type::Array)) {
                    convertedValue = "vb6_VariantToSafeArray1D(" + value + ")";
                }
                // 若已按 UDT 字段类型转换, 跳过已知变量集合检测
                if (convertedValue != value) {
                    // UDT 字段转换已完成
                } else {
                // 检查已知 Long/Double/ByteArray 变量
                std::string checkName = target;
                if (checkName.substr(0, 4) == "me->") checkName = checkName.substr(4);
                if (checkName.size() > 4 && checkName[0] == '(' && checkName[1] == '*'
                    && checkName.back() == ')') {
                    checkName = checkName.substr(2, checkName.size() - 3);
                }
                std::string lower = checkName;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                if (knownLongPtrVars_.count(lower)) {
                    convertedValue = "vb6_VariantToLongPtr(" + value + ")";
                } else if (knownLongVars_.count(lower)) {
                    convertedValue = "vb6_VariantToLong(" + value + ")";
                } else if (knownDoubleVars_.count(lower)) {
                    convertedValue = "vb6_VariantToDouble(" + value + ")";
                } else if (knownObjectVars_.count(lower)) {
                    // Fix 045: Object (void*) target + Variant value → extract object
                    convertedValue = "vb6_VariantToObjectVal(" + value + ")";
                } else if (knownByteArrayVars_.count(lower)
                           || classByteArrayMembers_.count(lower)) {
                    // Fix 062: Byte array target + Variant value → extract SafeArray1D*
                    // Fix 140: 改用 vb6_VariantToByteArray — Variant 持字符串时按
                    // VB6 语义转字节数组 (旧 SafeArray1D 版本对字符串返回 NULL,
                    // 会清空 m_Caption → 右侧卡片空白).
                    convertedValue = "vb6_VariantToByteArray(" + value + ")";
                } else if (knownClassVars_.count(lower)) {
                    // Fix 086 + Fix <vbeclipse> rev5: 类类型目标 (vb6_cls_X*) + Variant值
                    // → 提取对象指针。
                    // ⚠ rev5: 原式是 `(*)vb6_VariantToObjectVal(v)`, **漏了剥壳**。
                    // 工程类的 Variant 里装的是 `vb6_ComObject` 包装器 (AddPerspective
                    // 里 `vb6_ComObject_FromInstance(...)` 造的), `pdispVal` 指向包装器
                    // 而**不是** vb6_cls_Perspective 本体。直接当类指针用 = 把包装器
                    // 的第一个字段 (vtable 槽) 当成 m_Fields 之类去解引用。
                    // 实测 play78.exe: `Set l_Perspective = m_Perspectives.Item(id)`
                    // 之后 `l_Perspective.Folders.Contains(...)` 崩在
                    // vb6_List_Contains+0xB6, 而 vb6_UBound 打印出
                    // `sig=0x640045 et=7602281` —— 那是 BSTR "Ed..." 的字节,
                    // 因为 m_Folders 槽被前一个成员 (m_PerspectiveId 的 BSTR) 盖住。
                    // 正确口径与 cgen_base_type.cpp 的 deliverToClassSlot 同一口径:
                    // 先取 pdispVal, 再 vb6_ComObject_GetInstance 剥出裸实例。
                    // (GetInstance 对外部 COM 对象返回 NULL = 等效 Nothing, 与 VB6 一致。)
                    convertedValue = "(vb6_cls_" + cIdent(knownClassVars_[lower])
                                   + "*)vb6_ComObject_GetInstance("
                                     "vb6_VariantToObjectVal(" + value + "))";
                } else if (knownArrays_.count(lower)) {
                    // Fix 086: 数组目标 (vb6_SafeArray1D*) + Variant值 → 提取数组
                    convertedValue = "vb6_VariantToSafeArray1D(" + value + ")";
                } else if (node.target && node.target->kind == ASTNodeKind::IdentifierExpr) {
                    auto& id = static_cast<IdentifierExpr&>(*node.target);
                    std::string idLower = id.name;
                    std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
                    if (knownLongPtrVars_.count(idLower)) {
                        convertedValue = "vb6_VariantToLongPtr(" + value + ")";
                    } else if (knownLongVars_.count(idLower)) {
                        convertedValue = "vb6_VariantToLong(" + value + ")";
                    } else if (knownDoubleVars_.count(idLower)) {
                        convertedValue = "vb6_VariantToDouble(" + value + ")";
                    } else if (knownObjectVars_.count(idLower)) {
                        convertedValue = "vb6_VariantToObjectVal(" + value + ")";
                    } else if (knownByteArrayVars_.count(idLower)
                               || classByteArrayMembers_.count(idLower)) {
                        // Fix 140: 同前 — 字符串 Variant 需转字节数组.
                        convertedValue = "vb6_VariantToByteArray(" + value + ")";
                    } else if (knownClassVars_.count(idLower)) {
                        // Fix 086 + Fix <vbeclipse> rev5: 同上方 knownClassVars_ 分支
                        // —— 类槽必须经 vb6_ComObject_GetInstance 剥出裸实例, 不能把
                        // ComObject 包装器当类指针 (实测 play78.exe 的 l_Perspective)。
                        convertedValue = "(vb6_cls_" + cIdent(knownClassVars_[idLower])
                                       + "*)vb6_ComObject_GetInstance("
                                         "vb6_VariantToObjectVal(" + value + "))";
                    } else if (knownArrays_.count(idLower)) {
                        // Fix 086: 数组目标 (vb6_SafeArray1D*) + Variant值
                        convertedValue = "vb6_VariantToSafeArray1D(" + value + ")";
                    }
                }
                // Fix 161f: 上述集合在**每个过程入口被 clear** (见 cgen_decl_func.cpp:48
                // 起), 模块级标量 (如 `Private m_lvHWnd As Long`) 进过程后就没有了集合
                // 条目, 上述判定全落空 → 生成 `m_lvHWnd = vb6_VariantFromComResult(...)`
                // → C2440 "无法从 vb6_VARIANT 转换为 int32_t"
                // (extlist MListViewEx LVSortK `m_lvHWnd = lv.hWnd` 实测)。
                // 模块级变量仍在符号表里, 回查声明类型补一次提取。
                if (convertedValue == value && node.target
                    && node.target->kind == ASTNodeKind::IdentifierExpr) {
                    switch (inferExprType(*node.target)) {
                    case Vb6Type::Long: case Vb6Type::ULong: case Vb6Type::Integer:
                    case Vb6Type::Byte: case Vb6Type::Boolean:
                        convertedValue = "vb6_VariantToLong(" + value + ")"; break;
                    case Vb6Type::Double: case Vb6Type::Single:
                    case Vb6Type::Currency: case Vb6Type::Date:
                        convertedValue = "vb6_VariantToDouble(" + value + ")"; break;
                    case Vb6Type::String:
                        convertedValue = "vb6_VariantToString(" + value + ")"; break;
                    default: break;
                    }
                }
                }
            }
            // Fix 170: 整体数组目标 (`A() = expr`) 右侧若非新建载体则深拷贝, 不制造别名
            if (tgtIsWholeArray170)
                convertedValue = wrapWholeArrayAssign(target, convertedValue, node.target.get());
            // ai/009 5.10: 窄整型目标 (`Byte`/`Integer`/`Long`) 收窄赋值前查溢出。
            // 整体数组目标不是整型, narrowCheckAssign 内部按目标类型判掉, 不会套上。
            c_.emitLine(target + " = " + narrowCheckAssign(node.target.get(), node.value.get(),
                                                             convertedValue) + ";");
        } else {
            // Fix 091l: Variant 目标 ← 具体类型值 (SafeArray1D*/数组载体/标量/COM 结果).
            // C 侧目标是 vb6_VARIANT 结构体, 直接赋 SafeArray1D* 等 → C2440
            // (ToolsLogs: Dim EnumLevelNames As Variant ← Array(...) 物化 _arr_0).
            // vb6_VariantFromValue 是 _Generic 宏: 标量→Variant{Long,Int,Double,Bool,
            // Byte}, BSTR→VariantString, SafeArray1D*→VariantArray, 指针→VariantObject.
            std::string tgtKey091l = target;
            if (tgtKey091l.compare(0, 4, "me->") == 0) {
                tgtKey091l = tgtKey091l.substr(4);
            } else if (tgtKey091l.size() > 4 && tgtKey091l[0] == '(' && tgtKey091l[1] == '*'
                       && tgtKey091l.back() == ')') {
                tgtKey091l = tgtKey091l.substr(2, tgtKey091l.size() - 3);
            }
            std::string tgtLower091l = tgtKey091l;
            std::transform(tgtLower091l.begin(), tgtLower091l.end(),
                           tgtLower091l.begin(), ::tolower);
            bool tgtIsVariant091l = knownVariantVars_.count(tgtLower091l)
                                    || classVariantMembers_.count(tgtLower091l);
            if (!tgtIsVariant091l && !valueIsVariant) {
                tgtIsVariant091l = cExprIsVariant(target);
            }
            if (tgtIsVariant091l && !valueIsVariant && !value.empty()) {
                // Fix 198: 布尔值装箱走 VT_BOOL 而不是 _Generic 的 short 分支.
                value = boxToVariant(node.value.get(), value);
            }
            // Fix 092w: Byte 数组目标 + 具体值 (StrConv/字符串) → 字节数组语义
            if (tgtIsByteArray092w && !tgtIsVariant091l && !value.empty()) {
                value = rewriteByteArrayValue(value);
            }
            // Fix 170: 整体数组目标 (`A() = expr`) 右侧若非新建载体则深拷贝, 不制造别名
            if (tgtIsWholeArray170 && !value.empty())
                value = wrapWholeArrayAssign(target, value, node.target.get());
            // ai/009 5.10: 同上, Variant 目标分支同样收窄检查
            c_.emitLine(target + " = " + narrowCheckAssign(node.target.get(), node.value.get(),
                                                             value) + ";");
        }
    }