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

    // BSTR赋值检测: 如果目标是BSTR变量, 使用vb6_BSTR_Assign防止悬垂指针和双重释放
    bool targetIsBstr = false;
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
    }
    // P8.4: Variant赋值检测 - 如果目标是Variant变量, 用Variant构造函数包装值
    bool targetIsVariant = false;
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
    }
    if (targetIsBstr) {
        // Fix 038b-6: BSTR 目标 + Variant 值 → 先提取 BSTR
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
        if (valueIsVariant) {
            value = "vb6_VariantToString(" + value + ")";
        }
        // Fix 140: RHS 是 Byte() 数组 (类成员 m_Caption 或已知字节数组变量) 且目标是
        // BSTR → 按 UTF-16LE 字节还原为字符串. 不能把数组指针当 BSTR 交给
        // vb6_BSTR_Assign (SysStringLen 会读数组前的堆元数据当长度 → 堆损坏).
        // 场景: LabelPlus.ctl `Property Get Caption() As String: Caption = m_Caption`.
        // 记账这一支改没改成 (给下面 #122 那支看): 值本身是 Byte() 数组 ⇒ inferExprType 会
        // 报 Byte, 但那已经是**转好的字符串**了, 再套一层 CStrByte 是错的 (实测把
        // LabelPlus 的 `Caption = m_Caption` 发成 vb6_CStrByte(vb6_ByteArrayToString(..)))。
        bool wrappedByteArr140 = false;
        if (!valueIsVariant) {
            std::string rhs140 = value;
            if (rhs140.compare(0, 4, "me->") == 0) rhs140 = rhs140.substr(4);
            if (rhs140.size() > 4 && rhs140[0] == '(' && rhs140[1] == '*'
                && rhs140.back() == ')') {
                rhs140 = rhs140.substr(2, rhs140.size() - 3);
            }
            std::string rhs140l = rhs140;
            std::transform(rhs140l.begin(), rhs140l.end(), rhs140l.begin(), ::tolower);
            if (rhs140.find_first_of("( .") == std::string::npos
                && (knownByteArrayVars_.count(rhs140l)
                    || classByteArrayMembers_.count(rhs140l))) {
                value = "vb6_ByteArrayToString(" + value + ")";
                wrappedByteArr140 = true;
            }
        }
        // 账 #122: VB6 在**赋值那一刻**就做隐式 CStr —— `Dim s As String: s = 123` 合法，
        // 而这一格以前把值原样交给 vb6_BSTR_Assign。后果分两种，别当成同一种：
        //   整型那半 (Long/Integer/Boolean/Byte) 在 C 里只是警告 (int → BSTR 形参)，数字被
        //     当指针存进 String 变量，**第一次读/释放就崩** —— 实测修复前二进制 `s = n` 是
        //     BUILD-RC=0 而跑起来 SIGSEGV (rc=139)，所以它不是"编不过"，是静默埋雷;
        //   浮点/日期那半 (Double/Currency/Date) 才是响的 cl `error C2440 无法从 double
        //     转换为 BSTR`。
        // 这一支只在 inferExprType 认得出的**标量**类型上折 (isScalarImplicitStr159:
        // Integer/Long/Single/Double/Byte/Boolean/Currency/Date)，String/Variant/Object/
        // 未知一律不介入。正因为整型那半今天过得去，本批的护栏不是"现存工程零改动"而是
        // "改动必须逐条解释得清": 实测 7 枚存量工程里 6 枚逐字节相同，Charts 2020 只剩
        // `sDiplay = CLng(me->m_Value)` 那一行 —— 那是一处真实修复而非回归。
        // 复用形参那条路 (Fix 159-B) 的同一个 wrapToBSTR, 不另开一套折算表。
        // 上面两支 (Variant 提取 / Byte() 还原) 已经把值转成了字符串, 不能再套一层。
        bool wrapped122 = false;
        if (!valueIsVariant && !wrappedByteArr140 && node.value) {
            Vb6Type vt122 = inferExprType(*node.value);
            if (isScalarImplicitStr159(vt122)) {
                std::string w122 = wrapToBSTR(value, *node.value);
                if (w122 != value) {
                    value = w122;
                    wrapped122 = true;
                }
            }
        }
        // 账 #113 / #119: 值是不是"自有的临时串" —— **字符串**字面量 / 拼接 / 返回 String 的
        // 工程函数或 Property Get (返回变量 vb6_ret_* 在入口处已按值语义拷过)。变量、形参、
        // 成员字段、数组元素都是**借来的**, 只能深拷贝。
        // #119 之前这条判据只服务模块级目标 (#113 那格为了少翻发码面刻意加了
        // `bstrTargetViaTypeQuery` 那道闸), 局部/形参/UDT String 字段三条老路一律 Assign
        // ⇒ 每句从自有临时来的赋值都漏一只堆块 (`s = Left(x,2)` 就是 64 字符 + 块头那种
        // 146 字节级的漏, 循环里就是纯漏, 旧值要等下一次赋值才被释放)。
        // 铺开的底气是这一轮把 RTL 全扫了一遍: 77 枚 `BSTR` 返回函数里只有 3 枚发的是**借来
        // 的**指针 (vb6_ErrDescription / vb6_ErrSource 直接回 vb6_err 那两个字段、
        // vb6_ho_variantToBstr 回 v->bstrVal), 而它们都只能以成员访问的形态到达
        // (`Err.Description`), 下面的 IdentifierExpr 那一支根本捡不到 ⇒ 不会被误判自有。
        // 字面量那支限定成 LiteralKind::String: 数值/布尔字面量赋给 String 目标这一格发的是
        // 裸数字 (`vb6_BSTR_Assign(&s, 123)` → cl C2440, 已另立一账), 把它标成自有只会把
        // "当场编译不响"变成"存进一只假指针", 不能顺手动。
        bool valueOwned113 = false;
        if (node.value) {
            switch (node.value->kind) {
                case ASTNodeKind::LiteralExpr:
                    valueOwned113 = (static_cast<LiteralExpr&>(*node.value).literalKind
                                     == LiteralKind::String);
                    break;
                case ASTNodeKind::BinaryExpr: {
                    auto& vb113 = static_cast<BinaryExpr&>(*node.value);
                    valueOwned113 = (vb113.op == BinaryOp::Concat || vb113.op == BinaryOp::Add);
                    break;
                }
                case ASTNodeKind::IndexOrCallExpr: {
                    auto& vc113 = static_cast<IndexOrCallExpr&>(*node.value);
                    if (vc113.callee && vc113.callee->kind == ASTNodeKind::IdentifierExpr) {
                        auto* vSym113 = symTab_.lookup(
                            static_cast<IdentifierExpr&>(*vc113.callee).name);
                        valueOwned113 = (vSym113
                                         && (vSym113->kind == SymbolKind::Function
                                             || vSym113->kind == SymbolKind::PropertyGet));
                    }
                    break;
                }
                default:
                    break;
            }
        }
        // #122 折出来的 `vb6_CStrXxx(...)` 是刚 alloc 的新块, 与字面量/拼接同属"自有"。
        if (valueOwned113 || wrapped122) {
            c_.emitLine("vb6_BSTR_AssignMove(&" + target + ", " + value + ");");
        } else {
            c_.emitLine("vb6_BSTR_Assign(&" + target + ", " + value + ");");
        }
    } else if (targetIsVariant) {
        // P8.4: 包装值为vb6_VARIANT, 先释放旧BSTR
        std::string wrappedValue = wrapVariantValue(node.value.get(), value);
        // Fix <vbeclipse> rev16: Variant **数组元素**槽位走"接管"语义。
        // `Clear + 浅拷贝` 对"从别处借来的值"不足: RHS 是 ByRef Variant 实参或相邻
        // 元素时, wrapVariantValue 给出的是 vb6_VariantFromValue(...) —— 对 vb6_VARIANT
        // 是 identity **浅拷贝**, 槽与借用方共用同一只 BSTR / 同一个 Dispatch, 借用方
        // 一 Clear 槽即悬垂 (地址被后续同尺寸分配复用 ⇒ 每槽都读"最后一次写入的值")。
        // 实测 tests/ve_list (工程内 List.cls 的 m_Items) idx=0/1/2 全答 "Gamma";
        // 真工程 play78: Folder.Views.Item(0) 取回 1 字符垃圾 → 后续 NULL 解引用。
        // vb6_VariantAssign = Clear(dst) + 深拷贝(BSTR 另分配 / Dispatch AddRef),
        // 与同一个类的 m_Keys (String 数组, 走 vb6_BSTR_Assign) 同一口径。
        bool targetIsVariantArraySlot =
            (target.compare(0, 7, "VB6_SA_") == 0
             && target.find("(vb6_VARIANT,") != std::string::npos);
        if (targetIsVariantArraySlot) {
            c_.emitLine("vb6_VariantAssign(&" + target + ", " + wrappedValue + ");");
        } else {
            c_.emitLine("vb6_VariantClear(&" + target + ");");
            c_.emitLine(target + " = " + wrappedValue + ";");
        }
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
    } else if (target.find("vb6_VariantArrayGet(") == 0) {
        // Fix 038 Group 2: vb6_VariantArrayGet(&arr, idx) = value
        //   → vb6_VariantArraySet(&arr, idx, vb6_VariantFromValue(value))
        // VariantArrayGet 返回右值, 不能赋值; 改写为 VariantArraySet.
        size_t argStart = target.find('(');
        size_t argEnd = target.rfind(')');
        if (argStart != std::string::npos && argEnd != std::string::npos && argEnd > argStart) {
            std::string args = target.substr(argStart + 1, argEnd - argStart - 1);
            c_.emitLine("vb6_VariantArraySet(" + args + ", "
                        + boxToVariant(node.value.get(), value) + ");");
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