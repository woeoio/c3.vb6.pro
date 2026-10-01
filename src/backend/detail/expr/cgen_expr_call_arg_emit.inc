// cgen_expr_call_arg_emit.inc - visit(IndexOrCallExpr&) 函数体片段：位置实参发射（一）—— UDT 打包、Declare ANSI 与 ByRef 取地址（位于实参 for 循环内）
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1771~2136 行，逐行未改。

    for (size_t i = 0; i < node.positional.size(); i++) {
        // Fix 142: 省略的位置实参 (VB6 Foo a, , c) — 按形参缺省值填充, 不发射
        // 占位表达式 (占位 0/Empty 会覆盖缺省值, 且 _has_ 标志会被误判为已传).
        if (!calleeIsBuiltin && !calleeIsDeclare && i < calleeParams.size()
            && node.omittedArgs.count(i)) {
            const auto& op142 = calleeParams[i];
            std::string odefVal142;
            if (op142.hasDefaultValue && !op142.defaultValueExpr.empty()) {
                odefVal142 = op142.defaultValueExpr;
            } else {
                odefVal142 = defaultValue(op142.type);
            }
            if (op142.isByVal || op142.isParamArray) {
                args.push_back(odefVal142);
            } else {
                std::string ocType142 = mapType(op142.type);
                if (op142.type == Vb6Type::Variant || op142.type == Vb6Type::Empty
                    || op142.type == Vb6Type::Null || op142.type == Vb6Type::Object) {
                    args.push_back("&(" + ocType142 + "){0}");
                } else {
                    args.push_back("&(" + ocType142 + "){" + odefVal142 + "}");
                }
            }
            continue;
        }
        // Fix 090q: 实参是「返回 UDT 的函数调用」→ Declare As Any ByRef 打包须传
        // 临时 UDT 地址 (void*)&(vb6_type_X){...}; 否则强转 struct 值 → C2440
        // (cZipArchive pvVfsSetEof: SetFileTime ..., pvToFileTime(...) As Any)
        std::string argUdtRetCType090q;
        if (node.positional[i] && node.positional[i]->kind == ASTNodeKind::IndexOrCallExpr) {
            auto& nc090q = static_cast<IndexOrCallExpr&>(*node.positional[i]);
            if (nc090q.callee && nc090q.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& nid090q = static_cast<IdentifierExpr&>(*nc090q.callee);
                auto it090q = funcUdtRetCType_.find(Symbol::toLower(nid090q.name));
                if (it090q != funcUdtRetCType_.end()) argUdtRetCType090q = it090q->second;
            }
        }
        // 形参是对象类型 (Object/Control) 且实参是窗体控件标识符时, 抑制默认属性
        // 解析: VB6 把控件作为对象引用传递, 而非默认属性值. 否则 ByRef Object
        // 实参生成默认属性 (BSTR) 指针, 被调方 *objControl 读到 VARIANT.vt
        // (如 VT_BSTR=8) → 运行期崩溃 (BalloonTooltips cTT.CreateToolTip).
        bool savedSuppressProp = suppressDefaultProp_;
        if (i < calleeParams.size()
            && static_cast<Vb6Type>(static_cast<uint16_t>(calleeParams[i].type)
                 & ~static_cast<uint16_t>(Vb6Type::Array)) == Vb6Type::Object
            && node.positional[i] && node.positional[i]->kind == ASTNodeKind::IdentifierExpr) {
            std::string idLower = Symbol::toLower(
                static_cast<IdentifierExpr&>(*node.positional[i]).name);
            if (knownFormControls_.count(idLower) || knownWithEventsCtrlVars_.count(idLower)) {
                suppressDefaultProp_ = true;
            }
        }
        emitExpr(*node.positional[i]);
        suppressDefaultProp_ = savedSuppressProp;
      // COM属性标记残留: MsgBox dic.Count 等场景 — 参数是COM属性读取但标记未被消费
        // 统一用后期绑定(IDispatch), 避免vtable签名不匹配问题
        if (isComMarker_) {
            isComMarker_ = false;
            std::string objExpr = std::move(comObjExpr_);
            std::string memName = std::move(comMemberName_);
            // Fix <vbeclipse>: COM 成员访问的接收者是 Variant 时先提取对象指针。
            //   obj_dispatch 只在 knownVariantVars_ 分支 (P24-04) 包了
            //   vb6_VariantToObject(&var); 落到 typed/object 分支的 Variant 局部
            //   变量 (如 ucTabStrip.ctl For-Each 推断对象名?? l_Tab As Variant
            //   保存 Tab 对象, .Left/.Width 读) 会裸交给 vb6_ComGet*Prop → C2172
            //   "实参不是指针" (ucTabStrip.c 312/314/387). 函数调用结果也是
            //   Variant (vb6_List_Item(...)) 时用 vb6_VariantToObjectVal 提取.
            {
                std::string ojTrim = objExpr;
                while (ojTrim.size() >= 2 && ojTrim.front() == '(' && ojTrim.back() == ')')
                    ojTrim = ojTrim.substr(1, ojTrim.size() - 2);
                std::string ojLower = Symbol::toLower(ojTrim);
                if (knownVariantVars_.count(ojLower)) {
                    objExpr = "vb6_VariantToObject(&" + objExpr + ")";
                } else if (cExprIsVariant(objExpr)) {
                    objExpr = "vb6_VariantToObjectVal(" + objExpr + ")";
                }
            }
            if (isEarlyBoundCom_ && earlyBoundSym_) {
                isEarlyBoundCom_ = false;
                const Symbol* comSym = earlyBoundSym_;
                earlyBoundSym_ = nullptr;
                std::string memLower = memName;
                std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                auto it = comSym->comMethods.find(memLower);
                if (it != comSym->comMethods.end() && it->second.isPropertyGet) {
                    const auto& sig = it->second;
                    std::string returnType = mapType(sig.returnType);
                    if (returnType == "int32_t" || returnType == "int16_t") {
                        lastExpr_ = "vb6_ComGetIntProp(" + objExpr + ", L\"" + memName + "\")";
                    } else if (returnType == "BSTR") {
                        lastExpr_ = "vb6_ComGetStringProp(" + objExpr + ", L\"" + memName + "\")";
                    } else if (returnType == "double" || returnType == "float") {
                        lastExpr_ = "vb6_ComGetDoubleProp(" + objExpr + ", L\"" + memName + "\")";
                    } else if (returnType == "void*") {
                        lastExpr_ = "vb6_ComGetObjectProp(" + objExpr + ", L\"" + memName + "\")";
                    } else {
                        lastExpr_ = "vb6_VariantFromComResult(vb6_ComGetProp(" + objExpr + ", L\"" + memName + "\"))";
                    }
                } else {
                    lastExpr_ = "vb6_VariantFromComResult(vb6_ComGetProp(" + objExpr + ", L\"" + memName + "\"))";
                }
            } else {
                lastExpr_ = "vb6_VariantFromComResult(vb6_ComGetProp(" + objExpr + ", L\"" + memName + "\"))";  /* P25: late-bound VARIANT */
            }
        }
        std::string argVal = std::move(lastExpr_);

        // Fix 161b-decl-out: Declare A 版 String 出参回写 —— 判断实参是否为"可赋值左值"。
        // 变量 / 成员访问 / With 成员 = 左值; IndexOrCallExpr 要再分: callee 是**数组变量**
        // 才可赋值, 是 Sub/Function/Declare*/Property 则是函数调用 (如 `String$(256, 0)`),
        // VB6 里这种临时值收不到回写 —— 且生成 `&(vb6_String(...))` 会直接 C2102 "&"要求左值。
        auto isAssignableOutTarget161b = [&](Expr* pa) -> bool {
            if (!pa) return false;
            switch (pa->kind) {
                case ASTNodeKind::IdentifierExpr:
                case ASTNodeKind::MemberAccessExpr:
                case ASTNodeKind::WithMemberExpr:
                    return true;
                case ASTNodeKind::IndexOrCallExpr: {
                    auto* ice = static_cast<IndexOrCallExpr*>(pa);
                    if (!ice || !ice->callee) return false;
                    if (ice->callee->kind != ASTNodeKind::IdentifierExpr) {
                        // obj.Method(...) / obj.Prop(...) 也当调用处理 (成员方法不常见于
                        // 出参缓冲位); 数组的成员访问形态无实际用例, 保守判为不可赋值。
                        return false;
                    }
                    auto* ce = static_cast<IdentifierExpr*>(ice->callee.get());
                    const Symbol* cs = symTab_.lookup(ce->name);
                    if (!cs) return false;
                    switch (cs->kind) {
                        case SymbolKind::Sub:
                        case SymbolKind::Function:
                        case SymbolKind::DeclareSub:
                        case SymbolKind::DeclareFunc:
                        case SymbolKind::PropertyGet:
                        case SymbolKind::PropertyLet:
                        case SymbolKind::PropertySet:
                        case SymbolKind::ComClass:
                        case SymbolKind::ComInterface:
                        case SymbolKind::ComModule:
                        case SymbolKind::ComGlobalNs:
                        case SymbolKind::Class:
                            return false;   // 函数/过程调用, 非左值
                        default:
                            return true;    // 数组变量 arr(i) 等
                    }
                }
                default:
                    return false;
            }
        };

        // Fix 161c: 上面那条 AST 判据**只够判"语句层面是不是可赋值位置"**, 不足以保证
        // 生成出来的 C 表达式本身是左值。反例 (Charts 2020 / LabelPlus / PropPagLP):
        //   CreateWindowEx(..., "Edit", LabelPlus1.Caption, ...)
        // 第 3 个实参 AST 是 MemberAccessExpr (判 true), 但它是**后期绑定 COM 属性读**,
        // argVal 被换成 vb6_VariantToString(vb6_VariantFromComResult(vb6_ComGetProp(...)))
        // —— 纯右值。登记回写后生成
        //   vb6_BSTR_Assign(&(vb6_VariantToString(...)), ...)
        // → error C2102: "&"要求左值 (PropPagLP.c:279)。
        // 所以登记前**必须再看一眼生成串**: 只有当它是纯左值形态 (标识符 / 字段链 /
        // 解引用 / VB6_SA_AT / me->) 时才回写; 任何含函数调用的一律视为临时值, 与
        // VB6 "临时表达式收不到回写" 的语义一致。
        // 口径与下方 Fix 090k/090q 的 isLValue 判定保持同源 (标识符/./->/[]/解引用)。
        auto isCLvalueExpr161c = [](const std::string& e) -> bool {
            size_t s = 0;
            while (s < e.size() && (e[s] == ' ' || e[s] == '\t')) s++;
            if (s >= e.size()) return false;
            // 值宏 / 字面量: 非左值
            if (e.compare(s, 5, "NULL") == 0) return false;
            if (e[s] == '"' || e[s] == 'L' || std::isdigit(static_cast<unsigned char>(e[s]))) return false;
            // 允许的前导形态: 标识符 / (*p) 解引用 / &x / VB6_SA_AT(
            bool deref = false;
            if (e[s] == '&') { return false; }              // 已是取址结果, 不能再 &
            if (e[s] == '(' && s + 1 < e.size() && e[s + 1] == '*') {
                deref = true; s++;
            }
            if (!(std::isalpha(static_cast<unsigned char>(e[s])) || e[s] == '_')) return false;
            // 扫描: 括号外只允许 标识符字符 / . / -> / [] 下标; 出现 '(' 即函数调用 ⇒ 右值
            int sub = 0;
            for (size_t ci = s; ci < e.size(); ci++) {
                char c = e[ci];
                if (c == '[') { sub++; continue; }
                if (c == ']') { if (sub == 0) return false; sub--; continue; }
                if (sub > 0) continue;
                if (c == '(') return false;                  // 函数调用/强制转换 ⇒ 非左值
                if (std::isalnum(static_cast<unsigned char>(c)) || c == '_') continue;
                if (c == '.' || c == '-') continue;
                if (c == '>') continue;
                if (c == ' ' || c == '\t') continue;
                return false;
            }
            (void)deref;
            return true;
        };

        // M22: Declare ANSI函数 - ByVal String参数需要BSTR->ANSI转换
        // 生成临时char*变量, 调用后释放, 无内存泄露
        if (isDeclareAnsiCall && i < calleeParams.size() && calleeParams[i].isByVal
            && calleeParams[i].type == Vb6Type::String) {
            std::string ansiVar = "_ansi_" + std::to_string(ansiCounter_++);
            // Fix 110x: BSTR 形参的实参若是 Variant 表达式 (后期绑定控件属性读
            // vb6_VariantFromComResult(vb6_ComGetProp(...)) 等), 直接交给
            // vb6_BSTR_ToANSI(BSTR) → C2440 "无法从 vb6_VARIANT 转换为 BSTR"
            // (Charts 2020 PropPagLP.c 256: CreateWindowEx(..., LabelPlus1.Caption, ...)).
            if (cExprIsVariant(argVal)) {
                argVal = "vb6_VariantToString(" + argVal + ")";
            }
            c_.emitLine("char* " + ansiVar + " = vb6_BSTR_ToANSI_cap(" + argVal + ");  /* Fix 161b: ByVal String 作可写缓冲, 按 BSTR 全长转 (含内嵌 NUL) */");
            ansiTempsToFree_.push_back(ansiVar);
            // Fix 161b-decl-out: 登记出参回写。VB6 的 ByVal String 是"可写缓冲",
            // 被调用方写入的字节要回写成新 String (GetUserName/GetModuleFileName
            // 正是靠这条)。仅当实参是**可赋值左值**(变量/数组元素)才回写 ——
            // 临时表达式 (如 String$(256,0)) 在 VB6 里也收不到回写, 保持一致。
            // 容量取实参当前字符串长度 (VB6 缓冲语义的近似; cap<=0 则由 NUL 定长)。
            if (i < node.positional.size() && isAssignableOutTarget161b(node.positional[i].get())
                && !cExprIsVariant(argVal) && isCLvalueExpr161c(argVal)) {
                // 回写目标须是 BSTR 左值 (交给 vb6_BSTR_Assign 的 BSTR*)
                ansiOutParams_.push_back({ansiVar, argVal, "0"});  // cap=0: 由 NUL 定长
            }
            argVal = ansiVar;
        }

        // Fix 175: 用户 Sub/Function/Property 的 **ByVal String 形参**收到数值/日期实参
        // (`ShowStr(d)` / `ShowStr(n)`)。VB6 按 CStr 隐式转字符串; C 侧形参是 BSTR,
        // 裸传 double → C2440, 裸传 int32_t → 只报 C4047 警告, 运行期把整数当指针
        // 解引用 (0xC0000005 或读到垃圾字符串)。按形参类型补 wrapToBSTR。
        if (!isDeclareAnsiCall && i < calleeParams.size() && calleeParams[i].isByVal
            && calleeParams[i].type == Vb6Type::String && node.positional[i]
            && !cExprIsVariant(argVal)) {
            Vb6Type at175 = inferExprType(*node.positional[i]);
            if (at175 == Vb6Type::Date || at175 == Vb6Type::Double || at175 == Vb6Type::Single
                || at175 == Vb6Type::Currency || at175 == Vb6Type::Long
                || at175 == Vb6Type::Integer || at175 == Vb6Type::Byte
                || at175 == Vb6Type::Boolean || at175 == Vb6Type::LongPtr) {
                argVal = wrapToBSTR(argVal, *node.positional[i]);
            }
        }

        // Fix 210: 用户 Sub/Function/Property 的 **ByVal Variant 形参**收到标量实参
        // (`List.Item(i)` / `List.Item(sKey)`, 形参是 `ByVal Index As Variant`)。
        // C 侧形参是 vb6_VARIANT (按值), 裸传 int32_t → C2440 "无法从 int32_t
        // 转换为 vb6_VARIANT" (VbEclipse ucPerspective 实测 14 处:
        // `With l_Perspective.Folders` / `.Item(i)`)。VB6 这里要把实参装箱进 Variant。
        // 用 RTL 的 vb6_VariantFromValue: 它是 _Generic, 按实参 C 类型选构造器,
        // 且实参**已经是** vb6_VARIANT 时走 vb6_VariantIdentity (原样透传), 所以
        // 不必先判 cExprIsVariant —— 装了也无害, 还能顺带修掉推断成 Variant 的
        // 标量 (CStrDbl/LenB 那类), 与 Fix 024 的用法一致。
        if (!isDeclareAnsiCall && i < calleeParams.size() && calleeParams[i].isByVal
            && (calleeParams[i].type == Vb6Type::Variant
                || calleeParams[i].type == Vb6Type::Empty)
            // <vbeclipse>: 工程类形参的槽 p.type 也是 Variant (见 arg_variant 同一处注释),
            // 这里不能装箱 —— 判据共用 cParamClassPtrType, 与 dll 入口的 extern 同源。
            && cParamClassPtrType(calleeParams[i]).empty()
            // Fix <vbeclipse> <VBFlexGridDemo>: 只要 callee 形参**声明了具体类型名**
            // 且 mapTypeRef 那份权威不认它是 `vb6_VARIANT`, 这条槽**就不是 Variant
            // 容器** —— 装 vb6_VARIANT 会与 callee 侧 typed 槽错位 (VTableHandle.bas
            // 六处 `Call Replace*(This)` → cl 19.51 C2172; VBFlexGrid.c(3032) 与
            // MainForm.c(1044) 的 `As Variant` 形参如果被一起挡 → 反向 C2440)。
            // 关键: 判据必须是 **mapTypeRef 的 C 类型结果** 而不是"typeRefName 是否
            // 非空" —— `As Variant` 的形参 typeRefName 也是非空 ("Variant"), 前者
            // 会误伤。mapTypeRef 认 "Variant" → "vb6_VARIANT", 认 "IOleControl" (未
            // 命中) → "void*", 认工程类 → "vb6_cls_X*" —— 只有前者才是 Variant 槽。
            && (calleeParams[i].typeRefName.empty()
                || cTypeForDeclaredTypeName(calleeParams[i].typeRefName) == "vb6_VARIANT")
            && !cExprIsVariant(argVal)) {
            // 装箱一律走 boxToVariant 这一处权威 (按 VB 声明类型选 ctor)。
            // 直接手写 vb6_VariantFromValue 是让 C 的 _Generic 猜类型: VB 的 Boolean
            // 在 C 侧是 int32_t/int16_t, 于是 VarType(True) 报 3(Long) 而不是 11
            // (实测 DT43 / MV41 / test_bool_display 三条同时红)。
            argVal = boxToVariant(i < node.positional.size() ? node.positional[i].get()
                                                             : nullptr,
                                  argVal);
        }

        // ByVal String: 传 ANSI 缓冲的**地址**, 被调用方写入后按 NUL 截断回写。
        // C3 原先把 BSTR 变量的地址 (&buf, 即 BSTR*=wchar_t**) 直接交给期望 LPSTR
        // 的 A 版 API → 类型错配, 出参读不回来 (t5 实测 name=[)。这里改为生成
        // 临时 char* 缓冲 (先灌入原 BSTR 内容), 传其地址交给下面的 isByRef 加 &。
        if (isDeclareAnsiCall && i < calleeParams.size() && !calleeParams[i].isByVal
            && !calleeParams[i].isParamArray && calleeParams[i].type == Vb6Type::String
            && node.byvalOverrides.count(i) == 0) {
            std::string ansiVarR = "_ansi_" + std::to_string(ansiCounter_++);
            std::string srcExpr = argVal;
            if (cExprIsVariant(srcExpr)) srcExpr = "vb6_VariantToString(" + srcExpr + ")";
            c_.emitLine("char* " + ansiVarR + " = vb6_BSTR_ToANSI_cap(" + srcExpr + ");  /* Fix 161b: ByRef String 可写缓冲, 按 BSTR 全长转 */");
            ansiTempsToFree_.push_back(ansiVarR);
            // 登记回写: 目标为原左值 (仅当可赋值且非 Variant)
            if (i < node.positional.size() && isAssignableOutTarget161b(node.positional[i].get())
                && !cExprIsVariant(argVal) && isCLvalueExpr161c(argVal)) {
                ansiOutParams_.push_back({ansiVarR, argVal, "0"});
            }
            argVal = ansiVarR;  // 已是 char* —— 正是 A 版 API 要的 LPSTR (见下方 isByRef 抑制)
        }

        // ByRef参数: 调用点传指针. 如果实参已经是解引用形式(*x), 取地址还原为x;
        // 如果是普通变量, 加&取地址
        // Fix 072: VB6 允许 ByVal 覆盖 ByRef 声明 (如 SHCreateMemStream(ByVal 0, 0))
        bool argHasByValOverride = node.byvalOverrides.count(i) > 0;
        // Fix 078 rev2: ByRef array parameters are now vb6_SafeArray1D**,
        // so they DO need & at the call site — same as other ByRef params.
        // Removed isArrayParamType exclusion so arrays get & just like non-arrays.
        bool isByRef = (i < calleeParams.size() && !calleeParams[i].isByVal && !calleeParams[i].isParamArray)
                       && !argHasByValOverride;
        // Fix 161b-decl-out: Declare A 版 ByRef String —— VB6 传的是 ANSI 缓冲本身
        // (LPSTR, char*), 不是指针的指针。上面已把 argVal 换成 char* 临时缓冲,
        // 此处若再加 & 会得到 char** 传给 LPSTR → 段错误 (t5 实测 139)。
        if (isDeclareAnsiCall && i < calleeParams.size()
            && !calleeParams[i].isByVal && !calleeParams[i].isParamArray
            && calleeParams[i].type == Vb6Type::String
            && node.byvalOverrides.count(i) == 0) {
            isByRef = false;
        }
        bool calleeParamIsArray = (i < calleeParams.size()
            && (static_cast<uint16_t>(calleeParams[i].type) & static_cast<uint16_t>(Vb6Type::Array)));
        if (isByRef) {
            // Fix 158b: ByRef **数组**形参 (C 形参 vb6_SafeArray1D**) 收到"整体数组引用"
            // 实参. 1 维 arr() (空括号) 在 prelude 生成 VB6_SA_AT(T, arr, 0) (元素左值);
            // 直接作为 ByRef 数组实参会落入下方复合字面量 (&(vb6_SafeArray1D**){元素}),
            // 元素 T 与 ** 指针类型不符 → C2440 (VBFlexGrid.ctl SortGrid/MergeSort:
            //   TextDisplayMergeSortRec(..., Data(), ...) — ReDim Data(Row1 To Row2)
            //   As TINDIRECTMERGESORTDATA 局部数组). 正确语义: 传调用方数组指针
            //   变量/成员的地址 &arr / &me->field. 已是当前函数 ByRef 数组参数的
            //   (*Data) 形态由下方 (*x) 分支还原为 Data, 不在此处处理 (其还原走
            //   calleeParamIsArray 分支把 (*x) → x).
            bool wholeArr158 = false;
            if (calleeParamIsArray && argVal.rfind("VB6_SA_AT(", 0) == 0) {
                std::string inner158 = argVal.substr(10);
                if (!inner158.empty() && inner158.back() == ')') inner158.pop_back();
                std::vector<std::string> parts158;
                size_t cut158 = 0;
                for (size_t k158 = 0; k158 <= inner158.size(); k158++) {
                    if (k158 == inner158.size() || inner158[k158] == ',') {
                        parts158.push_back(inner158.substr(cut158, k158 - cut158));
                        cut158 = k158 + 1;
                    }
                }
                std::string name158, idx158;
                if (parts158.size() == 3) {
                    name158 = parts158[1];
                    idx158 = parts158[2];
                    size_t b158 = name158.find_first_not_of(" \t");
                    if (b158 == std::string::npos) b158 = 0;
                    size_t e158 = name158.find_last_not_of(" \t");
                    if (e158 != std::string::npos) name158 = name158.substr(b158, e158 - b158 + 1);
                    b158 = idx158.find_first_not_of(" \t");
                    if (b158 == std::string::npos) b158 = 0;
                    e158 = idx158.find_last_not_of(" \t");
                    if (e158 != std::string::npos) idx158 = idx158.substr(b158, e158 - b158 + 1);
                }
                // 仅字面 0 下标 (= 空括号整体引用) 时重写. 两种形态:
                //   1. 名字是标识符 (Data / me->field / _vb6_with_N->field) → &name.
                //   2. 名字是 (*X) — 当前函数内 ByRef 数组参数的还原形态 (X 已是
                //      vb6_SafeArray1D**), 直接传 X, 不加 & (否则成 *** C2440).
                //       (MergeSortRec 递归: Call MergeSortRec(..., Data(), Sort))
                if (idx158 == "0" && !name158.empty()) {
                    if (name158.size() > 3 && name158[0] == '('
                        && name158[1] == '*' && name158.back() == ')') {
                        // 判 X 是纯标识符(|->|链 也允许, 如 (*me->Arr))
                        std::string inner158 = name158.substr(2, name158.size() - 3);
                        bool innerOk158 = !inner158.empty()
                            && (std::isalpha(static_cast<unsigned char>(inner158[0])) || inner158[0] == '_');
                        if (innerOk158) {
                            argVal = inner158;
                            wholeArr158 = true;
                        }
                    } else if (std::isalpha(static_cast<unsigned char>(name158[0]))) {
                        argVal = "&" + name158;
                        wholeArr158 = true;
                    }
                }
            }
            // Task #44 根因3: UDT 内 inline 定长数组作为整体数组引用传给动态数组
            // 形参 (`sByteArrayToString(tLogFont.lfFaceName())`, 形参
            // `abBytes() As Byte` → C 形参 vb6_SafeArray1D**)。inline uint8_t[N]
            // 的地址被硬当 SafeArray 载体指针的指针 → 函数体内 (*abBytes) 把
            // lfFaceName 前 4 字节 "MS S" (0x5320534d, comdlg32 写回的字体名)
            // 当指针解引用 → 0xC0000005 (SSTabEx 字体对话框 IDOK 确认实测)。
            // 修复: 调用点包装栈上临时 SafeArray 头 (复合字面量, full-expression
            // 生存期覆盖调用, 与既有 &(vb6_VARIANT){...} 手法同族) 指向 inline
            // 存储 —— LBound/UBound/元素访问全部语义正确。
            // 判定三要件: 实参是空括号整体引用 + callee 是 UDT 字段访问 + 该成员
            // 在 UDT 符号表里 arraySize>0 (inline 定长)。UDT 内动态数组成员
            // (isArrayDynamic, C 字段是 vb6_SafeArray1D*) 与类实例数组字段
            // (me->X) 不命中, 走各自正确路径不受影响。
            // lBound=0 与 UDT 定长数组的元素访问口径一致 (cgen 生成 lfFaceName[i]
            // 直接 C 下标, 不做 lower 偏移), elemSize 用 sizeof(首元素) 对 UDT
            // 元素同样成立。
            if (calleeParamIsArray && !wholeArr158 && node.positional[i]
                && node.positional[i]->kind == ASTNodeKind::IndexOrCallExpr) {
                auto& callWrap44 = static_cast<IndexOrCallExpr&>(*node.positional[i]);
                if (callWrap44.positional.empty() && callWrap44.callee
                    && callWrap44.callee->kind == ASTNodeKind::MemberAccessExpr) {
                    auto& maWrap44 = static_cast<MemberAccessExpr&>(*callWrap44.callee);
                    if (maWrap44.object) {
                        std::string udtCType44 = inferUdtTypeOfExpr(*maWrap44.object);
                        const std::string prefix44 = "vb6_type_";
                        if (udtCType44.size() > prefix44.size()
                            && udtCType44.compare(0, prefix44.size(), prefix44) == 0) {
                            Symbol* udtSym44 = symTab_.lookupModule(
                                udtCType44.substr(prefix44.size()));
                            if (udtSym44 && udtSym44->kind == SymbolKind::UserDefinedType) {
                                std::string memWrap44 = Symbol::toLower(maWrap44.memberName);
                                for (auto& mi44 : udtSym44->udtMembers) {
                                    if (Symbol::toLower(mi44.name) != memWrap44) continue;
                                    if (mi44.arraySize > 0) {
                                        // 字段序: signature, elemType, elemSize,
                                        // lBound, uBound, count, data, isDynamic
                                        // (vb6rtl_array.h vb6_SafeArray1D)。
                                        // 个数用 sizeof(数组)/sizeof(元素) 由 C 编译器
                                        // 自算 —— 语义层 UdtMemberInfo.arraySize 目前
                                        // 只是"是定长数组"的占位 (=1), 不是元素个数。
                                        Vb6Type elemT44 = static_cast<Vb6Type>(
                                            static_cast<uint16_t>(calleeParams[i].type)
                                            & ~static_cast<uint16_t>(Vb6Type::Array));
                                        std::string cnt44 =
                                            "(int32_t)(sizeof(" + argVal
                                            + ")/sizeof((" + argVal + ")[0]))";
                                        // 形参是 vb6_SafeArray1D** (指向"载体指针变量"
                                        // 的指针), 函数体 (*abBytes) 先解一层。所以不能
                                        // 只传 &(结构字面量) (那是 SafeArray1D*, 解一层
                                        // 会把 signature 0x5A1D 当指针 → AV 读 0x5a1d,
                                        // 首版实测) —— 外面再包一层"1 元素指针数组字面量"
                                        // 取址, 得到真正的 SafeArray1D**。两个复合字面量
                                        // 都在最外层 full-expression 内生存, 覆盖调用。
                                        argVal = "&(vb6_SafeArray1D*[]){&(vb6_SafeArray1D){0x5A1D, "
                                            + mapSaElemType(elemT44)
                                            + ", (int32_t)sizeof((" + argVal + ")[0])"
                                            + ", 0, (" + cnt44 + " - 1), " + cnt44
                                            + ", (void*)(" + argVal + "), 0}}";
                                        wholeArr158 = true;
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            if (wholeArr158) {
                // 已重写为单个取址形式, 跳过下方 ByRef 分发.
            } else if (argVal.size() > 3 && argVal.substr(0, 2) == "(*" && argVal.back() == ')') {
                // (*x) → depends on callee param type:
                // If callee param is array (vb6_SafeArray1D**), arg (*x) means x is already vb6_SafeArray1D**,
                // so pass x directly (not &x which would be ***).
                // If callee param is non-array, (*x) → depends on whether x is a ByRef param of current function:
                //   - If x is a current-function ByRef param (C type T*), then (*x) is T, and callee expects T*.
                //     Since x itself is already T*, pass x directly (not &x which would be T**).
                //   - Otherwise (e.g. local variable holding a pointer), (*x) → &x to restore the pointer.
                std::string innerName = argVal.substr(2, argVal.size() - 3);
                if (calleeParamIsArray) {
                    argVal = innerName;
                } else {
                    // Fix 079: check if innerName is a ByRef param of the current function
                    bool isCurrentByRefParam = false;
                    Vb6Type curPType189 = Vb6Type::Unknown;
                    if (currentProc_) {
                        for (auto& p : currentProc_->params) {
                            if (Symbol::toLower(p.name) == Symbol::toLower(innerName) && !p.isByVal && !p.isParamArray) {
                                bool paramIsArray = (static_cast<uint16_t>(p.type) & static_cast<uint16_t>(Vb6Type::Array)) != 0;
                                if (!paramIsArray) {
                                    isCurrentByRefParam = true;
                                    curPType189 = p.type;
                                }
                                break;
                            }
                        }
                    }
                    if (isCurrentByRefParam) {
                        // ByRef param of current function → already T*, pass directly
                        // Fix 189: 例外 — 被调方形参是 ByRef Variant (C: vb6_VARIANT*)
                        // 时不能直传 T*: 被调方 vb6_VariantFromValue((*p)) 会把 T 的
                        // 字节当 VARIANT.vt 解析 → 字段错位 → 解引用垃圾 BSTR →
                        // 0xC0000005. 按当前形参的 VB 类型重打包为字段式 VARIANT.
                        std::string rep189;
                        if (i < calleeParams.size()
                            && (calleeParams[i].type == Vb6Type::Variant
                                || calleeParams[i].type == Vb6Type::Empty)) {
                            rep189 = wrapByRefVariantParamArg(innerName, curPType189);
                        }
                        argVal = rep189.empty() ? innerName : rep189;
                    } else {
                        // Local pointer variable → restore pointer with &
                        argVal = "&" + innerName;
                    }
                }
            } else if (argVal.size() > 2 && argVal.substr(0, 2) == "me" && argVal[2] == '-') {
                // me->field → &(me->field) (类成员字段取地址)
                argVal = "&(" + argVal + ")";
            } else {
                // Fix 064: As Any (Vb6Type::Unknown) ByRef 参数 — 传 void* 指针,
                // 不包装为 VARIANT. VB6 中 As Any 表示"任意类型指针",
                // 对数组元素应传 &(VB6_SA_AT(...)) 即首元素地址.
                // UDT ByRef 参数同理 — C 函数签名已是 UDT*, 传 &argVal 即可.
                bool isAsAnyParam = (i < calleeParams.size() && calleeParams[i].type == Vb6Type::Unknown);
                bool isUdtByRefParam = (i < calleeParams.size() && calleeParams[i].type == Vb6Type::UserDefinedType);

                // Fix 164x: 左值判定必须先剥离 UDT 字段链的注释后缀
                // "  /* udt objfield <CType> */" (cgen_util_classtype.cpp:392 追加).
                // 否则 Release.IUnk 判为字段链时因含空格/星号落入强转分支,
                // CopyMemory 的第一实参生成 (void*)(intptr_t)(Release.IUnk) —
                // 传字段值而非字段地址 → 运行时空指针写崩溃.
                // 仅剥离用于左值判断, 输出仍用原样 argVal 保留注释.
                std::string argValLv = argVal;
                size_t lvCmPos = argValLv.find("  /* udt objfield ");
                if (lvCmPos != std::string::npos) {
                    argValLv = argValLv.substr(0, lvCmPos);
                }

                // Fix 064b: VB6_SA_AT(...) 数组元素作为 ByRef 实参时,
                // 不应包装为 VARIANT 复合字面量, 应直接取地址.
                // 典型: PolyPolygon(hDC, uPoints(0), aSizes(0), nCount)
                //   uPoints(0) → VB6_SA_AT(vb6_type_POINTAPI, uPoints, 0)
                //   应生成 (void*)&(VB6_SA_AT(...)) 而非 (&(vb6_VARIANT){.bstrVal=...})
                bool isSaAtExpr = (argValLv.find("VB6_SA_AT(") == 0);
                if (isSaAtExpr && !isAsAnyParam && !isUdtByRefParam) {
                    // 数组元素作为 ByRef 参数: 当 calleeParams 缺失或类型不匹配时,
                    // 对 Declare 函数的 As Any 参数一律传 void*
                    if (calleeIsDeclare || i >= calleeParams.size()) {
                        isAsAnyParam = true;
                    }
                }

                if (isAsAnyParam || isUdtByRefParam) {
                    // As Any ByRef / UDT ByRef: 左值取地址, 非左值(字面量等)强转为void*
                    // Fix 069b: 对字面量(如 0)不能取地址 &(0) → C2101,
                    //   应直接强转 (void*)(intptr_t)(0).
                    bool isSimpleIdent = !argValLv.empty() && (std::isalpha(static_cast<unsigned char>(argValLv[0])) || argValLv[0] == '_');
                    if (isSimpleIdent) {
                        for (char c : argValLv) {
                            if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_') {
                                isSimpleIdent = false;
                                break;
                            }
                        }
                    }
                    // Fix 056b: UDT字段链 (uFile.Data.ftLastWriteTime) 也是左值,
                    // 作为 ByRef UDT/AsAny 实参应取地址 &(x) 而非强转 (void*)(intptr_t)(x)
                    // Fix 084m: 字段链须允许 -> — With变量展开 (._vb6_with_2->SendBuffer)
                    // 也是左值字段链, 若判为非左值会生成 (void*)(intptr_t)(udt) → C2440
                    // Fix 190: 再允许 [] 下标 — 定长数组元素也是左值。VBFlexGrid.ctl:28348
                    //   CopyMemory .szText(0), ByVal StrPtr(Text), LenB(Text)
                    // 经 With 展开成 _vb6_with_776->szText[0]; 旧判定因含方括号视其为非左值,
                    // 落入下方 (void*)(intptr_t)(expr) 兜底 → 把首元素的**值**当成目的地址。
                    // NMTTDI 已清零, 该值即 0 → RtlMoveMemory 写 0x0, 鼠标悬停即 0xC0000005。
                    // 方括号内是下标表达式(右值, 任意写法), 括号外仍只允许 标识符/./->
                    bool isUdtFieldChain = !argValLv.empty() && (std::isalpha(static_cast<unsigned char>(argValLv[0])) || argValLv[0] == '_');
                    int chainSubscriptDepth = 0;
                    if (isUdtFieldChain) {
                        for (size_t ci = 0; ci < argValLv.size(); ci++) {
                            char c = argValLv[ci];
                            if (c == '[') { chainSubscriptDepth++; continue; }
                            if (c == ']') {
                                if (chainSubscriptDepth == 0) { isUdtFieldChain = false; break; }
                                chainSubscriptDepth--; continue;
                            }
                            if (chainSubscriptDepth > 0) continue;
                            if (std::isalnum(static_cast<unsigned char>(c)) || c == '_' || c == '.') continue;
                            if (c == '-' && ci + 1 < argValLv.size() && argValLv[ci + 1] == '>') { ci++; continue; }
                            isUdtFieldChain = false;
                            break;
                        }
                        if (chainSubscriptDepth != 0) isUdtFieldChain = false;
                    }
                    bool isSaAt = (argValLv.find("VB6_SA_AT(") == 0);
                    // Fix 086: 宏常量/NULL 不是左值 — TLS_LOCAL_LEGACY_VERSION 是
                    // #define, &(宏) → C2101; NULL 同理.
                    bool isConstMacro = (argValLv == "NULL") || isConstIdent(argValLv);
                    // Fix 090k: (*uFile).BufferArray / (*uFile)->field — ByRef UDT
                    // 参数解引用的字段链也是左值 (As Any / CopyMemory 实参应取地址
                    // &(field), 而非下方 else 的 (void*)(intptr_t)(expr) 强转 —
                    // Variant 字段值强转指针 → C2440 "无法从 vb6_VARIANT 转 intptr_t").
                    bool isDerefFieldChain = false;
                    if (!argValLv.empty() && argValLv[0] == '(' && argValLv.size() > 2 && argValLv[1] == '*') {
                        size_t depth090k = 0;
                        for (size_t ci090k = 0; ci090k < argValLv.size(); ci090k++) {
                            char c090k = argValLv[ci090k];
                            if (c090k == '(') depth090k++;
                            else if (c090k == ')') {
                                depth090k--;
                                if (depth090k == 0) {
                                    if (ci090k + 1 < argValLv.size() &&
                                        (argValLv[ci090k + 1] == '.' ||
                                         (argValLv[ci090k + 1] == '-' && ci090k + 2 < argValLv.size()
                                          && argValLv[ci090k + 2] == '>'))) {
                                        isDerefFieldChain = true;
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    bool isLValue = (isSimpleIdent || isUdtFieldChain || isSaAt || argValLv.find("me->") == 0
                        || (argValLv.size() > 4 && argValLv[0] == '(' && argValLv[1] == '*' && argValLv.back() == ')')
                        || isDerefFieldChain)
                        && !isConstMacro;
                    if (isLValue) {
                        // 左值: 变量名、数组元素、me->field、(*ptr) 解引用 — 可以取地址
                        argVal = "(void*)&(" + argVal + ")";
                    } else if (!argUdtRetCType090q.empty()) {
                        // Fix 090q: 实参是返回 UDT 的函数调用 — VB6 把返回值拷入临时
                        // UDT 再传址 (As Any ByRef). C 不能直接强转 struct 值
                        // ((void*)(intptr_t)(struct) → C2440), 也不能用复合字面量
                        // 内联初始化 ((T){fnRetUdt()} MSVC C 报 C2440 "初始化…无法从
                        // T 转换") — 只能先求值到已声明临时再取址. 声明行挂 pending,
                        // 在包含本实参的语句行输出前落地 (CodeEmitter::flushPending),
                        // 保证 "声明先于引用". 例: cZipArchive pvVfsSetEof 内
                        // SetFileTime(..., pvToFileTime(dLastWriteTime)) (FILETIME).
                        std::string tmpAny090q = "_vb6_anytmp" + std::to_string(tempCounter_++);
                        c_.addPending(argUdtRetCType090q + " " + tmpAny090q
                                      + " = " + argVal + ";");
                        argVal = "(void*)&" + tmpAny090q;
                    } else if (isConstMacro && argVal != "NULL") {
                        // Fix 086: 宏常量 → 用对应类型的复合字面量承载地址
                        std::string ccType = mapType(constIdentType(argVal));
                        if (ccType.empty() || ccType == "vb6_VARIANT") ccType = "int32_t";
                        argVal = "(void*)&(" + ccType + "){" + argVal + "}";
                    } else if (argVal.compare(0, 20, "vb6_VariantArrayGet(") == 0) {
                        // Fix 158f: Variant 数组元素整体引用作为 As Any ByRef 实参
                        // (VTableHandle.bas: VariantCopy ArgListRev(i),
                        // ArgList(UBound(ArgList) - i)). 元素是右值 (vb6_VariantArrayGet
                        // 返回值), (void*)(intptr_t)(vb6_VARIANT) → C2440. 用 RTL
                        // vb6_VariantArrayElemPtr(数组, 下标) 取底层槽位地址 — dest/src
                        // 都指向真实元素, VariantCopy 读写落位正确.
                        std::string inner158f = argVal.substr(20);
                        if (!inner158f.empty() && inner158f.back() == ')') inner158f.pop_back();
                        argVal = "vb6_VariantArrayElemPtr(" + inner158f + ")";
                    } else if (isDefinitelyVariantExpr(*node.positional[i]) || cExprIsVariant(argVal)) {
                        // Fix 158g: Variant 值 (Variant 属性读取函数返回) 作为 As Any
                        // ByRef 实参 — VBFlexGrid.ctl: VariantCopy Cell, .CellTag
                        // (CellTag As Variant → vb6_VBFlexGrid_prop_get_CellTag(me)
                        // 返回 vb6_VARIANT). VB6 传临时副本地址; 复合字面量
                        // (&(vb6_VARIANT){expr}) 编译且运行时正确 (src 以值拷贝),
                        // (void*)(intptr_t)(vb6_VARIANT 结构值) → C2440.
                        // Fix 160-H: 上面那句"编译正确"只对**标量**实参成立。
                        // `(vb6_VARIANT){x}` 是**位置**初始化 → 用 x 初始化首个成员
                        // `.vt` (vb6_vartype 枚举), 而不是整体拷贝。实参本身就是
                        // vb6_VARIANT 时 → VBFlexGrid.c:22606
                        //   C2440 "初始化: 无法从 vb6_VARIANT 转换为 vb6_vartype"。
                        // 本文件其它分支要么用指定初始化器 (.vt=VT_BSTR,.bstrVal=…)
                        // 要么用 {0}, 只有此处漏了。改走 Fix 090q 的成熟套路:
                        // 求值到已声明临时再取址, 声明行挂 pending 保证先于引用。
                        if (cExprIsVariant(argVal)) {
                            std::string tmpVar160h = "_vb6_vartmp" + std::to_string(tempCounter_++);
                            c_.addPending("vb6_VARIANT " + tmpVar160h + " = " + argVal + ";");
                            argVal = "(void*)&" + tmpVar160h;
                        } else {
                            argVal = "(&(vb6_VARIANT){" + argVal + "})";
                        }
                    } else {
                        // 合并口径: 本 else 兜底两侧各自加了判别, 两者互斥可共存 ——
                        // 数值字面量取复合字面量地址 (上游 Fix 182), Variant 值取整型
                        // 载荷 (我们的 Fix 160w), 其余维持裸强转。
                        // Fix 182: As Any ByRef 的实参是数值字面量时, VB6 先把字面量求值到
                        // 临时存储, 再把该存储的地址传给形参. 直接强转 (void*)(intptr_t)(v)
                        // 是把数值本身当成指针 — cAsyncSocket 的
                        // ws_setsockopt(m_hSocket, ucsSolTCP, ucsSsoNoDelay, 1&, 4)
                        // 形参为 int32_t*, 收到地址 0x1 → WSAEFAULT(10014),
                        // Create 失败并关闭 socket, 后续 Bind 报 WSAENOTSOCK.
                        // 实测对照: optval=(char*)1 → ret=-1 wsaerr=10014;
                        //           optval=&(int32_t){1} → ret=0.
                        // 0 保持原有 (void*)0 行为 (Fix 069b 场景需要 NULL 指针),
                        // 其余复杂表达式维持原样, 控制影响面.
                        std::string litBody = argVal;
                        char litSuffix = 0;
                        if (!litBody.empty() && (litBody.back() == 'L' || litBody.back() == 'l'
                                                 || litBody.back() == 'f' || litBody.back() == 'F')) {
                            litSuffix = litBody.back();
                            litBody.pop_back();
                        }
                        bool isNumericLiteral = !litBody.empty();
                        for (char c : litBody) {
                            if (!std::isdigit(static_cast<unsigned char>(c))
                                && c != '.' && c != 'e' && c != 'E') {
                                isNumericLiteral = false;
                                break;
                            }
                        }
                        if (isNumericLiteral && litBody != "0") {
                            std::string ccType = "int32_t";
                            if (litBody.find('.') != std::string::npos
                                || litBody.find('e') != std::string::npos
                                || litBody.find('E') != std::string::npos) {
                                ccType = "double";
                            } else if (litSuffix == 'f' || litSuffix == 'F') {
                                ccType = "float";
                            }
                            argVal = "(void*)&(" + ccType + "){" + argVal + "}";
                        } else if (cExprIsVariant(argVal)) {
                            // Fix 160w: Variant 值 (COM 属性读取 vb6_VariantFromComResult 等)
                            // 作为 As Any ByRef 实参的非左值兜底 — (void*)(intptr_t)(struct)
                            // → C2440, 裸传 → C2172 (MainForm.c OleCreatePropertyFrame 参 8
                            // ByVal Pages.pElems: Web 后端该 Variant 的整型载荷即指针值).
                            // 用 vb6_VariantToLongPtr 取整型载荷再按值压回指针形参.
                            argVal = "(void*)(intptr_t)vb6_VariantToLongPtr(" + argVal + ")";
                        } else {
                            // Fix <vbeclipse>: **复杂右值表达式** 作为 ByRef(As Any/UDT)
                            // 实参 —— VB6 对 ByRef 形参一律传"一个存储的地址", 右值先求值
                            // 落临时再传址. 原先裸强转 `(void*)(intptr_t)(expr)` 是把
                            // **值**当指针, 被调方沿该"地址"读写 → 越界访问.
                            // 实测根因: MagneticWnd.cls
                            //   Private Sub zPatchRel(ByVal nAddr As Long, ...)
                            //       Call RtlMoveMemory(ByVal nAddr + nOffset, _
                            //                          nTargetAddr - nAddr - nOffset - 4, 4)
                            // 第 2 形参是 ByRef As Any, 拿到 (void*)(intptr_t)(rel) 后
                            // RtlMoveMemory 从地址 =rel 读 4 字节 (rel 常是负数/小值)
                            // → 0xC0000005 av read target=0xFEE3A55A; 补丁没写进机器码桩,
                            // 桩里跳 0 → 后续崩溃/堆损坏 (VbEclipse play78.exe 实测).
                            // 与 Fix 090q / Fix 160-H 同一套路: 声明行挂 pending, 保证先于引用.
                            // 例外: "0" 仍是 (void*)0 —— Fix 069b 场景要 NULL 指针语义.
                            if (argVal == "0" || litBody == "0") {
                                argVal = "(void*)(intptr_t)(" + argVal + ")";
                            } else {
                                std::string tmpAnyRel = "_vb6_anytmp" + std::to_string(tempCounter_++);
                                c_.addPending("intptr_t " + tmpAnyRel
                                              + " = (intptr_t)(" + argVal + ");");
                                argVal = "(void*)&" + tmpAnyRel;
                            }
                        }
                    }
                } else {
                // Fix 072b: ByVal 覆盖 As Any 参数 — 传 (void*)(intptr_t)(val),
                // 不走复合字面量路径. 典型: SHCreateMemStream(ByVal 0, 0) 中
                // pInit As Any 是 ByRef, 但 ByVal 0 覆盖为传值, 应生成 (void*)0.
                bool isAsAnyByVal = (i < calleeParams.size() && calleeParams[i].type == Vb6Type::Unknown)
                                    && argHasByValOverride;
                if (isAsAnyByVal) {
                    // Fix 160w: ByVal 覆盖 As Any 的实参恰为 vb6_VARIANT (COM 属性
                    // 结果, 如 MainForm.frm:579 OleCreatePropertyFrame(..., ByVal
                    // Pages.pElems, ...) → 实参 vb6_VariantFromComResult(vb6_ComGetProp(
                    // Pages, L"pElems"))). (void*)(intptr_t)(vb6_VARIANT 结构值) →
                    // C2440 "无法从 vb6_VARIANT 转换到 long", 且按位强转也丢指针.
                    // pElems 底层是 Long 载荷的指针值 (IFontDisp** 数组头), 用
                    // vb6_VariantToLongPtr 取整型载荷再按原样压回 void* 实参.
                    if (cExprIsVariant(argVal)) {
                        argVal = "(void*)(intptr_t)vb6_VariantToLongPtr(" + argVal + ")";
                    } else {
                        argVal = "(void*)(intptr_t)(" + argVal + ")";
                    }
                } else {
                // 变量/非左值 → 复合字面量取地址; 变量 → &变量
                // 检查是否是简单标识符 (变量名, 以字母/下划线开头)
                bool isSimpleIdent = !argVal.empty() && (std::isalpha(static_cast<unsigned char>(argVal[0])) || argVal[0] == '_');
                if (isSimpleIdent) {
                    for (char c : argVal) {
                        if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_') {
                            isSimpleIdent = false;
                            break;
                        }
                    }
                }
                if (isSimpleIdent && !argVal.empty()) {
                    // Fix 079: If the simple identifier is a ByRef param of the current function,
                    // it's already a pointer (T*) — pass directly, don't add &.
                    bool argIsCurrentByRefParam = false;
                    Vb6Type argCurPType189 = Vb6Type::Unknown;
                    if (currentProc_) {
                        for (auto& p : currentProc_->params) {
                            if (Symbol::toLower(p.name) == Symbol::toLower(argVal) && !p.isByVal && !p.isParamArray) {
                                argIsCurrentByRefParam = true;
                                argCurPType189 = p.type;
                                break;
                            }
                        }
                    }
                    if (argIsCurrentByRefParam) {
                        // ByRef param of current function → already T*, pass directly
                        // Fix 189: ByRef Variant 形参时需重打包 (同 (*x) 路径, 见上).
                        std::string rep189b;
                        if (i < calleeParams.size()
                            && (calleeParams[i].type == Vb6Type::Variant
                                || calleeParams[i].type == Vb6Type::Empty)) {
                            rep189b = wrapByRefVariantParamArg(argVal, argCurPType189);
                        }
                        if (!rep189b.empty()) argVal = rep189b;
                    } else if (isConstIdent(argVal)) {
                        // Fix 084aa: 常量是 #define 宏 (vb6_BSTR_FromStr(...) 或数值字面量),
                        // 取址 &ERR_X 展开为函数结果取址 → C2102. 按形参类型复合字面量包装.
                        Vb6Type pType = Vb6Type::Variant;
                        if (i < calleeParams.size()) pType = calleeParams[i].type;
                        argVal = wrapConstArgForByRef(argVal, pType);
                    } else if (argVal == "NULL" || argVal == "nullptr") {
                        // Fix 092l: NULL/nullptr 字面量不可取址 — &NULL 触发
                        // C2101 "常量上的 &" (cTlsSocket 440:
                        //   FireOnCertificate(me, &NULL) → 应为 (&(vb6_VARIANT){0})).
                        // 按形参类型改用 C11 复合字面量.
                        std::string ct092l = "vb6_VARIANT";
                        if (i < calleeParams.size()) ct092l = mapType(calleeParams[i].type);
                        argVal = "(&(" + ct092l + "){0})";
                    } else if (i < calleeParams.size()
                               && !calleeParams[i].isByVal && !calleeParams[i].isParamArray
                               && calleeParams[i].type == Vb6Type::Variant
                               && !cParamClassPtrType(calleeParams[i]).empty()) {                        // Fix <vbeclipse> (回归修复): `p.type == Variant` 有**两种**槽 ——
                        // 真 Variant 槽, 和**工程类槽** (`Sub UseCircle(c As Circle)`:
                        // resolveTypeOrDefault 把类名折成 Variant, 类名只留在 typeRefName,
                        // 而类模块发的是 `vb6_cls_ShapeAct** c`)。后者不是 Variant 槽:
                        // 落进下面的 092y default 分支会被装箱成 vb6_VARIANT*, 与形参
                        // 声明的 vb6_cls_ShapeAct** ABI 不符 → 被调方 `c.Move 1` 拿
                        // Variant 当类指针解引用 → 段错误 (实测 cc_act: 打印完 CC1:OK
                        // 即 rc=139, 门 #112/#116 两条红)。
                        // 判据用 cParamClassPtrType —— 它是"这个槽收不收类实例指针"的
                        // **唯一**口径 (走 cTypeForDeclaredTypeName → mapTypeRef, 与类模块
                        // 发定义那一份同一映射), ByVal 侧 Fix 091t 已在用; 这里补齐 ByRef 侧。
                        //
                        // Fix <vbeclipse> rev7: 交付方式与下游复合字面量路径的同名守卫**统一**
                        // —— 都先落临时再取址。原来这里直接 `"&" + deliverToClassSlot(...)`,
                        // 而 deliverToClassSlot 在实参是 Variant 值时会返回
                        // `(vb6_cls_X*)vb6_ComObject_GetInstance(vb6_VariantToObjectVal(v))`
                        // —— 那是**右值**, &(右值) = C2102。走落临时后两侧生成码同形。
                        {
                            const std::string clsSlot092r = cParamClassPtrType(calleeParams[i]);
                            const std::string tmpCls092r = classSlotTempName(clsSlot092r, tempCounter_++);
                            c_.addPending(clsSlot092r + " " + tmpCls092r + " = "
                                          + deliverToClassSlot(clsSlot092r, argVal) + ";");
                            argVal = "&" + tmpCls092r;
                        }
                    } else if (i < calleeParams.size()
                               && !calleeParams[i].isByVal && !calleeParams[i].isParamArray
                               && calleeParams[i].type == Vb6Type::Variant) {
                        // Fix 092y: ByRef Variant 形参 + 简单标识符实参.
                        // C 签名为 vb6_VARIANT*, 直接 &(BSTR) 会把字符串当 Variant
                        // 解析 → 字段错位 → lDataLen=0、QR 段为空 (mdQRTest2_url).
                        // 与 2039 M22 分支一致: 按实参 VB 类型构造字段式复合字面量.
                        Vb6Type argVbType92y = inferExprType(*node.positional[i]);
                        switch (argVbType92y) {
                            case Vb6Type::String:
                                argVal = "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=" + argVal + "})";
                                break;
                            case Vb6Type::Long:
                            case Vb6Type::Integer:
                                argVal = "(&(vb6_VARIANT){.vt=VT_I4, .lVal=(int32_t)(" + argVal + ")})";
                                break;
                            case Vb6Type::Double:
                            case Vb6Type::Single:
                                argVal = "(&(vb6_VARIANT){.vt=VT_R8, .dblVal=(double)(" + argVal + ")})";
                                break;
                            case Vb6Type::Boolean:
                                argVal = "(&(vb6_VARIANT){.vt=VT_BOOL, .boolVal=(int16_t)(" + argVal + ")})";
                                break;
                            case Vb6Type::Byte:
                                argVal = "(&(vb6_VARIANT){.vt=VT_UI1, .bVal=(uint8_t)(" + argVal + ")})";
                                break;
                            case Vb6Type::Date:
                                argVal = "(&(vb6_VARIANT){.vt=VT_DATE, .dblVal=(double)(" + argVal + ")})";
                                break;
                            default: {
                                // Fix <vbeclipse>: 实参是**晚绑定对象** (C 类型 void* =
                                // IDispatch*) 时, `&argVal` 交出去的是 `void**` —— 只有 4
                                // 字节, 而被调方按 `vb6_VARIANT*` 解析: x86 上 vb6_VARIANT
                                // 是 16 字节, 于是 `vb6_VariantFromValue((*Item))` 读到
                                // **越界的栈字节**; 若这些字节恰好凑出 vt=VT_BSTR/VT_DISPATCH,
                                // 随后的 vb6_VariantClear 会 free 一个垃圾指针 → 堆损坏
                                // (实测 play78.exe: 0xC0000374, 之前是 NULL 数组的 0xC0000005)。
                                // VB6 语义: 非 Variant 实参传给 ByRef Variant 形参要**拷进
                                // 临时 Variant** 再传址。用 vb6_VariantFromValue 装箱 ——
                                // 它是 _Generic, 且对 void* 选 vb6_VariantObject (VT_DISPATCH),
                                // 正是 knownObjectVars_ 那一族在 C 侧的表示。
                                // 实证: ucPerspective.AddView 的 `m_Views.Add ViewId, l_View`
                                // (List.Add(ByVal Key, ByRef Item As Variant, …)),
                                // l_View 由 Fix 179a 装成 IDispatch 后落进 knownObjectVars_。
                                std::string argLs92y = argVal;
                                std::transform(argLs92y.begin(), argLs92y.end(),
                                               argLs92y.begin(), ::tolower);
                                // 两条不同的装箱路径, 顺序不能反:
                                // (1) knownObjectVars_ = 晚绑定对象, C 侧就是 IDispatch* (void*)
                                //     → 直接 vb6_VariantFromValue。
                                // (2) 工程类实例: C 侧是**裸结构体指针** vb6_cls_<C>*, 不是
                                //     IDispatch。直接装箱成 VT_DISPATCH 的话, List/Variant 里
                                //     存的是裸结构体地址, 后续 vb6_VariantClear/Release 会把
                                //     它的首字段 (__comObj) 当 vtable 解引用 → **堆损坏**
                                //     (实测 Perspective.AddFolder 的
                                //      `m_Folders.Add FolderId, AddFolder`, 返回变量是
                                //      `vb6_cls_Folder*`)。必须先经 vb6_ComObject_FromInstance
                                //      包成真 IDispatch —— 与 cgen_setlet Fix 179a 同口径。
                                //     FromInstance 的 1 个引用按头文件说明直接转移给 VARIANT。
                                //     判据 inferClassTypeOfExpr (与 Fix 101 同源)。
                                const bool isObjVar92y = knownObjectVars_.count(argLs92y) > 0;
                                std::string argCls92y =
                                    isObjVar92y ? std::string()
                                                : inferClassTypeOfExpr(*node.positional[i]);
                                // Fix <vbeclipse>: **本函数的返回值变量** (实参 C 名正是
                                // currentReturnVar_, 如 vb6_ret_AddPerspective)。它的 C 类型是
                                // 原生 `vb6_cls_<C>*`, 不是 IDispatch, 而上面两条判据都够不着它:
                                //   - knownObjectVars_ 里没有 (它不是 void*, 是 vb6_cls_*)
                                //   - inferClassTypeOfExpr 按**裸 VB 名**查类, 返回值变量在符号表
                                //     里叫 vb6_ret_<N>, 名字对不上 → 返回空
                                // 于是回落 `&argVal` —— 把 4 字节槽当 16 字节 vb6_VARIANT*
                                // (x64 是 8 vs 24) 交出去, 数组里存进**越界读来的垃圾 Variant**。
                                // 实证: Perspective.AddPerspective 的
                                //   `m_Perspectives.Add PerspectiveId, AddPerspective`
                                // → ShowPerspective 里 `Contains(id)` 为真但 `Item(id)` 返回
                                //   Nothing ⇒ 项目自己 Err.Raise 1 "No perspective found!"
                                //   ⇒ 未处理 ⇒ ExitProcess(1)。
                                // 用 currentReturnCType_ (Fix 054 保存的返回类型 C 名) 判类,
                                // 与 cgen_with.cpp 的 With-返回值变量修复同口径 —— 天然只在本
                                // 过程内成立。
                                if (!isObjVar92y && argCls92y.empty()
                                    && !currentReturnVar_.empty() && argVal == currentReturnVar_
                                    && currentReturnCType_.rfind("vb6_cls_", 0) == 0) {
                                    argCls92y = currentReturnCType_.substr(8);
                                    if (!argCls92y.empty() && argCls92y.back() == '*') argCls92y.pop_back();
                                }
                                // Fix <VBFlexGridDemo>: 已知 **Variant 变量**不能走这两条
                                // 装箱。cExprIsVariant 只认函数表达式 (vb6_VariantXxx( 前缀),
                                // 裸变量名一律 false; 而 knownObjectVars_ 全程不清空, 别的
                                // 模块 `Dim Value As Object` 会把 "value" 留在里面 ⇒
                                // isObjVar92y 为真 ⇒ 把 vb6_VARIANT 强转 (void*) 装箱
                                // ⇒ C2440 "无法从 vb6_VARIANT 转换为 void *"
                                // (VBFlexGrid.c prop_set_CellTag / prop_let_CellText 等 6 处)。
                                // BASE(8c3e981) 发的是 `&Value` (形参本身可寻址), 这里补上
                                // knownVariantVars_ 判据后正好回到那条 else 分支。
                                if ((isObjVar92y || !argCls92y.empty())
                                    && !cExprIsVariant(argVal)
                                    && !knownVariantVars_.count(argLs92y)) {
                                    // 落一个**具名** Variant 临时再传址 ——
                                    // 不能用复合字面量 `(vb6_VARIANT){vb6_VariantFromValue(x)}`:
                                    // 那是聚合初始化, 第一个初始化式会去初始化首成员 `vt`
                                    // (类型 vb6_vartype) → C2440 "无法从 vb6_VARIANT 转换为
                                    // vb6_vartype" (实测 Perspective.c:115 等 7 处)。
                                    // 具名临时也正是 VB6 语义 (非 Variant 实参拷进临时 Variant)。
                                    std::string tmp92y = "_vb6_vararg_" +
                                        std::to_string(ansiCounter_++);
                                    std::string boxed92y = isObjVar92y
                                        ? ("(void*)(" + argVal + ")")
                                        : ("(void*)vb6_ComObject_FromInstance(vb6_FindCoClassDesc(\""
                                           + argCls92y + "\"), (void*)(" + argVal + "))");
                                    c_.emitLine("vb6_VARIANT " + tmp92y
                                        + " = vb6_VariantFromValue(" + boxed92y
                                        + ");  /* Fix <vbeclipse>: ByRef Variant 实参装箱 */");
                                    argVal = "&" + tmp92y;
                                } else {
                                    // Variant 变量 / UDT 字段: 值已是 vb6_VARIANT 且可寻址。
                                    argVal = "&" + argVal;
                                }
                                break;
                            }
                        }
                    } else if (calleeIsDeclare && i < calleeParams.size()
                               && !calleeParams[i].isByVal && !calleeParams[i].isParamArray
                               && calleeParams[i].type != Vb6Type::Variant
                               && calleeParams[i].type != Vb6Type::Unknown
                               && calleeParams[i].type != Vb6Type::UserDefinedType
                               && !(static_cast<uint16_t>(calleeParams[i].type) & static_cast<uint16_t>(Vb6Type::Array))
                               && argVal.size() > 1 && argVal[0] != '_'
                               && (std::isalpha(static_cast<unsigned char>(argVal[0])))
                               && (cExprIsVariant(argVal)
                                   || knownVariantVars_.count(Symbol::toLower(argVal)) > 0)) {
                        // Fix 133: Variant 变量作为 Declare 标量 ByRef 出参
                        // (Dim hFont: GdipCreateFont(..., hFont), hFont 是 Variant)。
                        // C 签名是 int32_t*, 直接 &hFont 会把字体句柄写进 VARIANT.vt
                        // → 出参永远读不回来 (Charts 2020: GdipMeasureString rc=2,
                        // 标题高度 0 → 标题与绘图区重叠)。
                        // VB6 语义: 传临时标量, 调用后回写 Variant。
                        // Fix 133x: ByRef LongPtr 形参时临时变量必须是 intptr_t,
                        // 否则 GdipCreateFont 把 64 位 GDI+ Font 指针写进 int32 → 截断,
                        // 后续 GdipMeasureString 解引用坏指针 → x64 下 0xC0000005
                        // (Charts 2020 ucChartArea/ucPieChart GetTextSize, 标题文字崩溃)。
                        bool isLp133x = (i < calleeParams.size()
                                        && (calleeParams[i].type == Vb6Type::LongPtr
                                            || calleeParams[i].type == Vb6Type::ULong));
                        std::string vref = "_vbvref" + std::to_string(tempCounter_++);
                        if (isLp133x) {
                            c_.addPending("intptr_t " + vref
                                          + " = vb6_VariantToLongPtr(" + argVal + ");");
                            c_.addPostLine("vb6_VariantSetI8(&" + argVal + ", " + vref + ");");
                        } else {
                            c_.addPending("int32_t " + vref
                                          + " = (int32_t)vb6_VariantToDouble(" + argVal + ");");
                            c_.addPostLine("vb6_VariantSetI4(&" + argVal + ", " + vref + ");");
                        }
                        argVal = "&" + vref;
                    } else {
                        argVal = "&" + argVal;
                    }
                } else {
                    // Fix 092z: 左值字段链 (With对象 _vb6_with_N->字段 / me->字段 /
                    // 虚UDT字段) 也是可寻址左值 → 生成 &(字段) 而非复合字面量.
                    // With 前 ByRef Long 字段 (QRCodegenMakeAlphanumeric →
                    // pvAppendBitsToBuffer ..., .Data, .BitLength) 若走复合字面量
                    // 会复制字段值到临时 int32, 被调方 (*lBitLen)++ 只改临时值 →
                    // 字段永不增长 → 位写入偏移全零 (HELLO WORLD 无法解码).
                    // Fix 164x: 同样先剥离 "  /* udt objfield */" 注释后缀再判左值.
                    std::string argValLv092z = argVal;
                    size_t lvCmPos092z = argValLv092z.find("  /* udt objfield ");
                    if (lvCmPos092z != std::string::npos) {
                        argValLv092z = argValLv092z.substr(0, lvCmPos092z);
                    }
                    bool isLvChain092z = !argValLv092z.empty() && (std::isalpha(static_cast<unsigned char>(argValLv092z[0])) || argValLv092z[0] == '_');
                    if (!isLvChain092z && argValLv092z.size() > 2 && argValLv092z[0] == '(' && argValLv092z[1] == '*') {
                        // Fix 090k: (*ptr)->field / (*ptr).field 解引用字段链也是左值
                        int depth092z = 0;
                        for (size_t ci092z = 0; ci092z < argValLv092z.size(); ci092z++) {
                            char c092z = argValLv092z[ci092z];
                            if (c092z == '(') depth092z++;
                            else if (c092z == ')') {
                                depth092z--;
                                if (depth092z == 0) {
                                    if (ci092z + 1 < argValLv092z.size() &&
                                        (argValLv092z[ci092z + 1] == '.' ||
                                         (argValLv092z[ci092z + 1] == '-' && ci092z + 2 < argValLv092z.size()
                                          && argValLv092z[ci092z + 2] == '>'))) {
                                        isLvChain092z = true;
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    // Fix 113f: VB6_SA_AT(TYPE, arr, idx).field / ->field — UDT 数组
                    // 元素字段访问也是**可寻址左值** (VB6_SA_AT 是 RTL 中唯一产出左值
                    // 的函数式宏, 展开为 ((T*)arr->data)[idx - arr->lBound] 元素).
                    // 必须生成 &(expr) 而非复合字面量**副本**: 否则被调方对 ByRef 数组
                    // 形参 ReDim/写回只改临时副本, 调用方字段永不更新.
                    // Charts 2020 ucTreeMaps.Draw: Squarified m_Serie(i).Values,
                    //   m_Serie(i).Rects, ... 内部 ReDim Rects(...) 写 (*Rects) →
                    //   原走 (&(vb6_SafeArray1D**){...Rects}) 临时副本 →
                    //   me->m_Serie[i].Rects 仍为 NULL → 随后绘制循环
                    //   VB6_SA_AT(vb6_type_RectF, _vb6_with_N->Rects, j) 空指针崩溃
                    //   (0xC0000005 @ ucTreeMaps_Draw, edx=0).
                    // 括号配对后须紧跟 '.' 或 '->' 才算字段链; 且**必须跳过**下方的
                    // 简单字符扫描 — 该扫描不认括号, 会把本形态又判回非左值 (曾因此
                    // 使本 Fix 首次落地无效, 生成代码未变而困惑).
                    bool isSaAtField113f = false;
                    if (argVal.rfind("VB6_SA_AT(", 0) == 0) {
                        int depth113f = 1;
                        for (size_t ci113f = 10; ci113f < argVal.size(); ci113f++) {
                            char c113f = argVal[ci113f];
                            if (c113f == '(') depth113f++;
                            else if (c113f == ')') {
                                depth113f--;
                                if (depth113f == 0) {
                                    if (ci113f + 1 < argVal.size()
                                        && (argVal[ci113f + 1] == '.'
                                            || (argVal[ci113f + 1] == '-'
                                                && ci113f + 2 < argVal.size()
                                                && argVal[ci113f + 2] == '>'))) {
                                        isSaAtField113f = true;
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    if (isSaAtField113f) isLvChain092z = true;
                    if (isLvChain092z && !isSaAtField113f) {
                        for (size_t ci092z = 0; ci092z < argVal.size(); ci092z++) {
                            char c092z = argVal[ci092z];
                            if (std::isalnum(static_cast<unsigned char>(c092z)) || c092z == '_' || c092z == '.') continue;
                            if (c092z == '-' && ci092z + 1 < argVal.size() && argVal[ci092z + 1] == '>') { ci092z++; continue; }
                            isLvChain092z = false;
                            break;
                        }
                    }
                    if (isLvChain092z) {
                        argVal = "&(" + argVal + ")";
                    } else {
                    // 字面量或复杂表达式: 使用C11复合字面量
                    // &(int32_t){10} 或 &(double){3.14}
                    std::string cType = "int32_t";
                    // Fix <vbeclipse> rev7: 命中 ByRef 工程类槽 = 本实参已按"类指针的地址"
                    // 交付完毕 (下面已落临时并 &temp), 后续所有装箱/复合字面量加工都必须跳过。
                    bool isClassSlot092r = false;
                    if (i < calleeParams.size()) {
                        cType = mapType(calleeParams[i].type);
                        // Fix 092e: ByRef 数组形参 (如 cToolsUtf8.Decode(ByRef
                        // Utf() As Byte)) 的 C 形参是 vb6_SafeArray1D** (Fix 078
                        // rev2), 复合字面量类型必须一致 — mapType(Byte|Array)
                        // 只给出 "uint8_t*" → (&(uint8_t*){Variant}) C2440
                        // "无法从 vb6_VARIANT 转换为 uint8_t *"
                        // (cHttpClient 418: ToolsUtf8.Decode(Inst.ResponseBody)).
                        bool isArrP092e = (static_cast<uint16_t>(calleeParams[i].type)
                                           & static_cast<uint16_t>(Vb6Type::Array)) != 0;
                        if (isArrP092e && !calleeParams[i].isByVal
                            && !calleeParams[i].isParamArray) {
                            cType = "vb6_SafeArray1D**";
                        }
                        // Fix <vbeclipse> rev7: **ByRef 工程类形参 + 非左值实参** ——
                        // 092e 数组守卫的对称 counterpart, 判据同一个口径
                        // (cParamClassPtrType = "这个槽收不收类实例指针")。
                        //
                        // 为什么必须在这里再判一次: 上游 787 那个类槽分支只在
                        // **isSimpleIdent**(实参是裸标识符) 时才有机会命中; 实参一旦是
                        // 函数调用/属性读, 就落到本复合字面量路径, 而 mapType(Variant)
                        // 给出的 "vb6_VARIANT" 会把它装成
                        //   (&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=vb6_VariantToString(...)})
                        // 送进本该收 vb6_cls_View** 的槽 → 被调方把 vb6_VARIANT* 当
                        // 类指针解引用 → 堆损坏 (实测 vbeclipse ucPerspective.CreateFolder
                        // → ucFolder.AddView: 触发 vb6_heapCorruptVEH, 崩点漂到
                        // vb6_List_Item+0x23A, av read target=0x3c4000)。
                        //
                        // 口径对称性就是本条的判据纪律: 一个"ByRef + 类槽"的形参,
                        // 无论实参是标识符还是表达式, 都不该走进 M22 装箱 switch。
                        const std::string clsSlotP092r =
                            (i < calleeParams.size() && !calleeParams[i].isByVal
                             && !calleeParams[i].isParamArray)
                                ? cParamClassPtrType(calleeParams[i]) : std::string();
                        if (!clsSlotP092r.empty()) {
                            // ByRef 类形参 = 传**类指针的地址**。实参是右值(函数调用等),
                            // 不能直接 & —— 必须先落一个临时再取址, 否则 &(右值) C2102。
                            // 落临时这一层与 Fix 161d 的 ByRef(As Any/UDT) 复杂右值同源。
                            const std::string tmpCls092r = classSlotTempName(clsSlotP092r, tempCounter_++);
                            c_.addPending(clsSlotP092r + " " + tmpCls092r + " = "
                                          + deliverToClassSlot(clsSlotP092r, argVal) + ";");
                            argVal = "&" + tmpCls092r;
                            // ⚠ argVal 就位后必须让**下面整段** (Fix 120 → M22/090d → 收尾
                            // `&(cType){argVal}`) 全部跳过, 否则它们会再包一层:
                            //   (&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=&vb6_argtmp_...})
                            // 或 (&(vb6_cls_View*){&vb6_argtmp_...}) —— 类指针的地址被当
                            // BSTR 塞进 Variant, 或者把"地址"再取址一次。实测第一版只改
                            // argVal 时就是 (&(vb6_VARIANT){...bstrVal=&vb6_argtmp_...})。
                            //
                            // 做法 = 把这一整段包进 if (!isClassSlot092r) —— 与数组槽 092e
                            // 的处理方式同构 (它靠 cType 改成 vb6_SafeArray1D** 让所有
                            // 分支自然落空; 这里类槽类型不含 "**" 后缀, 与任何既有
                            // cType 都不重合, 所以只能显式短路)。
                            isClassSlot092r = true;
                            cType = clsSlotP092r;   // 供下方 090d 的 void* 判定等留个一致的读数
                        }
                    }
                    // Fix 120: ByRef 具体数值形参 ← 裸 COM 结果 (void* 承载 COM VARIANT*)
                    // 不能把指针直接截断成数值 —— 取到的是指针低位, 于是颜色/数值全成噪声:
                    //   Charts 2020 ucChartArea.AddLineSeries(..., ByRef SerieColor As OLE_COLOR)
                    //   调用处 cPalette(j) 是 vb6_ComCall(...) → 生成
                    //   &(int32_t){vb6_ComCall(...)} → 面积图系列变成洋红/紫随机色.
                    // 先解出 Variant 的值再放进复合字面量 (Fix 029 只覆盖 ByVal 路径).
                    {
                        bool comRes120 = (argVal.compare(0, 12, "vb6_ComCall(") == 0)
                            || (argVal.compare(0, 15, "vb6_ComGetProp(") == 0)
                            || (argVal.compare(0, 21, "vb6_ComGetObjectProp(") == 0);
                        if (comRes120) {
                            std::string ex120 = "vb6_VariantFromComResult(" + argVal + ")";
                            if (cType == "int32_t")
                                argVal = "(int32_t)vb6_VariantToDouble(" + ex120 + ")";
                            else if (cType == "int16_t")
                                argVal = "(int16_t)vb6_VariantToDouble(" + ex120 + ")";
                            else if (cType == "double")
                                argVal = "vb6_VariantToDouble(" + ex120 + ")";
                            else if (cType == "float")
                                argVal = "(float)vb6_VariantToDouble(" + ex120 + ")";
                        }
                    }
                    // M22: VARIANT类型需要指定字段初始化 (.vt=VT_xxx, .field=value)
                    // Fix <vbeclipse> rev7: `!isClassSlot092r` 是显式护栏 —— 理论上命中类槽时
                    // cType 已被改成 vb6_cls_X*, 本 if 自然不成立; 但这里是"绝不能装箱"的
                    // 语义关口, 不该只依赖 cType 的取值巧合 (类槽类型与任何既有 cType 都不
                    // 重合, 可一旦将来新增某档 cType 恰好撞上, 就会静默回归成装 Variant)。
                    if (cType == "vb6_VARIANT" && !isClassSlot092r) {
                        // 推断实参的VB6类型来决定VARIANT字段
                        Vb6Type argVbType = inferExprType(*node.positional[i]);
                        switch (argVbType) {
                            case Vb6Type::String:
                                // Fix 027b: 外层加括号让预处理器把 {a,b} 内的逗号视为同一参数
                                // (避免宏调用时 #define RtlCopyMemory(D,s,l) memcpy(D,s,l) 把复合字面量
                                //  内部的逗号错算成宏实参分隔符 → C4002 参数过多)
                                argVal = "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=" + argVal + "})";
                                break;
                            case Vb6Type::Long:
                            case Vb6Type::Integer:
                                argVal = "(&(vb6_VARIANT){.vt=VT_I4, .lVal=(int32_t)(" + argVal + ")})";
                                break;
                            case Vb6Type::Double:
                            case Vb6Type::Single:
                                argVal = "(&(vb6_VARIANT){.vt=VT_R8, .dblVal=(double)(" + argVal + ")})";
                                break;
                            case Vb6Type::Boolean:
                                argVal = "(&(vb6_VARIANT){.vt=VT_BOOL, .boolVal=(int16_t)(" + argVal + ")})";
                                break;
                            case Vb6Type::Byte:
                                argVal = "(&(vb6_VARIANT){.vt=VT_UI1, .bVal=(uint8_t)(" + argVal + ")})";
                                break;
                            case Vb6Type::Date:
                                argVal = "(&(vb6_VARIANT){.vt=VT_DATE, .dblVal=(double)(" + argVal + ")})";
                                break;
                            default:
                                // Fix 054: Variant 或未知类型
                                // 1. 实参是 UDT 的 Variant 字段 (With 成员如 .SourceFile,
                                //    .FileName 等 As Variant) → 值已是 vb6_VARIANT,
                                //    ByRef Variant 形参直接取地址 &(...), 不要包装成
                                //    {.vt=VT_BSTR, .bstrVal=...} (Variant→BSTR C2440).
                                // 2. C 表达式已是 Variant (vb6_VariantArrayGet 等),
                                //    提取 BSTR 值到 .bstrVal (避免 VARIANT→BSTR C2440)
                                // 3. 其他未知类型保持原样 (UDT 等单独处理)
                                if (i < node.positional.size()
                                    && inferUdtFieldVb6Type(node.positional[i].get()) == Vb6Type::Variant) {
                                    argVal = "&(" + argVal + ")";
                                } else if (cExprIsVariant(argVal)) {
                                    argVal = "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=vb6_VariantToString(" + argVal + ")})";
                                } else if (argVal.compare(0, 12, "vb6_ComCall(") == 0
                                           || argVal.compare(0, 15, "vb6_ComGetProp(") == 0
                                           || argVal.compare(0, 21, "vb6_ComGetObjectProp(") == 0) {
                                    // Fix 113l: **裸 COM 结果** (void*, 承载 COM
                                    // VARIANT*) 填入 VT_BSTR 槽时必须先
                                    // vb6_VariantFromComResult 再 vb6_VariantToString;
                                    // 原样把 void* 塞进 .bstrVal 会让 GDI+ 把垃圾当
                                    // 字符串 → 渲染成一排"豆腐块"空框.
                                    //   Charts 2020 ucTreeMaps 项名标签:
                                    //     DrawText(..., Labels.Item(j), ...) (text As Variant)
                                    //   → 原生成 (&(vb6_VARIANT){.vt=VT_BSTR,
                                    //       .bstrVal=vb6_ComCall(...Labels..., L"Item", ...)}).
                                    argVal = "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal="
                                           + std::string("vb6_VariantToString(vb6_VariantFromComResult(")
                                           + argVal + "))})";
                                } else {
                                    argVal = "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=" + argVal + "})";
                                }
                                break;
                        }
                    } else {
                        // Fix 090d: ByRef 具体类型形参 + 实参是 vb6_VARIANT 值 —
                        // 复合字面量初始化 (&(double){vb6_VARIANT}) 触发 C2440
                        // (vb6_VARIANT→double/BSTR/... 初始化失败). 按 cType 先提取.
                        // 例: cHttpServerCookies Encode: FormatHttpDate(CK.Expires)
                        //     (形参 ByRef Date, Expires As Variant 属性 →
                        //     vb6_cHttpServerCookieAttr_prop_get_Expires 返回 vb6_VARIANT);
                        //     cHttpServerResponse File: ComputeETag(fileInfo.DateLastModified,
                        //     ...) (COM 属性 → vb6_VariantFromComResult(...) vb6_VARIANT).
                        bool argIsVariant090d = isDefinitelyVariantExpr(*node.positional[i]);
                        if (!argIsVariant090d) argIsVariant090d = cExprIsVariant(argVal);
                        // Fix <vbeclipse> rev7: 类槽已交付 (&temp) —— 整段 Variant 提取
                        // 跳过。cType 此时是 vb6_cls_X* / vb6_ivref_X*, 上面的 case 都不
                        // 匹配, 但留显式护栏: 将来若给类指针加 case 也不会静默吃掉它。
                        if (argIsVariant090d && !isClassSlot092r) {
                            if (cType == "double" || cType == "float") {
                                argVal = "vb6_VariantToDouble(" + argVal + ")";
                            } else if (cType == "BSTR") {
                                argVal = "vb6_VariantToString(" + argVal + ")";
                            } else if (cType == "int32_t" || cType == "int16_t"
                                       || cType == "uint8_t" || cType == "int64_t"
                                       || cType == "LONG") {
                                argVal = "vb6_VariantToLong(" + argVal + ")";
                            } else if (cType == "vb6_SafeArray1D*"
                                       || cType == "vb6_SafeArray1D**") {
                                argVal = "vb6_VariantToSafeArray1D(" + argVal + ")";
                            } else if (cType == "void*") {
                                argVal = "vb6_VariantToObjectVal(" + argVal + ")";
                            }
                        }
                        // Fix 027b: 同样加外层括号防止预处理器把复合字面量 `{a, b}` 内的逗号算成宏实参分隔符.
                        // Fix <vbeclipse> rev7: 类槽已交付 (&temp, 见上) —— 绝不能再包一层
                        // (&(vb6_cls_View*){&temp}) : 那是"把地址的地址"当类指针传。
                        if (!isClassSlot092r) {
                            argVal = "(&(" + cType + "){" + argVal + "})";
                        }
                    }
                    }
                }
                } // end Fix 072b else (As Any ByVal override)
                    } // end Fix 064 else
            }
        } // Fix 024 P2: ByVal Variant 参数 — 调用点用 vb6_VariantFromValue 包装实参.