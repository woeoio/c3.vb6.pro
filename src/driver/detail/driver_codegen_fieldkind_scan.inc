// driver_codegen_fieldkind_scan.inc - runCodeGeneration 函数体片段：
// 账 #305 §B168 (第五十九刀) 建「类 → vb6_VARIANT 字段」表, §B169 (第六十刀) 扩成
// 「类 → 标量数据字段 → 声明档」。在函数体内被 #include, 非独立编译单元。
// 位置: typedfield_scan 之后 —— 与 classVoidFieldMap / classTypedFieldMap 同一趟预扫描,
// 同一批模块清单。

    // 账 #305 §B168/§B169: className(小写) → (字段名(小写) → 该字段的**声明档** Vb6Type)。
    // 字段名两种拼法都记 (原名小写 + "m_" 前缀小写), 与 classVoidFieldMap_ 同格。
    // 表里**只放标量数据字段**; 数组、定长串、对象名 (项目类/COM/未知)、UDT 一律不入表 ——
    // 那些各有自己的表 (classVoidFieldMap_ / classTypedFieldMap_) 或另格账。
    //
    // 为什么要有它 (§B168 那一问): 赋值站判「这个左值到底是不是 Variant 容器」的四条既有
    // 判据 —— 裸名 knownVariantVars_ / 本模块 classVariantMembers_ / 文本 `VB6_SA_*(vb6_VARIANT,` /
    // UDT 字段类型 —— 一条都认不出 `<别的对象>.<公共 Variant 字段>`, 于是整句**完全不装箱**
    // 裸发 (实测 use3.bas + holder2.cls):
    //   h.PV = 7    → `h->PV = 7;`                      → cl C2440 "无法从 int32_t 转换为 vb6_VARIANT"
    //   h.PV = "s"  → `h->PV = vb6_BSTR_FromStr(L"s");` → 同形 C2440
    // 类模块**内**的 `Me.PV = x` 走的是同一条路 (发 `me->PV = (*x);`, 既不装箱也不回收
    // 上一份带权值)。C2440 是硬错误 ⇒ 整棵工程编不过, 语料实测 94 个写点 / 28 份文件。
    //
    // 为什么要扩成整档 (§B169 那两问): 同一个「这枚字段是什么档」在另外两站还各答过一遍,
    // 而两份答案都只认得自己那一半:
    //   形① 直接路 (`h.PS = s` / `Me.PS = v`) 的 BSTR 档判据 (knownBstrVars_ / UDT 字段类型)
    //        同样够不着**别类**的 String 字段 ⇒ 裸发 `h->PS = s;` —— 存进去的是别人那只 BSTR
    //        的地址 (别名), 旧值既不释放, 而类析构 `vb6_BSTR_Free(me->PS)` 会把它再 free 一遍
    //        (实测 MSNMacro CDLG.cls 的 `Me.SelectedPrinterName = myP.DeviceName`);
    //   形② With 块里 `.PS = s` 走的是属性写那一站, 它在「查不到 Property Let/Set」之后**凭空
    //        拼一个名字** `vb6_<Cls>_prop_let_<F>` (实测 DllPrint 工程发 281 句这种调用) ——
    //        那函数在任何地方都不存在 ⇒ cl C4013 + LNK2019, 整棵工程连不上。
    //        这一站的 `isDataField090c` 只查 classTypedFieldMap_ (对象字段), 所以**所有标量
    //        字段**都被它当成"不是数据字段"。
    // 一张表答三档 (Variant / String / 其它标量), 三站都问它 —— 见 classFieldKind()。
    //
    // 为什么这张表必须是**工程级**: 字段的 C 类型由**那个类自己的声明**决定, 与谁在写它
    // 无关; 后端每模块一张 SymbolTable, 既看不见别类的字段, 也会被同名成员遮蔽
    // (说明见 driver_codegen_voidfield_scan.inc 的 projectClassNames 注释)。
    // 键取**小写**类名 (比 void/typed 两张表严一点): 调用侧交来的类名来自
    // variableTypeName / knownClassVars_, 拼法与模块名的大小写不保证一致, 而字段名
    // 本来就是小写存的 —— 这张表答"有没有", 不该由拼法决定。
    //
    // 判据 = "类模块发 struct 时这条字段写的是什么 C 类型", 与 mapTypeRef 同源
    // (isVoidFieldType 用的也是这份内置类型表):
    //   无 As → VB6 默认 Variant (Fix 089j 口径), C 字段是 vb6_VARIANT;
    //   SimpleTypeRef 且 resolveTypeName == Variant → vb6_VARIANT;
    //   resolveTypeName == String 且不是 FixedStringTypeRef → BSTR;
    //   resolveTypeName 是内置标量档 (Long/Integer/Boolean/Single/Double/Currency/Date/
    //     Byte/Decimal/LongPtr/LongLong/Error) → 该档原样入表;
    //   resolveTypeName 交回 Unknown 的名字 (工程枚举 `Public Enum typeField` 就是这一档,
    //     struct 发 int32_t) → 只有符号表指着 SymbolKind::EnumType 才入表, 记 **Long**
    //     (与 Fix 050/098 给 COM 枚举、XxxConstants 定的底层档同一个答案);
    //   其余一律不入表: 数组 (Public PA(2) As Variant 的 C 字段是 vb6_SafeArray1D*,
    //     元素赋值是另一格账) / 定长串 (As String * N 是 FixedStringTypeRef, 收口宽度
    //     另格账 §B118) / 项目类·COM·未知对象名 (void* 或类指针, Fix 090bx 那批) /
    //     UDT (vb6_type_x) / 查不到符号的名字。
    std::unordered_map<std::string, std::unordered_map<std::string, Vb6Type>> classFieldKindMap;
    // 账 #476 §B190: 同一趟扫描、同一个来源 (VariableDecl), 另运一份**动态数组成员**的声明
    // 语法 (只运 `As <T>` 的 AST + 秩, 不运档位 —— 档位到 C 元素类型的映射在发码层那一枚
    // 出口里做, 与发射类结构体字段同一枚)。上面那张表刻意不收数组字段, 而 `clsreg_` 在无
    // 继承的工程里根本不建表 ⇒ 类字段那一腿以前在发码侧没有答案 (§B190 的 11 处 C2064)。
    ClassMemberArrayMap classMemberArrayMap;
    auto classFieldBucket = [](const VariableDecl& var, const TypeSystem& typeSys,
                               SymbolTable& symTab) -> Vb6Type {
        if (var.isDynamicArray || !var.dimensions.empty()) return Vb6Type::Unknown;
        if (!var.asType) return Vb6Type::Variant;                 // 无 As = 默认 Variant
        // 定长串不是 BSTR 槽 (As String * N 发的是带宽度的另一档), 本表刻意不收
        if (var.asType->kind == ASTNodeKind::FixedStringTypeRef) return Vb6Type::Unknown;
        if (var.asType->kind != ASTNodeKind::SimpleTypeRef) return Vb6Type::Unknown;
        const std::string& typeName =
            static_cast<const SimpleTypeRef*>(var.asType.get())->name;
        const Vb6Type t = typeSys.resolveTypeName(typeName);
        if (t != Vb6Type::Unknown && t != Vb6Type::Object && t != Vb6Type::Variant
            && t != Vb6Type::String && t != Vb6Type::UserDefinedType
            && t != Vb6Type::Empty && t != Vb6Type::Null && t != Vb6Type::Void) {
            return t;                                             // 内置标量档原样交回
        }
        if (t == Vb6Type::Variant) return Vb6Type::Variant;
        if (t == Vb6Type::String) return Vb6Type::String;
        // Unknown / Object 的名字: 只有工程枚举才是标量字段, 对象名与 UDT 都不入表。
        Symbol* typeSym = symTab.lookupModule(typeName);
        if (typeSym && typeSym->kind == SymbolKind::EnumType) return Vb6Type::Long;
        return Vb6Type::Unknown;
    };
    for (size_t i = 0; i < modules_.size(); i++) {
        const auto& module = modules_[i];
        // 类模块和窗体模块有结构体字段; 标准模块无 struct, 不需要扫描。
        if (!module->isClassModule && !module->isFormModule) continue;
        if (module->moduleName.empty()) continue;
        std::unordered_map<std::string, Vb6Type> kindByName;
        for (const auto& decl : module->declarations) {
            if (decl->kind != ASTNodeKind::VariableDecl) continue;
            const auto& var = static_cast<const VariableDecl&>(*decl);
            // 账 #476 §B190: 动态数组成员**另外**记一份语法 (与 kindByName 同一循环、同一来源,
            // 不新开一趟扫描)。只收 `Field() As T` 这一形; 定长数组 (`Field(2)`) 不归 §B190
            // —— 那一形的元素访问走 Pattern A (C 多维数组), 出口也只认动态。
            if (var.isDynamicArray) {
                MemberArraySyntax syn;
                syn.asType = var.asType.get();
                syn.rank = (int)var.dimensions.size();
                const std::string oLower = Symbol::toLower(var.name);
                classMemberArrayMap[Symbol::toLower(module->moduleName)][oLower] = syn;
                classMemberArrayMap[Symbol::toLower(module->moduleName)]["m_" + oLower] = syn;
            }
            const Vb6Type bucket = classFieldBucket(
                var, analyzers_[i]->typeSystem(), analyzers_[i]->symbolTable());
            if (bucket == Vb6Type::Unknown) continue;
            std::string oLower = Symbol::toLower(var.name);
            kindByName[oLower] = bucket;
            kindByName["m_" + oLower] = bucket;
        }
        if (!kindByName.empty()) {
            classFieldKindMap[Symbol::toLower(module->moduleName)] = std::move(kindByName);
        }
    }
