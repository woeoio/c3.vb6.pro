// cgen_builtin_return_type_table.inc - 「内置函数返回档」的后端兜底表
//
// 为什么要有这张表:
//   语义层 src/semantics/builtin/builtin_funcs.inc 登记了 160 枚内置函数, 但 cgen
//   的发射表 (cgen_expr_ident_builtin_table.inc) 里还有 16 枚**没有符号** —— 它们能
//   发码 (LNK 能找到 RTL 实现), 却因为 symTab_.lookup() 查不到符号, 让
//   inferExprType 落到 default 答 Variant。Variant 是个**兜底答案**, 不是断言:
//   消费点 (Print# / Write# / 实参包装 / 比较 / 收窄检查) 会据此以为那是整只
//   vb6_VARIANT 槽, 于是把 int32_t / BSTR 塞进去 —— C2440 或静默错值。
//
// 口径 (每一行都对着 RTL 声明抄, 不是猜):
//   * 返回档 = RTL 里那枚函数的**返回值声明**, 与 Vb6Type 的语义档一一对应。
//   * 只登记「语义层没有符号」的名字; 有符号的那 160 枚由 symTab_ 回答, 这里不许
//     重复 —— 两张表各答一份同一个问题, 迟早会走散。
//   * 明确返回 vb6_VARIANT 的 (LoadRes* 三枚) 就答 Variant; 这不是兜底, 是事实
//     (vb6rtl_runtime.h:78-80 的签名)。它们在 cExprIsVariant 的前缀表里也有登记。
//   * 不登记多态内置 (IIf/Choose/Switch/If) —— 它们的 C 返回档由发码现场决定
//     (vb6_IIfBSTR/Long/Double/Variant 或裸三目), 固定成任何一档都会答错。
//
// 维护: 语义层补齐这些名字之后, 对应行应当**删掉**, 让 symTab_ 成为唯一来源。
// 当前状态与移除前提见 docs/print_write_numeric.md。

static Vb6Type builtinReturnTypeFallback(const std::string& name) {
    // 调用点传进来的名字可能带 $ (parser 在使用点不一定剥掉), 与
    // cgen_expr_ident_builtin_table.inc 的 lookupName 同口径: 去掉尾缀再查。
    std::string n = name;
    std::transform(n.begin(), n.end(), n.begin(), ::tolower);
    if (!n.empty() && n.back() == '$') n.pop_back();

    static const std::unordered_map<std::string, Vb6Type> table = {
        // cbyt: CByte 的兼容别名 (cgen 发射表里与 cbyte 双键, 语义层只登记了 CByte)。
        {"cbyt",            Vb6Type::Byte},
        // Kill / Randomize 在 VB6 里是语句; 这里给的是 RTL 返回值档
        // (vb6_Kill 返回 int32_t; vb6_Randomize 返回 void)。它们不会出现在
        // Print# 的数值位, 但答错成 Variant 会让实参包装走错路。
        {"kill",            Vb6Type::Long},
        {"randomize",       Vb6Type::Void},
        // 字符串族: RTL 声明 BSTR (vb6rtl_builtin.h:44-47)。
        {"input",           Vb6Type::String},   // Input$/Input 语句值 (vb6_InputString)
        {"inputb",          Vb6Type::String},
        {"leftb",           Vb6Type::String},
        {"rightb",          Vb6Type::String},
        {"midb",            Vb6Type::String},
        {"loadresstring",   Vb6Type::Variant},  // vb6_VARIANT vb6_LoadResString(...)
        // 数值族。
        {"instrb",          Vb6Type::Long},      // int32_t vb6_InStrB(...)
        {"lenb",            Vb6Type::Long},      // vb6_LenB 宏每一档都交 int32_t
        {"typeof",          Vb6Type::Boolean},   // int32_t vb6_TypeOf(...); VB 语义是 Boolean
        // LoadResData / LoadResPicture 返回整只槽 (vb6rtl_runtime.h:78-79)。
        {"loadresdata",     Vb6Type::Variant},
        {"loadrespicture",  Vb6Type::Variant},
        // Command$ 与 Command 同一枚 RTL 函数; 语义层已登记 command, 这里留着只为
        // 名字形状完整 (若将来语义层撤了 command, 这一行仍能答对)。
        {"command",         Vb6Type::String},
    };
    auto it = table.find(n);
    if (it == table.end()) return Vb6Type::Unknown;
    return it->second;
}
