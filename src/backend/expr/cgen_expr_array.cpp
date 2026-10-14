#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>
#include <cstdio>

namespace vb6c3 {

// --- cgen_expr_array.cpp: SafeArray / 数组元素类型辅助方法 ---

std::string CCodeGen::mapSaElemType(Vb6Type type) const {
    switch (type) {
        case Vb6Type::Boolean:  return "vb6_sa_bool";
        case Vb6Type::Byte:     return "vb6_sa_byte";
        case Vb6Type::Integer:  return "vb6_sa_int";
        case Vb6Type::Long:     return "vb6_sa_long";
        case Vb6Type::Single:   return "vb6_sa_single";
        case Vb6Type::Double:   return "vb6_sa_double";
        case Vb6Type::Date:     return "vb6_sa_double";
        case Vb6Type::Currency: return "vb6_sa_currency";  // P1-9修复: 8-byte Currency
        case Vb6Type::String:   return "vb6_sa_bstr";
        case Vb6Type::Variant:  return "vb6_sa_variant";
        case Vb6Type::Object:   return "vb6_sa_ptr";
        // Fix 191: LongPtr/LongLong 是指针宽度整数, 不能落到 default 的 Variant 载体。
        // VTableHandle.bas 的手写 vtable 子类化要求 `VTableIPAO(0 To 9) As LongPtr`
        // 是紧密排布的 8 字节槽 —— 它把 VarPtr(VTableIPAO(0)) 直接当 COM vtable 用。
        // 按 Variant 分配后步长变 16 字节、槽内容是 vb6_VARIANT 头部, 伪 vtable 的
        // 第 5 槽读出 3 (VT_I4), 于是 (3 + 0x28) 读 → 0xC0000005 (左键点网格即崩)。
        case Vb6Type::LongPtr:  return "vb6_sa_ptr";
        default:                return "vb6_sa_variant";
    }
}

std::string CCodeGen::mapSaElemCType(Vb6Type type) const {
    switch (type) {
        case Vb6Type::Boolean:  return "int16_t";
        case Vb6Type::Byte:     return "uint8_t";
        case Vb6Type::Integer:  return "int16_t";
        case Vb6Type::Long:     return "int32_t";
        case Vb6Type::Single:   return "float";
        case Vb6Type::Double:   return "double";
        case Vb6Type::Date:     return "double";
        case Vb6Type::String:   return "BSTR";
        case Vb6Type::Variant:  return "vb6_VARIANT";
        case Vb6Type::Currency: return "int64_t";
        case Vb6Type::Object:   return "void*";
        case Vb6Type::LongPtr:  return "intptr_t";  // Fix 191: 见 mapSaElemType 同处注释
        default:                return "vb6_VARIANT";
    }
}

Vb6Type CCodeGen::resolveArrayElemType(ASTNode* typeRef) const {
    if (!typeRef) return Vb6Type::Variant;
    if (typeRef->kind == ASTNodeKind::ArrayTypeRef) {
        auto& arrType = static_cast<ArrayTypeRef&>(*typeRef);
        return resolveArrayElemType(arrType.elementType.get());
    }
    if (typeRef->kind == ASTNodeKind::SimpleTypeRef) {
        auto& simple = static_cast<SimpleTypeRef&>(*typeRef);
        Vb6Type t = typeSys_.resolveTypeName(simple.name);
        if (t != Vb6Type::Unknown) return t;
        // Fix 090by: VB6 内置对象类型 (VBA 标准对象, 无 typelib 符号) —
        // Collection/Forms/ErrObject/App/Screen/Printer/Clipboard.
        // 与 semantic_analyzer.cpp vb6BuiltinObjTypes 一致, 处理为 Object.
        // 若不识别, resolveTypeName/lookup 均 miss → 回落 Variant → 类字段
        // (Private x As Collection) 被误注册 classVariantMembers_ → 类内
        // m_Col(i) 生成 vb6_VariantArrayGet(&me->m_Col, i) (C2172 实参不是
        // 指针), 而非 void* COM 字段的 vb6_ComCall(me->m_Col, L"Item", ...).
        {
            std::string nm0 = simple.name;
            std::string nmLower0;
            nmLower0.resize(nm0.size());
            std::transform(nm0.begin(), nm0.end(), nmLower0.begin(), ::tolower);
            if (nmLower0.size() > 4 && nmLower0.compare(0, 4, "vba.") == 0) {
                nmLower0 = nmLower0.substr(4);
            }
            static const std::unordered_set<std::string> vb6BuiltinObjTypes = {
                "collection", "forms", "errobject", "app", "screen", "printer", "clipboard",
                "control", "form", "console"  // Fix 161: Console (twinBASIC 口径)
            };
            if (vb6BuiltinObjTypes.count(nmLower0)) return Vb6Type::Object;
        }
        // Fix 049b: 项目内 UDT/Enum/类名 需查符号表 (typeSys_ 只含内置类型)。
        // 否则 Dim x() As 某UDT (且该 UDT 声明位于 Dim 之后, Fix 049 预扫描已注册)
        // 的元素类型退回 Variant → arrayElemTypes_ 记录 Variant → 元素访问生成
        // VB6_SA_AT(vb6_VARIANT, ...).字段 (C2039/C2223) → With 对象类型也变 Variant,
        // 成员解析退化为跨模块类查找 (如 .Pos 误解析到 cToast.Pos, C2198)。
        // 与 resolveArrayUdtElemCType (Fix 055b) 的符号表回退保持一致。
        // Fix 107: 用 lookupTypeSymbol — 类型与过程同名时 (Fix 103 把类型存到
        // <name>$ty), 通用 lookup 命中过程符号 → 枚举/UDT 字段回落 Variant,
        // 与 mapTypeRef 的 int32_t/vb6_type_x 不一致 → C2440.
        if (auto* sym = lookupTypeSymbol(simple.name)) {
            if (sym->kind == SymbolKind::UserDefinedType) return Vb6Type::UserDefinedType;
            if (sym->kind == SymbolKind::EnumType) return Vb6Type::Long;
            if (sym->kind == SymbolKind::Class || sym->kind == SymbolKind::ComClass ||
                sym->kind == SymbolKind::ComInterface || sym->kind == SymbolKind::ComModule ||
                sym->kind == SymbolKind::ComGlobalNs) {
                return Vb6Type::Object;
            }
        }
        // Fix 097: 无 vb 前缀的 VBA 枚举别名 — CompareMethod 是 Long
        // (与 type_system.cpp resolveTypeName 的 Fix 097 对齐, 否则字段声明
        // 映射 int32_t 而此处回落 Variant, 类工厂初始化生成
        // me->m_CompareMode = vb6_VariantEmpty() → C2440, Dictionary.cls:45).
        {
            std::string nm097 = simple.name;
            std::transform(nm097.begin(), nm097.end(), nm097.begin(), ::tolower);
            if (nm097 == "comparemethod") return Vb6Type::Long;
        }
        return Vb6Type::Variant;
    }
    return Vb6Type::Variant;
}

// 账 #305 §B154 (第四十二刀): ReDim 的元素档只有一个答案来源 —— 声明。
// 旧写法问的是 `resolveArrayElemType(node.asType.get())`, 而 ReDim 不带 As 子句时
// asType 为空, 于是**标量/字符串/对象**那一族一律回落 vb6_sa_variant, 只有 UDT 有回退
// (:109 查 arrayUdtElemTypes_)。可这张描述符的**访问侧**问的是另一张表 (arrayElemTypes_,
// 由六条声明路填、VB6_SA_AT / VB6_SA_ND_AT2 吃的就是它), 固定数组那条路也按声明给档
// (`Dim f(5) As String` 发 vb6_SafeArrayCreate1D(vb6_sa_bstr,…), 实测)。⇒ 同一个决定
// 两处问、动态那条漏问, 后果是造出来 24 字节一步而写进去 8 字节一步: 直接读写自己一致
// (所以看起来是好的), 而凡按描述符自己的档办事的消费方 (Destroy / Erase / 装箱元素读 /
// ArrayAssign) 全走错步长 —— 实测 `Dim a() As String` + `ReDim a(1 To 3,1 To 4)` +
// `a(2,3)="X"` 之后 `v = a` 再读 v(2,3) 交**空串**, VarType 交 8204 而不是声明那一档。
// 收成一处 = 这一枚出口: As 子句在场时**仍然优先** (VB6 语义, 且语料实测 declared==As
// 占 1580/1588, 两处相异的形状今天没有实物), 缺席时问声明表, 表里没有才回落 Variant
// (隐式 Variant 数组 / 成员链 / With 成员 —— 沿用 Fix 180 的「不去猜」纪律)。
Vb6Type CCodeGen::resolveReDimElemType(const std::string& lowerVar, ASTNode* typeRef) const {
    if (typeRef) return resolveArrayElemType(typeRef);
    if (!lowerVar.empty()) {
        auto it = arrayElemTypes_.find(lowerVar);
        if (it != arrayElemTypes_.end()) return it->second;
    }
    return Vb6Type::Variant;
}

std::string CCodeGen::resolveArrayUdtElemCType(ASTNode* typeRef) const {
    if (!typeRef) return "";
    if (typeRef->kind == ASTNodeKind::ArrayTypeRef) {
        auto& arrType = static_cast<ArrayTypeRef&>(*typeRef);
        return resolveArrayUdtElemCType(arrType.elementType.get());
    }
    if (typeRef->kind == ASTNodeKind::SimpleTypeRef) {
        auto& simple = static_cast<SimpleTypeRef&>(*typeRef);
        // 检查类型系统
        Vb6Type vtype = typeSys_.resolveTypeName(simple.name);
        if (vtype == Vb6Type::UserDefinedType) {
            return "vb6_type_" + cIdent(simple.name);
        }
        // Fix 055b: resolveTypeName 对项目内 UDT 返回 Unknown,
        // 需要在符号表中查找是否为 UserDefinedType 符号
        auto* sym = symTab_.lookup(simple.name);
        if (sym && (sym->type == Vb6Type::UserDefinedType || sym->kind == SymbolKind::UserDefinedType)) {
            return "vb6_type_" + cIdent(simple.name);
        }
        sym = symTab_.lookupModule(simple.name);
        if (sym && (sym->type == Vb6Type::UserDefinedType || sym->kind == SymbolKind::UserDefinedType)) {
            return "vb6_type_" + cIdent(simple.name);
        }
    }
    return "";
}

// Fix 192: 数组元素为**项目类**时返回该类名 (否则空串)。
// 与 resolveArrayUdtElemCType 对称: 后者只认 UDT, 而 `As ShapeAct` 这类元素类型
// 在 resolveArrayElemType 里被压成 Vb6Type::Object, 类名就此丢失 —— arr(i).Move
// 的接收者推断不出类, 落到通用成员访问 → VB6_SA_AT(void*, arr, i).Move(...) →
// MSVC `C2224: ".Move" 的左侧必须具有结构/联合类型` (void* 不是结构体)。
// 只认**项目内 Class** (和 COM 类): 内置对象 (Collection/Forms 等) 走 COM 晚绑定,
// 不需要也不该进这张表 (它们没有 vb6_<Cls>_<Method> 早绑定符号)。
std::string CCodeGen::resolveArrayClassElemType(ASTNode* typeRef) const {
    if (!typeRef) return "";
    if (typeRef->kind == ASTNodeKind::ArrayTypeRef) {
        auto& arrType = static_cast<ArrayTypeRef&>(*typeRef);
        return resolveArrayClassElemType(arrType.elementType.get());
    }
    if (typeRef->kind != ASTNodeKind::SimpleTypeRef) return "";
    auto& simple = static_cast<const SimpleTypeRef&>(*typeRef);
    if (simple.name.empty()) return "";
    // Fix 107 的教训: 类型与过程可能同名, 必须用 lookupTypeSymbol (查 <name>$ty)
    // 而不是通用 lookup, 否则命中的是过程符号。
    const Symbol* sym = lookupTypeSymbol(simple.name);
    if (!sym) return "";
    // 只认**项目内 Class**。ComClass (typelib 导入的外部 COM 类) 的成员访问走
    // vb6_ComCall 晚绑定, 把它登记进来会让 inferClassTypeOfExpr 返回类名、把调用
    // 引向并不存在的 vb6_<Cls>_<Method> 早绑定符号 —— 按"宁少列不可多列"排除。
    if (sym->kind == SymbolKind::Class) return simple.name;
    return "";
}

} // namespace vb6c3
