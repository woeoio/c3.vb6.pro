#pragma once
// VB6类型系统 - 类型解析/推导/兼容性判断
// VB6类型特点: 不区分大小写, Variant万能类型, 隐式转换

#include "common/types.hpp"
#include "common/diagnostics.hpp"
#include <string>
#include <unordered_map>
#include <optional>

namespace vb6c3 {

class ASTNode;  // 前向声明(TypeRef)

// ============================================================
// 类型系统
// ============================================================

class TypeSystem {
public:
    TypeSystem();

    // 从类型名解析Vb6Type (如"Long"→Vb6Type::Long)
    Vb6Type resolveTypeName(const std::string& name) const;

    // Vb6Type转字符串
    static const char* typeToString(Vb6Type t);

    // 类型是否为数值型
    static bool isNumeric(Vb6Type t);

    // 类型是否为整型 (含Byte)
    static bool isIntegral(Vb6Type t);

    // 类型是否为浮点型
    static bool isFloat(Vb6Type t);

    // 类型是否为字符串型
    static bool isString(Vb6Type t);

    // 类型是否为对象型
    static bool isObject(Vb6Type t);

    // 两个类型是否可以隐式转换
    static bool canImplicitConvert(Vb6Type from, Vb6Type to);

    // 两个类型的运算结果类型 (VB6 widened rules)
    static Vb6Type promote(Vb6Type a, Vb6Type b);

    // 位运算 And/Or/Xor/Eqv/Imp 的结果类型 —— 全仓唯一口径 (账 #216)
    static Vb6Type bitwiseResult(Vb6Type a, Vb6Type b);

    // 一元 Not 的结果类型 —— 全仓唯一口径 (账 #216)
    static Vb6Type logicalNotResult(Vb6Type t);

    // 类型占用的字节数
    static int typeSize(Vb6Type t);

    // 整数字面量的默认类型
    static Vb6Type defaultIntType(int64_t val);

    // 布尔值对应的VB6类型
    static Vb6Type boolType() { return Vb6Type::Boolean; }

    // 字符串字面量类型
    static Vb6Type stringType() { return Vb6Type::String; }

    // Variant类型
    static Vb6Type variantType() { return Vb6Type::Variant; }

    // 无返回值类型 (Sub)
    static Vb6Type voidType() { return Vb6Type::Void; }

    // 未知类型
    static Vb6Type unknownType() { return Vb6Type::Unknown; }

private:
    // 类型名到Vb6Type的映射 (小写)
    std::unordered_map<std::string, Vb6Type> builtinTypes_;
};

} // namespace vb6c3
