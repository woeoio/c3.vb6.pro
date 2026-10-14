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

    // ---- 账 #230 §B139: 目标架构贯通到语义层 ----
    // 背景: 这一层此前**拿不到目标架构**, 于是所有"宽度随架构"的地方只能拿宿主
    // 的 sizeof(void*) 顶替 —— typeSize(LongPtr) 就是这么答的, 而 C3 是 x64 进程,
    // 编译 --arch x86 目标时会答 8 (应为 4), 属 Fix 081e 留下的既有隐患。
    // C3 一次运行只服务一个目标架构 (--arch 全局开关), 故这里用进程级设置,
    // 由驱动在开局按 options.arch 注入一次; 默认 x64 与既有行为一致。
    static void setTargetPtrBits(int bits) { targetPtrBits_ = (bits == 32 ? 32 : 64); }
    static int  targetPtrBits() { return targetPtrBits_; }
    static bool targetIsX64() { return targetPtrBits_ == 64; }

    // 整型的"位宽 + 符号" (ai/032 rev2)。返回 false 表示这个类型不是"位宽与符号都
    // 确定"的整型 —— Boolean (VB 语义上强制升到 Short) 刻意不收, promote 对它走档位兜底。
    //
    // 账 #230 §B141: **LongPtr 已收进来**, 宽度 = targetPtrBits() (目标架构, 与 VBA 一致)。
    //   此前它因"宽度随架构"被排除、只能走 promote 的 rank 兜底 (而 rank 表里它又
    //   不登记) ⇒ `LongPtr + Long` 恒答 Long, x64 上把 64 位缩回 32 位。
    //   注意**不要**顺手去动 promote 的 rank 表: 那张表里 LongPtr 不登记是有负控
    //   依据的 (登记 6 档后 CStr(p + 1.5) 出 4), 两条路不是一回事 ——
    //   intShape 走的是"两侧都定宽"的宽度正确提升, rank 是"位宽不确定"时的兜底。
    // 只此一份: promote 与发码层 (cgen_expr_binary 的混符号加宽) 都问这里,
    // 免得两侧各写一套 8/16/32/64 的对照表而漂移。
    static bool intShape(Vb6Type t, int* bitsOut, bool* signedOut);

    // 两个类型的运算结果类型 (VB6 widened rules + VB.NET 二进制数值提升, ai/032 rev2)
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
    // 账 #230 §B139: 目标架构指针位宽 (32 / 64), 默认 64。见上方 setTargetPtrBits。
    static int targetPtrBits_;
};

} // namespace vb6c3
