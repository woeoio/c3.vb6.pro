#pragma once
// 整数字面量的形状 —— 全编译器只有一个出口 (账 #194)
//
// 为什么要单开一个头: VB 的整数类型后缀 (`0&` = Long、`3%` = Integer) 留在 token 的 rawText 里,
// 而 C 不认识它们。发码侧 (visit(LiteralExpr)) 本来就是按**数值**重打, 从不抄 rawText;
// 但语义层那份「Optional 默认值求值」(evalOptionalDefault) 是**另写的一份** —— Long/Integer 两支
// 直接 return rawText ⇒ `Optional ByVal FontIndex As Long = 0&` 发成 `(*FontIndex) = 0&;`
// → MSVC C2059 "bad suffix on number" (实测 Charts 2020/ucTreeMaps 的 PropPagFMR.c 723/923)。
//
// 用法: 要 Long 档就 wantLong=true (一律带 L, 装不进 32 位带 LL), Integer 档带 false。
// 别在调用点再自己拼 "L" / 抄 rawText —— 后缀那一步漏一次就是整片工程编不过。

#include <cstdint>
#include <string>

namespace vb6c3 {

// Integer 档 → 十进制无后缀; Long 档 → 一律带 L。
// LL 那一步不是装饰: MSVC 的 long 是 32 位, `2147483648L` 装不下会退成 unsigned long,
// 而负号在 VB 里是**另一个 token** (一元 Negate) ⇒ 既可能 C2105/C4146, 也会把值算错。
inline std::string intLiteralText(int64_t value, bool wantLong) {
    if (!wantLong) return std::to_string(static_cast<int>(static_cast<int32_t>(value)));
    return std::to_string(value)
         + ((value < INT32_MIN || value > INT32_MAX) ? "LL" : "L");
}

} // namespace vb6c3
