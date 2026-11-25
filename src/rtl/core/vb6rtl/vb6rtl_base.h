#pragma once
// vb6rtl_base.h - VB6 运行时公共基础头：标准库 / 平台 include
// 由 vb6rtl.h 伞头按固定顺序 include，不要单独使用

#include <stdint.h>
#include <stddef.h>
#include <wchar.h>
#include <stdlib.h>
#include <string.h>
#include <setjmp.h>
#include <stdio.h>
#include <math.h>
#include <limits.h>

#ifdef _WIN32
#include <windows.h>
#include <oleauto.h>
#endif

// ============================================================
// 账 #305 §11.159 (第一百七十六刀): double → Currency 的 4 位定点缩放**只有这一位主人**。
//
// VB6 的 Currency 是定点 4 位 (VT_CY = 值×10000), C3 用 double 承载**值**。把一个
// double 换回 4 位小数时，四种写法曾经各写一遍，且都是「先乘 10000 再取整」——
// 乘积本身先被舎进 double，于是**两次舍入**：探针 `.temp/k176/p1` 在真 VB6 里量到
// `3.00005` 存成 3.0000，而 `round(3.00005 * 10000.0)` 得 30001 (乘积恰好变成 30000.5)。
// 同一个偏差在 2.00005 / -3.00005 / 2.00025 / 4.00005 上各出现一次，
// 而 `0.00005` (真值本来就在半分之上) 两边都是 1 —— 这是「舍了乘积」而不是「取整方向变了一半」的铁证。
//
// 改法: 把整数部分与小数部分**拆开**缩放。
//   ip = trunc(v),  fp = v - ip                 (fp 精确: 同量级相减)
//   ip * 10000.0                                 (整数×10000, 在 Currency 值域内 < 2^53 ⇒ 精确)
//   round(fp * 10000.0)                          (只看 |x|<10000 的小量, 一次舍入)
// 二者相加仍精确，于是 .5 的判定对着**原值**而不是乘积。
// 注: double 里不存在形如 (2m+1)/20000 的值 (分母含 625), 所以真“半分”永不发生，
//     round() 的半值方向在这里根本不会被问到。
static inline int64_t vb6_CurrencyScale(double v) {
    if (!isfinite(v)) return 0;
    const double ip = trunc(v);
    const double fp = v - ip;
    return (int64_t)(ip * 10000.0 + round(fp * 10000.0));
}
