#pragma once
// cgen_pch.hpp - §56.53 PCH: 模块前奏 include 块的**唯一排法**
//
// 为什么要有这一份 (§56.53 的全部要害都在这里):
//   预编译头 (PCH) 的内容必须与被编译 TU 自己的前奏**逐字一致** —— 不一致 ⇒
//   套用 PCH 后的预处理状态 (声明/宏的集合与顺序) 与 TU 自己解析出来的状态**不同构**
//   ⇒ 编译结果可能与不开 PCH 时不同 (多一份声明 ⇒ 重载/遮蔽变了; 少一份 ⇒ C2061)。
//   差一行就是这种风险, 所以前奏的 include 块只有一处排法:
//     * 模块 .h 的前奏  —— cgen_base_generate_header_open.inc (发码器)
//     * PCH 头的文件内容 —— driver 侧写盘 (msvc_driver.cpp)
//   两边都从这里取, 于是"逐字一致"是**结构上**成立的, 不是靠人肉同步。
//
// 判据唯一权威 = 发码器: 「这个模块用哪一副前奏」只有发码器知道
//   (module.isFormModule || isDesignerModule_, 见 cgen_base_generate_prologue.inc)。
//   driver 只**转述**这个判定去分组, 不自己重判 (isDesignerModule_ 是发码器内部状态)。
//
// ⚠ PCH 头里**只有 include 块**: Fix 020 的 #undef 宏块、Fix 059 的 IID 重映射、
//   tB 接口契约 typedef 都**留在 TU 里** —— 它们的落点必须在 include 之后 (顺序语义),
//   而接口契约的守卫是模块级的 (每模块一份)。留在 TU 里 ⇒ 预处理状态与不开 PCH 时
//   完全同构 (PCH 只把"include 那一段"提前算好, 后面照旧)。

#include <string>
#include <vector>

namespace vb6c3 {

// 一副前奏的 include 块 (普通档 / 窗体档)。
// formPrelude = true ⇒ 窗体/设计期档: 额外要 windows.h + vb6forms.h。
inline std::vector<std::string> c3PreludeIncludeLines(bool formPrelude) {
    std::vector<std::string> lines;
    lines.push_back("#include <stdint.h>");
    lines.push_back("#include <stddef.h>");
    lines.push_back("#include <stdbool.h>");
    lines.push_back("#include <wchar.h>");
    if (formPrelude) {
        // P7: 窗体模式下 windows.h 必须在 vb6rtl.h 之前, 避免 VARIANT 重定义冲突
        // Fix 110: .ctl/.pag 也需要 vb6forms.h — 设计期子控件数组用 vb6_CtrlArr.
        lines.push_back("#include <windows.h>");
        lines.push_back("#include \"vb6forms.h\"");
    }
    lines.push_back("#include \"vb6rtl.h\"");
    return lines;
}

} // namespace vb6c3
