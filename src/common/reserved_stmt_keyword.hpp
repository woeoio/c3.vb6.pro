// 账 §56.50: 「这个名字是不是 VB6 **保留语句关键字**」的唯一权威。
//
// VB6 语义: 保留**语句**关键字 (只有语句形态, 没有函数形态) 在**语句位置**永远是内建语句,
// 工程里同名的成员顶不掉它 —— 那枚成员只能靠**限定接收者**调。PD 真实语料:
//   tests/photodemon/Controls/pdRandomizeUI.ctl:211  `Public Sub Randomize()`
//   tests/photodemon/Controls/pdRandomizeUI.ctl:301  `Me.Randomize`   ← 成员只能这样调
// 而 pdPaletteChild.cls:291-292 / pdRandomize.cls:118 的**裸** `Randomize` 是内建语句。
//
// 顶不掉的后果有两副面孔, 同一个根 (裸写的名字被解析成那枚 0 参成员):
//   · **语义层**: 把它当成被调方 ⇒ 假 VB3005「参数数量错误: 'Randomize' 需要 0-0 个参数」
//     (同模块那枚; 本机探针 .temp/probe/rnd4 复现)
//   · **发码层**: 拿它的**空形参表**当被调方 ⇒ Fix 041b「实参多于形参就截断」把种子
//     实参全砍掉 ⇒ `vb6_Randomize()` ⇒ cl C2198「too few arguments for call」
//     (跨模块那枚; 本机探针 .temp/probe/rnd3 与 PD 三处复现)
//
// 口径只此一处: 语义层与发码层都问这一个函数 (名字→小写后查表, 大小写不敏感)。
//
// 名单只收**语句形态独占**的那些: 有函数形态的一律不收 ——
//   · `Input` / `Seek`  既有语句形态 (`Input #f, x` / `Seek #1, 0`) 也有函数形态
//     (`Input(n, #f)` / `Seek(1)`), 裸写按函数理解才是常态 ⇒ 不收。
//   · `Rnd` / `Mid` / `Left` / `Right` 等是纯函数, 不是语句关键字 ⇒ 不收。
#ifndef C3_COMMON_RESERVED_STMT_KEYWORD_HPP
#define C3_COMMON_RESERVED_STMT_KEYWORD_HPP

#include <string>
#include <unordered_set>

namespace vb6c3 {

// 名字是不是 VB6 保留语句关键字 (C3 有内建出口的那几枚)。
inline bool isReservedStmtKeyword(const std::string& name) {
    static const std::unordered_set<std::string> kReservedStmtKeywords = {
        "randomize",  // Randomize [seed]      → vb6_Randomize / vb6_RandomizeTimer
        "kill",       // Kill pathname         → vb6_Kill
        "load",       // Load form             → vb6_LoadForm
        "unload",     // Unload form           → vb6_UnloadForm
    };
    std::string lower;
    lower.reserve(name.size());
    for (char c : name) {
        lower.push_back((c >= 'A' && c <= 'Z') ? static_cast<char>(c - 'A' + 'a') : c);
    }
    return kReservedStmtKeywords.count(lower) != 0;
}

}  // namespace vb6c3

#endif  // C3_COMMON_RESERVED_STMT_KEYWORD_HPP
