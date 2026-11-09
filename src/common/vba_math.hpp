#pragma once
// VBA 标准库那两张名单的唯一出口（语义层与发码层共用）。
//
// 背景：PhotoDemon 的 `Math` 定案（ai/C3_PHOTODEMON_HANDOFF.md §10.4）。真 VB6 探针读数：
//   * `VBA.Math` 是 VB6 自带 VBA 类型库里的**标准模块**，不是类、不是对象；
//     成员只是内建数学函数的限定写法，真实成员集合 = Abs/Atn/Cos/Exp/Log/Rnd/Round/Sgn/Sin/Sqr/Tan。
//   * `Math.Fix(1.7)` / `Math.Int(1.7)` 在真 VB6 里是**编译期**「未找到方法或数据成员」——
//     Fix/Int 是全局内建函数，不是 VBA.Math 的成员。
//   * `Dim m As VBA.Math` / `Dim m As Math` 在真 VB6 里报「Visual Basic 中不支持自动化类型」——
//     标准模块不能当类型用。
// 本表此前只有发码层的一份副本（cgen_expr_member_precheck.inc 里的 vbaMathMembers），
// 语义层问不到，于是上面两条都漏成了「编过、链接期 LNK2019」或「静默当 Variant」。
// 收成一张表后，两层问的是同一个事实。

#include <cctype>
#include <string>
#include <unordered_set>

namespace vb6c3 {

// 把名字归到小写。common 不向上依赖 semantics 的 Symbol::toLower（与 host_pseudo.hpp 同一做法）。
inline std::string vbaLibLower(const std::string& s) {
    std::string o(s);
    for (char& ch : o) ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
    return o;
}

// VBA.Math 的**真实成员**（小写）。真 VB6 逐成员探针读出的集合，见 handoff §10.4。
// 刻意不含 fix/int：那是内建函数，真 VB6 在 `Math.Fix` 上报"未找到方法或数据成员"。
inline const std::unordered_set<std::string>& vbaMathMemberNames() {
    static const std::unordered_set<std::string> kNames = {
        "abs", "atn", "cos", "exp", "log", "rnd", "round", "sgn", "sin", "sqr", "tan"
    };
    return kNames;
}

// VBA 类型库里的**标准模块**名（小写，不带 "VBA." 前缀）。这些名字只能站在限定符位；
// 出现在 As 位置的合法答复是 VB6 的「Visual Basic 中不支持自动化类型」。
// 只收模块：VBA 类型库里的类/接口名（Collection / ErrObject / StdFont …）与枚举类型名
// 不在此列，`As VBA.VbMsgBoxStyle` 这类合法写法不受影响。
inline const std::unordered_set<std::string>& vbaStandardModuleNames() {
    static const std::unordered_set<std::string> kNames = {
        "math", "strings", "interaction", "information", "conversion",
        "datetime", "financial", "constants"
    };
    return kNames;
}

// 类型位置上的 VBA 标准模块判定（`VBA.Math` / `Math` 两种拼法都算，`MathHelper` 不算）。
inline bool isVbaStandardModuleName(const std::string& typeName) {
    std::string n = vbaLibLower(typeName);
    if (n.size() > 4 && n.compare(0, 4, "vba.") == 0) n = n.substr(4);
    return vbaStandardModuleNames().count(n) != 0;
}

} // namespace vb6c3
