// keeps_calib.cpp —— 账 #458 §B228 的**校准探针**（K17b）
//
// 用途：把哨兵里 PowerShell 侧重演的那条规则，与**头里那位主人**（C++ 侧
// `ownedTempNominationVerdict`）逐条比对 ⇒ 两侧同源是**跑出来的**，不是断言。
// 用法（要有 cl / VS）：
//     cl /nologo /utf-8 /I<repo>\src scripts\keeps_calib.cpp /Fe:keeps_calib.exe
//     keeps_calib.exe
// ⚠ **必须带 /utf-8**：本仓的源码是 UTF-8，而 cl 默认按 ANSI（cp936）读 ⇒ 中文注释的字节
//   在 cp936 下可能拼出**续行反斜杠** ⇒ 下一行被吞掉 ⇒ 报出来的错是**误导性的**
//   （实测：不带 /utf-8 时报「'vb6c3' 不是命名空间」，而头其实好好的）。
//   仓里 build.bat 的编译行本来就带 /utf-8 —— 这一条与本仓其它头同款纪律。
#include "common/owned_temp_keeps.hpp"
#include <cstdio>
using namespace vb6c3;

struct Shape { const char* id; const char* consumer; int slot; const char* mint; };

int main() {
    const Shape shapes[] = {
        { "concat-arg",   "vb6_Len",        0, "vb6_BSTR_Concat" },
        { "concat-arg2",  "vb6_Len",        0, "vb6_BSTR_Concat" },
        { "varianttoStr", "vb6_Len",        0, "vb6_VariantToString" },
        { "variant-arg",  "vb6_VarType",    0, "vb6_VariantFromValue" },
        { "variant-arg2", "vb6_VarAddV",    0, "vb6_CDec" },
        { "TRAP-borrowed","vb6_Len",        0, "vb6_ErrDescription" },
        { "TRAP-alias",   "vb6_Len",        0, "vb6_VariantArrayGetND" },
        { "TRAP-closed",  "vb6_SaveSetting",0, "vb6_BSTR_FromStr" },
        { "TRAP-comfour", "vb6_Len",        0, "vb6_ComCall" },
        { "TRAP-notmint", "vb6_Len",        0, "" },
    };
    for (const Shape& s : shapes) {
        char v = ownedTempNominationVerdict(s.consumer, s.slot, s.mint);
        printf("%s=%c\n", s.id, v ? v : '0');
    }
    return 0;
}
