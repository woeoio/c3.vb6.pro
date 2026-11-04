#pragma once

// .frm 源码的"某过程是否真的定义"判据。
//
// 背景 (ucSegmentedControl): 窗体上的事件处理器常常被整段注释掉，例如
//
//     'Private Sub ucSegmentedControl2_ItemMouseEnter(Index As Long)
//     '    ...
//     'End Sub
//
// 旧实现只在整份小写源码里 find("sub <name>(")，注释行同样命中 ⇒ 为不存在
// 的处理函数发 thunk ⇒ 链接期 LNK2019。这里要求：命中位置所在行在去掉行首
// 空白后不是以 '\'' 开头（即不是注释行），并且过程名之后紧跟 '(' 或 " ("。

#include <cctype>
#include <string>

namespace vb6c3 {

inline bool frmHasSubDefinition(const std::string& lowerCode, const std::string& lowerName) {
    const std::string needle = "sub " + lowerName;
    std::size_t pos = 0;
    while ((pos = lowerCode.find(needle, pos)) != std::string::npos) {
        // 回退到本行行首，跳过空白，若第一个字符是 '\'' 则是注释行。
        std::size_t lineStart = lowerCode.rfind('\n', pos);
        lineStart = (lineStart == std::string::npos) ? 0 : lineStart + 1;
        std::size_t k = lineStart;
        while (k < pos && std::isspace(static_cast<unsigned char>(lowerCode[k]))) ++k;
        const bool commentedLine = (k < pos && lowerCode[k] == '\'');
        if (!commentedLine) {
            std::size_t p = pos + needle.size();
            if (p < lowerCode.size()
                && (lowerCode[p] == '(' || (std::isspace(static_cast<unsigned char>(lowerCode[p]))
                                            && p + 1 < lowerCode.size() && lowerCode[p + 1] == '('))) {
                return true;
            }
        }
        pos += needle.size();
    }
    return false;
}

} // namespace vb6c3
