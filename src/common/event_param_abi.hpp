// 账 #305 第二百五十六刀: 「这一位事件形参要不要按地址递」的唯一主人。
//
// VB6 的 `Public Event GetEmptyMarkup(ByRef Text As String, ByRef Center As Boolean)` 里
// ByRef 那一档是**出参**: 处理器写进去的值要回到递出方。实测 ComCtlsDemo
// Builds/ListView/ListView.ctl:8946 递出之后紧接着 `If Not Text = vbNullString Then …` ⇒
// 值回不来就是"整块空表提示文本不画"(两头证人: 插桩真跑打出处理器被调到, 而产物 typedef 是
// (void* handler, BSTR Text, int16_t Center))。
//
// 从前这一问在六处各自硬答"按值" (emitEventSink 的回调 typedef / 窗体 prelude 的 thunk
// 形参表 / applyEventHandlerAbi 把处理器形参全翻 ByVal / visit(RaiseEventStmt) 的实参位 /
// 模块级 Dim WithEvents 那两站包装器的 .h 声明与 .c 定义)。UC 那一路六答彼此自洽 ⇒ 编得过、
// 跑不崩, 只是出参的值回不来 (第二百五十五刀: ListView 空表提示文本永远不画); 模块级那一路
// 从来就不自洽 —— 处理器自己的 VB 签名是 ByRef (C 侧 BSTR*), 包装器却按值 (BSTR), 生成的 C
// 是「把 wchar_t* 递给 BSTR* 槽」, C 模式下 cl 只给一条警告, 运行期处理器多解一层引用
// ⇒ 实测 0xC0000005 (2026-10-27 拿 BASE 编译器跑 tests/test_events 的读数)。
// 本文件把那一条问句收在一处, 六站同问。
#ifndef C3_COMMON_EVENT_PARAM_ABI_HPP
#define C3_COMMON_EVENT_PARAM_ABI_HPP

#include <string>

namespace vb6c3 {

// 数组形参刻意排除: `ByRef State() As Boolean` 的 C 侧本来就是载体指针 vb6_SafeArray1D*,
// 再加一层星号换的是"载体的地址"而不是"元素可写" —— 而 VB6 那一侧要的就是元素可写, 已经成立。
inline bool eventParamByAddress(bool declaredByRef, bool arrayCarrier) {
    return declaredByRef && !arrayCarrier;
}

// 形参声明文本（typedef 与 thunk 两处共用，别各自拼一遍）
inline std::string eventParamCDecl(const std::string& baseCType, bool byAddress,
                                   const std::string& name) {
    return byAddress ? (baseCType + "* " + name) : (baseCType + " " + name);
}

}  // namespace vb6c3

#endif  // C3_COMMON_EVENT_PARAM_ABI_HPP
