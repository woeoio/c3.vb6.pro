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
//
// 账 §56.47: **两路两口径**, 各路内部必须自洽, 判定仍只此一处:
//   · **类事件 / 模块级 WithEvents 那一路** (isUserControl = false) —— **VB6 默认**:
//     没写 ByVal 就是 ByRef。与同一枚处理器的**定义** (过程那一路, 读 isByVal) 同一答案
//     ⇒ 包装器与处理器不再一个按值一个按地址 ⇒ 不再 cl C2440 (PD 实测: pdCBZ.c 的
//     `m_ZipArchive_BeforeExtract`, 事件声明
//     `Event BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean)`
//     里 `File` 没写 ByVal ⇒ VB6 默认 ByRef)。
//   · **UC (.ctl) 那一路** (isUserControl = true) —— 仍只认**显式** ByRef。这一路六站
//     本来自洽 (typedef / 抬事件 / prelude thunk / 容器侧原型 / 处理器 / 出参槽),
//     语料里几千处形状靠它冻结 ⇒ 本刀**不动**。
//     ⚠ 本刀第一版漏了这一条: typedef 与抬事件那一侧改了口径, prelude 的 thunk 与
//     容器侧处理器仍是按值, 两边靠一枚显式 cast 才编得过 ⇒ 运行期把地址当 24B 槽读。
//     判据 `ucevt_thunk_type_same_authority` (tests/run_tests.ps1) 当时就红了, 是它抓到的。
#ifndef C3_COMMON_EVENT_PARAM_ABI_HPP
#define C3_COMMON_EVENT_PARAM_ABI_HPP

#include <string>

namespace vb6c3 {

// 数组形参刻意排除: `ByRef State() As Boolean` 的 C 侧本来就是载体指针 vb6_SafeArray1D*,
// 再加一层星号换的是"载体的地址"而不是"元素可写" —— 而 VB6 那一侧要的就是元素可写, 已经成立。
//
// 账 §56.47: **两路两口径**, 各路内部必须自洽, 判定仍只此一处 (出口就这一枚 inline):
//   · **类事件 / 模块级 WithEvents 那一路** (isUserControl = false) —— **VB6 默认**:
//     没写 ByVal 就是 ByRef。与同一枚处理器的**定义** (过程那一路, 读 isByVal) 同一答案
//     ⇒ 包装器与处理器不再一个按值一个按地址 ⇒ 不再 cl C2440 (PD 实测: pdCBZ.c 的
//     `m_ZipArchive_BeforeExtract`, 事件声明
//     `Event BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean)`
//     里 `File` 没写 ByVal ⇒ VB6 默认 ByRef)。
//   · **UC (.ctl) 那一路** (isUserControl = true) —— 仍只认**显式** ByRef。这一路六站
//     本来自洽 (typedef / 抬事件 / prelude thunk / 容器侧原型 / 处理器 / 出参槽),
//     语料里几千处形状靠它冻结 ⇒ 本刀**不动**。
//     ⚠ 本刀第一版漏了这一条: typedef 与抬事件那一侧改了口径, prelude 的 thunk 与
//     容器侧处理器仍是按值, 两边靠一枚显式 cast 才编得过 ⇒ 运行期把地址当 24B 槽读。
//     判据 `ucevt_thunk_type_same_authority` (tests/run_tests.ps1) 当时就红了, 是它抓到的。
//
// 形参声明的两个标志 (parser 置真, 见 ast_decl.hpp 的 ParameterDecl):
//   isByVal —— 显式 ByVal 为 true; 显式 ByRef / 什么都没写 为 false。
//   isByRef —— **仅显式** ByRef 为 true。
// isUserControl —— 这枚事件声明在**用户控件 (.ctl)** 模块里 (UC 那一路)。
inline bool eventParamByAddress(bool declaredByRef, bool isByVal,
                                bool isUserControl, bool arrayCarrier) {
    const bool byRef = isUserControl ? declaredByRef : !isByVal;
    return byRef && !arrayCarrier;
}

// 形参声明文本（typedef 与 thunk 两处共用，别各自拼一遍）
inline std::string eventParamCDecl(const std::string& baseCType, bool byAddress,
                                   const std::string& name) {
    return byAddress ? (baseCType + "* " + name) : (baseCType + " " + name);
}

}  // namespace vb6c3

#endif  // C3_COMMON_EVENT_PARAM_ABI_HPP
