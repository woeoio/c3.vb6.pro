#pragma once
// vb6rtl_crash.h - 账 #181 = C29-CH-f: 崩溃轨迹的**重复收费**闸
// 账 #305 §B218 第三百十五刀: 口径从「每进程每槽一次」改成「每槽每个**不同的故障地址**一次」。
//
// **零依赖**是刻意的: 三个出口分别在 vb6rtl.c 与 vb6forms.c 里, 而 vb6forms.c 的
// include 顺序吃不了 vb6rtl_runtime.h (那个头自己要求排在 array/builtin 之后,
// 提前吃它 = vb6_SafeArray1D 未声明)。诊断面的东西不该挑宿主。
//
// 为什么必须有这个闸: 记轨迹这件事本身会再触发异常 —— 栈溢出时 fprintf 要拿 CRT 锁、
// CaptureStackBackTrace 与栈扫描要再读同一批页, 于是处理器套处理器。实测 (czUI x64
// 启动期 AV 那份 c3_crash.txt): 同一份递归栈被写了 **4 遍**, 真正的第一现场
// (gdiplus 那条 AV) 反而排在最后 —— 诊断工具把现场盖掉了。B23/#176 那种堆损坏归因
// 读的也是这份文件, 所以这个闸是所有后续崩溃调查的前置。
//
// 第三百十五刀把「一次」问的**对象**改掉: 那个病是同一个地址被**重放**, 所以按故障地址
// 去重就挡得住; 而"每进程只记一次"顺手把**另一发完全不同的** AV 也挡了。实测靶工程
// ComCtlsDemo 的 Command2: vb6_ComIsDispatchable 里 IsBadReadPtr 那发**被守卫接住的**
// first-chance AV (READ 0C70C70C) 吃掉 FILE 槽唯一名额, 真打死进程的那一发 ole32 AV
// 只剩 filter 帧、一条应用栈都没有 ⇒ 账 #599 归因不动。口径照 §B188 的 vb6_TermErrClaim
// (进程内定长小表 + 按来路去重), 表满则不再记 —— 轨迹文件不许无界长。
//
// 槽位互不相干 (不是全局一份): 三个出口写的是**不同**落点、各有价值 ——
//   STDERR : vb6_CrashTraceVEH  → stderr (没有可写文件时也能看见)
//   FILE   : vb6_heapCorruptVEH → c3_crash.txt 的 === AV / === 堆损坏 段
//   FILTER : vb6_crashFilter    → c3_crash.txt 末尾的 === EXCEPTION 段
//            (过滤器天然只跑一次, 给它一个槽是为了让"只记一次"在同一处能读全)
// 返回 1 = 本槽位上这个**故障地址**第一次, 该写; 0 = 该地址已记过 / 表满, 让路。
// faultAddress 传 NULL = 退回旧口径 (本槽位只记第一次)。
#define VB6_CRASH_CLAIM_STDERR 0
#define VB6_CRASH_CLAIM_FILE   1
#define VB6_CRASH_CLAIM_FILTER 2

#ifdef __cplusplus
extern "C" {
#endif
int vb6_CrashTraceClaim(int slot, const void* faultAddress);
#ifdef __cplusplus
}
#endif
