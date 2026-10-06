#pragma once
// vb6rtl_crash.h - 账 #181 = C29-CH-f: 崩溃轨迹**每进程只记一次**的闸
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
// 槽位各记一次 (不是全局一次): 三个出口写的是**不同**落点、各有价值 ——
//   STDERR : vb6_CrashTraceVEH  → stderr (没有可写文件时也能看见)
//   FILE   : vb6_heapCorruptVEH → c3_crash.txt 的 === AV / === 堆损坏 段
//   FILTER : vb6_crashFilter    → c3_crash.txt 末尾的 === EXCEPTION 段
//            (过滤器天然只跑一次, 给它一个槽是为了让"只记一次"在同一处能读全)
// 返回 1 = 本槽位第一次, 该写; 0 = 已写过, 让路。
#define VB6_CRASH_CLAIM_STDERR 0
#define VB6_CRASH_CLAIM_FILE   1
#define VB6_CRASH_CLAIM_FILTER 2

#ifdef __cplusplus
extern "C" {
#endif
int vb6_CrashTraceClaim(int slot);
#ifdef __cplusplus
}
#endif
