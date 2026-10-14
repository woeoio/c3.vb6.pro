#pragma once
// phase_timing.hpp - 编译相位计时 (ai/030 续, 2026-10-14)
//
// 为什么: "稍微大一点的项目编译太费时间" 这条只能靠读数收敛。C3 的墙钟拆成
// 若干相位 (vbp/lex/preprocess/parse/typelib/prepasses/semantic/crossmod/
// unpack-rtl/codegen/link + 逐模块 gen/io + 逐过程 proc), 每个 mark 打**距上一个
// mark** 的毫秒。
//
// 口径:
//   · C3_TIMING=1 才输出, 未设时除一次 getenv 外零成本、零输出 (不动任何既有读数);
//   · 行格式 "C3: TIMING <label> <ms> ms" —— ASCII 可断言 (与 OBJCACHE 行同族);
//   · 打 stderr; 多次编译共用一个进程时时钟连续, 看差值即可;
//   · c3PhaseMarkIf 的"未打印也推时钟": 逐过程这种高频 mark 不打小值, 但每次都要
//     把时钟前推, 否则下一个 mark 会把前面吞掉的时间全算进去。
#include <chrono>
#include <cstdio>
#include <cstdlib>
#include <string>
#include <vector>

namespace vb6c3 {

namespace detail {
struct PhaseClock {
    bool on;
    std::chrono::steady_clock::time_point last;
    PhaseClock()
        : on(std::getenv("C3_TIMING") != nullptr),
          last(std::chrono::steady_clock::now()) {}
};
inline PhaseClock& phaseClock() { static PhaseClock c; return c; }
} // namespace detail

inline void c3PhaseMark(const std::string& label) {
    auto& clk = detail::phaseClock();
    if (!clk.on) return;
    auto now = std::chrono::steady_clock::now();
    double ms = std::chrono::duration<double, std::milli>(now - clk.last).count();
    std::fprintf(stderr, "C3: TIMING %s %.0f ms\n", label.c_str(), ms);
    std::fflush(stderr);
    clk.last = now;
}

// 只打 >= minMs 的 mark; 但无论打不打, 时钟都前推。
inline void c3PhaseMarkIf(const std::string& label, double minMs) {
    auto& clk = detail::phaseClock();
    if (!clk.on) return;
    auto now = std::chrono::steady_clock::now();
    double ms = std::chrono::duration<double, std::milli>(now - clk.last).count();
    if (ms >= minMs) {
        std::fprintf(stderr, "C3: TIMING %s %.0f ms\n", label.c_str(), ms);
        std::fflush(stderr);
    }
    clk.last = now;
}

// ---- 第二层: 热函数累计统计 (emitLine / inferExprType / emitExpr) ----
// 口径: C3_TIMING=1 时按模块 dump 一行 STATS 并清零 (c3StatDump); 高调用量函数里
// 用 RAII (C3ScopeTimer) 记账 —— 两次时钟读 ~60ns/次, 万级调用下忽略不计。
namespace detail {
struct PhaseStats {
    // 24 档上限: 0-7 热函数自时 / 8-12 call 分段 / 13-19 member 分段 / 20-22 ident 分段。
    // (曾把 kSlots 留 20 而用到 22 ⇒ ms[] 越界写进 n[] 区, 计数列打成垃圾数 —— 加档位
    //  必须同步扩这里。)
    static const int kSlots = 24;
    double ms[kSlots] = {};
    unsigned long long n[kSlots] = {};
};
inline PhaseStats& phaseStats() { static PhaseStats s; return s; }
} // namespace detail

inline void c3StatAdd(int slot, double ms) {
    auto& s = detail::phaseStats();
    s.ms[slot] += ms;
    s.n[slot] += 1;
}

// 分段打点 (v6): 不做 RAII 帧 —— 把 `now - prev` 记进 slot 并把 prev 推到 now。
// 给「同一函数体里顺序排列的片段组」归因用 (call/member 的 .inc 组), 组与组之间是
// 纯 wall 差值 (组内含子节点, 本就不求自时); 早退路径只是少打几个点, 不破坏别的档。
inline void c3StatSplit(int slot, std::chrono::steady_clock::time_point& prev) {
    auto& clk = detail::phaseClock();
    if (!clk.on) return;
    auto now = std::chrono::steady_clock::now();
    c3StatAdd(slot, std::chrono::duration<double, std::milli>(now - prev).count());
    prev = now;
}

inline void c3StatDump(const std::string& tag) {
    auto& clk = detail::phaseClock();
    if (!clk.on) return;
    auto& s = detail::phaseStats();
    std::fprintf(stderr,
                 "C3: STATS %s emitLine=%.0f/%llu inferExprType=%.0f/%llu emitExpr=%.0f/%llu "
                 "defVariant=%.0f/%llu ident=%.0f/%llu member=%.0f/%llu call=%.0f/%llu binary=%.0f/%llu\n",
                 tag.c_str(),
                 s.ms[0], s.n[0], s.ms[1], s.n[1], s.ms[2], s.n[2], s.ms[3], s.n[3],
                 s.ms[4], s.n[4], s.ms[5], s.n[5], s.ms[6], s.n[6], s.ms[7], s.n[7]);
    std::fprintf(stderr,
                 "C3: STATS2 %s callPre=%.0f callCallee=%.0f callParams=%.0f callArgs=%.0f callFixup=%.0f "
                 "mPre=%.0f mForm=%.0f mObj=%.0f mCls=%.0f mM22=%.0f mGen=%.0f mVoid=%.0f "
                 "idDisp=%.0f idSym=%.0f idBlt=%.0f\n",
                 tag.c_str(),
                 s.ms[8], s.ms[9], s.ms[10], s.ms[11], s.ms[12],
                 s.ms[13], s.ms[14], s.ms[15], s.ms[16], s.ms[17], s.ms[18], s.ms[19],
                 s.ms[20], s.ms[21], s.ms[22]);
    std::fflush(stderr);
    for (int i = 0; i < detail::PhaseStats::kSlots; i++) { s.ms[i] = 0; s.n[i] = 0; }
}

// RAII 记账器 —— **自时**口径 (ai/030 续 v5):
// 每条计时器只把「不在任何子计时器里」的那段时间记进自己的档位; 子计时器的整段时长
// 从父帧里扣 (childMs 累积)。于是 8 档读数互不重叠、相加 ≈ 全时, 每档都是"该函数
// 自己干活"的钱 (emitExpr 的 self 近零 = 它的活都在子节点/子函数里; call 的 self 大
// = 调用节点自己体内那句代码重)。单线程假设成立 (C3 前端不并行), 栈直接全局一张。
namespace detail {
struct TimerFrame {
    int slot;
    std::chrono::steady_clock::time_point start;
    double childMs;
};
inline std::vector<TimerFrame>& timerStack() {
    static std::vector<TimerFrame> s;
    return s;
}
} // namespace detail

struct C3ScopeTimer {
    bool on;
    explicit C3ScopeTimer(int s) : on(false) {
        static const bool c3t = (std::getenv("C3_TIMING") != nullptr);
        on = c3t;
        if (!on) return;
        detail::timerStack().push_back({s, std::chrono::steady_clock::now(), 0.0});
    }
    ~C3ScopeTimer() {
        if (!on) return;
        auto& st = detail::timerStack();
        if (st.empty()) return;
        detail::TimerFrame frame = st.back();
        st.pop_back();
        double ms = std::chrono::duration<double, std::milli>(
                        std::chrono::steady_clock::now() - frame.start).count();
        double selfMs = ms - frame.childMs;
        if (selfMs < 0.0) selfMs = 0.0;
        c3StatAdd(frame.slot, selfMs);
        if (!st.empty()) st.back().childMs += ms;
    }
};

} // namespace vb6c3
