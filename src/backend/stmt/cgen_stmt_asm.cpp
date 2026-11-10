#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_stmt_asm.cpp: 混排 (项2) + Asm 片段引用 VB 局部变量 (项3) ---
//
// ai/vb-asm-extension-spec §12.5。目标语义:
//
//   Sub Foo(ByVal n As Long)
//       Dim acc As Long
//       acc = n * 2
//       Asm                       ' ← 片段与 VB 语句混排
//           mov eax, [acc]
//           add eax, 1
//           mov [acc], eax
//       End Asm
//       Debug.Print acc
//   End Sub
//
// 两条后端怎么落地:
//
//   x86 —— MSVC 有 `__asm { }` 内联汇编, 且**按名解析** C 变量。于是就地发一个
//          `__asm { }` 块, `[acc]` → `acc` 即可: 编译器自己决定 acc 在寄存器还是
//          栈槽, 名字天然指向它。这是项3 在 x86 上的"免费"解。
//
//   x64 —— MSVC 完全不支持 x64 内联汇编。片段降级为一个**独立**的 MASM 过程,
//          被引用的 VB 变量以**地址**作参数传入: `[acc]` → `[rcx]` (rcx = &acc)。
//          调用点发 `vb6_asm_Foo_0((intptr_t)&acc)`。
//
// 明确语义边界 (与真 FB 内联汇编的差别, 必须写进文档):
//   * 每个片段是**独立单元** —— 寄存器/标志位状态不跨片段传播 (x64 下它们甚至是
//     不同的函数)。片段内部可自由使用任意寄存器 (callee-saved 由编译器自动保护)。
//   * x64 单个片段最多引用 **4 个** VB 变量 (Win64 只有 RCX/RDX/R8/R9 四个整型
//     参数寄存器; 地址必须常驻寄存器才能用 `[reg]` 表示"变量本身") → 超了报 3041。
//     x86 无此限制 (名字解析走栈帧, 不占参数寄存器)。
//   * 不能引用类成员 (`me->x`) —— 类方法整体不支持 Asm (3037)。
//   * 参数与局部变量的可见性以**发码时**的已知集合为准: 过程内 Dim 出来的标量局部
//     变量已在 hoistLocalDecls 阶段登记, 故片段写在 Dim 之前也能引用。

// ============================================================
// 引用扫描: 从汇编文本里挑出 `[名字]` / `[名字±N]` 形态
// ============================================================
CCodeGen::AsmRefScan CCodeGen::scanAsmRefs(const AsmStmt& node) {
    AsmRefScan out;
    std::vector<std::string> seen;   // 小写去重

    // 收集所有 `[` 后紧跟标识符的位置
    auto consider = [&](const std::string& name) {
        std::string lower = Symbol::toLower(name);
        // [Function] 是返回值占位, 与变量同形但不查符号表
        if (lower == "function") { out.function = true; return; }
        // 寄存器 (含别名) → 与 VB 变量无关, 不算引用
        if (asmRegWidthBits(lower) != 0) return;
        // MASM 的 ptr 关键字 / 大小标注不是名字
        if (lower == "ptr" || lower == "byte" || lower == "word" ||
            lower == "dword" || lower == "qword" || lower == "tbyte")
            return;
        if (std::find(seen.begin(), seen.end(), lower) != seen.end()) return;

        // 1) 是不是"当前作用域可见的 VB 变量/参数"?
        //    ByRef 与否必须查过程符号的 ParameterInfo (Symbol 的 Parameter 节点
        //    自己不记录 isByVal) —— 顺带把参数名单也建出来, 免得只靠 knownByRefParams_
        //    (那张表只装 ByRef 的, ByVal 参数要靠符号表局部作用域查)。
        bool isVar = false;
        bool isByRef = false;
        if (knownByRefParams_.count(lower)) { isVar = true; isByRef = true; }
        if (!isVar && currentProc_) {
            for (auto& pi : currentProc_->P().params) {
                if (Symbol::toLower(pi.name) == lower) { isVar = true; isByRef = !pi.isByVal; break; }
            }
        }
        if (!isVar && knownLocalVars_.count(lower)) isVar = true;
        if (!isVar) {
            // 本模块的模块级变量/数组 (isArray 是 Variable 上的标志, 不是独立 kind)
            Symbol* modSym = symTab_.lookupModule(name);
            if (modSym && modSym->kind == SymbolKind::Variable) isVar = true;
        }

        if (!isVar) { out.unknown.push_back(name); return; }

        seen.push_back(lower);
        out.vbVars.push_back(name);
        if (isByRef) out.byRefVars.push_back(name);
    };

    for (auto& raw : node.lines) {
        std::string s = raw;
        for (size_t i = 0; i < s.size(); i++) {
            if (s[i] != '[') continue;
            size_t j = i + 1;
            while (j < s.size() && (s[j] == ' ' || s[j] == '\t')) j++;
            if (j >= s.size()) break;
            unsigned char c0 = static_cast<unsigned char>(s[j]);
            if (!(std::isalpha(c0) || s[j] == '_')) continue;
            size_t k = j;
            while (k < s.size() &&
                   (std::isalnum(static_cast<unsigned char>(s[k])) || s[k] == '_'))
                k++;
            consider(s.substr(j, k - j));
            i = k;
        }
    }
    return out;
}

// ============================================================
// 混排降级主入口
// ============================================================
void CCodeGen::emitAsmStmtMixed(AsmStmt& node) {
    const bool x86 = (targetArch_ == "x86");

    if (isClassModule_) {
        diag_.error(DiagnosticID::SemAsmMixedBody, node.loc,
                    "类方法暂不支持 Asm 块 (仅标准模块过程)");
        return;
    }

    AsmRefScan refs = scanAsmRefs(node);

    // 小工具: 该名字是不是 ByRef 参数 (C 变量本身就是指针 → 传它本身, 不用取地址)。
    auto isByRefName = [&](const std::string& v) {
        std::string lower = Symbol::toLower(v);
        if (knownByRefParams_.count(lower)) return true;
        if (currentProc_) {
            for (auto& pi : currentProc_->P().params)
                if (Symbol::toLower(pi.name) == lower) return !pi.isByVal;
        }
        return false;
    };

    if (!refs.unknown.empty()) {
        diag_.error(DiagnosticID::SemAsmMixedRefUnresolved, node.loc,
                    "Asm 片段里引用的 `[" + refs.unknown.front() +
                    "]` 既不是寄存器, 也不是当前作用域可见的 VB 变量/参数; "
                    "混排片段只能引用过程的参数、局部变量或本模块的模块级变量");
        return;
    }

    // 序号: 同一过程内多个片段各自独立
    const int seq = asmMixedBlockCounter_++;
    std::string procTag = currentProc_ ? currentProc_->name : std::string("unnamed");

    // ---------------- x86: MSVC 内联汇编 ----------------
    if (x86) {
        std::string retName = currentReturnVar_.empty() ? std::string("eax") : currentReturnVar_;

        // ByRef 参数: C 变量是 `T*`, 而 VB 侧 `[x]` 指的是**调用者的那个变量**。
        // MSVC 内联汇编里 `[x]` ≡ `x` ≡ "x 这块内存", 对指针变量就是指针值 —— 不是
        // 调用者的变量。所以给每个被引用的 ByRef 参数造一个**本地副本**:
        //     T vb6_asmref_<blk>_<name> = (*x);
        //     __asm { ... [x] → vb6_asmref_... }
        //     (*x) = vb6_asmref_<blk>_<name>;
        // 这样 x86 与 x64 (地址参数) 的片段写法完全一致: `[x]` 都是"那个变量"。
        // 用本地副本而不是 `(*x)`: 内联汇编的操作数文法只认 `[reg±常量]` 与符号名,
        // 不接受 C 表达式。
        struct RefAlias { std::string ctype; std::string alias; std::string byRefPtr; };
        std::vector<RefAlias> aliases;

        std::vector<std::pair<std::string, std::string>> subs;
        for (auto& v : refs.vbVars) {
            if (isByRefName(v)) {
                // 指针的指向类型 = 参数的 C 类型去掉尾随 `*` (参数列表里 ByRef 已带 *)
                std::string pointee;
                if (currentProc_) {
                    for (auto& pi : currentProc_->P().params) {
                        if (Symbol::toLower(pi.name) == Symbol::toLower(v)) {
                            pointee = mapType(pi.type);
                            break;
                        }
                    }
                }
                while (!pointee.empty() && pointee.back() == ' ') pointee.pop_back();
                if (!pointee.empty() && pointee.back() == '*')
                    pointee.pop_back();
                while (!pointee.empty() && pointee.back() == ' ') pointee.pop_back();
                if (pointee.empty()) pointee = "int32_t";   // 兜底 (理论上到不了)
                std::string alias = "vb6_asmref_" + std::to_string(seq) + "_" + cIdent(v);
                aliases.push_back({ pointee, alias, cIdent(v) });
                subs.push_back({ "[" + v + "]", alias });
            } else {
                // ByVal 参数 / 局部变量 / 模块级变量: `[X]` 就是该变量本身,
                // MSVC 按名解析 (编译器自己决定它在寄存器还是栈槽)。
                subs.push_back({ "[" + v + "]", cIdent(v) });
            }
        }
        subs.push_back({ "[function]", retName });

        std::vector<std::string> body = asmRewriteLines(node.lines, subs, procTag);

        // 宽度校验 (把 cl 的 C2443/A2022 前移成 3038)
        bool widthBad = false;
        asmCheckRegWidths(body, [&](int idx, const std::string& dst, const std::string& src,
                                    int dw, int sw) {
            if (widthBad) return;
            widthBad = true;
            diag_.error(DiagnosticID::SemAsmOperandWidthMismatch, node.loc,
                        "Asm 第 " + std::to_string(idx + 1) + " 行操作数宽度不一致: `" +
                        node.lines[idx] + "` (" + dst + " 是 " + std::to_string(dw) +
                        " 位, " + src + " 是 " + std::to_string(sw) +
                        " 位); 请统一宽度 —— 32 位值用低 32 位寄存器 (如 ebx/edi), "
                        "或改用 movsxd/movzx");
        });
        if (widthBad) return;

        // 项1: 隐含累加器别名 → 3042 (x86 内联块同样适用)
        {
            bool aliasBad = false;
            asmCheckAccumAlias(body, [&](int idx, int wLine, int lLine,
                                         const std::string& mnem, const std::string& fam) {
                if (aliasBad) return;
                aliasBad = true;
                diag_.error(DiagnosticID::SemAsmAccumAliasClobber, node.loc,
                            "Asm 第 " + std::to_string(idx + 1) + " 行 `" + node.lines[idx] +
                            "`: 该指令的隐含累加器 " + fam + " 已被第 " +
                            std::to_string(wLine + 1) + " 行 `" + node.lines[wLine] +
                            "` 写坏 (之后第 " + std::to_string(lLine + 1) + " 行 `" +
                            node.lines[lLine] + "` 只写了它的低位)。cmpxchg/mul/div 的累加器"
                            "就是 AX/DX 家族 —— 别再把它当指针/基址用 (模板见 spec §11)");
            }, /*x64=*/false);
            if (aliasBad) return;
        }

        // 块内踩到的 callee-saved (ebx/esi/edi, 含 clobber 声明) 成对 push/pop
        std::vector<std::string> saved =
            asmSavedRegsForArch(node.lines, node.clobbers, /*x64=*/false);

        // ByRef 副本: 进块前拷入, 出块后拷回 (见上) —— 与 x64 的"地址参数"语义对齐
        for (auto& a : aliases)
            c_.emitLine(a.ctype + " " + a.alias + " = (*" + a.byRefPtr + ");");

        c_.emitLine("/* ai/vb-asm-extension-spec: Asm 片段 (混排, x86 __asm 内联) */");
        c_.emitLine("__asm {");
        c_.indent();
        for (auto& r : saved) c_.emitLine("push " + r);
        for (auto& l : body) c_.emitLine(l);
        for (auto it = saved.rbegin(); it != saved.rend(); ++it) c_.emitLine("pop " + *it);
        c_.dedent();
        c_.emitLine("}");

        for (auto& a : aliases)
            c_.emitLine("(*" + a.byRefPtr + ") = " + a.alias + ";");
        return;
    }

    // ---------------- x64: 独立 MASM 过程 + 地址参数 ----------------
    // 变量 → 参数槽 (按出现顺序)。地址必须常驻寄存器: 第 5 个起会落到栈上,
    // 而 `[X+4]` 这种带偏移的形态对栈槽地址无法就地表达 → 限 4 个。
    if (refs.vbVars.size() > 4) {
        diag_.error(DiagnosticID::SemAsmMixedRefLimit, node.loc,
                    "x64 混排片段最多引用 4 个 VB 变量 (Win64 只有 RCX/RDX/R8/R9 可用作"
                    "地址寄存器), 本片段引用了 " + std::to_string(refs.vbVars.size()) +
                    " 个: 请拆分片段, 或先把值搬到更少的变量里");
        return;
    }

    static const char* kReg64[4] = { "rcx", "rdx", "r8", "r9" };

    AsmProcInfo info;
    info.cName = "vb6_asm_" + cIdent(procTag) + "_" + std::to_string(seq);
    info.retCType = "void";
    info.naked = false;
    info.clobbers = node.clobbers;
    info.linesFinal = true;

    std::vector<std::pair<std::string, std::string>> addrSubs;   // VB名 → 寄存器
    std::vector<std::string> callArgs;
    for (size_t i = 0; i < refs.vbVars.size(); i++) {
        const std::string& v = refs.vbVars[i];
        std::string pname = "vb6_a" + std::to_string(i);
        info.params.push_back("intptr_t " + pname);
        addrSubs.push_back({ v, kReg64[i] });
        // ByRef 参数: C 变量就是指针 → 传它本身; 其余取地址。
        callArgs.push_back(std::string("(intptr_t)") + (isByRefName(v) ? "" : "&") + cIdent(v));
    }

    // [Function] → 返回变量的地址 (也是最后一个参数), 只有 Function/Property Get 有
    if (refs.function && !currentReturnVar_.empty()) {
        std::string pname = "vb6_ret";
        info.params.push_back("intptr_t " + pname);
        addrSubs.push_back({ "Function", kReg64[refs.vbVars.size()] });
        callArgs.push_back("(intptr_t)&" + currentReturnVar_);
    } else if (refs.function) {
        // Sub 里 [Function] → 随便一个可写寄存器 (与 x86 路径的 eax 口径一致; 用户
        // 写进它的值会被丢弃, 不会有更坏的后果)。
        addrSubs.push_back({ "Function", "rax" });
    }

    // 地址引用重写: [X] → [rcx], [X+4] → [rcx+4]
    std::vector<std::string> body = asmRewriteAddrRefs(node.lines, addrSubs);

    // 宽度校验用重写后的 body (asmRewriteAddrRefs 不增删行, 下标与源码行一一对应)
    bool widthBad = false;
    asmCheckRegWidths(body, [&](int idx, const std::string& dst, const std::string& src,
                                int dw, int sw) {
        if (widthBad) return;
        widthBad = true;
        diag_.error(DiagnosticID::SemAsmOperandWidthMismatch, node.loc,
                    "Asm 第 " + std::to_string(idx + 1) + " 行操作数宽度不一致: `" +
                    node.lines[idx] + "` (" + dst + " 是 " + std::to_string(dw) +
                    " 位, " + src + " 是 " + std::to_string(sw) +
                    " 位); 请统一宽度 —— 32 位值用低 32 位寄存器 (如 ecx/eax), "
                    "或改用 movsxd/movzx");
    });
    if (widthBad) return;

    // 项1: 隐含累加器别名 (cmpxchg×RAX/EAX 等) → 3042。
    // 混排 x64 片段里地址在 rcx/rdx/r8/r9, 理论上不碰 rax —— 但用户完全可能在片段内
    // 自己 `mov rax, [q]` 之类, 所以同一套检查照做。
    {
        bool aliasBad = false;
        asmCheckAccumAlias(body, [&](int idx, int wLine, int lLine,
                                     const std::string& mnem, const std::string& fam) {
            if (aliasBad) return;
            aliasBad = true;
            diag_.error(DiagnosticID::SemAsmAccumAliasClobber, node.loc,
                        "Asm 第 " + std::to_string(idx + 1) + " 行 `" + node.lines[idx] +
                        "`: 该指令的隐含累加器 " + fam + " 已被第 " + std::to_string(wLine + 1) +
                        " 行 `" + node.lines[wLine] + "` 写坏 (之后第 " +
                        std::to_string(lLine + 1) + " 行 `" + node.lines[lLine] +
                        "` 只写了它的低 32 位)。cmpxchg/mul/div 的累加器是 RAX/EAX 一族, "
                        "而 EAX 就是 RAX 的低 32 位 —— 二者不可兼得; 请把指针/基址改放到 "
                        "R10/R11 等无关寄存器 (模板见 spec §11)");
        }, /*x64=*/true);
        if (aliasBad) return;
    }

    // 剩下的 `[名字]` 若既非地址参数也非寄存器, 就是打错了 —— 由 3040 已挡,
    // 这里再兜一次 (不报错, 交给 ml64; 但重写后残留 [x] 一定是漏网之鱼)。
    info.lines = body;

    std::string argList = "void";
    if (!callArgs.empty()) {
        argList.clear();
        for (size_t i = 0; i < callArgs.size(); i++) {
            if (i) argList += ", ";
            argList += callArgs[i];
        }
    }

    c_.emitLine("/* ai/vb-asm-extension-spec: Asm 片段 (混排, x64 → "
                + info.cName + ") */");
    if (info.params.empty())
        c_.emitLine("extern void " + info.cName + "(void);");
    else {
        std::string decl;
        for (size_t i = 0; i < info.params.size(); i++) {
            if (i) decl += ", ";
            decl += info.params[i];
        }
        c_.emitLine("extern void " + info.cName + "(" + decl + ");");
    }
    c_.emitLine(info.cName + "(" + argList + ");");

    asmProcs_.push_back(std::move(info));
}

} // namespace vb6c3
