#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_decl.cpp: 通用声明辅助（过程数组追踪清理、过程/参数签名、枚举/类型/常量声明） ---

// Forward declaration from cgen_base.cpp
 FrmControlType controlTypeFromName(const std::string& name);

// --- cgen_decl.cpp: 声明生成 + 签名 + Property/Event ---


// Fix 056b: 过程开始时清理数组注册, 但保留模块级/类成员数组 (跨过程需要)
// 原逻辑 knownArrays_.clear() 会丢失模块级UDT数组的 arrayUdtElemTypes_ 注册,
// 导致方法体内访问元素时类型回退 vb6_VARIANT (MSVC C2440: 无法从LONG转换为vb6_VARIANT等).
void CCodeGen::clearProcArrayTracking() {
    std::vector<std::string> toErase;
    for (auto& name : knownArrays_) {
        // 模块级/类成员数组 → 符号表模块作用域有对应Variable符号 → 保留
        if (!symTab_.lookupModule(name)) toErase.push_back(name);
    }
    for (auto& name : toErase) {
        knownArrays_.erase(name);
        arrayElemTypes_.erase(name);
        arrayUdtElemTypes_.erase(name);
        arrayClassElemTypes_.erase(name);
        arrayDimCounts_.erase(name);
        knownByteArrayVars_.erase(name);
    }
    knownNDArraysInProc_.clear();
}

// ============================================================
// ai/vb-asm-extension-spec: Asm 块过程降级
//   命中: 过程体恰好是 1 条 AsmStmt (块形式或单行形式)。
//   x86 (32 位): 降级为 MSVC `__asm { }` 内联块 —— 共享 C 函数栈帧, 按名引用天然成立
//                ([var] → C 参数名), [Function] → 返回临时变量; 非 Naked 时自动
//                push/pop 块内踩到的 callee-saved (ebx/esi/edi, 及 clobber 里的 ebp)。
//   x64: 不发 C 体, 只发 `extern` 原型 + 把元数据记进 asmProcs_
//        (driver 侧生成 .asm, ml64 汇编后进链接)。
//   **体里还有其它语句 (混排, 项2) 时本函数返回 false** —— 交给常规过程生成流程,
//   其中每条 AsmStmt 由 emitStmtList 的 AsmStmt 分支就地降级 (见 cgen_stmt_asm.cpp)。
//   边界 (spec §10): 类方法 / 无 Optional/ParamArray / 参数为整型或指针或浮点 /
//   x64 参数 ≤4 (寄存器传参上限)。
// ============================================================
bool CCodeGen::tryEmitAsmProc(const std::string& procName, AccessLevel access,
                              std::vector<std::unique_ptr<ParameterDecl>>& params,
                              ASTNode* returnType, const StmtList& body, SourceLocation loc,
                              bool isNaked) {
    AsmStmt* asmNode = nullptr;
    for (auto& st : body) {
        if (st && st->kind == ASTNodeKind::AsmStmt) {
            if (!asmNode) asmNode = static_cast<AsmStmt*>(st.get());
        }
    }
    if (!asmNode) return false;   // 与 Asm 无关的常规过程

    // 项2 混排: 体里除了 Asm 块还有别的语句 → 不在这里接管。
    // (整过程降级的 x64 路只发 extern 原型, 没法表达"块之间还有 VB 语句";
    //  混排由 emitStmtList 逐条 AsmStmt 就地降级, 见 cgen_stmt_asm.cpp。)
    if (body.size() != 1) return false;

    auto fail = [&](DiagnosticID id, const std::string& msg) -> bool {
        diag_.error(id, loc, msg);
        return true;              // 已接管: 不再发 C 体 (编译已失败)
    };

    const bool x86 = (targetArch_ == "x86");

    if (isClassModule_)
        return fail(DiagnosticID::SemAsmMixedBody, "类方法暂不支持 Asm 块 (仅标准模块过程)");

    for (auto& p : params) {
        if (p->isOptional)   return fail(DiagnosticID::SemAsmMixedBody, "Asm 过程暂不支持 Optional 参数");
        if (p->isParamArray) return fail(DiagnosticID::SemAsmMixedBody, "Asm 过程暂不支持 ParamArray 参数");
    }

    AsmProcInfo info;
    info.cName = cProcName(procName, access, "");
    info.retCType = returnType ? mapTypeRef(returnType) : "void";
    for (auto& p : params) info.params.push_back(makeParamCType(p.get()));

    // 参数/返回类型白名单: 整型 / 任意指针 / 浮点 (float,double —— 走 xmm/浮点栈, spec §6)。
    // 其余 (VARIANT/结构体按值/String 等) 仍拒绝 —— 它们各有复杂编组约定, 待后续。
    static const char* kIntTypes[] = {"int8_t", "int16_t", "int32_t", "int64_t", "intptr_t", "unsigned", "VBABOOL"};
    auto typeOk = [](const std::string& t) {
        if (t.find('*') != std::string::npos) return true;    // 指针
        if (asmIsFloatCType(t)) return true;                  // float / double
        for (const char* k : kIntTypes) if (t == k) return true;
        return false;
    };
    for (auto& ps : info.params) {
        std::string t = ps.substr(0, ps.find(' '));
        if (!typeOk(t)) return fail(DiagnosticID::SemAsmMixedBody,
                                    "Asm 过程参数暂只支持整型/指针/浮点: " + ps);
    }
    if (!(info.retCType == "void") && !typeOk(info.retCType))
        return fail(DiagnosticID::SemAsmMixedBody,
                    "Asm 过程返回类型暂只支持整型/指针/浮点/void: " + info.retCType);

    info.lines = asmNode->lines;
    info.naked = isNaked;
    info.clobbers = asmNode->clobbers;

    std::string paramsC = makeParamList(params);

    // ---------------- x86: MSVC 内联汇编 ----------------
    if (x86) {
        // 注释状态: `<Naked>` 走 __declspec(naked) (无 prologue/epilogue, 用户自写 ret);
        // 普通块由 MSVC 保留 C 函数帧, 我们只在块入口/出口成对 push/pop callee-saved。
        if (info.naked)
            c_.emitLine("/* ai/vb-asm-extension-spec: <Naked> Asm 过程 → MSVC __asm 块 (x86, 无 prologue/epilogue) */");
        else
            c_.emitLine("/* ai/vb-asm-extension-spec: Asm 块 → MSVC __asm 块 (x86) */");

        std::string head = info.retCType + " " + info.cName + "(" + paramsC + ")";
        c_.emitLine((info.naked ? "__declspec(naked) " : "") + head + " {");

        const char* kRetName = "vb6_asm_ret_";
        // 注意: 用 `= 0` 而不是 `{}` —— `/std:c11` 下空花括号初始化标量是 C23 才有的语法,
        // 会让 cl 报 C2143 (实测踩过)。
        if (!info.naked && info.retCType != "void")
            c_.emitLine("    " + info.retCType + " " + kRetName + " = 0;");

        // `[param]` → C 参数名 (MSVC 内联汇编按名解析, ByRef 参数名本身就是指针 →
        // "变量即其地址"语义与 x64 侧一致); [Function] → 返回临时变量 (naked 下为 eax)。
        // x86 `<Naked>` 例外: 没有栈帧, 参数在调用者栈上, 名字无从解析 → 报 3036。
        for (auto& ps : info.params) {
            size_t sp = ps.find(' ');
            if (sp == std::string::npos) continue;
            std::string name = ps.substr(sp + 1);
            while (!name.empty() && name.back() == ' ') name.pop_back();
            if (info.naked) {
                std::string token = "[" + name + "]";
                std::string tokenLower; for (char ch : token) tokenLower += static_cast<char>(::tolower(static_cast<unsigned char>(ch)));
                for (auto& l : info.lines) {
                    std::string lower; for (char ch : l) lower += static_cast<char>(::tolower(static_cast<unsigned char>(ch)));
                    if (lower.find(tokenLower) != std::string::npos)
                        return fail(DiagnosticID::SemAsmFormUnsupported,
                                    "x86 <Naked> 过程无法按名引用参数 " + token +
                                    " (naked 无栈帧, 参数在调用者的栈上); 请去掉 <Naked> 或改用寄存器/立即数");
                }
            }
        }
        auto subs = asmBuildX86Subs(info, info.naked ? std::string("eax") : std::string(kRetName));

        std::vector<std::string> body = asmRewriteLines(info.lines, subs, info.cName);

        // 宽度校验 (把 cl 的 C2443/A2022 前移成 VB 诊断 3038)
        {
            bool widthBad = false;
            asmCheckRegWidths(body, [&](int idx, const std::string& dst, const std::string& src,
                                        int dw, int sw) {
                if (widthBad) return;   // 只报第一处
                widthBad = true;
                fail(DiagnosticID::SemAsmOperandWidthMismatch,
                     "Asm 第 " + std::to_string(idx + 1) + " 行操作数宽度不一致: `" +
                     info.lines[idx] + "` (" + dst + " 是 " + std::to_string(dw) +
                     " 位, " + src + " 是 " + std::to_string(sw) +
                     " 位); 请统一宽度 —— 32 位值用低 32 位寄存器 (如 ebx/edi), 或改用 movsxd/movzx");
            });
            if (widthBad) return true;   // 诊断已报, 不再发射
        }

        // 项1: 隐含累加器别名 (cmpxchg×EAX/EDX 等) → 3042。x86 同为 32 位寄存器:
        // `cmpxchg [eax], ecx` 的地址寄存器与隐含累加器同为 EAX 时同样互毁。
        {
            bool aliasBad = false;
            asmCheckAccumAlias(body, [&](int idx, int wLine, int lLine,
                                         const std::string& mnem, const std::string& fam) {
                if (aliasBad) return;
                aliasBad = true;
                fail(DiagnosticID::SemAsmAccumAliasClobber,
                     "Asm 第 " + std::to_string(idx + 1) + " 行 `" + info.lines[idx] +
                     "`: 该指令的隐含累加器 " + fam + " 已被第 " + std::to_string(wLine + 1) +
                     " 行 `" + info.lines[wLine] + "` 写坏 (之后第 " + std::to_string(lLine + 1) +
                     " 行 `" + info.lines[lLine] + "` 只写了它的低位)。cmpxchg/mul/div 的"
                     "累加器就是 AX/DX 家族 —— 别再把它当指针/基址用 (模板见 spec §11)");
            }, /*x64=*/false);
            if (aliasBad) return true;
        }

        std::vector<std::string> saved =
            info.naked ? std::vector<std::string>()
                       : asmSavedRegsForArch(info.lines, info.clobbers, /*x64=*/false);

        // 返回值收尾: 块里没写 `[Function]` 时, 值按 VB 约定留在返回寄存器 —— 落到 C 返回变量。
        // (x64 那条后端不需要这步: 值本来就在 RAX 里, 过程直接 ret 即可。)
        // 单行形式 (`Asm mov eax, [num]`) 正是靠这条拿到返回值。
        //   浮点返回 (spec §6): __stdcall 下返回值在 ST(0); 64 位整型在 EDX:EAX。
        // 判据用**重写后**的 body: 用户写了 [Function] 就会看到 kRetName。
        bool wroteRet = false;
        for (auto& l : body)
            if (l.find(kRetName) != std::string::npos) { wroteRet = true; break; }
        if (!info.naked && info.retCType != "void" && !wroteRet) {
            if (asmIsFloatCType(info.retCType))
                body.push_back(std::string("fstp ") + kRetName);      // ST(0) → 返回变量
            else if (info.retCType == "int64_t")
                body.push_back(std::string("mov dword ptr [") + kRetName + "], eax");  // 低 32 位
            else
                body.push_back(std::string("mov ") + kRetName + ", eax");
            // 注: int64 的高 32 位在 edx, 单独再发一条 (见下)。
            if (info.retCType == "int64_t")
                body.push_back(std::string("mov dword ptr [") + kRetName + "+4], edx");
        }

        c_.emitLine("    __asm {");
        for (auto& r : saved) c_.emitLine("        push " + r);
        for (auto& l : body) c_.emitLine("        " + l);
        for (auto it = saved.rbegin(); it != saved.rend(); ++it) c_.emitLine("        pop " + *it);
        c_.emitLine("    }");
        if (!info.naked && info.retCType != "void")
            c_.emitLine("    return " + std::string(kRetName) + ";");
        c_.emitLine("}");
        return true;
    }

    // ---------------- x64: 独立 MASM 过程 (driver 侧) ----------------
    // 宽度校验 (把 ml64 的 A2022 前移成 VB 诊断 3038): 用与 driver 完全同一张替换表
    // (asmBuildX64Subs) 模拟代入后查两个纯寄存器操作数的宽度。
    {
        auto xsubs = asmBuildX64Subs(info);
        auto xbody = asmRewriteLines(info.lines, xsubs, info.cName);
        bool widthBad = false;
        asmCheckRegWidths(xbody, [&](int idx, const std::string& dst, const std::string& src,
                                     int dw, int sw) {
            if (widthBad) return;   // 只报第一处
            widthBad = true;
            fail(DiagnosticID::SemAsmOperandWidthMismatch,
                 "Asm 第 " + std::to_string(idx + 1) + " 行操作数宽度不一致: `" +
                 info.lines[idx] + "` (" + dst + " 是 " + std::to_string(dw) +
                 " 位, " + src + " 是 " + std::to_string(sw) +
                 " 位); 请统一宽度 —— 32 位值用低 32 位寄存器 (如 ecx/eax), 或改用 movsxd/movzx");
        });
        if (widthBad) return true;   // 诊断已报, 不再收集 (编译到此失败)

        // 项1: 隐含累加器别名 (cmpxchg×RAX/EAX 等静态可见的踩法) → 3042
        bool aliasBad = false;
        asmCheckAccumAlias(xbody, [&](int idx, int wLine, int lLine,
                                      const std::string& mn, const std::string& fam) {
            if (aliasBad) return;   // 只报第一处
            aliasBad = true;
            fail(DiagnosticID::SemAsmAccumAliasClobber,
                 "Asm 第 " + std::to_string(idx + 1) + " 行 `" + info.lines[idx] +
                 "`: 该指令的隐含累加器 " + fam + " 已被第 " + std::to_string(wLine + 1) +
                 " 行 `" + info.lines[wLine] + "` 写坏 (之后第 " + std::to_string(lLine + 1) +
                 " 行 `" + info.lines[lLine] + "` 只写了它的低 32 位, 高 32 位回不来了)。"
                 "cmpxchg/mul/div 的累加器是 RAX/EAX 一族, 而 EAX 就是 RAX 的低 32 位 —— "
                 "把指针/基址放在 RAX 又让它当累加器, 必然互毁: 轻则永不相等死循环, 重则"
                 "把值当地址访问而崩溃。请把指针/基址改放到 R10/R11 等无关寄存器 "
                 "(模板见 spec §11)");
        }, /*x64=*/true);
        if (aliasBad) return true;
    }

    c_.emitLine("/* ai/vb-asm-extension-spec: 过程体为 Asm 块; 实现在 ml64 汇编的 "
                + info.cName + " (见 .asm) */");
    c_.emitLine("extern " + info.retCType + " " + info.cName + "(" + paramsC + ");");
    asmProcs_.push_back(std::move(info));
    return true;
}

std::string CCodeGen::makeProcSignature(SubDecl& node) {
    // 类模块方法始终带 vb6_<ClassName>_ 前缀 (与 dll_entry.c / resolveClassMemberCall 调用一致)
    // 重载组内按声明位置取本变体, 非 head 变体名带 _ov<fp> 后缀 (O2; 无重载时为空串)
    std::string name = cProcName(node.name, node.access, isClassModule_ ? moduleName_ : "")
                       + ovlCSuffix(symTab_.lookupModuleOverloadByLoc(node.name, node.loc));
    std::string params;
    if (isClassModule_) {
        params = classMeParam();
        std::string userParams = makeParamList(node.params);
        if (userParams != "void") {
            params += ", " + userParams;
        }
    } else {
        params = makeParamList(node.params);
    }
    return "void " + name + "(" + params + ")";
}

std::string CCodeGen::makeProcSignature(FunctionDecl& node) {
    // 类模块方法始终带 vb6_<ClassName>_ 前缀 (与 dll_entry.c / resolveClassMemberCall 调用一致)
    std::string name = cProcName(node.name, node.access, isClassModule_ ? moduleName_ : "")
                       + ovlCSuffix(symTab_.lookupModuleOverloadByLoc(node.name, node.loc));
    std::string params;
    if (isClassModule_) {
        params = classMeParam();
        std::string userParams = makeParamList(node.params);
        if (userParams != "void") {
            params += ", " + userParams;
        }
    } else {
        params = makeParamList(node.params);
    }
    std::string retType = mapTypeRef(node.returnType.get());
    return retType + " " + name + "(" + params + ")";
}

// Fix 084k: 单个参数的C类型+名字, 与makeParamList逐参数逻辑完全一致
//
// ai/024 T02 `staticEntry`: 静态库 (归档) 直连路径专用。唯一差别是 `ByVal <x> As String`:
//   动态路 (DLL 导入)   → `BSTR`  —— Fix 187 起调用点发的是 ANSI `char*`, 声明发 BSTR
//                                    本来就不一致, 靠 MSVC 只报 C4047 容忍。
//   静态路 (归档直连)   → `char*`  —— 没有转发桩做中间转换, `char*` 直接进真实函数,
//                                    必须与调用点口径一致, 否则警告噪声 + 固化不一致。
// 只改 ByVal String 这一种形态 (024 §五之三)。ByRef String 在静态路下仍是 `BSTR*`,
// 语义未定义, 属 v1 文档化边界, 不在这里猜。
std::string CCodeGen::makeParamCType(ParameterDecl* p, bool isDeclare, bool staticEntry) {
    // P14.1.5: ParamArray → SAFEARRAY* (always Variant array)
    if (p->isParamArray) {
        return "SAFEARRAY* " + cIdent(p->name);
    }

    std::string cType = mapTypeRef(p->asType.get());
    std::string cName = cIdent(p->name);

    // Fix 081e: Declare函数中ByVal Long/LongPtr参数映射为intptr_t
    // VB6 Long在Declare中常用于传句柄/指针 (ByVal hdc As Long等),
    // VB6是32位环境,Long=4字节=指针大小; 但x64下指针8字节,int32_t不够。
    // 将Declare中ByVal Long和ByVal LongPtr都映射为intptr_t:
    //   x86: intptr_t=4字节, 与VB6 Long兼容
    //   x64: intptr_t=8字节, 可容纳指针/句柄值
    // 纯值参数(如CodePage)传入intptr_t也不影响正确性(低32位包含值)。
    // ByRef Long参数不受影响(已映射为int32_t*,指针大小由架构决定)。
    if (isDeclare && p->isByVal && cType == "int32_t") {
        // 只对SimpleTypeRef中的Long/LongPtr提升为intptr_t
        if (p->asType && p->asType->kind == ASTNodeKind::SimpleTypeRef) {
            auto& simpleP = static_cast<SimpleTypeRef&>(*p->asType);
            if (simpleP.name == "Long" || simpleP.name == "LongPtr") {
                cType = "intptr_t";
            }
        }
    }

    // Fix 010: As Any 参数 — VB6中Any仅用于Declare, ByRef/ByVal均映射为void*
    // 不额外添加ByRef指针 (void*已是"指向任意类型的指针")
    bool isAnyType = false;
    if (p->asType && p->asType->kind == ASTNodeKind::SimpleTypeRef) {
        auto& simpleType = static_cast<SimpleTypeRef&>(*p->asType);
        if (simpleType.name == "Any" || simpleType.name == "any") {
            isAnyType = true;
            cType = "void*";
        }
    }

    // ai/024 T02: 静态归档直连路径 —— ByVal String 声明为 char* (见函数头注释)。
    // 放在这里而不是 mapTypeRef 里: mapTypeRef 是全局类型映射, 静态路只是"声明口径"
    // 不同, 不该污染全局 (同样的理由: 调用点的 ANSI 编组仍由 knownDeclareAnsi_ 单点驱动)。
    if (staticEntry && p->isByVal && cType == "BSTR") {
        cType = "char*";
    }

    // Fix 010r-6 rev2: ByRef array parameters need vb6_SafeArray1D** (double pointer)
    // so the callee can assign a new SafeArray (e.g. ReDim) and the caller sees it.
    // ByVal array params and As Any params stay as single pointer.
    bool isArrayParam = (p->asType && p->asType->kind == ASTNodeKind::ArrayTypeRef);

    if (p->isByVal || isAnyType) {
        return cType + " " + cName;
    }
    // ByRef → C pointer (ByRef array同: vb6_SafeArray1D** — callee can modify the caller's pointer)
    (void)isArrayParam;
    return cType + "* " + cName;
}

std::string CCodeGen::makeParamList(std::vector<std::unique_ptr<ParameterDecl>>& params, bool isDeclare,
                                    bool staticEntry) {
    if (params.empty()) return "void";

    std::string result;
    for (size_t i = 0; i < params.size(); i++) {
        if (i > 0) result += ", ";
        auto& p = params[i];
        result += makeParamCType(p.get(), isDeclare, staticEntry);
    }
    // P20-36: IsMissing support - append _has_ flags for Optional params
    // Fix 042c: Declare functions are __declspec(dllimport) — external DLL imports
    // that don't use the _has_ convention. Skip _has_ flags for Declare functions
    // so all modules agree on the same signature without _has_ params.
    if (!isDeclare) {
        for (size_t i = 0; i < params.size(); i++) {
            auto& p = params[i];
            if (p->isOptional && !p->isParamArray) {
                result += ", int _has_" + cIdent(p->name);
            }
        }
    }
    return result;
}

void CCodeGen::visit(EnumDecl& node) {
    std::string enumName = cIdent(node.name);

    // Fix 010b: 多个VB6模块可能定义同名枚举 (如LongPtr), 用#ifndef防止C2011重定义
    std::string guardName = "VB6_ENUM_" + enumName + "_DEFINED";
    h_.emitLine("#ifndef " + guardName);
    h_.emitLine("#define " + guardName);
    h_.emitLine("typedef enum vb6_enum_" + enumName + " {");
    h_.indent();

    // 逐成员推进: 每定下一枚就把 (小写名 → 值) 记进兄弟表, 下一个成员的值表达式里就能
    // 折叠到它。VB6 里 `A = 1 : B = A Or 2` 的 A 必须是**已声明**的兄弟, 顺序推进与语义一致。
    // 声明序之外的引用 (含前置引用) 仍按原有回退处理, 不静默改语义。
    //
    // 键必须走 `cIdent` 而不是原名: VB6 里枚举成员常写成方括号转义 (`[MSG_AFTER]`),
    // AST 里保留着括号, 而值表达式里引用它时是**裸名** (`MSG_AFTER`) —— 用原名做键
    // 查不到, 这正是本修复最初不生效的原因 (sib 表有 2 项仍 NOT-FOLDED)。
    std::map<std::string, int64_t> siblings;
    enumSiblingConsts_ = siblings;

    int64_t nextVal = 0;
    for (auto& member : node.members) {
        std::string memName = enumName + "_" + cIdent(member->name);
        int64_t settled = 0;
        bool haveSettled = false;
        if (member->value) {
            // Fix 010b: 尝试常量折叠enum成员值 (如2^0 → 1, 2^1|2^2 → 6)
            // C语言enum值必须是编译期常量, 不能用vb6_Pow()等函数调用
            int64_t constVal;
            if (tryEvalConstInt(member->value.get(), constVal)) {
                h_.emitLine("vb6_enum_" + memName + " = " + std::to_string(constVal) + ",");
                nextVal = constVal;
                settled = constVal;
                haveSettled = true;
            } else {
                // 回退: 使用表达式 (可能在C中编译失败)
                emitExpr(*member->value);
                h_.emitLine("vb6_enum_" + memName + " = " + lastExpr_ + ",");
                nextVal = 0;
                if (auto* lit = dynamic_cast<LiteralExpr*>(member->value.get())) {
                    if (lit->literalKind == LiteralKind::Long ||
                        lit->literalKind == LiteralKind::LongPtr) nextVal = lit->longValue;
                    else if (lit->literalKind == LiteralKind::Integer) nextVal = lit->intValue;
                } else if (auto* unary = dynamic_cast<UnaryExpr*>(member->value.get())) {
                    if (auto* inner = dynamic_cast<LiteralExpr*>(unary->operand.get())) {
                        int64_t v = 0;
                        if (inner->literalKind == LiteralKind::Long ||
                            inner->literalKind == LiteralKind::LongPtr) v = inner->longValue;
                        else if (inner->literalKind == LiteralKind::Integer) v = inner->intValue;
                        nextVal = (unary->op == UnaryOp::Negate) ? -v : v;
                    }
                }
            }
        } else {
            h_.emitLine("vb6_enum_" + memName + " = " + std::to_string(nextVal) + ",");
            settled = nextVal;
            haveSettled = true;
        }
        // 定下来的值进兄弟表, 供后续成员折叠。回退路径没算出确定值时不登记 ——
        // 拿 nextVal(=0) 冒充真实值会让后续成员静默算错。
        if (haveSettled) {
            std::string low = cIdent(member->name);  // 去方括号转义, 与引用侧同形
            for (auto& c : low) c = (char)tolower(c);
            siblings[low] = settled;
            enumSiblingConsts_ = siblings;
        }
        nextVal++;
    }
    enumSiblingConsts_.clear();  // 离开枚举作用域, 后续表达式不再拿兄弟表

    h_.dedent();
    h_.emitLine("} vb6_enum_" + enumName + ";");
    h_.emitLine("#endif");
    h_.emitBlank();
}

void CCodeGen::visit(EnumMember& node) {
    // 由EnumDecl内部处理
}

void CCodeGen::visit(TypeDecl& node) {
    // 泛型模板 (tB, G2): 模板本体不发码 (特化副本由泛型器注入, 是普通 TypeDecl)
    if (!node.typeParams.empty()) return;
    std::string typeName = cIdent(node.name);
    bool hasPack = (node.packingAlignment > 0);
    if (hasPack) {
        h_.emitLine("#pragma pack(push, " + std::to_string(node.packingAlignment) + ")");
    }
    // Fix 010b: 多个VB6模块可能定义同名UDT (如SYSTEMTIME, FILETIME), 用#ifndef防止C2011重定义
    std::string guardName = "VB6_TYPE_" + typeName + "_DEFINED";
    h_.emitLine("#ifndef " + guardName);
    h_.emitLine("#define " + guardName);
    h_.emitLine("typedef struct vb6_type_" + typeName + " {");
    h_.indent();

    for (auto& member : node.members) {
        std::string memType = mapTypeRef(member->type.get());
        std::string memName = cIdent(member->name);
        // 账 #118: 定长串字段 (`f As String * 8`) 的长度登记 —— 结构体这边只发
        // `BSTR f;`, 长度不进符号表 (语义层把 FixedStringTypeRef 折成 Vb6Type::String,
        // typeRefName 为空), 于是 `r.f = "ef"` 与 `Len(r.f)` 都无从知道该收口到 8。
        // 在这里按**声明**留一份 (键 = 结构体的 C 类型名), 供 fixedStrLenOfExpr 查。
        // 注意 emitExpr 会覆写 lastExpr_, 故先取长度再发字段行 (下面 mapTypeRef 已算完)。
        if (member->type && member->type->kind == ASTNodeKind::FixedStringTypeRef) {
            auto& fsField = static_cast<FixedStringTypeRef&>(*member->type);
            emitExpr(*fsField.length);
            if (!lastExpr_.empty())
                udtFixedStrFieldLen_["vb6_type_" + typeName][Symbol::toLower(member->name)] = lastExpr_;
        }
        // P15.2: 固定大小数组成员 (如 Buf(0 To 255) As Byte)
        if (member->arraySize) {
            // Fix 010d: 优先用常量折叠将数组维度求值为字面量
            // Private Const 只发射到.c, 不在.h中, 跨模块#include时会变成未声明标识符(C2065/C2057/C2229)
            // tryEvalConstInt覆盖: LiteralExpr, UnaryExpr, BinaryExpr(算术/位运算), IdentifierExpr(跨模块Const/EnumMember)
            // 账 #262: 每一维都走同一条折叠出口，所以 `M(0 To 4, 0 To 4)` 发得出
            // `float M[5][5]` —— 100 字节连片，正是 GDI+ ColorMatrix 要的那个形状
            // (以前只发得出第一维，第二维整个不见)。折不动的维沿用旧形 `(expr) + 1`。
            // 下界与第一维同口径: 只留语法、折算按 0..upper (声明 1 To 4 会多空一格)。
            int64_t arrVal;
            std::string brackets;
            if (tryEvalConstInt(member->arraySize.get(), arrVal)) {
                brackets = "[" + std::to_string(arrVal + 1) + "]";
            } else {
                emitExpr(*member->arraySize);
                brackets = "[(" + lastExpr_ + ") + 1]";
            }
            for (auto& md : member->moreDims) {
                if (tryEvalConstInt(md.upper.get(), arrVal)) {
                    brackets += "[" + std::to_string(arrVal + 1) + "]";
                } else {
                    emitExpr(*md.upper);
                    brackets += "[(" + lastExpr_ + ") + 1]";
                }
            }
            h_.emitLine(memType + " " + memName + brackets + ";");
        } else if (member->isArrayDynamic) {
            // Fix 037 Pattern B: 动态数组成员 (`Data() As Byte`) emit `vb6_SafeArray1D* Member;`
            // 之前 bug: arraySize==nullptr 与无括号成员无法区分, emit `uint8_t Data;` (单标量字段),
            // 运行时不正确且导致 obj.Data(i) 调用变成 C2064.
            h_.emitLine("vb6_SafeArray1D* " + memName + ";  /* dynamic array member */");
        } else {
            h_.emitLine(memType + " " + memName + ";");
        }
    }

    h_.dedent();
    h_.emitLine("} vb6_type_" + typeName + ";");
    h_.emitLine("#endif");
    if (hasPack) {
        h_.emitLine("#pragma pack(pop)");
    }
    h_.emitBlank();
}

void CCodeGen::visit(TypeMember& node) {
    // 由TypeDecl内部处理
}

void CCodeGen::visit(ConstDecl& node) {
    std::string cType = mapTypeRef(node.asType.get());
    std::string cName = cIdent(node.name);

    // Fix 049: Register module-level constant to type-specific known*Vars_ sets.
    // Module-level constants are emitted as #define macros; when used in
    // expressions, the codegen writes the constant name. Without registration,
    // inferExprType falls back to Variant, causing wrapToBSTR to generate
    // vb6_CStr(BSTR_const) → C2440 (BSTR→VARIANT).
    {
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (cType == "BSTR") {
            knownBstrVars_.insert(lower);
        } else if (cType == "int32_t" || cType == "int16_t" || cType == "VBABOOL") {
            knownLongVars_.insert(lower);
        } else if (cType == "float") {
            knownSingleVars_.insert(lower); knownDoubleVars_.insert(lower);   // Fix 117c: VT_R4
        } else if (cType == "double") {
            knownDoubleVars_.insert(lower);
        } else if (cType == "vb6_VARIANT") {
            // Fix 161f: 无 As Type 的 Const (Private Const LVM_GETHEADER = ...) 的
            // cType 是 mapTypeRef(nullptr) = "vb6_VARIANT"。但 VB6 语义上无类型
            // **整型** Const 在数值上下文就是数值 — 落 knownVariantVars_ 会让
            // Declare 调用的实参发射包 vb6_VariantToLong(<int 宏>) → C2440
            // "无法从 int 转换为 vb6_VARIANT" (extlist MListView SendMessage 实测)。
            // 整数可折叠的登记 knownLongVars_; 字符串/浮点仍按 Variant (旧路径)。
            int64_t cv161f = 0;
            if (node.value && tryEvalConstInt(node.value.get(), cv161f)) {
                knownLongVars_.insert(lower);
                // Fix 161f (续): 过程级集合会被 clear, 必须另存一份模块级表,
                // 供标识符内联数值使用。
                moduleIntConstValues_[lower] = cv161f;
            } else {
                knownVariantVars_.insert(lower);
            }
        }
    }

    if (node.value) {
        // Fix 091d: 整型常量表达式折叠 — 避免 Variant 语义包装操作数.
        // 例: BIF_USENEWUI = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE →
        // ((vb6_VariantToLong(64) | vb6_VariantToLong(16))) → C2440
        // "int → vb6_VARIANT" (cDialog.c 36; 使用处报 532). 折叠为字面量后
        // 与 VB6 常量语义一致 (整型位运算).
        int64_t cv091d = 0;
        if (tryEvalConstInt(node.value.get(), cv091d)) {
            std::string folded = "(" + std::to_string(cv091d) + ")";
            if (isPublicModuleDecl(node)) {
                h_.emitLine("#define " + cName + " " + folded);
            } else {
                c_.emitLine("#define " + cName + " " + folded);
            }
        } else {
            emitExpr(*node.value);
            // 公共常量 → .h, 私有 → .c
            if (isPublicModuleDecl(node)) {
                h_.emitLine("#define " + cName + " (" + lastExpr_ + ")");
            } else {
                c_.emitLine("#define " + cName + " (" + lastExpr_ + ")");
            }
        }
    }
}

} // namespace vb6c3
