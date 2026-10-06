#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>

namespace vb6c3 {

// --- cgen_delegate.cpp: 调用桩后端 (约定差异收敛在桩上) ---
// 两类桩共用同一套签名拆分/转发口径:
//   1) Delegate (tB 扩展): 值表示 = 按委托约定生成的桩地址, 与 LongPtr 位兼容.
//      - typedef: 委托声明的调用约定 + 参数 C 映射 (cast 调用点与 API 传参共用)
//      - thunk:   对外签名 = 委托约定; 参数逐字复制目标过程实际 C 签名, 体内纯
//                 转发转调 cdecl 目标 — 约定差异被桩吸收, 零 ABI/参数漂移.
//   2) AddressOf 取址的过程 (账 #184): VB6 的 AddressOf 交出去的就是 __stdcall 桩,
//      不是本体地址。x86 下本体是 __cdecl, 把本体交给 OS/COM 等于每次回调少弹
//      N*4 字节 —— 实测 VBFlexGrid 的 SUBCLASSPROC 起窗即 0xC0000374 (账 B23)。
//      这里为被标记的过程另发一枚 __stdcall 桩, AddressOf 站点取桩的地址。
// x64 下 MSVC 接受并忽略 __stdcall, 两种架构用同一份生成代码.

namespace {

// "RET name(params)" → 三段. 返回 false = 不是 "返回类型 名字(...)" 的形状.
// 参数逐字保留 (含括号), 桩签名靠它保证与本体**参数个数/宽度/顺序完全一致** ——
// 这是"x86 少弹 N*4 字节"这类事故不会再从这里长出来的前提。
bool splitProcSignature(const std::string& sig, std::string* ret,
                        std::string* target, std::string* params) {
    size_t lp = sig.find('(');
    if (lp == std::string::npos) return false;
    size_t sp = sig.rfind(' ', lp);
    if (sp == std::string::npos) return false;
    *ret = sig.substr(0, sp);
    *target = sig.substr(sp + 1, lp - sp - 1);
    *params = sig.substr(lp);
    return true;
}

// 从 "(type a, type b)" 提转发实参名 (与形参逐字同名)
std::string thunkArgsFromParams(const std::string& params) {
    std::string args;
    if (params.size() < 2) return args;
    std::string inner = params.substr(1, params.size() - 2);
    if (inner == "void" || inner.empty()) return args;
    size_t pos = 0;
    bool first = true;
    int depth = 0;  // 不切函数指针参数里的逗号 (防御)
    for (size_t i = 0; i <= inner.size(); i++) {
        char ch = (i < inner.size()) ? inner[i] : ',';
        if (ch == '(') depth++;
        else if (ch == ')') depth--;
        if (ch != ',' || depth != 0) continue;
        std::string p = inner.substr(pos, i - pos);
        pos = i + 1;
        size_t nameEnd = p.find_last_not_of(" \t");
        if (nameEnd == std::string::npos) continue;
        size_t nameStart = p.rfind(' ', nameEnd);
        std::string nm = (nameStart == std::string::npos)
            ? p.substr(0, nameEnd + 1)
            : p.substr(nameStart + 1, nameEnd - nameStart);
        if (!first) args += ", ";
        first = false;
        args += nm;
    }
    return args;
}

} // namespace

static std::string delLower(const std::string& s) {
    std::string r = s;
    std::transform(r.begin(), r.end(), r.begin(),
                   [](unsigned char c) { return (char)std::tolower(c); });
    return r;
}

std::string CCodeGen::delegateTyName(const std::string& delName) {
    return "vb6_del_t_" + delLower(delName);
}

std::string CCodeGen::delegateThunkName(const std::string& delName,
                                        const std::string& procName) {
    return "vb6_delthunk_" + delLower(delName) + "_" + delLower(procName);
}

// 在本模块声明里按 VB 名找 Sub/FunctionDecl (大小写不敏感)。
// procFp 非空 (O2 重载组) 时用符号表指纹精确匹配该变体。
Decl* CCodeGen::findDelegateTargetDecl(Module& module, const std::string& name,
                                        const std::string& procFp) {
    for (auto& d : module.declarations) {
        if (d->kind != ASTNodeKind::SubDecl && d->kind != ASTNodeKind::FunctionDecl)
            continue;
        const std::string dn = (d->kind == ASTNodeKind::SubDecl)
            ? static_cast<SubDecl&>(*d).name : static_cast<FunctionDecl&>(*d).name;
        if (delLower(dn) != delLower(name)) continue;
        if (procFp.empty()) return d.get();
        auto* sym = symTab_.lookupModuleOverloadByLoc(dn, d->loc);
        if (sym && !sym->overloadFp.empty() && sym->overloadFp == procFp) return d.get();
    }
    return nullptr;
}

void CCodeGen::emitDelegateDecls(Module& module) {
    for (auto& d : module.declarations) {
        if (d->kind != ASTNodeKind::DelegateDecl) continue;
        auto& del = static_cast<DelegateDecl&>(*d);
        auto* delSym = symTab_.lookupModuleByKind(del.name, SymbolKind::Delegate);

        std::string conv = del.callingConv == CallConv::CDecl ? "__cdecl" : "__stdcall";
        std::string retC = del.procKind == ProcKind::Function
                           ? mapTypeRef(del.returnType.get()) : "void";
        c_.emitLine("typedef " + retC + " (" + conv + " *" + delegateTyName(del.name) +
                    ")(" + makeParamList(del.params, true) + ");");

        if (!delSym) continue;
        for (auto& t : delSym->delegateTargets) {
            Decl* proc = findDelegateTargetDecl(module, t.procName, t.procFp);
            if (!proc) continue;  // 语义层已校验目标存在, 防御性跳过
            std::string sig = (proc->kind == ASTNodeKind::FunctionDecl)
                ? makeProcSignature(static_cast<FunctionDecl&>(*proc))
                : makeProcSignature(static_cast<SubDecl&>(*proc));
            // sig = "RET name(params)". 提取三段后用委托约定重组为桩签名.
            std::string ret, targetName, params;
            if (!splitProcSignature(sig, &ret, &targetName, &params)) continue;
            // 重载组 (O2): 桩名带变体指纹后缀, 与 visit(AddressOfExpr) 侧同规则
            std::string thunk = delegateThunkName(del.name, t.procName)
                              + (t.procFp.empty() ? "" : ovlCSuffixFromKey("$ov$" + t.procFp));
            std::string thunkSig = ret + " " + conv + " " + thunk + params;
            c_.emitLine("static " + thunkSig + ";");

            // 转发实参名 = 桩形参声明的尾标识符 (与目标逐字一致)。
            std::string args = thunkArgsFromParams(params);
            std::string body;
            if (proc->kind == ASTNodeKind::FunctionDecl) {
                body = thunkSig + " { return " + targetName + "(" + args + "); }";
            } else {
                body = thunkSig + " { " + targetName + "(" + args + "); }";
            }
            delegateThunkDefs_.push_back(std::move(body));
        }
        c_.emitLine("");
    }
}

void CCodeGen::emitDelegateThunks() {
    if (delegateThunkDefs_.empty()) return;
    c_.emitLine("// === Delegate call thunks (generated) ===");
    for (auto& def : delegateThunkDefs_) c_.emitLine(def);
    c_.emitLine("");
    delegateThunkDefs_.clear();
}

// --- AddressOf 回调桩 (账 #184) ---

std::string CCodeGen::addressOfThunkName(const std::string& procCName) {
    return "aoThunk_" + procCName;
}

// AddressOf 站点交出去的 C 名: 有桩取桩, 无桩取本体 (口径只在这一处)。
std::string CCodeGen::addressOfTargetCName(const Symbol* sym, const std::string& procCName) {
    if (sym && sym->addressOfCallback) return addressOfThunkName(procCName);
    return procCName;
}

// 被 AddressOf 取过址的标准模块过程 → 前向声明一枚 __stdcall 桩, 定义留到 epilogue。
// 桩的参数表逐字复制本体 (makeProcSignature), 体内原样转调 ⇒ 只有约定不同,
// 参数个数/宽度/顺序与本体完全一致 (与委托桩同一套拆/转发口径)。
void CCodeGen::emitAddressOfThunkDecls(Module& module) {
    if (module.isClassModule || module.isFormModule) return;  // 成员本体带 me, 桩造不出来
    for (auto& d : module.declarations) {
        bool isSub = (d->kind == ASTNodeKind::SubDecl);
        bool isFunc = (d->kind == ASTNodeKind::FunctionDecl);
        if (!isSub && !isFunc) continue;
        SubDecl* sub = isSub ? &static_cast<SubDecl&>(*d) : nullptr;
        FunctionDecl* fn = isFunc ? &static_cast<FunctionDecl&>(*d) : nullptr;
        const std::string vbName = sub ? sub->name : fn->name;
        const AccessLevel acc = sub ? sub->access : fn->access;
        Symbol* s = symTab_.lookupModule(vbName);
        if (!s || !s->addressOfCallback) continue;
        if (s->kind != SymbolKind::Sub && s->kind != SymbolKind::Function) continue;

        std::string sig = sub ? makeProcSignature(*sub) : makeProcSignature(*fn);
        std::string ret, target, params;
        if (!splitProcSignature(sig, &ret, &target, &params)) continue;
        if (target == addressOfThunkName(target)) continue;  // 防御: 别给桩再套桩

        std::string proto = ret + " __stdcall " + addressOfThunkName(target) + params;
        bool nonStatic = (acc == AccessLevel::Public || acc == AccessLevel::Friend ||
                          acc == AccessLevel::Protected);
        std::string prefix = nonStatic ? std::string() : "static ";
        h_.emitLine(prefix + proto + ";");

        std::string args = thunkArgsFromParams(params);
        std::string body = isFunc
            ? prefix + proto + " { return " + target + "(" + args + "); }"
            : prefix + proto + " { " + target + "(" + args + "); }";
        addressOfThunkDefs_.push_back(std::move(body));
    }
}

void CCodeGen::emitAddressOfThunkDefs() {
    if (addressOfThunkDefs_.empty()) return;
    c_.emitLine("// === AddressOf call thunks (generated, __stdcall) ===");
    for (auto& def : addressOfThunkDefs_) c_.emitLine(def);
    c_.emitLine("");
    addressOfThunkDefs_.clear();
}

// 委托变量直调 op(a, b): ((vb6_del_t_<del>)(op))(a, b)
void CCodeGen::emitDelegateCall(IndexOrCallExpr& node) {
    auto* ident = dynamic_cast<IdentifierExpr*>(node.callee.get());
    if (!ident) { lastExpr_ = "0"; return; }  // 语义层只标记 Identifier callee
    Symbol* del = symTab_.lookupModuleByKind(node.delegateTypeName, SymbolKind::Delegate);
    std::string cast = "((" + delegateTyName(node.delegateTypeName) + ")(" +
                       cIdent(ident->name) + "))";
    std::string args;
    for (size_t i = 0; i < node.positional.size(); i++) {
        emitExpr(*node.positional[i]);
        std::string e = lastExpr_;
        // ByRef 形参 + 变量实参 → 取地址 (与常规调用同形).
        if (del && i < del->params.size() && !del->params[i].isByVal &&
            node.positional[i]->kind == ASTNodeKind::IdentifierExpr) {
            e = "&" + cIdent(static_cast<IdentifierExpr&>(*node.positional[i]).name);
        }
        if (i) args += ", ";
        args += e;
    }
    lastExpr_ = cast + "(" + args + ")";
}

} // namespace vb6c3
