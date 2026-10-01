// vb6c3 - 新式 Interface 契约比对 (tB 扩展; ai/022 D5, 批次 B02)
//
// 与 legacy VB6 Implements 的关系 (D5 分叉): 同一个 `Implements <Name>` 语句, 名字
// 命中 stage 2.7 登记表 = 新式接口 → 走这里的 **error 级**严格比对; 否则原样交给
// semantic_analyzer.cpp 的旧路径 (warn 级 + IFace_M 命名约定), 一行不改。
//
// 实现侧取名与接口侧共用 interface_sig.hpp 的规范函数, 因此"接口怎么写、实现就得
// 怎么写"只有一份定义。比对口径 = 源码签名 (见该文件头注释)。

#include "semantics/semantic_analyzer.hpp"

#include "semantics/interface_sig.hpp"
#include "semantics/interfaces_registry.hpp"

#include <map>
#include <set>
#include <string>
#include <vector>

namespace vb6c3 {
namespace {

// 槽键还原成可读成员名 (属性三槽去掉 get_/put_/putref_ 前缀)
std::string ifaceBareName(const std::string& slotKey) {
    for (const char* p : {"putref_", "get_", "put_"}) {
        std::string k(p);
        if (slotKey.size() > k.size() && slotKey.compare(0, k.size(), k) == 0) {
            return slotKey.substr(k.size());
        }
    }
    return slotKey;
}

// `Project.IFoo` 式限定名的末段 (与模块级 parseImplements 的点号拼接对称, Fix 083)
std::string ifaceLastSegment(const std::string& s) {
    size_t p = s.rfind('.');
    return p == std::string::npos ? s : s.substr(p + 1);
}

bool ifaceNameMatches(const std::string& written, const std::string& name) {
    const std::string lower = ifaceLower(name);
    return ifaceLower(written) == lower || ifaceLower(ifaceLastSegment(written)) == lower;
}

// 子句 `I.M` 在 view 展平槽表里对应的槽 (B02b)。两个自由度都要覆盖:
//   - M 写成员名 (`I.Name`) → 按实现成员的 procKind 补 get_/put_/putref_ 前缀;
//     同时允许直接写槽名 (`I.get_Name`)。
//   - I 写链上任一接口名 (`Extends` 继承来的槽也算), 见 IfaceSlotView::ownerIface。
const IfaceSlotView* findClauseSlot(const IfaceView& view, const std::string& writtenIface,
                                    const Decl& implDecl, const std::string& memberName) {
    const std::string bare = ifaceLower(memberName);
    const std::string prefixed = ifaceSlotPrefix(implDecl) + bare;
    for (const std::string& key : {prefixed, bare}) {
        for (const auto& slot : view.slots) {
            if (slot.key != key) continue;
            if (ifaceNameMatches(writtenIface, slot.ownerIface) ||
                ifaceNameMatches(writtenIface, view.name)) {
                return &slot;
            }
        }
    }
    return nullptr;
}

const IfaceView* lookupWrittenIface(const IfaceRegistry* reg, const std::string& written) {
    if (!reg) return nullptr;
    auto it = reg->find(ifaceLower(written));
    if (it == reg->end()) it = reg->find(ifaceLower(ifaceLastSegment(written)));
    return it == reg->end() ? nullptr : &it->second;
}

// (本类, 本接口) 是否被 stage 2.7 Pass D 判成委托式实现 `Implements I Via m_h` (B10)。
// 判定不在这里做 (要看别的模块的 Implements 列表, 语义层看不到), 这里只查结果。
bool viaDelegates(const ViaRegistry* reg, const Module& m, const IfaceView& view) {
    if (!reg) return false;
    auto it = reg->find(ifaceLower(m.moduleName));
    if (it == reg->end()) return false;
    const std::string want = ifaceLower(view.name);
    for (const ViaView& vv : it->second) {
        if (vv.ifaceKey == want) return true;
    }
    return false;
}

} // namespace

void SemanticAnalyzer::checkNewStyleInterface(const Module& module, const IfaceView& view,
                                              const std::string& writtenName,
                                              const SourceLocation& loc,
                                              std::set<IfaceClauseRef>& boundClauses) {
    // D11 v1 边界: 泛型类不得实现接口 (泛型器已拒 Implements, 这里兜住新式路径)
    if (!module.classTypeParams.empty()) {
        diag_.error(DiagnosticID::SemInterfaceNotImplemented, loc,
            "Generic class '" + module.moduleName +
            "' cannot implement interface '" + writtenName + "' (not supported yet)");
        return;
    }
    if (view.chainBroken) return;  // 建表阶段 (stage 2.7) 已就该接口报过错, 不再级联

    // tB 委托式实现 (B10): `Implements I Via m_h` 把整份契约转交给 m_h 那个对象, 所以
    // 逐槽"未实现"不报 —— 但本类自家写了的成员仍按接口槽校签名 (写了还对不上必然是笔误)。
    const bool delegated = viaDelegates(viaReg_, module, view);

    // 实现侧成员表: 槽键 → 签名 (两席争同一槽 = 契约歧义, 报错)
    std::map<std::string, IfaceProcSig> impl;
    auto bindSlot = [&](const std::string& key, const IfaceProcSig& sig,
                        const SourceLocation& where) {
        auto res = impl.emplace(key, sig);
        if (!res.second) {
            diag_.error(DiagnosticID::SemInterfaceSignatureMismatch, where,
                "Class '" + module.moduleName + "' has two members bound to interface '" +
                view.name + "' slot '" + key + "'");
        }
    };

    // --- 显式绑定 (B02b): 写了子句的成员**只**按子句入座, 不再参与同名隐式匹配 ---
    // VB6 风格接口 (.cls 宿主) 的实现成员带 `<接口名>_` 前缀, 槽键要摘掉前缀再比
    // (见 interface_sig.hpp 的 ifaceImplSlotKeyVb6)。
    const bool vb6Names = view.clsHost != nullptr;
    std::set<const Decl*> explicitDecls;
    for (const auto& d : module.declarations) {
        if (!d) continue;
        const std::vector<ImplementsClause>* clauses = ifaceProcClauses(*d);
        if (!clauses || clauses->empty()) continue;
        explicitDecls.insert(d.get());
        IfaceProcSig sig;
        if (!ifaceSigFromDecl(*d, sig)) continue;
        for (size_t i = 0; i < clauses->size(); i++) {
            const ImplementsClause& c = (*clauses)[i];
            const IfaceSlotView* slot = findClauseSlot(view, c.ifaceName, *d, c.memberName);
            if (!slot) continue;  // 不属于本接口: 留给别的接口认领, 都没认领则由兜底诊断报错
            boundClauses.insert(IfaceClauseRef(d.get(), i));
            bindSlot(slot->key, sig, d->loc);
        }
    }

    // --- 隐式匹配: 同名即入席 (大小写不敏感, 含属性三槽) ---
    for (const auto& d : module.declarations) {
        if (!d || explicitDecls.count(d.get())) continue;
        IfaceProcSig sig;
        if (!ifaceSigFromDecl(*d, sig)) continue;
        bindSlot(sig.slotKey, sig, d->loc);
        if (vb6Names) {
            // `IScheme_ActiveCaptionForeColor` -> 槽 `get_activecaptionforecolor`
            const std::string vb6Key = ifaceImplSlotKeyVb6(*d, sig.memberName, view.name, true);
            if (vb6Key != sig.slotKey) bindSlot(vb6Key, sig, d->loc);
        }
    }

    for (const auto& slot : view.slots) {
        IfaceProcSig want;
        if (slot.sig) ifaceSigFromDecl(*slot.sig, want);
        const std::string member = slot.ownerIface + "." +
            (slot.memberName.empty() ? ifaceBareName(slot.key) : slot.memberName);

        auto it = impl.find(slot.key);
        if (it == impl.end()) {
            if (!delegated) {
                diag_.error(DiagnosticID::SemInterfaceNotImplemented, loc,
                    "Implements " + writtenName + ": member '" + member + "' (" + want.text +
                    ") is not implemented by class '" + module.moduleName + "'");
            }
            continue;
        }
        if (!ifaceSigEqual(want, it->second)) {
            diag_.error(DiagnosticID::SemInterfaceSignatureMismatch, it->second.decl
                            ? it->second.decl->loc
                            : loc,
                "Implements " + writtenName + ": member '" + member +
                "' signature mismatch (interface: " + want.text + ", class: " +
                it->second.text + ")");
        }
    }
}

// ============================================================
// 成员级 Implements 子句的兜底诊断 (B02b)
// ============================================================
//
// 子句只有在"本类确实实现了那个接口"时才有意义。checkNewStyleInterface 每处理一个
// 新式接口就把认领掉的子句记进 boundClauses; 这里剩下的就是永远不会生效的写法,
// 三种成因分开给一句话, 免得使用者猜: 宿主不是类模块 / 接口名不存在或是 legacy
// 类 / 类没实现该接口 / 接口里没有这个成员。
void SemanticAnalyzer::checkMemberImplementsClauses(
        const Module& module, const std::set<IfaceClauseRef>& boundClauses) {
    // 泛型宿主: D11 边界已由契约比对报过错, 不再级联
    if (!module.classTypeParams.empty()) return;

    // 本类模块级 Implements 命中的新式接口集合 (按登记表节点身份比对, 规避大小写/限定名)
    std::set<const IfaceView*> implemented;
    for (const auto& impl : module.implements) {
        if (!impl) continue;
        if (const IfaceView* v = lookupWrittenIface(ifaceReg_, impl->interfaceName)) {
            implemented.insert(v);
        }
    }

    for (const auto& d : module.declarations) {
        if (!d) continue;
        const std::vector<ImplementsClause>* clauses = ifaceProcClauses(*d);
        if (!clauses || clauses->empty()) continue;
        IfaceProcSig sig;
        if (!ifaceSigFromDecl(*d, sig)) continue;
        for (size_t i = 0; i < clauses->size(); i++) {
            const ImplementsClause& c = (*clauses)[i];
            if (!module.isClassModule) {
                diag_.error(DiagnosticID::SemInterfaceClauseUnbound, c.loc,
                    "Member-level Implements is only valid in a class module (found one on " +
                    sig.text + ")");
                continue;
            }
            if (boundClauses.count(IfaceClauseRef(d.get(), i))) continue;
            const IfaceView* tgt = lookupWrittenIface(ifaceReg_, c.ifaceName);
            std::string msg;
            if (!tgt) {
                Symbol* legacy = symTab_.lookupModule(ifaceLastSegment(c.ifaceName));
                msg = (legacy && legacy->kind == SymbolKind::Class)
                    ? "Member-level Implements cannot bind to class '" + c.ifaceName +
                      "' (only Interface declarations carry a checked contract)"
                    : "Member-level Implements: interface '" + c.ifaceName + "' not found";
            } else if (!implemented.count(tgt)) {
                msg = "Member-level Implements: class '" + module.moduleName +
                      "' does not implement interface '" + c.ifaceName + "'";
            } else {
                msg = "Member-level Implements: interface '" + c.ifaceName + "' has no member '" +
                      c.memberName + "' for " + sig.procKind + " '" + sig.memberName + "'";
            }
            diag_.error(DiagnosticID::SemInterfaceClauseUnbound, c.loc, msg);
        }
    }
}

} // namespace vb6c3
