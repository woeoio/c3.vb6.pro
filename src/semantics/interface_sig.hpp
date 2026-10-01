#pragma once
// vb6c3 - Interface 槽键与签名比对 (tB 扩展; ai/022 D2, 批次 B02)
//
// 同一套规范函数同时服务两侧: driver 的 stage 2.7 建槽表、语义层的 Implements 契约
// 比对。两侧走同一函数 = "接口怎么写、实现就得怎么写" 只有一个定义, 不会漂移。
//
// 比对口径 = **源码签名** (类型引用原文小写 + 参数个数 + ByVal/Optional/ParamArray
// + 返回类型), 不做 Vb6Type 归一: 跨模块的 Enum/UDT 名在 stage 3 早期尚未注入,
// 归一后比反而会把"其实正确的实现"误判成不符 (Fix 047 同源问题)。

#include "ast/ast.hpp"
#include <string>
#include <vector>

namespace vb6c3 {

inline std::string ifaceLower(const std::string& s) {
    std::string out;
    out.reserve(s.size());
    for (char c : s) out += (c >= 'A' && c <= 'Z') ? static_cast<char>(c + 32) : c;
    return out;
}

// 类型引用的规范文本 (nullptr = 未写 As = Variant)
inline std::string ifaceTypeText(const ASTNode* t) {
    if (!t) return "variant";
    switch (t->kind) {
        case ASTNodeKind::SimpleTypeRef:
            return ifaceLower(static_cast<const SimpleTypeRef*>(t)->name);
        case ASTNodeKind::ArrayTypeRef:
            return ifaceTypeText(static_cast<const ArrayTypeRef*>(t)->elementType.get()) + "()";
        case ASTNodeKind::FixedStringTypeRef:
            return "fixedstring";
        default:
            return "?";
    }
}

// COM 惯例槽键 (D2): 属性拆三槽, Sub/Function 用原名
inline std::string ifaceSlotKey(const Decl& d) {
    std::string n = ifaceLower(
        d.kind == ASTNodeKind::SubDecl      ? static_cast<const SubDecl&>(d).name :
        d.kind == ASTNodeKind::FunctionDecl ? static_cast<const FunctionDecl&>(d).name :
        d.kind == ASTNodeKind::PropertyDecl ? static_cast<const PropertyDecl&>(d).name : std::string());
    if (d.kind == ASTNodeKind::PropertyDecl) {
        switch (static_cast<const PropertyDecl&>(d).propKind) {
            case ProcKind::PropertyGet: return "get_" + n;
            case ProcKind::PropertyLet: return "put_" + n;
            case ProcKind::PropertySet: return "putref_" + n;
            default: break;
        }
    }
    return n;
}

struct IfaceParamSig {
    std::string type;
    bool isByVal = false;
    bool isOptional = false;
    bool isParamArray = false;
};

struct IfaceProcSig {
    std::string slotKey;
    std::string memberName;            // 声明原样名 (诊断可读)
    std::string procKind;              // "Sub" / "Function" / "Property Get" / ...
    std::string retType;               // 无返回值 (Sub/Let/Set) = ""
    std::vector<IfaceParamSig> params;
    std::string text;                  // 诊断用可读签名
    const Decl* decl = nullptr;        // 出处节点 (诊断定位)
};

inline std::string ifaceDeclName(const Decl& d) {
    return d.kind == ASTNodeKind::SubDecl      ? static_cast<const SubDecl&>(d).name :
           d.kind == ASTNodeKind::FunctionDecl ? static_cast<const FunctionDecl&>(d).name :
           d.kind == ASTNodeKind::PropertyDecl ? static_cast<const PropertyDecl&>(d).name
                                               : std::string();
}

// 仅 Sub/Function/Property 是接口成员/实现成员; 其余声明返回 false
inline bool ifaceSigFromDecl(const Decl& d, IfaceProcSig& out) {
    out = IfaceProcSig();
    out.decl = &d;
    out.memberName = ifaceDeclName(d);
    const std::vector<std::unique_ptr<ParameterDecl>>* params = nullptr;
    const ASTNode* ret = nullptr;
    switch (d.kind) {
        case ASTNodeKind::SubDecl: {
            const auto& n = static_cast<const SubDecl&>(d);
            out.procKind = "Sub"; out.slotKey = ifaceSlotKey(d); params = &n.params;
            break;
        }
        case ASTNodeKind::FunctionDecl: {
            const auto& n = static_cast<const FunctionDecl&>(d);
            out.procKind = "Function"; out.slotKey = ifaceSlotKey(d);
            params = &n.params; ret = n.returnType.get(); out.retType = ifaceTypeText(ret);
            break;
        }
        case ASTNodeKind::PropertyDecl: {
            const auto& n = static_cast<const PropertyDecl&>(d);
            out.slotKey = ifaceSlotKey(d); params = &n.params;
            if (n.propKind == ProcKind::PropertyGet) {
                out.procKind = "Property Get";
                out.retType = ifaceTypeText(n.returnType.get());
            } else {
                out.procKind = n.propKind == ProcKind::PropertyLet ? "Property Let"
                                                                     : "Property Set";
            }
            break;
        }
        default:
            return false;
    }
    out.text = out.procKind + " " + out.memberName + "(";
    for (const auto& p : *params) {
        IfaceParamSig ps;
        ps.type = ifaceTypeText(p->asType.get());
        ps.isByVal = p->isByVal;
        ps.isOptional = p->isOptional;
        ps.isParamArray = p->isParamArray;
        out.text += (out.params.empty() ? "" : ", ");
        out.text += (ps.isParamArray ? "ParamArray " : (ps.isByVal ? "ByVal " : "ByRef "));
        if (ps.isOptional) out.text += "Optional ";
        out.text += ps.type;
        out.params.push_back(std::move(ps));
    }
    out.text += ")";
    if (!out.retType.empty()) out.text += " As " + out.retType;
    return true;
}

inline bool ifaceSigEqual(const IfaceProcSig& a, const IfaceProcSig& b) {
    if (a.procKind != b.procKind || a.retType != b.retType) return false;
    if (a.params.size() != b.params.size()) return false;
    for (size_t i = 0; i < a.params.size(); i++) {
        if (a.params[i].type != b.params[i].type) return false;
        if (a.params[i].isByVal != b.params[i].isByVal) return false;
        if (a.params[i].isOptional != b.params[i].isOptional) return false;
        if (a.params[i].isParamArray != b.params[i].isParamArray) return false;
    }
    return true;
}

// ============================================================
// 成员级 `Implements I.M` 子句的共用口径 (tB 扩展, ai/022 B02b/B04)
// 语义层比对与后端绑定查找都用这两个函数, 保证"检查到什么就发什么码".
// ============================================================

// 过程声明尾部的子句列表; 非过程声明 = nullptr
inline const std::vector<ImplementsClause>* ifaceProcClauses(const Decl& d) {
    switch (d.kind) {
        case ASTNodeKind::SubDecl:      return &static_cast<const SubDecl&>(d).implementsClauses;
        case ASTNodeKind::FunctionDecl: return &static_cast<const FunctionDecl&>(d).implementsClauses;
        case ASTNodeKind::PropertyDecl: return &static_cast<const PropertyDecl&>(d).implementsClauses;
        default:                        return nullptr;
    }
}

// 属性成员槽键前缀 (Sub/Function = 空), 与 ifaceSlotKey 同一套 COM 惯例 (D2)
inline std::string ifaceSlotPrefix(const Decl& d) {
    if (d.kind != ASTNodeKind::PropertyDecl) return std::string();
    switch (static_cast<const PropertyDecl&>(d).propKind) {
        case ProcKind::PropertyGet: return "get_";
        case ProcKind::PropertyLet: return "put_";
        case ProcKind::PropertySet: return "putref_";
        default:                    return std::string();
    }
}

// 子句 `I.M` 的目标槽键: 按实现成员的 Get/Let/Set 补前缀 (也允许直接写槽名,
// 那由调用方的候选集回退处理)
inline std::string ifaceClauseSlotKey(const Decl& implDecl, const std::string& memberName) {
    return ifaceSlotPrefix(implDecl) + ifaceLower(memberName);
}

// ============================================================
// VB6 风格接口的 `<接口名>_<成员名>` 命名约定
// ============================================================
//
// VB6 没有 `Interface ... End Interface` 语法 (那是 tB 扩展), 它的接口就是一个
// `VB_Creatable = False`、成员全是**无体签名**的普通 .cls (如 VbEclipse 的 IScheme)。
// 实现方写 `Implements IScheme` 之后, 每个实现成员**必须**命名为 `IScheme_<成员名>` ——
// 这是 VB6 编译器强制的, 本仓库的 legacy 实路径 (semantic_analyzer.cpp 里
// `required = ifaceName + "_" + m`) 也正是这么比对的。
//
// 于是槽键不能直接拿实现成员的 slotKey 去比: `IScheme_ActiveCaptionForeColor` 算出来是
// `get_ischeme_activecaptionforecolor`, 而接口槽是 `get_activecaptionforecolor`, 永远
// 对不上, 结果是"一个槽都没实现"。所以给 clsHost 接口补一条摘前缀的回退匹配。
//
// 只对 clsHost (VB6 .cls) 接口生效: tB Interface 块的成员名本身就是契约, 不带这层前缀。
// 返回空串 = 该成员名没有这层前缀。
inline std::string ifaceVb6StrippedMember(const std::string& memberName,
                                          const std::string& ifaceName) {
    const std::string pre = ifaceLower(ifaceName) + "_";
    const std::string low = ifaceLower(memberName);
    if (low.size() > pre.size() && low.compare(0, pre.size(), pre) == 0) {
        return memberName.substr(pre.size());
    }
    return std::string();
}

// 实现成员 → 槽键候选 (按顺序试)。普通键第一; clsHost 接口再试摘掉 `<接口名>_` 的键。
inline std::string ifaceImplSlotKeyVb6(const Decl& implDecl, const std::string& memberName,
                                       const std::string& ifaceName, bool clsHost) {
    const std::string plain = ifaceClauseSlotKey(implDecl, memberName);
    if (!clsHost) return plain;
    const std::string bare = ifaceVb6StrippedMember(memberName, ifaceName);
    return bare.empty() ? plain : ifaceSlotPrefix(implDecl) + ifaceLower(bare);
}

} // namespace vb6c3
