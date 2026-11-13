// vb6c3 - 类继承链登记 + 继承成员合并 (tB 扩展; ai/022 记录 D6/D24/D26, 批次 B07a/B07b)
//
// stage 2.8 (runClassChainPrepass): 所有模块解析完、语义分析开始**之前**, 把工程内的类与其
// `Inherits` 子句建成只读登记表 ClassChainRegistry (小写类名 → 链视图)。必须早于 stage 3:
// 派生类的成员合并要读基类的声明, 而每个模块各有一张符号表 (登记表挂 Driver 的理由见
// class_chain_registry.hpp)。
//
// stage 3.4 (mergeInheritedMembers): 语义分析**之后**、跨模块链接 (3.5) **之前**, 把祖先自己
// 声明的成员并进派生类的 Class 符号 —— 晚于 3.5 就得再抄一遍外部工程副本。
//
// 两个阶段的早退条件都是"工程里既没有 Inherits / 虚修饰符, 也没有 Protected 成员"
// (D24 护栏口径 + B08c 的 D34-2): 无新语法时不改任何生成物、符号表或诊断。

#include "driver/driver.hpp"

#include "ast/ast.hpp"
#include "semantics/class_chain_registry.hpp"
#include "semantics/interface_sig.hpp"  // ifaceLower: tB 这一条线共用的小写键函数
#include "semantics/semantic_analyzer.hpp"  // stage 3.4: 取各模块 SymbolTable 并成员表

#include <algorithm>
#include <set>
#include <string>
#include <vector>

namespace vb6c3 {

namespace {

// 一个类模块**自己声明**的成员, 按用途分两桶 (键小写)。为什么不用它自己的 Symbol 表:
// 那张表在合并后含祖先条目 (以及祖先的祖先), 按它并表会把同一条目按不同层级重复计入,
// 而"谁赢"的裁决必须只看本类写了什么 (D26)。
struct OwnMemberKeys {
    std::vector<std::pair<std::string, Decl*>> fields;  // 数据字段 (键 → 声明)
    std::vector<std::pair<std::string, Decl*>> procs;   // Sub/Function/Property
};

OwnMemberKeys collectOwnMembers(Module& mod) {
    OwnMemberKeys out;
    for (const auto& d : mod.declarations) {
        if (!d) continue;
        std::string name;
        switch (d->kind) {
            case ASTNodeKind::VariableDecl: name = static_cast<const VariableDecl&>(*d).name; break;
            case ASTNodeKind::SubDecl:      name = static_cast<const SubDecl&>(*d).name; break;
            case ASTNodeKind::FunctionDecl: name = static_cast<const FunctionDecl&>(*d).name; break;
            case ASTNodeKind::PropertyDecl: name = static_cast<const PropertyDecl&>(*d).name; break;
            default: continue;  // Enum/Type/Const/Declare/Event 不参与合并
        }
        if (name.empty()) continue;
        const std::string key = ifaceLower(name);
        if (d->kind != ASTNodeKind::VariableDecl) {
            // 过程桶**不去重**: 同名 Property 的 Get/Let/Set 是三个方向、一个成员键, 转发桩
            // 要按方向各发一份 (成员表则按键只并一次, 见 mergeInheritedMembers)。
            out.procs.emplace_back(key, d.get());
            continue;
        }
        // 字段桶按键去重: 同名字段声明两次是非法输入, 结构体也只能有一个成员
        bool dup = false;
        for (const auto& p : out.fields) {
            if (p.first == key) { dup = true; break; }
        }
        if (!dup) out.fields.emplace_back(key, d.get());
    }
    return out;
}

AccessLevel memberAccess(const Decl& d) {
    switch (d.kind) {
        case ASTNodeKind::VariableDecl: return static_cast<const VariableDecl&>(d).access;
        case ASTNodeKind::SubDecl:      return static_cast<const SubDecl&>(d).access;
        case ASTNodeKind::FunctionDecl: return static_cast<const FunctionDecl&>(d).access;
        case ASTNodeKind::PropertyDecl: return static_cast<const PropertyDecl&>(d).access;
        default: return AccessLevel::Private;
    }
}

bool keyIn(const std::vector<std::string>& names, const std::string& key) {
    for (const auto& n : names) {
        if (ifaceLower(n) == key) return true;
    }
    return false;
}

// --- B08b: 虚方法修饰位的读法 (只有过程声明能带, 见 parser_decl.cpp) ---

ProcVirt procVirtOf(const Decl& d) {
    switch (d.kind) {
        case ASTNodeKind::SubDecl:      return static_cast<const SubDecl&>(d).virt;
        case ASTNodeKind::FunctionDecl: return static_cast<const FunctionDecl&>(d).virt;
        case ASTNodeKind::PropertyDecl: return static_cast<const PropertyDecl&>(d).virt;
        default: return ProcVirt::None;
    }
}

const char* procVirtText(ProcVirt v) {
    switch (v) {
        case ProcVirt::Overridable: return "Overridable";
        case ProcVirt::Overrides: return "Overrides";
        case ProcVirt::NotOverridable: return "NotOverridable";
        default: break;
    }
    return "";
}

std::string procDeclName(const Decl& d) {
    switch (d.kind) {
        case ASTNodeKind::SubDecl:      return static_cast<const SubDecl&>(d).name;
        case ASTNodeKind::FunctionDecl: return static_cast<const FunctionDecl&>(d).name;
        case ASTNodeKind::PropertyDecl: return static_cast<const PropertyDecl&>(d).name;
        default: break;
    }
    return std::string();
}

// 本模块有没有出现过任一虚修饰符 (2.8 的早退判据之一, 与 Inherits 子句并列)
bool moduleHasVirtualMods(const Module& m) {
    for (const auto& d : m.declarations) {
        if (d && procVirtOf(*d) != ProcVirt::None) return true;
    }
    return false;
}

// B08c: 本模块有没有声明过任一 Protected 成员。越权判定要认"接收者声明的那个类是不是
// 工程类", 而这只能读链视图 —— 所以带 Protected 的工程即便没有 Inherits 也得建表。
bool moduleHasProtectedMember(const Module& m) {
    if (!m.isClassModule) return false;   // 非类模块的 Protected 没有家族可言
    for (const auto& d : m.declarations) {
        if (d && memberAccess(*d) == AccessLevel::Protected) return true;
    }
    return false;
}

void addKeyOnce(std::vector<std::string>& keys, const std::string& k) {
    if (k.empty()) return;
    for (const auto& n : keys) {
        if (n == k) return;
    }
    keys.push_back(k);
}

// --- B08d: 类虚表槽的读法 (见 Driver::buildVirtualSlotTables 与 class_chain_registry.hpp) ---

// 虚表字段名 = 槽键的 C 写法。属性键自带方向 (get_/put_/putref_), C 名用 prop_get_/
// prop_let_/prop_set_ 前缀 (与 makePropertySignature、inheritedStubSig 同一口径); 一律小写,
// 因为槽键本身就是小写 —— 类型、填表、调用点三处都按它拼, 再没有各写一套大小写的机会。
std::string cvtblFieldOf(const std::string& slotKey) {
    if (slotKey.compare(0, 4, "get_") == 0)    return "prop_get_" + slotKey.substr(4);
    if (slotKey.compare(0, 4, "put_") == 0)    return "prop_let_" + slotKey.substr(4);
    if (slotKey.compare(0, 7, "putref_") == 0) return "prop_set_" + slotKey.substr(7);
    return slotKey;
}

// 本类该槽的入口声明: 先看本类自己的声明 (遮蔽或 Overrides), 再看 stage 3.4 判定要发转发桩
// 的那批祖先声明。**都没有 = 这个槽在本类面上没有可调用入口** → v1 判死, 不静默少发一张表。
Decl* virtFaceDecl(const ClassChainView& v, const std::string& slotKey) {
    if (!v.mod || slotKey.empty()) return nullptr;
    for (const auto& d : v.mod->declarations) {
        // 只有过程声明能带虚修饰符 → 槽键非空即 Sub/Function/Property (ifaceSlotKey 对其余返回空)
        if (d && ifaceSlotKey(*d) == slotKey) return d.get();
    }
    for (const auto& ip : v.inhProcs) {
        if (ip.decl && ifaceSlotKey(*ip.decl) == slotKey) return ip.decl;
    }
    return nullptr;
}

// --- Pass D 的 v1 边界判据 (只看 AST, 不依赖语义层) ---

bool hasEventDecl(const Module& m) {
    for (const auto& d : m.declarations) {
        if (d && d->kind == ASTNodeKind::EventDecl) return true;
    }
    return false;
}

// 基类的过程名里有没有重载组 ($ov$ 指纹方案)。无重载的工程恒 false → 判据本身零影响。
// 注意属性的三个方向: 同名 Property Get/Let/Set 是**一个**成员 (三张节点、一个键),
// 不能按声明节点数判重载 —— 否则每个带读写属性的类都被误判。
bool hasOverloadedProcs(const Module& m) {
    std::vector<std::string> seen;
    std::vector<std::string> propSeen;
    for (const auto& d : m.declarations) {
        if (!d) continue;
        std::string name;
        switch (d->kind) {
            case ASTNodeKind::SubDecl:      name = static_cast<const SubDecl&>(*d).name; break;
            case ASTNodeKind::FunctionDecl: name = static_cast<const FunctionDecl&>(*d).name; break;
            case ASTNodeKind::PropertyDecl: name = static_cast<const PropertyDecl&>(*d).name; break;
            default: continue;
        }
        const std::string key = ifaceLower(name);
        if (key.empty()) continue;
        if (d->kind == ASTNodeKind::PropertyDecl) {
            if (std::find(propSeen.begin(), propSeen.end(), key) != propSeen.end()) continue;
            propSeen.push_back(key);
        }
        if (std::find(seen.begin(), seen.end(), key) != seen.end()) return true;
        seen.push_back(key);
    }
    return false;
}

} // namespace

// B08b E0: 虚修饰符的**位置**合法性。只读 modules_, 不依赖类链登记表 —— 这样"工程里没有一条
// Inherits 但写了 Overridable"也能查, 而且不必为它建登记表 (建了就等于给无 Inherits 的工程
// 新开一条后端路径, 破零回归护栏)。
void Driver::checkVirtualPlacement() {
    for (auto& mod : modules_) {
        if (!mod) continue;
        for (const auto& d : mod->declarations) {
            if (!d) continue;
            const ProcVirt pv = procVirtOf(*d);
            if (pv == ProcVirt::None) continue;
            const std::string vt = procVirtText(pv);
            const std::string dn = procDeclName(*d);
            if (!mod->isClassModule) {
                diag_->error(DiagnosticID::SemVirtualNotSupported, d->loc,
                    "'" + vt + "' member '" + dn + "' is only allowed in a class module (.cls):"
                    " standard modules have no derived domain to override into");
            } else if (!mod->classTypeParams.empty()) {
                // 与 Inherits 同一理由: 特化克隆不携带 virt 位, 收了就是静默丢语义
                diag_->error(DiagnosticID::SemVirtualNotSupported, d->loc,
                    "'" + vt + "' is not allowed inside a generic class template ('" +
                    mod->moduleName + "')");
            } else if (mod->isInterfaceModule) {
                diag_->error(DiagnosticID::SemVirtualNotSupported, d->loc,
                    "'" + vt + "' is not allowed in an interface host module ('" +
                    mod->moduleName + "'): a contract block has no implementation to dispatch");
            }
            if (pv == ProcVirt::Overrides && mod->inherits.empty()) {
                diag_->error(DiagnosticID::SemOverrideTargetUnknown, d->loc,
                    "Overrides member '" + dn + "' has no base class to override (class '" +
                    mod->moduleName + "' has no Inherits clause)");
            }
        }
    }
}

// B08b E1/E2/E3: 覆盖契约 + dynamicKeys 汇总。需要 2.8 的类链登记表, 所以只在有 Inherits 时跑。
//
// 契约口径 (与 Interface 的 D16 同源): 签名按 interface_sig.hpp 的**源码签名**比, 不比归一后的
// Vb6Type —— 跨模块 Enum/UDT 在这一步还没注入。
//
// v1 边界 (SemVirtualNotSupported): 今天 B07b 的发码是**静态绑定** —— 派生类自己声明的成员抢键,
// 所以 d.M() 直调已经天然是派生实现; 但"基类体内调 Me.M"仍会绑到基类实现, 那是假虚派发。
// 类虚表在 B08d, 所以这里把 dynamicKeys 算出来交给语义层, 让那些调用点**报错**而不是静默绑错。
void Driver::runVirtualContractChecks() {
    // 一遍: 先把"哪个祖先声明过某成员的虚位"记下来 (供遮蔽裁决与诊断)
    for (const std::string& key : classOrder_) {
        auto self = classes_.find(key);
        if (self == classes_.end()) continue;
        ClassChainView& v = self->second;
        v.dynamicKeys.clear();
    }

    for (const std::string& key : classOrder_) {
        auto self = classes_.find(key);
        if (self == classes_.end()) continue;
        ClassChainView& v = self->second;
        if (v.chainBroken || v.chain.size() < 2) continue;  // 无祖先: 位置检查已给说法
        bool logged = false;  // 一个类只报第一条契约错 (D18-3: 级联文本没有信息量)
        auto reject = [&](SourceLocation loc, DiagnosticID id, const std::string& msg) {
            diag_->error(id, loc, msg);
            logged = true;
        };

        for (const auto& d : v.mod->declarations) {
            if (!d || logged) continue;
            const ProcVirt pv = procVirtOf(*d);
            if (pv != ProcVirt::Overrides) continue;
            const std::string slot = ifaceSlotKey(*d);
            const std::string dn = procDeclName(*d);

            // 自近到远找**声明过这个槽**的祖先 (没声明过的中间类继续往上走: 它只是继承了那条槽)
            const Decl* target = nullptr;
            const ClassChainView* tv = nullptr;
            for (size_t i = v.chain.size(); i-- > 1;) {
                auto it = classes_.find(v.chain[i - 1]);
                if (it == classes_.end() || !it->second.mod) continue;
                for (const auto& ad : it->second.mod->declarations) {
                    if (!ad || ad->kind != d->kind) continue;  // Sub 只覆盖 Sub, 属性按方向配
                    if (slot != ifaceSlotKey(*ad)) continue;
                    target = ad.get();
                    tv = &it->second;
                    break;
                }
                if (target) break;
            }
            if (!target) {
                reject(d->loc, DiagnosticID::SemOverrideTargetUnknown,
                    "Overrides member '" + dn + "' of class '" + v.name +
                    "' has no matching member in the inherited class chain");
                continue;
            }
            const ProcVirt pv2 = procVirtOf(*target);
            if (memberAccess(*target) == AccessLevel::Private) {
                reject(d->loc, DiagnosticID::SemOverrideNotOverridable,
                    "Overrides member '" + dn + "' cannot replace the Private member declared by"
                    " class '" + tv->name + "' (Private members are not visible to derived classes)");
                continue;
            }
            if (pv2 != ProcVirt::Overridable && pv2 != ProcVirt::Overrides) {
                reject(d->loc, DiagnosticID::SemOverrideNotOverridable,
                    "Overrides member '" + dn + "' targets member '" + procDeclName(*target) +
                    "' of class '" + tv->name + "', which is not declared Overridable");
                continue;
            }
            IfaceProcSig mine, base;
            if (!ifaceSigFromDecl(*d, mine) || !ifaceSigFromDecl(*target, base) ||
                !ifaceSigEqual(mine, base)) {
                reject(d->loc, DiagnosticID::SemOverrideSignatureMismatch,
                    "Overrides member '" + dn + "' signature (" + mine.text +
                    ") does not match the Overridable member of class '" + tv->name + "' (" +
                    base.text + ")");
                continue;
            }
            // 契约通过 → 本类与全部祖先的体内调用都要走虚槽 (B08d 前由语义层拒绝)
            const std::string nk = ifaceLower(ifaceDeclName(*d));
            for (size_t i = 0; i + 1 < v.chain.size(); i++) {
                auto it = classes_.find(v.chain[i]);
                if (it == classes_.end()) continue;
                addKeyOnce(it->second.dynamicKeys, nk);
            }
        }
    }
}

// v1 深度上限: 链上类数 (含自身)。VB6/tB 都没规定上限, 但转发桩数与链长成正比, 且环检测
// 之外的病态输入要有个兜底 → 取一个远超真实代码的值。
static const size_t kMaxInheritsDepth = 16;

bool Driver::runClassChainPrepass() {
    classes_.clear();
    classOrder_.clear();

    // 早退: 工程里既没有一条 Inherits、也没有一个虚修饰符 → 不建表、不发诊断, 生成物与
    // B07a 之前逐字节相同 (零回归护栏的按 feature 门禁口径, D24 末两条)。
    bool anyClause = false;
    bool anyVirtual = false;
    bool anyProtected = false;   // tB B08c
    for (auto& mod : modules_) {
        if (!mod) continue;
        if (!mod->inherits.empty()) anyClause = true;
        if (!anyVirtual && moduleHasVirtualMods(*mod)) anyVirtual = true;
        if (!anyProtected && moduleHasProtectedMember(*mod)) anyProtected = true;
        if (anyClause && anyVirtual && anyProtected) break;
    }
    if (!anyClause && !anyVirtual && !anyProtected) return true;

    // B08b E0 只要 modules_ → 放在建表之前, 且**不**因为"没有 Inherits"而跳过
    checkVirtualPlacement();
    if (!anyClause && !anyProtected) return !diag_->hasErrors();
    // 只有 Protected、一条 Inherits 都没有的工程也往下走 Pass A/C: 越权判定需要视图。
    // 这些视图全是单元素链, 后端 classChainOf() 的 `chain.size() < 2` 守卫照旧返回
    // nullptr, 成员合并 (3.4) 与槽表 (3.4b) 也都在同一个守卫上空转 —— 也就是这条只喂
    // 语义层, 不给存量工程开任何新的发码路径 (D34-2)。

    // --- Pass A: 登记可继承的工程类 + 子句侧边界拒绝 ---
    for (auto& mod : modules_) {
        if (!mod) continue;
        const std::string key = ifaceLower(mod->moduleName);
        // 子句侧的边界先查 (与"这个类能否当基类"无关, 写了就要给说法):
        //   泛型模板类内的 Inherits 会随特化克隆被复制一份, 与 IfaceRegistry 当年同一坑;
        //   非类模块 (.bas/.frm) 没有类实例语义。
        if (!mod->inherits.empty() && key.empty()) {
            diag_->error(DiagnosticID::SemInheritsNotSupported, mod->inherits[0].loc,
                "Inherits clause has no host class name to attach to");
        } else if (!mod->inherits.empty()) {
            if (!mod->classTypeParams.empty()) {
                diag_->error(DiagnosticID::SemInheritsNotSupported, mod->inherits[0].loc,
                    "Inherits is not allowed inside a generic class template (" +
                    mod->moduleName + ")");
            } else if (!mod->isClassModule) {
                diag_->error(DiagnosticID::SemInheritsNotSupported, mod->inherits[0].loc,
                    "Inherits is only allowed in a class module (.cls); '" + mod->moduleName +
                    "' is not a class module");
            }
        }
        if (!mod->isClassModule || !mod->classTypeParams.empty() || key.empty()) continue;
        // 接口宿主 (.cls 里那一个同名 Interface 块) 没有实例, 不能当基类, 也不登记.
        if (mod->isInterfaceModule) continue;

        auto res = classes_.emplace(key, ClassChainView{});
        if (!res.second) {
            diag_->error(DiagnosticID::SemDuplicateDeclaration, mod->loc,
                "Duplicate class name '" + mod->moduleName + "' in the project");
            continue;
        }
        ClassChainView& v = res.first->second;
        v.name = mod->moduleName;
        v.mod = mod.get();
        classOrder_.push_back(key);

        if (mod->inherits.size() > 1) {
            // v1 = 单继承: 一条子句里写 `Inherits A, B` 在 parse 期就撞 VB2003, 这里是
            // "分两行各写一条" 的形态。取第一条继续, 免得基名解析全线级联。
            diag_->error(DiagnosticID::SemInheritsNotSupported, mod->inherits[1].loc,
                "Class '" + mod->moduleName + "' has more than one Inherits clause ("
                "single inheritance only; VB6/tB have no multiple class inheritance)");
        }
        if (!mod->inherits.empty() && !mod->inherits[0].baseName.empty()) {
            v.clause = &mod->inherits[0];
            v.baseText = mod->inherits[0].baseName;
        }
    }

    // --- Pass B: 基名解析 (未知基) + 环检测 ---
    for (const std::string& key : classOrder_) {
        auto self = classes_.find(key);
        if (self == classes_.end() || !self->second.clause) continue;
        ClassChainView& v = self->second;

        std::string bk = ifaceLower(v.baseText);
        if (classes_.find(bk) == classes_.end()) {
            // 点号限定名 (`Project.Base`) 的末段兜底, 与发码侧 ivLastSegment 同思路。
            size_t dot = bk.rfind('.');
            if (dot != std::string::npos) {
                std::string tail = bk.substr(dot + 1);
                if (classes_.find(tail) != classes_.end()) bk = tail;
            }
        }
        if (classes_.find(bk) == classes_.end()) {
            // 基名是个 CoClass 块: 挡下的动作一样 (它不是类, 没有成员表可继承), 但理由得说对
            // —— 报"unknown base class"会让人去查拼写, 而真正的问题是拿组合同的块当基类
            // (ai/022 D49 留给本批的文案活; 026 五-6 第一条)。号保持 VB3020, 既有断言不动。
            // 折算记录除外 (D52-3): 它的名字**恒等于**一个类模块名, 走到这一支说明那个模块
            // 本来就没登记成基类 (接口宿主/泛型) —— 那句"是个 CoClass 块"是凭空捏造的理由。
            auto ccIt = coclassIds_.find(bk);
            if (ccIt != coclassIds_.end() && !ccIt->second.legacyFolded) {
                diag_->error(DiagnosticID::SemInheritsUnknownBase, v.clause->loc,
                    "Class '" + v.name + "' inherits '" + v.baseText + "', which is a CoClass "
                    "block: a coclass groups contracts and owns no member table to inherit "
                    "(Inherits takes a class module; bind to the coclass' [Implementation] "
                    "class instead)");
                v.chainBroken = true;
                continue;
            }
            diag_->error(DiagnosticID::SemInheritsUnknownBase, v.clause->loc,
                "Class '" + v.name + "' inherits unknown base class '" + v.baseText +
                "' (the target must be a class module in this project)");
            v.chainBroken = true;
            continue;
        }
        v.baseKey = bk;

        std::set<std::string> seen;
        seen.insert(key);
        std::string chain = v.name;
        for (std::string cur = bk; !cur.empty();) {
            auto it = classes_.find(cur);
            if (it == classes_.end()) break;
            chain += " -> " + it->second.name;
            if (seen.count(cur)) {
                diag_->error(DiagnosticID::SemCircularDependency, v.clause->loc,
                    "Circular Inherits chain on class '" + v.name + "': " + chain);
                v.chainBroken = true;
                break;
            }
            seen.insert(cur);
            cur = it->second.baseKey;
        }
    }

    // --- Pass C: 链展开 (父先己后) + 深度上限 ---
    for (const std::string& key : classOrder_) {
        auto self = classes_.find(key);
        if (self == classes_.end()) continue;
        ClassChainView& v = self->second;
        if (v.chainBroken) continue;

        std::set<std::string> guard;  // Pass B 已查过环, 这里只兜底防无限循环
        for (std::string cur = key; !cur.empty();) {
            auto it = classes_.find(cur);
            if (it == classes_.end()) break;
            if (it->second.chainBroken) { v.chainBroken = true; break; }
            if (!guard.insert(cur).second) break;
            v.chain.insert(v.chain.begin(), cur);  // 小写键; 成员合并按此序取声明
            cur = it->second.baseKey;
        }
        if (v.chainBroken) { v.chain.clear(); continue; }
        if (v.chain.size() > kMaxInheritsDepth) {
            std::string text;
            for (const std::string& k : v.chain) {
                auto it = classes_.find(k);
                text += (text.empty() ? "" : " -> ") + (it == classes_.end() ? k : it->second.name);
            }
            diag_->error(DiagnosticID::SemInheritsTooDeep, v.clause->loc,
                "Class '" + v.name + "' Inherits chain is deeper than the supported limit (" +
                std::to_string(kMaxInheritsDepth) + " classes): " + text);
            v.chainBroken = true;
            v.chain.clear();
        }
    }

    // --- Pass D: B07b 的 v1 边界 ---
    // 判死在这里、而不是让发码层悄悄降级的理由: 四条全都源于"派生实例必须能被基类方法按
    // 基类布局使用" (D24③ + D19 的偏移 0 不变式), 任一条件不满足时**生成的代码是错的**:
    //   ① 事件: `events` sink 指针挂在结构体尾部, 派生类一连字段就与基类偏移不一致;
    //   ② 新式 Interface: __refcount/__iv_<I> 跟着复制 = 一个对象两份计数, 撞 D21-1 单门禁;
    //   ③ 重载成员: 转发桩要带 $ov$ 变体后缀, 与 B08 的 Overridable/去虚化一起设计才不返工;
    //   ④ 同名字段: C 结构体容不下两个同名成员, 而基类方法按基类偏移读它 → 不能静默择一。
    // 一个类只报第一条 (与 D18-3 的宿主违规同口径: 级联文本没有信息量)。
    for (const std::string& key : classOrder_) {
        auto self = classes_.find(key);
        if (self == classes_.end()) continue;
        ClassChainView& v = self->second;
        if (v.chainBroken || v.chain.size() < 2) continue;  // 没有祖先 = 没有 v1 边界问题
        const SourceLocation loc = v.clause ? v.clause->loc : SourceLocation{};
        bool blocked = false;
        auto reject = [&](const std::string& what) {
            diag_->error(DiagnosticID::SemInheritsNotSupported, loc,
                "Class '" + v.name + "' cannot use Inherits in this build: " + what +
                " (v1 supports only event-free, new-style-interface-free, overload-free bases"
                " with no inherited field of the same name)");
            v.chainBroken = true;
            v.chain.clear();
            blocked = true;
        };
        // 这个模块是否实现**新式** Interface (命中 2.7 登记表)。判据与发码侧
        // ivImplementedIfaces 一致, 但那边是 CCodeGen 成员 → 各写一次而不是抽公共头。
        auto implNewStyle = [this](const Module& m) {
            for (const auto& impl : m.implements) {
                if (!impl) continue;
                const std::string lk = ifaceLower(impl->interfaceName);
                if (ifaces_.count(lk)) return true;
                const size_t dot = lk.rfind('.');
                if (dot != std::string::npos && ifaces_.count(lk.substr(dot + 1))) return true;
            }
            return false;
        };
        if (implNewStyle(*v.mod)) {
            reject("the derived class itself implements a new-style Interface");
            continue;
        }
        std::vector<std::string> selfFieldKeys;
        for (const auto& p : collectOwnMembers(*v.mod).fields) selfFieldKeys.push_back(p.first);
        for (size_t i = 0; i + 1 < v.chain.size() && !blocked; i++) {
            auto it = classes_.find(v.chain[i]);
            if (it == classes_.end() || !it->second.mod) continue;
            Module& am = *it->second.mod;
            const std::string an = it->second.name;
            if (hasEventDecl(am)) {
                reject("base class '" + an + "' declares an Event");
                break;
            }
            if (implNewStyle(am)) {
                reject("base class '" + an + "' implements a new-style Interface");
                break;
            }
            if (hasOverloadedProcs(am)) {
                reject("base class '" + an + "' has overloaded members");
                break;
            }
            for (const auto& p : collectOwnMembers(am).fields) {
                if (std::find(selfFieldKeys.begin(), selfFieldKeys.end(), p.first) ==
                    selfFieldKeys.end()) continue;
                std::string orig = p.first;
                if (p.second && p.second->kind == ASTNodeKind::VariableDecl) {
                    orig = static_cast<const VariableDecl&>(*p.second).name;
                }
                reject("class '" + v.name + "' redeclares inherited field '" + orig +
                       "' declared by base class '" + an + "'");
                break;
            }
        }
    }

    // --- Pass E (B08b): 覆盖契约 + dynamicKeys (需要上面解好的 chain, 所以排在 Pass D 之后) ---
    runVirtualContractChecks();

    return !diag_->hasErrors();
}

// ============================================================
// stage 3.4 (B07b): 继承成员合并 —— 把祖先"自己声明"的成员并进派生类的 Class 符号
//
// 为什么做在符号表而不是发码层 (D26 的核心收益): resolveClassMemberCall 的 Fix 014 兜底、
// findClassMemberCallParams、getClassMethodReturnType、canonicalClassMemberName 全部读
// Class 符号的成员表 —— 并表之后这些消费点一行都不用改。
//
// v1 并 8 张表: memberNames / memberProcKinds / memberReturnTypes / memberParams /
// memberLetParams / memberSetParams (过程面) + memberFieldNames / memberFieldTypes (字段面)。
// 刻意不并: publicFieldNames + memberFieldDispids (Fix 099 的 COM 暴露清单会去 dll_entry 找
// "祖先 Public 字段"的读写访问器, 而派生 TU 不发那份访问器 → 链接期才炸; 归 P6/B13)、
// eventNames (带事件的基类已在 2.8 判死)。字段不进 memberNames 是 Fix 099 的硬规定: 那张表
// 用来判定"成员访问是否为属性调用", 并字段会把 `Foo(obj.Field)` 改写成 prop_get_ 调用。
bool Driver::mergeInheritedMembers() {
    if (classes_.empty()) return true;  // 工程无 Inherits (2.8 早退) → 一行都不做

    // modules_ 与 analyzers_ 在 runSemanticAnalysis 里同序构建, 但 3.5b 的泛型 fixpoint 会
    // 追加模块 → 按指针找下标, 不假设两个向量永远平行。
    auto classSymOf = [&](const Module* mod, const std::string& name) -> Symbol* {
        for (size_t i = 0; i < modules_.size(); i++) {
            if (modules_[i].get() != mod) continue;
            if (i >= analyzers_.size()) return nullptr;
            Symbol* s = analyzers_[i]->symbolTable().lookupModule(name);
            return (s && s->kind == SymbolKind::Class) ? s : nullptr;
        }
        return nullptr;
    };

    for (const std::string& key : classOrder_) {
        auto self = classes_.find(key);
        if (self == classes_.end()) continue;
        ClassChainView& v = self->second;
        v.inhFields.clear();
        v.inhProcs.clear();
        if (v.chainBroken || v.chain.size() < 2) continue;  // 无祖先 = 无合并
        Symbol* dst = classSymOf(v.mod, v.name);
        if (!dst) continue;
        SymbolPayload& dstP = dst->W();   // 载荷 COW: 合并写本类自己的那份

        // 祖先序列: 自根到叶 (chain 末位是自身, 不含)
        std::vector<const ClassChainView*> anc;
        std::vector<OwnMemberKeys> own;
        for (size_t i = 0; i + 1 < v.chain.size(); i++) {
            auto it = classes_.find(v.chain[i]);
            if (it == classes_.end() || !it->second.mod) continue;
            anc.push_back(&it->second);
            own.push_back(collectOwnMembers(*it->second.mod));
        }
        if (anc.empty()) continue;

        // --- ① 过程面: 近者 (叶方向) 抢键, 落表按根→叶序 (memberNames 次序 = 声明序稳定序) ---
        // 一个成员键可能对应多个方向节点 (Property Get/Let/Set) → 表按键并一次、桩按节点发多份。
        struct Pick {
            std::string key;
            size_t ai;
            Decl* node;
        };
        std::set<std::string> taken;  // 已被更近的类占用的键
        for (const auto& p : collectOwnMembers(*v.mod).procs) taken.insert(p.first);  // 子优先
        std::vector<Pick> picks;
        for (size_t i = anc.size(); i-- > 0;) {  // 叶 → 根
            std::set<std::string> mine;          // 本层抢到的键 (层内多个方向都放行, 层间才是遮蔽)
            for (const auto& p : own[i].procs) {
                if (taken.count(p.first) || !p.second) continue;
                if (memberAccess(*p.second) == AccessLevel::Private) continue;  // 私有不进可见面
                mine.insert(p.first);
                picks.push_back({p.first, i, p.second});
            }
            taken.insert(mine.begin(), mine.end());
        }
        std::stable_sort(picks.begin(), picks.end(),
            [](const Pick& a, const Pick& b) { return a.ai < b.ai; });

        std::set<std::string> mergedKeys;  // 本类已并入过成员表的键
        for (const auto& w : picks) {
            const ClassChainView& av = *anc[w.ai];
            Symbol* src = classSymOf(av.mod, av.name);
            const std::string& k = w.key;
            bool firstTime = mergedKeys.insert(k).second;
            if (firstTime && src) {
                if (auto it = src->P().memberProcKinds.find(k); it != src->P().memberProcKinds.end()) {
                    dstP.memberProcKinds.emplace(k, it->second);
                }
                if (auto it = src->P().memberAccessLevels.find(k); it != src->P().memberAccessLevels.end()) {
                    dstP.memberAccessLevels.emplace(k, it->second);  // tB B08a
                }
                if (auto it = src->P().memberReturnTypes.find(k); it != src->P().memberReturnTypes.end()) {
                    dstP.memberReturnTypes.emplace(k, it->second);
                }
                if (auto it = src->P().memberParams.find(k); it != src->P().memberParams.end()) {
                    dstP.memberParams.emplace(k, it->second);
                }
                if (auto it = src->P().memberLetParams.find(k); it != src->P().memberLetParams.end()) {
                    dstP.memberLetParams.emplace(k, it->second);
                }
                if (auto it = src->P().memberSetParams.find(k); it != src->P().memberSetParams.end()) {
                    dstP.memberSetParams.emplace(k, it->second);
                }
                // memberNames 存祖先写的**原名** (大小写); 查表侧一律走小写键
                for (const auto& n : src->P().memberNames) {
                    if (ifaceLower(n) != k) continue;
                    if (!keyIn(dstP.memberNames, k)) dstP.memberNames.push_back(n);
                    break;
                }
            }
            v.inhProcs.push_back({w.node, av.mod});
        }

        // --- ② 字段面: 全部祖先自有字段, 根→叶串成结构体前缀 ---
        // 不按可见性过滤这份清单: 派生类**访问**不到祖先私有字段 (上面并表按可见性裁决过),
        // 但结构体必须把它们**复制**进前缀 —— 基类方法拿到的是 (vb6_cls_<Base>*)me,
        // 少一个字段全体错位 (D24③)。同名字段冲突已在 2.8 Pass D 判死, 这里无需再裁决。
        for (size_t i = 0; i < anc.size(); i++) {
            const ClassChainView& av = *anc[i];
            Symbol* src = classSymOf(av.mod, av.name);
            for (const auto& p : own[i].fields) {
                if (p.second) v.inhFields.push_back(p.second);
                if (!src) continue;
                const std::string& k = p.first;
                if (auto it = src->P().memberFieldNames.find(k); it != src->P().memberFieldNames.end()) {
                    dstP.memberFieldNames.emplace(k, it->second);
                }
                if (auto it = src->P().memberFieldTypes.find(k); it != src->P().memberFieldTypes.end()) {
                    dstP.memberFieldTypes.emplace(k, it->second);
                }
                if (auto it = src->P().memberAccessLevels.find(k); it != src->P().memberAccessLevels.end()) {
                    dstP.memberAccessLevels.emplace(k, it->second);  // tB B08a
                }
            }
        }
    }

    return !diag_->hasErrors();
}

// ============================================================
// stage 3.4b (B08d): 类虚表槽表 —— 把"链上被覆盖过的可覆盖成员"排成每个类的有序槽清单
//
// 为什么另起一张**按槽键**的表 (B08b 的 dynamicKeys 不够用): 那份只有成员名小写, 属性方向
// (Get/Let/Set) 在分析器里拿不到; 发码要的是"第几槽叫什么、签名取自哪个声明、本类的入口是哪个"。
//
// 定序 = 首次声明位置 根→叶、每槽键一份; 筛选集 = 链**根**的 dynamicKeys (祖先的键集必然包含
// 后代的键集, 因为标记是从每个最深类向上打的) → 同一条链上任何祖先的槽表都是更深层那张表的前缀,
// `__cvtbl` 这一字段的偏移在整条链上一致 —— 不满足这条就会静默读错字段 (D19 同一族的坑)。
//
// 排在 3.4 之后: 判"本类有没有这个槽的入口"要读 inhProcs (3.4 的产物), 在 2.8 判会漏掉遮蔽。
bool Driver::buildVirtualSlotTables() {
    if (classes_.empty()) return true;  // 工程无 Inherits (2.8 早退) → 一行都不做 (护栏)

    for (auto& kv : classes_) kv.second.virtSlots.clear();

    for (const std::string& key : classOrder_) {
        auto self = classes_.find(key);
        if (self == classes_.end()) continue;
        ClassChainView& v = self->second;
        // 注意**不能**用 chain.size() < 2 早退: 链根 (谁都不继承的 Overridable 声明者) 的 chain
        // 就只有它自己一个元素, 而它正是最需要视图的那一个 —— 判据换成下面那条 dyn.empty()。
        if (v.chainBroken) continue;
        auto rootIt = classes_.find(v.chain.empty() ? key : v.chain[0]);
        if (rootIt == classes_.end()) continue;
        const std::vector<std::string>& dyn = rootIt->second.dynamicKeys;
        if (dyn.empty()) continue;  // 整条链没人 Overrides → 全员无表 (零新语法零改动)

        bool logged = false;  // 一个类只报第一条 (D18-3: 级联文本没有信息量)
        for (size_t ci = 0; ci < v.chain.size() && !logged; ci++) {
            auto cit = classes_.find(v.chain[ci]);
            if (cit == classes_.end() || !cit->second.mod) continue;
            Module* cm = cit->second.mod;
            for (const auto& d : cm->declarations) {
                if (!d || logged) continue;
                if (procVirtOf(*d) != ProcVirt::Overridable) continue;
                const std::string nk = ifaceLower(ifaceDeclName(*d));
                if (!keyIn(dyn, nk)) continue;  // 链上没人覆盖这条槽 → 静态绑定就是对的
                const std::string sk = ifaceSlotKey(*d);
                bool seen = false;
                for (const auto& s : v.virtSlots) {
                    if (s.slotKey == sk) { seen = true; break; }
                }
                if (seen) continue;  // 更靠根的类已声明过这条槽 (首份胜出; 签名一致由 E3 保证)

                // v1 边界: 只有**读上下文**的槽 (Sub/Function/Property Get) 能派发。写上下文
                // (`Me.Props(1) = v`) 的发码在 tryRewriteCOMLvalue 那条路上, 本批不碰 → 与其让它
                // 静默绑成基类实现 (D27-13 判例), 不如在契约期判死。
                if (d->kind == ASTNodeKind::PropertyDecl &&
                    static_cast<const PropertyDecl&>(*d).propKind != ProcKind::PropertyGet) {
                    diag_->error(DiagnosticID::SemVirtualNotSupported, d->loc,
                        "Overridable member '" + procDeclName(*d) + "' of class '" + cit->second.name +
                        "' is a Property Let/Set, which this build cannot dispatch (only Sub/Function"
                        "/Property Get slots are emitted; ai/022 B08d)");
                    logged = true;
                    break;
                }
                Decl* impl = virtFaceDecl(v, sk);
                if (!impl) {
                    // 本类面上没有这个槽的入口: 祖先那份是 Private (3.4 不发桩)、或本类用**另一个
                    // 方向**的同名成员遮蔽了它。两种都会让表项指向一个不存在的函数 → 判死。
                    diag_->error(DiagnosticID::SemVirtualNotSupported, d->loc,
                        "Overridable member '" + procDeclName(*d) + "' has no implementation visible to"
                        " class '" + v.name + "' (a Private base member, or a same-named member of"
                        " another direction, leaves no callable entry); v1 cannot emit its class"
                        " vtable slot (ai/022 B08d)");
                    logged = true;
                    break;
                }
                v.virtSlots.push_back({sk, nk, cvtblFieldOf(sk), d.get(), cm, impl});
            }
        }
    }

    return !diag_->hasErrors();
}

} // namespace vb6c3
