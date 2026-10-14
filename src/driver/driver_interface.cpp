// vb6c3 - Interface 契约登记 (tB 扩展; ai/022 D2/D4, 批次 B02)
//
// stage 2.7: 在所有 Interface 块解析完、语义分析开始**之前**, 把工程级接口契约建成
// 只读登记表 IfaceRegistry (名字 → 展平槽表)。必须早于 stage 3: 跨模块引用接口名是
// 常态, 而每个模块各有自己的符号表, 因此登记表挂在 Driver 上而不是任何符号表里。
//
// 本批只做"编译期契约"侧: 建表 + Extends 链求解 + 槽序展平 + 冲突诊断。
// Implements 的严格比对在语义层 (semantic_analyzer_iface.cpp), 发码在 B04。

#include "driver/driver.hpp"

#include "ast/ast.hpp"
#include "driver/coclass_activate.hpp"   // stage 2.7 Pass G: 组内名字激活 (ai/022 D54, B11/C05)
#include "semantics/interface_sig.hpp"
#include "semantics/interfaces_registry.hpp"
#include "semantics/coclass_identity.hpp"   // CoClass 身份 (tB, B11/C02)

#include <iostream>
#include <map>
#include <set>
#include <string>
#include <unordered_map>
#include <vector>

namespace vb6c3 {

bool Driver::runInterfacePrepass(const CompileOptions& options) {
    ifaces_.clear();
    ifaceOrder_.clear();

    std::set<std::string> moduleKeys;
    for (auto& mod : modules_) moduleKeys.insert(ifaceLower(mod->moduleName));

    // --- Pass A: 登记接口名 (工程级唯一) ---
    for (auto& mod : modules_) {
        // v1 边界 (D11): 泛型模板类内的 Interface 块会随特化克隆被登记两次
        // (模板模块本体 + 特化副本都在 modules_ 里), 表现为莫名其妙的重名错.
        // 与其让使用者困惑, 不如在这里明确拒绝.
        if (!mod->classTypeParams.empty() && !mod->interfaces.empty()) {
            diag_->error(DiagnosticID::SemInterfaceNotSupported, mod->interfaces.front()->loc,
                "Interface declarations are not allowed inside a generic class template (" +
                mod->moduleName + ")");
            continue;
        }
        // 头行宿主识别 (B03, D17): VB6 "一文件一接口" 写成 `IFoo.cls` + 体内唯一的
        // `Interface IFoo … End Interface`. 模块名要到 driver_frontend 才定得下来,
        // 所以识别放在这里而不是 parser 里.
        if (mod->isClassModule && mod->interfaces.size() == 1 && mod->interfaces.front() &&
            ifaceLower(mod->interfaces.front()->name) == ifaceLower(mod->moduleName)) {
            mod->isInterfaceModule = true;
            for (const auto& d : mod->declarations) {
                if (!d) continue;
                diag_->error(DiagnosticID::SemInterfaceNotSupported, d->loc,
                    "Interface host module '" + mod->moduleName +
                    "' may contain only the Interface block (declaration here is not allowed)");
                break;  // 一条宿主违规只报一次
            }
        }
        for (auto& d : mod->interfaces) {
            if (!d) continue;
            const std::string key = ifaceLower(d->name);
            if (key.empty()) continue;  // 无名 = parse 阶段已报错, 不再级联
            // 头行宿主的接口名天然等于它自己的模块名, 这不是撞车 (B03)
            const bool hostOwnName =
                mod->isInterfaceModule && d.get() == mod->interfaces.front().get();
            if (moduleKeys.count(key) && !hostOwnName) {
                diag_->error(DiagnosticID::SemDuplicateDeclaration, d->loc,
                    "Interface name '" + d->name +
                    "' collides with a module of the same name (module names and "
                    "interface names share one project-wide namespace)");
                continue;
            }
            auto res = ifaces_.emplace(key, IfaceView{});
            if (!res.second) {
                diag_->error(DiagnosticID::SemDuplicateDeclaration, d->loc,
                    "Duplicate interface name '" + d->name +
                    "' (interface names are project-unique)");
                continue;
            }
            IfaceView& v = res.first->second;
            v.name = d->name;
            v.decl = d.get();
            v.extendsKey = ifaceLower(d->extendsName);
            for (const auto& a : d->attributes) {
                if (ifaceLower(a.name) == "interfaceid" && a.hasStr) v.guid = a.strValue;
            }
            ifaceOrder_.push_back(key);
        }
    }

    // --- Pass A2: VB6 风格接口 (.cls 宿主) 登记 ---
    //
    // VB6 没有 `Interface ... End Interface` (那是 tB 扩展): 它的接口就是一个
    // `VB_Creatable = False`、成员全是**无体签名**的普通 .cls, 实现方写 `Implements I`
    // 再把每个实现成员命名为 `I_<成员>` (IScheme 就是这么写的)。这批接口过去完全不在
    // 登记表里, 连锁后果有三处:
    //   - 语义层退到 legacy 路径 (只 warn, 不逐槽校验);
    //   - 发码层 `ivLookupIface` 返回 nullptr, tB 那条路整条跳过, 退到 legacy
    //     `emitInterfaceVtable` **按实现类自己 harvest 槽** —— 各实现类槽数/槽序/`me` 类型
    //     并不一致 (实测 25/26/27), 于是每个类各发一份同名 `vb6_vtbl_I`, C2011 撞 290 次;
    //   - `As I` 引用类型没有 `vb6_ivref_I` 可落。
    // 登记之后 tB 那套 (canonical 槽表 + `vb6_ivtbl_I` + `vb6_ivref_I`) 自然接管,
    // 一个接口只发一份表, 槽序由接口自己说了算, 与实现类怎么排无关。
    //
    // 判据收紧到"只有真接口进表": 误把普通类认成接口, 会让 Implements 校验变严、动态派发
    // 行为变掉, 比不管更糟。四条都要过:
    //   1) 被某个模块 `Implements` 按名字指到 (VB6 里只有接口会被 Implements);
    //   2) 是 .cls 类模块, 且自己**没有** tB Interface 块 (有的话 Pass A 已登记);
    //   3) 显式写了 `Attribute VB_Creatable` 就必须为 False; 没写不拦, 靠第 4 条兜;
    //   4) 全部 Sub/Function/Property 成员都是**无体**签名 —— 这是 VB6 接口类的定义性特征,
    //      只要有一个带实现体就说明它是普通类, 不认。
    {
        std::set<std::string> wanted;
        for (const auto& mod : modules_) {
            for (const auto& impl : mod->implements) {
                if (impl) wanted.insert(ifaceLower(impl->interfaceName));
            }
        }
        std::unordered_map<std::string, Module*> byName;
        for (auto& mod : modules_) byName.emplace(ifaceLower(mod->moduleName), mod.get());

        // `Attribute X = False/0` 读法与本文件 Pass E 的 boolLiteral 同一口径
        auto attrLiteral = [](const Module& m, const char* want, std::string& out) -> bool {
            for (const auto& a : m.attributes) {
                if (!a || ifaceLower(a->attrName) != want) continue;
                if (!a->value || a->value->kind != ASTNodeKind::LiteralExpr) return false;
                out = static_cast<const LiteralExpr&>(*a->value).rawText;
                return true;
            }
            return false;
        };
        // 全部成员都是无体签名 => 是接口类.
        //
        // Fix 213: "空体" 的判据不能只看 `body->empty()`. VB6 接口宿主的每个成员都带
        // 过程内属性行 —— `Attribute Message.VB_Description = "..."` /
        // `Attribute X.VB_UserMemId = ...` —— 而 parser 把 Attribute 行收成
        // `AttributeStmt` 语句放进过程体 (parseAttributeInBody, parser_stmt_io.cpp:331).
        // 于是 `Public Function Message(...) … End Function` 这种**真正的无体签名**体里
        // 恰好有 1 条 AttributeStmt, `body->empty()` 为假 ⇒ 宿主被误判成普通类 ⇒
        // Pass A2 不登记 ⇒ 发码层 `ivLookupIface("ISubclass")` 恒 nullptr ⇒
        // `emitInterfaceVtable` 按 35 个实现类各发一份同名 `vb6_vtbl_ISubclass` /
        // `vb6_iface_ISubclass` ⇒ 111 个 TU × 68 份重定义 = 7548×C2011 + 111×C2079,
        // 且 35×VB3001 "Implements: interface not found" (连工程内的 ISubclass.cls 都
        // "找不到")。判据改为"体里除了 Attribute 行之外没有任何语句".
        auto bodyHasImplementation = [](const StmtList& body) {
            for (const auto& s : body) {
                if (!s) continue;
                if (s->kind == ASTNodeKind::AttributeStmt) continue;  // 纯元数据, 不是实现
                return true;
            }
            return false;
        };
        auto allMembersBodyless = [&](const Module& m) {
            for (const auto& d : m.declarations) {
                if (!d) continue;
                const StmtList* body = nullptr;
                if (d->kind == ASTNodeKind::SubDecl) {
                    body = &static_cast<const SubDecl&>(*d).body;
                } else if (d->kind == ASTNodeKind::FunctionDecl) {
                    body = &static_cast<const FunctionDecl&>(*d).body;
                } else if (d->kind == ASTNodeKind::PropertyDecl) {
                    body = &static_cast<const PropertyDecl&>(*d).body;
                } else {
                    continue;  // Type/Enum/Declare 之类不是可实现成员
                }
                if (bodyHasImplementation(*body)) return false;
            }
            return true;
        };

        for (const std::string& want : wanted) {
            if (ifaces_.count(want)) continue;   // Pass A 已登记 (tB Interface 块)
            auto it = byName.find(want);
            if (it == byName.end()) continue;    // 工程里没有同名模块
            Module* host = it->second;
            if (!host->isClassModule || !host->interfaces.empty()) continue;
            std::string creatable;
            if (attrLiteral(*host, "vb_creatable", creatable)) {
                // 显式写了 VB_Creatable 就必须为 False
                if (creatable != "False" && creatable != "false" && creatable != "0") continue;
            }
            if (!allMembersBodyless(*host)) continue;
            auto res = ifaces_.emplace(want, IfaceView{});
            if (!res.second) continue;
            IfaceView& v = res.first->second;
            v.name = host->moduleName;
            v.clsHost = host;                    // 槽成员取自宿主自己的 declarations
            ifaceOrder_.push_back(want);
        }
    }

    // --- Pass B: Extends 链求解 (未知父 / 环) ---
    for (const std::string& key : ifaceOrder_) {
        IfaceView& v = ifaces_[key];
        if (v.extendsKey.empty()) continue;
        if (!ifaces_.count(v.extendsKey)) {
            diag_->error(DiagnosticID::SemInterfaceUnknownParent, v.decl->loc,
                "Interface '" + v.name + "' extends unknown interface '" + v.decl->extendsName + "'");
            v.chainBroken = true;
            continue;
        }
        std::set<std::string> seen;
        seen.insert(key);
        std::string chain = v.name;
        std::string cur = v.extendsKey;
        while (!cur.empty()) {
            auto it = ifaces_.find(cur);
            if (it == ifaces_.end()) break;
            chain += " -> " + it->second.name;
            if (seen.count(cur)) {
                diag_->error(DiagnosticID::SemCircularDependency, v.decl->loc,
                    "Circular Extends chain on interface '" + v.name + "': " + chain);
                v.chainBroken = true;
                break;
            }
            seen.insert(cur);
            cur = it->second.extendsKey;
        }
    }

    // --- Pass C: 槽表展平 (父先己后, 同层声明序) + 槽名冲突 ---
    for (const std::string& key : ifaceOrder_) {
        IfaceView& v = ifaces_[key];
        if (v.chainBroken) continue;

        // 自根到叶的接口序列 (根在前)
        std::vector<const IfaceView*> chain;
        for (std::string cur = key; !cur.empty();) {
            auto it = ifaces_.find(cur);
            if (it == ifaces_.end()) break;
            if (it->second.chainBroken) { v.chainBroken = true; break; }
            chain.insert(chain.begin(), &it->second);
            cur = it->second.extendsKey;
        }
        if (v.chainBroken) continue;

        std::set<std::string> used;      // 已占用的槽键
        std::set<std::string> ownNames;  // 本接口内的成员名 (禁重载)
        for (const IfaceView* iv : chain) {
            const bool own = (iv == &v);
            // 槽成员来源有两种, 统一摊平成"声明序的一串 Decl"再往下走:
            //   - tB Interface 块 -> InterfaceDecl::members;
            //   - VB6 .cls 宿主  -> 该宿主模块自己的 declarations (Pass A2 登记进来的)。
            // 两者都只取签名 (ifaceSigFromDecl 拒带体的成员), 槽序都是声明序。
            std::vector<const Decl*> memberDecls;
            if (iv->clsHost) {
                for (const auto& md : iv->clsHost->declarations) {
                    if (md) memberDecls.push_back(md.get());
                }
            } else if (iv->decl) {
                for (const auto& m : iv->decl->members) {
                    if (m.decl) memberDecls.push_back(m.decl.get());
                }
            }
            for (const Decl* md : memberDecls) {
                IfaceProcSig sig;
                if (!ifaceSigFromDecl(*md, sig)) continue;
                if (own) {
                    if (!ownNames.insert(sig.slotKey).second) {
                        diag_->error(DiagnosticID::SemInterfaceSlotConflict, md->loc,
                            "Interface '" + v.name + "' declares member '" + sig.memberName +
                            "' more than once (interface members cannot be overloaded)");
                        continue;
                    }
                }
                if (!used.insert(sig.slotKey).second) {
                    std::string holder;
                    for (const auto& s : v.slots) {
                        if (s.key == sig.slotKey) { holder = s.ownerIface; break; }
                    }
                    diag_->error(DiagnosticID::SemInterfaceSlotConflict, md->loc,
                        "Interface '" + iv->name + "' slot '" + sig.slotKey +
                        "' conflicts with the inherited slot from '" + holder +
                        "' (COM vtables have no shadowing)");
                    continue;
                }
                IfaceSlotView slot;
                slot.key = sig.slotKey;
                slot.memberName = sig.memberName;
                slot.ownerIface = iv->name;
                slot.sig = md;
                slot.index = static_cast<int32_t>(v.slots.size());
                v.slots.push_back(std::move(slot));
            }
        }
    }

    // --- Pass D: 委托式实现 `Implements I Via m_holder` (ai/022 D42/D43, 批次 B10) ---
    //
    // 为什么在这里、不在语义层: 判定要看"字段类型那个类实现了接口没有", 而那些类的
    // 符号要到 stage 3.5 才注入本模块作用域 —— 只有此刻整工程的模块表看得见。
    // 裁决结果 vias_ 有两个消费方: 语义层据此免掉逐槽 VB3012 (契约由被委托对象满足),
    // 发码层据此给没有自家实现的槽转调持有对象的接口槽 (cgen_iface_vtbl.cpp)。
    vias_.clear();
    // 模块名小写 -> 模块. Pass D 与 Pass F 共用 (CoClass 的 [Implementation] 也按名字找类模块).
    std::unordered_map<std::string, Module*> byName;
    for (auto& mod : modules_) byName[ifaceLower(mod->moduleName)] = mod.get();
    {
        // 与语义层 lookupWrittenIface 同一口径: 先按写的全名, 再按点号末段 (Fix 083)
        auto resolveIface = [this](const std::string& written) -> const IfaceView* {
            auto it = ifaces_.find(ifaceLower(written));
            if (it == ifaces_.end()) {
                size_t dot = written.rfind('.');
                if (dot != std::string::npos)
                    it = ifaces_.find(ifaceLower(written.substr(dot + 1)));
            }
            return it == ifaces_.end() ? nullptr : &it->second;
        };
        auto findField = [](Module& m, const std::string& name) -> VariableDecl* {
            for (const auto& d : m.declarations) {
                if (!d || d->kind != ASTNodeKind::VariableDecl) continue;
                auto& v = static_cast<VariableDecl&>(*d);
                if (ifaceLower(v.name) == ifaceLower(name)) return &v;
            }
            return nullptr;
        };

        for (auto& mod : modules_) {
            for (const auto& impl : mod->implements) {
                if (!impl || impl->viaField.empty()) continue;  // 非委托式: 本 Pass 不管
                const std::string written = impl->interfaceName + " Via " + impl->viaField;
                auto reject = [&](DiagnosticID id, const std::string& msg) {
                    diag_->error(id, impl->loc, "'" + written + "': " + msg);
                };
                if (!mod->isClassModule) {
                    reject(DiagnosticID::SemViaTargetUnknown,
                           "Via is only allowed in a class module");
                    continue;
                }
                const IfaceView* v = resolveIface(impl->interfaceName);
                if (!v) {
                    reject(DiagnosticID::SemViaTargetUnknown,
                           "the delegated name is not an Interface block (Via is only "
                           "defined for interfaces with a checked contract)");
                    continue;
                }
                VariableDecl* fld = findField(*mod, impl->viaField);
                auto* ref = fld && fld->asType
                    ? dynamic_cast<SimpleTypeRef*>(fld->asType.get()) : nullptr;
                if (!ref) {
                    reject(DiagnosticID::SemViaTargetUnknown,
                           "holder '" + impl->viaField + "' is not a module-level field "
                           "declared As <Class> in this module");
                    continue;
                }
                Module* holder = nullptr;
                auto hit = byName.find(ifaceLower(ref->name));
                if (hit != byName.end()) holder = hit->second;
                if (!holder || !holder->isClassModule || holder->isInterfaceModule) {
                    reject(DiagnosticID::SemViaTargetUnknown,
                           "field '" + impl->viaField + "' is not of a project class type ("
                           "'" + ref->name + "' is not a class module)");
                    continue;
                }
                // v1 边界 (D43-4): 持有类必须**自己**实现同一个接口。它自己又是委托
                // (A Via f, f:B; B Via g, g:A) 的话运行期能构成无限回环, 这里直接拒。
                bool holderImplements = false;
                bool chained = false;
                for (const auto& hi : holder->implements) {
                    if (!hi || resolveIface(hi->interfaceName) != v) continue;
                    if (!hi->viaField.empty()) {
                        reject(DiagnosticID::SemViaHolderNotImplemented,
                               "class '" + holder->moduleName + "' delegates interface '" +
                               v->name + "' too (chained Via is not supported)");
                        chained = true;
                    } else {
                        holderImplements = true;
                    }
                    break;
                }
                if (!holderImplements && !chained) {
                    reject(DiagnosticID::SemViaHolderNotImplemented,
                           "class '" + holder->moduleName + "' (type of field '" +
                           impl->viaField + "') does not implement interface '" + v->name + "'");
                    continue;
                }
                ViaView vv;
                vv.ifaceKey = ifaceLower(v->name);
                vv.fieldName = fld->name;
                vv.holderModule = holder->moduleName;
                vias_[ifaceLower(mod->moduleName)].push_back(std::move(vv));
            }
        }
    }

    // --- Pass E0: 存量 header attribute 只读折算 (ai/026 六节 C04 / ai/022 D52, 批次 B11/C04) ---
    //
    // VB6 不把 COM 身份写在源码里，而写在 `.cls` 头部那几行 `Attribute VB_*` 上。本 Pass 把它们
    // 折成**同一条** CoClass 记录塞进 Module::coclasses —— 于是 Pass E 仍是全工程唯一的身份出口、
    // coclassIds_ 仍是唯一读数（D47 之后这条是硬规矩：不许另立第二套身份通道）。
    //
    // 三条按 D52 实测定的口径：
    //  1) 只折 `.cls` 类模块、且属性名**不带点**。语料实测成员级行（`Attribute m_oSocket.
    //     VB_VarHelpID` 一类）另有 40+ 种名字，混进来就是把成员元数据当身份读；`.ctl`/`.pag`
    //     虽然 `isClassModule` 也算真，但 026 六节只让折类模块，按文件名后缀挡掉。
    //  2) 折算记录**只记信息、不开判死**：Pass F 的形状/名字校验与 stage 3.4c 的契约聚合都按
    //     foldedFromAttributes() 跳过它。实测 143 条 `VB_Creatable` 里 134 条是 True、且全在
    //     EXE 工程 —— 一视同仁就是给 C03a 新立的 VB3033 送一批"改码前一声不吭"的存量工程当红。
    //  3) 手写块优先：同一模块两者都有时**不折**，只报一条信息行（026 六节"以手写块为准"）。
    //     `Instancing` 不在这里折：它的唯一真相是 `Module::instancing`（parse 期从 BEGIN 头读），
    //     复制进记录就是造两处真相。
    {
        auto boolLiteral = [](const AttributeStmt& a, bool& out) {
            if (!a.value || a.value->kind != ASTNodeKind::LiteralExpr) return false;
            const std::string& raw = static_cast<const LiteralExpr&>(*a.value).rawText;
            if (raw == "True" || raw == "true" || raw == "-1") { out = true; return true; }
            if (raw == "False" || raw == "false" || raw == "0") { out = false; return true; }
            return false;
        };
        // { 源码里的属性名, 折进记录后的属性名 } —— 只有第一枚有现成归宿（身份求解读
        // [ComCreatable]），其余三枚按原名带着：今天无人读，B13/B15 的类型库标志位从这里取。
        struct FoldKey { const char* legacy; const char* target; };
        static const FoldKey kFoldKeys[] = {
            { "VB_Creatable",       "ComCreatable"       },
            { "VB_Exposed",         "VB_Exposed"         },
            { "VB_PredeclaredId",   "VB_PredeclaredId"   },
            { "VB_GlobalNameSpace", "VB_GlobalNameSpace" },
        };

        for (auto& mod : modules_) {
            if (!mod || !mod->isClassModule || mod->isInterfaceModule) continue;
            if (!mod->classTypeParams.empty()) continue;   // 泛型模板: 这些属性行本身在 2.6 就被拒
            const std::string fname = ifaceLower(mod->filename);
            size_t dot = fname.rfind('.');
            if (dot == std::string::npos || fname.substr(dot) != ".cls") continue;

            // 每个键取首行（与 Pass E 对同名属性行"首值胜"同一条口径）
            const AttributeStmt* hit[4] = { nullptr, nullptr, nullptr, nullptr };
            for (const auto& attr : mod->attributes) {
                if (!attr || attr->attrName.find('.') != std::string::npos) continue;
                for (size_t i = 0; i < 4; i++) {
                    if (!hit[i] && ifaceLower(attr->attrName) == ifaceLower(kFoldKeys[i].legacy)) {
                        hit[i] = attr.get();
                        break;
                    }
                }
            }
            int nHit = 0;
            const AttributeStmt* first = nullptr;
            for (size_t i = 0; i < 4; i++) if (hit[i]) { nHit++; if (!first) first = hit[i]; }
            if (!nHit) continue;   // 只写 VB_Name 的模块（含本工程全部 .bas 用例）到此为止

            bool hasHandwritten = false;
            for (const auto& cc : mod->coclasses) {
                if (cc && !cc->foldedFromAttributes()) { hasHandwritten = true; break; }
            }
            if (hasHandwritten) {
                std::cerr << "C3: class '" << mod->moduleName << "' has both a CoClass block and "
                          << nHit << " legacy header attribute line(s): the block wins, the "
                          << "attributes are not folded" << std::endl;
                continue;
            }

            auto cc = std::make_unique<CoClassDecl>(first->loc, mod->moduleName);
            InterfaceAttr impl;
            impl.name = "Implementation";
            impl.strValue = mod->moduleName;   // 类自己就是那个实现
            impl.hasStr = true;
            impl.loc = cc->loc;
            cc->attributes.push_back(std::move(impl));
            for (size_t i = 0; i < 4; i++) {
                if (!hit[i]) continue;
                bool truthy = false;
                const bool gotBool = boolLiteral(*hit[i], truthy);
                InterfaceAttr a;
                a.name = kFoldKeys[i].target;
                a.loc = hit[i]->loc;
                a.hasNum = true;
                a.numValue = (gotBool && truthy) ? 1 : 0;
                cc->legacyFoldKeys.push_back(std::string(kFoldKeys[i].legacy) + "=" +
                                             (a.numValue ? "True" : "False"));
                cc->attributes.push_back(std::move(a));
            }
            mod->coclasses.push_back(std::move(cc));
        }
    }

    // --- Pass E: CoClass 身份求解 (tB 扩展, ai/026 三节 / ai/022 D46, 批次 B11/C02) ---
    // 求解本身在 src/semantics/coclass_identity.cpp 这个唯一入口里；这里只给上下文、缓存结果。
    // 报告走 stderr 的 "C3: ..." 信息行（driver_compile.cpp 的 "C3: 加载工程" 是同族先例）：
    // 诊断通道今天只在**阶段失败**时才整体打印，note 级在成功的编译里根本看不见；而 --emit-c
    // 的 stdout 是 C 文本，不能混。D44/D45 那条"每层的可观测面不一样"的第三次应验。
    coclassIds_.clear();
    {
        CoClassEnv env;
        // <Proj>: vbp 的 Name= > 工程基名 > 兜底字面量（第三条沿用 com_entry 那侧已有的兜序）
        env.project = !vbpProjectName_.empty() ? vbpProjectName_
                    : (!projectBaseName_.empty() ? projectBaseName_ : std::string("VB6EXE"));
        env.vbpClsids = &classClsidMap_;
        env.ifaces = &ifaces_;
        for (auto& mod : modules_) {
            for (auto& cc : mod->coclasses) {
                if (!cc || cc->name.empty()) continue;  // 无名块: parse 期已报错, 不再级联
                CoClassIdentity id = resolveCoClassIdentity(*cc, env);
                std::cerr << "C3: CoClass '" << id.name << "' identity: CLSID=" << id.clsid
                          << " (" << identitySourceName(id.clsidSource) << ")"
                          << " IID=" << (id.iid.empty() ? std::string("-") : id.iid)
                          << " (" << identitySourceName(id.iidSource) << ")"
                          << " ProgID=" << id.progId
                          << " (" << identitySourceName(id.progIdSource) << ")"
                          << " impl='" << id.implName << "'"
                          << " comCreatable=" << (id.comCreatable ? "True" : "False");
                if (cc->foldedFromAttributes()) {
                    // 折算来的记录在**同一行**里说清楚: 消费者是谁、折了哪几行属性。
                    std::cerr << " folded-from-legacy:";
                    for (const std::string& k : cc->legacyFoldKeys) std::cerr << ' ' << k;
                }
                std::cerr << std::endl;
                // 同名两个块: 首值胜。重复名/引用是否存在这类校验按 D44 整片归 C03。
                coclassIds_.emplace(ifaceLower(id.name), std::move(id));
            }
        }
        // 接口 IID 表 (ai/022 B13c): 同一个 <Proj>、同一组种子，块内 [Default] 那条与
        // 新式接口 vtable 用的是同一个函数 ⇒ 结构上不可能分叉。
        ifaceIds_ = buildIfaceIdMap(env.project, ifaces_);
        // 读数按源码序打（unordered_map 的桶序不稳定，会让同一工程两次编译的信息行换顺序）
        for (auto& mod : modules_) {
            for (auto& d : mod->interfaces) {
                if (!d || d->name.empty()) continue;
                auto it = ifaceIds_.find(ifaceLower(d->name));
                if (it == ifaceIds_.end()) continue;  // 父链已错的接口不发身份
                std::cerr << "C3: Interface '" << d->name << "' IID=" << it->second << std::endl;
            }
        }
    }

    // --- Pass F: CoClass 块的形状与名字校验 (tB 扩展, ai/022 D48, 批次 B11/C03a) ---
    // 放在这里（而不是语义层）的理由与 Pass D 同一条：只有此刻整工程的模块表与接口登记表同时可见。
    // **契约聚合不在这一格**：实现类到底满不满足每个条目，要等 stage 3.4 的成员合并才判得准
    // （基类实现了、派生类没重写 的情况在 2.7 看不见）→ 归 C03b，见 D48-3。
    {
        std::unordered_map<std::string, const CoClassDecl*> firstBlock;
        for (auto& mod : modules_) {
            for (auto& cc : mod->coclasses) {
                if (!cc || cc->name.empty()) continue;   // 无名块: parse 期已报错, 不再级联
                // 折算记录不参与 (D52-2): 语料里 134 条 `VB_Creatable = True` 全在 EXE 工程,
                // 让新校验打它们 = 把改码前编得过的存量工程当场打死。
                if (cc->foldedFromAttributes()) continue;
                const std::string key = ifaceLower(cc->name);

                // 1) 块名撞车（重复块名 / 撞模块名 / 撞接口名）
                if (firstBlock.count(key)) {
                    diag_->error(DiagnosticID::SemCoClassDuplicate, cc->loc,
                        "CoClass name '" + cc->name + "' is declared twice "
                        "(CoClass names are project-wide unique)");
                    continue;
                }
                if (moduleKeys.count(key) && ifaceLower(mod->moduleName) != key) {
                    // 例外（与 B03 的接口宿主同一条理由）：VB6 最自然的写法就是把
                    // `CoClass Widget` 写在 `Widget.cls` 里（块名 = 宿主模块名），这不是撞车。
                    diag_->error(DiagnosticID::SemCoClassDuplicate, cc->loc,
                        "CoClass name '" + cc->name + "' collides with a module of the same name "
                        "(module, interface and CoClass names share one project-wide namespace)");
                    continue;
                }
                if (ifaces_.count(key)) {
                    diag_->error(DiagnosticID::SemCoClassDuplicate, cc->loc,
                        "CoClass name '" + cc->name + "' collides with an Interface block of the "
                        "same name (a contract and the class that groups contracts cannot share one name)");
                    continue;
                }
                firstBlock[key] = cc.get();

                // 2) 契约条目
                std::unordered_map<std::string, int> entrySeen;
                int defaults = 0;
                for (const auto& r : cc->ifaces) {
                    const std::string rk = ifaceLower(r.ifaceName);
                    if (rk.empty()) continue;            // 无名条目: parse 期已报错
                    if (!ifaces_.count(rk)) {
                        auto cls = byName.find(rk);
                        if (cls != byName.end() && cls->second && cls->second->isClassModule) {
                            // 026 五-6 第四类: VB6 把 .cls 当接口用的旧习惯, 不是新式契约
                            diag_->error(DiagnosticID::SemCoClassEntryInvalid, r.loc,
                                "CoClass '" + cc->name + "' lists '" + r.ifaceName + "' as a contract, "
                                "but that is a class module (the legacy 'use a .cls as an interface' "
                                "habit is not a contract; declare an Interface block instead)");
                        } else {
                            diag_->error(DiagnosticID::SemCoClassEntryInvalid, r.loc,
                                "CoClass '" + cc->name + "' lists contract entry '" + r.ifaceName +
                                "' which is not an Interface block in this project");
                        }
                    }
                    if (++entrySeen[rk] > 1) {
                        diag_->error(DiagnosticID::SemCoClassEntryInvalid, r.loc,
                            "CoClass '" + cc->name + "' lists interface '" + r.ifaceName +
                            "' more than once (the contract set is a set)");
                    }
                    if (r.isDefault) defaults++;
                }
                if (defaults > 1) {
                    diag_->error(DiagnosticID::SemCoClassEntryInvalid, cc->loc,
                        "CoClass '" + cc->name + "' marks " + std::to_string(defaults) +
                        " interfaces [Default]; a coclass has exactly one default interface");
                }

                // 3) v1 边界: [Implementation] 必须指本工程的类模块; EXE 工程不能声明可注册为 COM 服务器
                auto idit = coclassIds_.find(key);
                const std::string impl = idit == coclassIds_.end() ? std::string() : idit->second.implName;
                if (!impl.empty()) {
                    auto it = byName.find(ifaceLower(impl));
                    Module* target = it == byName.end() ? nullptr : it->second;
                    if (!target || !target->isClassModule || target->isInterfaceModule) {
                        diag_->error(DiagnosticID::SemCoClassNotSupported, cc->loc,
                            "CoClass '" + cc->name + "' binds [Implementation(\"" + impl +
                            "\")] but '" + impl + "' is not a class module of this project");
                    }
                }
                if (!options.isDll && idit != coclassIds_.end() && idit->second.comCreatable) {
                    diag_->error(DiagnosticID::SemCoClassNotSupported, cc->loc,
                        "CoClass '" + cc->name + "' marks [ComCreatable(True)] in an EXE project: only "
                        "ActiveX DLL projects register a COM server (an EXE keeps the in-project half)");
                }
            }
        }
    }

    // --- Pass G: 组内名字激活 (tB 扩展, ai/026 五-3/4/5, ai/022 D54, 批次 B11/C05) ---
    // 把写在类型位置上的**块名**就地换成块绑的类名, 于是 `As <块名>`/`New <块名>`/
    // `Set x = CreateObject("<本工程 ProgID>")` 全部落到"工程类"这条已经实测通了的路上
    // (D54 探针 M4)。为什么是就地改名而不是给语义层+cgen 各开一个别名入口, 以及不改之前
    // 那三份坏读数 (`void* a = 0` + 晚绑定 DISPID 调用 / `vb6_NewObject(L"块名")` /
    // 走注册表的 CreateObject) 都记在 D54; 实现口径见 src/driver/coclass_activate.hpp 头注。
    // 三条准入 (D54-③): 折算记录不算 (它的块名恒等于一个类模块名, 抢的就是那条路的字节)、
    // 被同名模块占住的名字不算 (类/模块赢 —— `As Widget` 今天的含义不许被一块新语法改掉)、
    // 绑不上类模块的不算 (Pass F 的 VB3033 已经判死, 这里不重复报)。
    {
        std::vector<CoClassActivation> acts;
        std::vector<std::string> implLess;
        for (const auto& kv : coclassIds_) {
            const CoClassIdentity& id = kv.second;
            if (id.legacyFolded || id.name.empty()) continue;
            const std::string key = ifaceLower(id.name);
            if (byName.count(key)) continue;         // 同名模块占位: 类/模块赢
            if (id.implName.empty()) { implLess.push_back(id.name); continue; }
            auto impl = byName.find(ifaceLower(id.implName));
            if (impl == byName.end() || !impl->second || !impl->second->isClassModule) continue;
            acts.push_back(CoClassActivation{id.name, id.implName, id.progId});
        }
        activateCoClassNames(modules_, acts, implLess, *diag_);
    }

    return !diag_->hasErrors();
}

// ============================================================
// stage 3.4c: CoClass 契约聚合 (tB 扩展; ai/022 D50, 批次 B11/C03b)
// ============================================================
//
// 判"块里列出的每个接口, [Implementation] 那个类满足不满足"。为什么不在 stage 2.7
// (Pass F 就在旁边) 而要多一个阶段: 实现类**自己不写成员、由祖先提供**那份实现是合法
// 形状 (D50 探针 pd: 祖先只声明成员、不写 Implements, 派生 Inherits 它 —— 这条路今天
// 唯一被 VB3022 挡住的是"祖先或派生自己写了新式 Implements"那半)。祖先的成员表要到
// stage 2.8 的链表 + 3.4 的成员合并才看得全, 所以比对排在 3.4b 之后。
//
// 不新造比对器 (026 五-1): 槽键与签名口径全部走 interface_sig.hpp 那一份 inline 函数,
// 与语义层的 checkNewStyleInterface 共用同一套词汇, 两侧不可能对"什么叫同一槽"给出不同
// 答案。这里重复的只有"把声明表按槽键建索引"这一小段 —— 刻意不去改 B02 那条已发货的
// 路径 (它的子句记账与遮蔽规则牵动 VB3019), 回归面比省下的二十行贵。
// 诊断复用既有 VB3012/VB3017 两个号: 缺槽与签名不符的语义和 Implements 那边完全同族,
// 只是主语从"类实现了接口"换成"CoClass 绑定了接口" (D48-2)。
bool Driver::runCoClassContractCheck() {
    bool anyBlock = false;
    for (const auto& mod : modules_) {
        // 折算记录没有契约条目、也不该被契约判死 (D52-2) ⇒ 早退条件只认手写块
        for (const auto& cc : mod->coclasses) {
            if (cc && !cc->foldedFromAttributes()) { anyBlock = true; break; }
        }
        if (anyBlock) break;
    }
    if (!anyBlock) return true;   // 零新语法护栏: 工程里没有 CoClass 块就到此为止

    std::unordered_map<std::string, Module*> byName;
    for (auto& mod : modules_) byName[ifaceLower(mod->moduleName)] = mod.get();

    // 槽键还原成可读成员名 (与语义层 ifaceBareName 同一条口径)
    auto bareName = [](const std::string& slotKey) {
        for (const char* p : {"putref_", "get_", "put_"}) {
            std::string k(p);
            if (slotKey.size() > k.size() && slotKey.compare(0, k.size(), k) == 0)
                return slotKey.substr(k.size());
        }
        return slotKey;
    };
    // (实现类, 本接口) 是否委托式实现 `Implements I Via m_h`: 是则契约由持有对象满足,
    // 逐槽"未实现"不报 (B10 同一口径; 签名不符照报, 写了还对不上必然是笔误)。
    auto delegatedTo = [this](const std::string& classKey, const std::string& ifaceKey) {
        auto it = vias_.find(classKey);
        if (it == vias_.end()) return false;
        for (const ViaView& vv : it->second)
            if (vv.ifaceKey == ifaceKey) return true;
        return false;
    };

    for (auto& mod : modules_) {
        for (auto& cc : mod->coclasses) {
            if (!cc || cc->name.empty()) continue;
            auto idit = coclassIds_.find(ifaceLower(cc->name));
            // 没写 [Implementation] 的块跳过 (Pass F 今天不要求非有不可, 见 D50-4)
            if (idit == coclassIds_.end() || idit->second.implName.empty()) continue;
            auto tgt = byName.find(ifaceLower(idit->second.implName));
            Module* impl = tgt == byName.end() ? nullptr : tgt->second;
            // 指向不存在/不是类模块: Pass F 已经报过 VB3033, 这里不级联
            if (!impl || !impl->isClassModule || impl->isInterfaceModule) continue;
            // D11 v1 边界: 泛型类不实现接口, 语义层已就该类报过错
            if (!impl->classTypeParams.empty()) continue;

            // 链 = 自根到叶 (stage 2.8 解好). 类链登记表只收"能当基类用"的模块,
            // 表里没有 = 这个类没有基类可继承, 就它自己一张声明表。
            std::vector<Module*> chain{ impl };
            auto cv = classes_.find(ifaceLower(impl->moduleName));
            if (cv != classes_.end() && !cv->second.chainBroken && !cv->second.chain.empty()) {
                std::vector<Module*> full;
                for (const std::string& k : cv->second.chain) {
                    auto hit = byName.find(k);
                    if (hit != byName.end() && hit->second) full.push_back(hit->second);
                }
                if (!full.empty()) chain.swap(full);
            }

            for (const auto& r : cc->ifaces) {
                auto vit = ifaces_.find(ifaceLower(r.ifaceName));
                if (vit == ifaces_.end()) continue;   // 条目名不存在: Pass F 已报 VB3031
                const IfaceView& view = vit->second;
                if (view.chainBroken) continue;       // 父链有错, 建表期已报, 不再级联

                // 实现侧槽表: **叶优先**入席 (派生遮蔽祖先), 所以链倒着走、首见者胜。
                // 与 B02 同一取舍: 访问级别不参与判定 (Private 成员也算入席)。
                std::map<std::string, IfaceProcSig> bound;
                for (size_t i = chain.size(); i-- > 0;) {
                    for (const auto& d : chain[i]->declarations) {
                        if (!d) continue;
                        IfaceProcSig sig;
                        if (!ifaceSigFromDecl(*d, sig)) continue;
                        std::string key = sig.slotKey;
                        if (const auto* clauses = ifaceProcClauses(*d)) {
                            if (!clauses->empty()) {
                                // 写了子句的成员只按子句入座 (B02b 同一条规矩)
                                key.clear();
                                for (const ImplementsClause& c : *clauses) {
                                    bool mine = ifaceLower(c.ifaceName) == ifaceLower(view.name);
                                    for (const IfaceSlotView& s : view.slots) {
                                        if (ifaceLower(c.ifaceName) == ifaceLower(s.ownerIface)) mine = true;
                                    }
                                    if (!mine) continue;
                                    key = ifaceSlotPrefix(*d) + ifaceLower(c.memberName);
                                    break;
                                }
                                if (key.empty()) continue;   // 子句指向别的接口: 不参与本契约
                            }
                        }
                        bound.emplace(key, sig);
                    }
                }

                const bool via = delegatedTo(ifaceLower(impl->moduleName), ifaceLower(view.name));
                for (const auto& slot : view.slots) {
                    IfaceProcSig want;
                    if (slot.sig) ifaceSigFromDecl(*slot.sig, want);
                    const std::string member = slot.ownerIface + "." +
                        (slot.memberName.empty() ? bareName(slot.key) : slot.memberName);

                    auto it = bound.find(slot.key);
                    if (it == bound.end()) {
                        if (!via) {
                            diag_->error(DiagnosticID::SemInterfaceNotImplemented, r.loc,
                                "CoClass '" + cc->name + "' binds interface '" + view.name +
                                "': member '" + member + "' (" + want.text + ") is not implemented by "
                                "class '" + impl->moduleName + "' or any of its ancestors");
                        }
                        continue;
                    }
                    if (!ifaceSigEqual(want, it->second)) {
                        diag_->error(DiagnosticID::SemInterfaceSignatureMismatch,
                            it->second.decl ? it->second.decl->loc : r.loc,
                            "CoClass '" + cc->name + "' binds interface '" + view.name +
                            "': member '" + member + "' signature mismatch (interface: " +
                            want.text + ", class: " + it->second.text + ")");
                    }
                }
            }
        }
    }
    return !diag_->hasErrors();
}

} // namespace vb6c3
