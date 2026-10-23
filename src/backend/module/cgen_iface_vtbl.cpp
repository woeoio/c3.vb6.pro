// vb6c3 - tB 式 Interface 契约发码 (ai/022 D2/D3, 批次 B04)
//
// 本文件只负责"新式 `Interface ... End Interface`"这条路径，与 legacy VB6 的
// `IFoo_M` 前缀扫描 + 胖对 `vb6_iface_<I>{vtbl,obj}` (cgen_com.cpp) 完全隔离：
//
//   类型（工程级，只发一次，#ifndef 守卫）
//     typedef struct vb6_ivref_I { ... } vb6_ivref_I;   // 薄指针的指向类型（单词）
//     typedef struct vb6_ivtbl_I { QI, AddRef, Release, <自有槽…> } vb6_ivtbl_I;
//   实现类
//     vb6_cls_C { void* __comObj; vb6_ivref_I __iv_I; … }   // __comObj 必须仍是第 0 字段 (D19)
//     static <R> vb6_iimpl_C_I_<slot>(vb6_ivref_I* self, …) // container_of 后直调成员
//     static const vb6_ivtbl_I vb6_ivtbl_I_for_C = { … };
//     vb6_cls_C_New(): me->__iv_I.vt = &vb6_ivtbl_I_for_C;
//   接口值 = &obj->__iv_I（D3 的薄指针口径），派发见 cgen_expr_call_ivref 分支.
//
// IUnknown 三件套：B04 是**占位**（E_NOTIMPL / 常量计数），槽号因此从 B04 起永久固定。
// B05 把 AddRef/Release 换成真计数（引用计数头 `__refcount` 只落在实现新式接口的类上，
// 无新语法的工程一个字节都不变），QueryInterface 仍占位到 B06（TypeOf/转换同批）.
//
// B16: 整条薄面对（IUnknown 前缀 + 契约槽）一律加 `__stdcall`。x86 下不加就是 `__cdecl`
// （B16 测量①：`--arch x86 --emit-c` 的产物里一个约定修饰都没有），而 COM 规范是
// `CC_STDCALL`；类型库里那一档要如实写 `callconv`，它就必须真的成立。x64 下 MSVC 接受并
// 忽略 `__stdcall`（同 cgen_delegate/cgen_com_events 的先例），所以两种架构共用一份生成码，
// 只有"实现新式接口"的工程这一段会变（其余工程逐字节不变）。
// 同批还给出薄面的出入口 `vb6_iv_thin_<C>` / `vb6_iv_claim_<C>`（见 emitIfaceImplTables
// 末尾）：COM 包装器按 IID 交回薄指针、并在自己最后一次 Release 时把底座引用交还实例，
// 于是 `IClassFactory::CreateInstance` 那种"QI 完就 Release 包装器"的规范姿势下，交出去
// 的接口指针不会悬空。

#include "backend/cgen.hpp"

#include "semantics/interface_sig.hpp"
#include "semantics/interfaces_registry.hpp"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <map>
#include <set>
#include <string>
#include <vector>

namespace vb6c3 {
namespace {

// 限定名末段：`Project.IFoo` → `IFoo`（与 parseImplements 的点号拼接对称, Fix 083）
std::string ivLastSegment(const std::string& s) {
    size_t p = s.rfind('.');
    return p == std::string::npos ? s : s.substr(p + 1);
}

bool ivNameMatches(const std::string& written, const std::string& name) {
    const std::string lower = ifaceLower(name);
    return ifaceLower(written) == lower || ifaceLower(ivLastSegment(written)) == lower;
}

std::vector<std::unique_ptr<ParameterDecl>>* ivParamsOf(Decl& d) {
    switch (d.kind) {
        case ASTNodeKind::SubDecl:      return &static_cast<SubDecl&>(d).params;
        case ASTNodeKind::FunctionDecl: return &static_cast<FunctionDecl&>(d).params;
        case ASTNodeKind::PropertyDecl: return &static_cast<PropertyDecl&>(d).params;
        default:                        return nullptr;
    }
}

// B06a: 接口 IID —— 16 字节，按真实 GUID 内存序摆（Data1/2/3 小端、Data4 原序），
// 这样 P6 把它交给真 COM 时不必再翻字节。取值优先级：源码里的 [InterfaceId("...")]
// （IfaceView.guid，B01 起只写不读，本批开始有消费者）→ 否则按 4 词 FNV-1a 从接口
// 小写名确定性派生（与 cgen_util_dllentry_prelude.inc 的 generateIid 同族；D8 禁止随机）。
uint32_t ivFnv1a(const std::string& s, uint32_t seed) {
    uint32_t h = seed;
    for (unsigned char c : s) { h ^= c; h *= 0x01000193u; }
    return h;
}

bool ivParseGuidText(const std::string& text, unsigned char out[16]) {
    auto digitVal = [](char c) -> int { return c <= '9' ? c - '0' : (c - 'a' + 10); };
    std::string hex;
    for (char ch : text) {
        char lo = (char)tolower((unsigned char)ch);
        if (isdigit((unsigned char)lo) || (lo >= 'a' && lo <= 'f')) hex.push_back(lo);
    }
    if (hex.size() != 32) return false;
    unsigned char b[16];
    for (int i = 0; i < 16; i++) {
        b[i] = (unsigned char)(digitVal(hex[(size_t)i * 2]) * 16 + digitVal(hex[(size_t)i * 2 + 1]));
    }
    // 可读序: Data1=b[0..3] Data2=b[4..5] Data3=b[6..7] Data4=b[8..15]
    out[0] = b[3]; out[1] = b[2]; out[2] = b[1]; out[3] = b[0];
    out[4] = b[5]; out[5] = b[4];
    out[6] = b[7]; out[7] = b[6];
    for (int i = 8; i < 16; i++) out[i] = b[i];
    return true;
}

void ivDeriveIid(const IfaceView& v, unsigned char out[16]) {
    if (!v.guid.empty() && ivParseGuidText(v.guid, out)) return;
    const std::string key = "iviface:" + ifaceLower(v.name);
    const uint32_t h1 = ivFnv1a(key, 0xa1b2c3d4u);
    const uint32_t h2 = ivFnv1a(key, 0xe5f60718u);
    const uint32_t h3 = ivFnv1a(key, 0x9a0b1c2du);
    const uint32_t h4 = ivFnv1a(key, 0x3e4f5061u);
    const uint32_t d1 = h1;
    const uint32_t d2 = (h2 >> 16) & 0xFFFFu;
    const uint32_t d3 = ((h2 & 0xFFFFu) | 0x4000u) & 0xFFFFu;   // version 4
    const uint32_t d4a = ((((h3 >> 16) & 0xFFFFu) | 0x8000u)) & 0xFFFFu;  // variant 1
    const uint32_t d4b = (h3 & 0xFFFFu);
    out[0] = (unsigned char)(d1 >> 24); out[1] = (unsigned char)(d1 >> 16);
    out[2] = (unsigned char)(d1 >> 8);  out[3] = (unsigned char)d1;
    out[4] = (unsigned char)(d2 >> 8);  out[5] = (unsigned char)d2;
    out[6] = (unsigned char)(d3 >> 8);  out[7] = (unsigned char)d3;
    out[8] = (unsigned char)(d4a >> 8); out[9] = (unsigned char)d4a;
    out[10] = (unsigned char)(d4b >> 8); out[11] = (unsigned char)d4b;
    out[12] = (unsigned char)(h4 >> 24); out[13] = (unsigned char)(h4 >> 16);
    out[14] = (unsigned char)(h4 >> 8);  out[15] = (unsigned char)h4;
}

std::string ivIidInitializer(const unsigned char b[16]) {
    static const char* hex = "0123456789ABCDEF";
    std::string out = "{ ";
    for (int i = 0; i < 16; i++) {
        if (i) out += ",";
        out += std::string("0x") + hex[(b[i] >> 4) & 0xF] + hex[b[i] & 0xF];
    }
    out += " }";
    return out;
}

} // namespace

// ============================================================
// 查找
// ============================================================

const IfaceView* CCodeGen::ivLookupIface(const std::string& written) {
    if (!ivreg_) return nullptr;
    auto it = ivreg_->find(ifaceLower(written));
    if (it == ivreg_->end()) it = ivreg_->find(ifaceLower(ivLastSegment(written)));
    if (it == ivreg_->end()) return nullptr;
    if (it->second.chainBroken || it->second.slots.empty()) return nullptr;
    return &it->second;
}

std::string CCodeGen::ivrefCType(const std::string& typeName) {
    const IfaceView* v = ivLookupIface(typeName);
    return v ? "vb6_ivref_" + cIdent(v->name) + "*" : std::string();
}

// 本模块（类）实现了哪些新式接口：按 `Implements` 书写序，同名去重
std::vector<const IfaceView*> CCodeGen::ivImplementedIfaces(Module& module) {
    std::vector<const IfaceView*> out;
    std::vector<std::string> seen;
    for (const auto& impl : module.implements) {
        if (!impl) continue;
        const IfaceView* v = ivLookupIface(impl->interfaceName);
        if (!v) continue;
        const std::string key = ifaceLower(v->name);
        if (std::find(seen.begin(), seen.end(), key) != seen.end()) continue;
        seen.push_back(key);
        out.push_back(v);
    }
    return out;
}

// 槽的实现成员：显式子句优先，其次同名隐式匹配（与语义层同一套口径, D16-2）
Decl* CCodeGen::ivFindImplMember(Module& module, const IfaceView& v, const IfaceSlotView& slot) {
    for (auto& d : module.declarations) {
        if (!d) continue;
        const std::vector<ImplementsClause>* clauses = ifaceProcClauses(*d);
        if (!clauses || clauses->empty()) continue;
        for (const auto& c : *clauses) {
            if (!ivNameMatches(c.ifaceName, slot.ownerIface) &&
                !ivNameMatches(c.ifaceName, v.name)) {
                continue;
            }
            // 允许写成员名 (`I.Name`) 或直接写槽名 (`I.get_Name`)
            if (ifaceClauseSlotKey(*d, c.memberName) == slot.key ||
                ifaceLower(c.memberName) == slot.key) {
                return d.get();
            }
        }
    }
    for (auto& d : module.declarations) {
        if (!d) continue;
        const std::vector<ImplementsClause>* clauses = ifaceProcClauses(*d);
        if (clauses && !clauses->empty()) continue;  // 写了子句 = 只按子句入座
        IfaceProcSig sig;
        if (!ifaceSigFromDecl(*d, sig)) continue;
        if (sig.slotKey == slot.key) return d.get();
        // VB6 风格接口 (.cls 宿主): 实现成员按 VB6 强制约定命名为 `<接口名>_<成员名>`,
        // 槽键得摘掉前缀再比。与语义层 checkNewStyleInterface 的隐式匹配同一口径
        // (两边必须挑中**同一个**成员, 否则语义说实现了、发码却把槽留成 NULL)。
        if (v.clsHost) {
            const std::string vb6Key = ifaceImplSlotKeyVb6(*d, sig.memberName, v.name, true);
            if (vb6Key == slot.key) return d.get();
        }
    }
    return nullptr;
}

// tB 委托式实现 (ai/022 B10): 本类是否把接口 v 的整份契约转交给一个持有字段。
// 裁决 (字段存不存在、字段类型那个类实现没实现 v) 全在 stage 2.7 Pass D —— 那时才
// 看得见整工程的模块表; 这里只取结果, 不再判一遍。
const ViaView* CCodeGen::ivViaFor(const Module& module, const IfaceView& v) {
    if (!ivviareg_) return nullptr;
    auto it = ivviareg_->find(ifaceLower(module.moduleName));
    if (it == ivviareg_->end()) return nullptr;
    const std::string want = ifaceLower(v.name);
    for (const ViaView& vv : it->second) {
        if (vv.ifaceKey == want) return &vv;
    }
    return nullptr;
}

// ============================================================
// 签名文本
// ============================================================
// 调用点成员名 -> 槽键：先按 Sub/Function 的裸名匹配，再按属性三槽前缀回退
// （`s.Name` 读属性 → get_name；写属性由赋值路径自己带前缀，B04b 之后补）
std::string CCodeGen::ivSlotKeyForMember(const IfaceView& v, const std::string& member) {
    const std::string bare = ifaceLower(member);
    for (const IfaceSlotView& s : v.slots) {
        if (s.key == bare) return s.key;
    }
    for (const char* p : {"get_", "put_", "putref_"}) {
        const std::string prefixed = std::string(p) + bare;
        for (const IfaceSlotView& s : v.slots) {
            if (s.key == prefixed) return s.key;
        }
    }
    return std::string();
}

std::string CCodeGen::ivSlotRetType(const Decl* sig) {
    if (!sig) return "void";
    Decl& d = *const_cast<Decl*>(sig);
    if (d.kind == ASTNodeKind::FunctionDecl) {
        return mapTypeRef(static_cast<FunctionDecl&>(d).returnType.get());
    }
    if (d.kind == ASTNodeKind::PropertyDecl) {
        PropertyDecl& p = static_cast<PropertyDecl&>(d);
        return p.propKind == ProcKind::PropertyGet ? mapTypeRef(p.returnType.get())
                                                   : std::string("void");
    }
    return "void";
}

// 形参声明文本（不含 self），与 makeParamList 同口径：ByRef→指针，Optional 追加 _has_ 尾标记
std::string CCodeGen::ivParamDecls(Decl& d) {
    std::string out;
    std::vector<std::unique_ptr<ParameterDecl>>* ps = ivParamsOf(d);
    if (!ps) return out;
    for (auto& p : *ps) out += ", " + makeParamCType(p.get(), false);
    for (auto& p : *ps) {
        if (p->isOptional && !p->isParamArray) out += ", int _has_" + cIdent(p->name);
    }
    return out;
}

// 适配器转调实参（不含 me），顺序必须与 ivParamDecls 严格一致
std::string CCodeGen::ivForwardArgs(Decl& d) {
    std::string out;
    std::vector<std::unique_ptr<ParameterDecl>>* ps = ivParamsOf(d);
    if (!ps) return out;
    for (auto& p : *ps) out += ", " + cIdent(p->name);
    for (auto& p : *ps) {
        if (p->isOptional && !p->isParamArray) out += ", _has_" + cIdent(p->name);
    }
    return out;
}

// 实现成员对应的 C 过程名（与 cgen_com.cpp 的 legacy 取名口径一致）
std::string CCodeGen::ivImplCName(Decl& d) {
    const std::string mod = isClassModule_ ? moduleName_ : std::string();
    if (d.kind == ASTNodeKind::SubDecl) {
        SubDecl& n = static_cast<SubDecl&>(d);
        return cProcName(n.name, n.access, mod);
    }
    if (d.kind == ASTNodeKind::FunctionDecl) {
        FunctionDecl& n = static_cast<FunctionDecl&>(d);
        return cProcName(n.name, n.access, mod);
    }
    if (d.kind == ASTNodeKind::PropertyDecl) {
        PropertyDecl& n = static_cast<PropertyDecl&>(d);
        std::string prefix = "prop_let_";
        if (n.propKind == ProcKind::PropertyGet) prefix = "prop_get_";
        else if (n.propKind == ProcKind::PropertySet) prefix = "prop_set_";
        return cProcName(prefix + n.name, n.access, mod);
    }
    return std::string();
}

// ============================================================
// 头文件：工程级接口类型（#ifndef 守卫，任何模块用到都自足）
// ============================================================

void CCodeGen::emitIfaceContractTypedefs() {
    if (!ivreg_ || ivreg_->empty()) return;  // 无新式接口的工程：零输出（护栏）

    std::vector<const IfaceView*> views;
    for (const auto& kv : *ivreg_) views.push_back(&kv.second);
    // unordered_map 迭代顺序不稳定 → 按小写名排序，保证生成物可复现
    std::sort(views.begin(), views.end(), [](const IfaceView* a, const IfaceView* b) {
        return ifaceLower(a->name) < ifaceLower(b->name);
    });

    bool headerDone = false;
    bool iidBaseDone = false;
    for (const IfaceView* v : views) {
        if (v->chainBroken || v->slots.empty()) continue;
        if (!headerDone) {
            h_.emitLine("// === tB-style Interface contract slot tables (ai/022 B04) ===");
            headerDone = true;
        }
        const std::string id = cIdent(v->name);
        const std::string tbl = "vb6_ivtbl_" + id;
        const std::string ref = "vb6_ivref_" + id;
        std::string guard = "VB6_IVTBL_" + id;
        for (char& ch : guard) {
            if (ch >= 'a' && ch <= 'z') ch = static_cast<char>(ch - 'a' + 'A');
        }
        // IUnknown 的 IID 全工程唯一一份（同样每 TU 一份 static 常量 → 只能按值比）
        if (!iidBaseDone) {
            iidBaseDone = true;
            h_.emitLine("#ifndef VB6_IV_IID_IUNKNOWN");
            h_.emitLine("#define VB6_IV_IID_IUNKNOWN");
            h_.emitLine("/* IID of IUnknown {00000000-0000-0000-C000-000000000046}, GUID memory order */");
            h_.emitLine("static const unsigned char vb6_iv_iid_IUnknown[16] = "
                        "{ 0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46 };");
            h_.emitLine("#endif");
            h_.emitBlank();
        }
        h_.emitLine("#ifndef " + guard);
        h_.emitLine("#define " + guard);
        h_.emitLine("typedef struct " + ref + " " + ref + ";");
        h_.emitLine("typedef struct " + tbl + " {");
        h_.indent();
        h_.emitLine("/* IUnknown prefix slots: AddRef/Release real in B05, QueryInterface in B06a */");
        h_.emitLine("long (__stdcall *QueryInterface)(void* self, const void* riid, void** ppv);");
        h_.emitLine("unsigned long (__stdcall *AddRef)(void* self);");
        h_.emitLine("unsigned long (__stdcall *Release)(void* self);");
        for (const IfaceSlotView& slot : v->slots) {
            h_.emitLine(ivSlotRetType(slot.sig) + " (__stdcall *" + cIdent(slot.key) + ")(" +
                        ref + "* self" + ivParamDeclsRef(slot.sig) + ");");
        }
        h_.dedent();
        h_.emitLine("} " + tbl + ";");
        h_.emitLine("struct " + ref + " { const " + tbl + "* vt; };");
        // B06a: 本接口的 IID（QueryInterface 按值比这个 16 字节块）
        {
            unsigned char iid[16];
            const std::string* mapIid = nullptr;
            if (iidreg_) {
                auto it = iidreg_->find(ifaceLower(v->name));
                if (it != iidreg_->end()) mapIid = &it->second;
            }
            std::string srcNote = "derived from name";
            if (mapIid && ivParseGuidText(*mapIid, iid)) {
                // ai/022 B13c: 一枚接口在一次编译里只许一枚 GUID —— 值来自 stage 2.7 Pass E
                // 的接口 IID 表，与 dll_entry 的 IID 表、类型库读同一份。
                srcNote = "from interface IID map";
            } else {
                ivDeriveIid(*v, iid);
                if (!v->guid.empty()) srcNote = "from InterfaceId attribute";
            }
            h_.emitLine("static const unsigned char vb6_iv_iid_" + id + "[16] = " +
                        ivIidInitializer(iid) + ";  /* " + srcNote + " */");
        }
        h_.emitLine("#endif");
        h_.emitBlank();
    }
}

// 头文件槽声明用的形参文本（接口侧签名可能是 const 节点）
std::string CCodeGen::ivParamDeclsRef(const Decl* sig) {
    if (!sig) return std::string();
    return ivParamDecls(*const_cast<Decl*>(sig));
}

// ============================================================
// 账 #305 新账 #464: legacy VB6 接口 (`.cls` 宿主、没进 tB 登记表) 的槽表
// ============================================================
//
// 从前这一发住在 `emitInterfaceVtable` 里, 由**每枚实现类**按自家成员声明顺序 + 自家
// `classMeParam()` 现 harvest ⇒ 一枚接口被 N 枚类实现就发 N 份同名 `vb6_vtbl_<I>`
// (真语料 Ray.vbp 实测 54 发 C2011, 四份表里槽序有两套)。现在形状读 driver 预扫描的
// 那张只读表, 表类型只由接口自己这一趟发一份, `me` 一律 `void*` —— 派发点
// (`<view>.vtbl-><成员>(<view>.obj, …)`) 交的本来就是 `void*`, 与实现类是谁无关。
//
// 与 `emitIfaceContractTypedefs` 那套薄面对 (`vb6_ivref_<I>` / `vb6_ivtbl_<I>`) 刻意并存、
// 不合并: 两套类型名、两套所有权, 合并会把已经进登记表的接口重新拖回胖值结构的口径。

std::vector<CCodeGen::LegacyIfaceSlot> CCodeGen::legacyIfaceSlots(const std::string& ifaceName) {
    std::vector<LegacyIfaceSlot> out;
    if (!legacyIfaceShapes_) return out;
    const std::string key = ifaceLower(ifaceName);
    auto it = legacyIfaceShapes_->find(key);
    if (it == legacyIfaceShapes_->end()) return out;
    const std::string prefix = key + "_";
    for (const Decl* sig : it->second.slots) {
        if (!sig) continue;
        Decl& d = *const_cast<Decl*>(sig);
        std::string written;
        if (d.kind == ASTNodeKind::SubDecl) {
            written = static_cast<SubDecl&>(d).name;
        } else if (d.kind == ASTNodeKind::FunctionDecl) {
            written = static_cast<FunctionDecl&>(d).name;
        } else if (d.kind == ASTNodeKind::PropertyDecl) {
            written = static_cast<PropertyDecl&>(d).name;
        } else {
            continue;
        }
        // 宿主在工程外时形状是按实现类那批 `<接口名>_成员` 建的, 摘掉前缀才是成员名
        const std::string lower = ifaceLower(written);
        LegacyIfaceSlot s;
        if (lower.size() > prefix.size() && lower.compare(0, prefix.size(), prefix) == 0) {
            written = written.substr(prefix.size());
        }
        // 键带档 = 同一位主人 ifaceSlotKeyNamed (interface_sig.hpp); 字段名只给 Let/Set
        // 那两档加后缀 => Get/Sub/Function 的产物逐字节不动 (账 #305 新账 #489)。
        const ProcKind dir = ifaceDeclProcKind(d);
        s.slotKey = ifaceSlotKeyNamed(ifaceLower(written), d.kind, dir);
        s.field = dir == ProcKind::PropertyLet  ? cIdent(written) + "_put"
                : dir == ProcKind::PropertySet ? cIdent(written) + "_putref"
                                               : cIdent(written);
        s.retType = ivSlotRetType(&d);
        s.params = ivParamDecls(d);   // 带前置 ", "; 无参 = 空串
        out.push_back(std::move(s));
    }
    return out;
}

void CCodeGen::emitLegacyIfaceContractTypedefs() {
    // 无 legacy 接口的工程: 零输出 (护栏)。std::map 按键有序 ⇒ 发射顺序与模块顺序、
    // 与编译器二进制都无关 (同 §B97 那条「定序不许交给 unordered」的口径)。
    if (!legacyIfaceShapes_ || legacyIfaceShapes_->empty()) return;

    bool headerDone = false;
    for (const auto& kv : *legacyIfaceShapes_) {
        const std::string& writtenName = kv.second.writtenName;
        const std::vector<LegacyIfaceSlot> slots = legacyIfaceSlots(writtenName);
        if (slots.empty()) continue;
        if (!headerDone) {
            h_.emitLine("// === legacy VB6 interface slot tables (账 #305 新账 #464: 形状归接口自己定) ===");
            headerDone = true;
        }
        const std::string id = cIdent(writtenName);
        const std::string vtblName = "vb6_vtbl_" + id;
        const std::string ifaceType = "vb6_iface_" + id;
        std::string guard = "VB6_IFACE_" + id + "_TBL";
        for (char& ch : guard) {
            if (ch >= 'a' && ch <= 'z') ch = static_cast<char>(ch - 'a' + 'A');
        }
        h_.emitLine("#ifndef " + guard);
        h_.emitLine("#define " + guard);
        h_.emitLine("typedef struct " + vtblName + " {");
        h_.indent();
        for (const LegacyIfaceSlot& s : slots) {
            h_.emitLine(s.retType + " (*" + s.field + ")(void* me" + s.params + ");");
        }
        h_.dedent();
        h_.emitLine("} " + vtblName + ";");
        h_.emitLine("typedef struct " + ifaceType + " {");
        h_.indent();
        h_.emitLine(vtblName + "* vtbl;");
        h_.emitLine("void* obj;");
        h_.dedent();
        h_.emitLine("} " + ifaceType + ";");
        h_.emitLine("#endif");
        h_.emitBlank();
    }
}

// wrap 名的唯一出口: 一枚接口的表是工程级一份, 而 `_wrap` 绑的是**实现类** ——
// 名字不带实现者, 多实现类就会发 N 份同名函数 (C2084)。
std::string CCodeGen::legacyIfaceWrapFn(const std::string& ifaceName,
                                        const std::string& implName) const {
    return "vb6_iface_" + cIdent(ifaceName) + "_wrap_" + cIdent(implName);
}

// 账 #305 新账 #471 (第一百七十四刀): 「运行期手里那一枚实例是哪一枚工程类 → 交回这一枚接口的
// 视图」这一发的名字唯一出口。与上面那位主人同一格 —— `vb6_iface_` 那串前缀不许在这里长第三份。
std::string CCodeGen::legacyIfaceViewFromFn(const std::string& ifaceName) const {
    return "vb6_iface_" + cIdent(ifaceName) + "_view_from";
}

// 铸在**接口自己**那份 TU 里 (原型进它的 .h、定义进它的 .c —— 视图怎么造出来归接口自己定,
// 与 `emitLegacyIfaceWrapDefs` 那一批同处)。
// 为什么必须在运行期问: `For Each Obj In 集合` 的循环变量声明成视图 (`Dim Obj As <接口>`) 时,
// 元素是集合交回来的一枚 COM VARIANT, 里面装的是哪一枚实现类**编译期不知道** —— 而视图的 vtbl
// 是 `static` 在各实现类那份 TU 里的 (第一百六十六刀), 消费侧手里只有 wrap 的原型。
// 交回的是**具名结构体值**, 与 wrap 那一族同形 ⇒ 不赌「返回结构体」那一族的调用约定。
// 认不出 (不是本族包装器 / 外部 COM 对象 / 控件句柄 / 名字不在这份名单里) 一律 `{0}` 静默 ——
// VB6 那一形报 13 类型不匹配, 那是这条线在册的欠项口径 (#471 那一格), 本刀不引错误腿。
void CCodeGen::emitLegacyIfaceViewFrom(Module& module) {
    if (!legacyIfaceShapes_) return;
    const std::string ifaceName = module.moduleName;
    if (legacyIfaceShapes_->find(ifaceLower(ifaceName)) == legacyIfaceShapes_->end()) return;
    if (legacyIfaceSlots(ifaceName).empty()) return;   // 与表类型那一趟同一个闸

    const std::string ifaceType = "vb6_iface_" + cIdent(ifaceName);
    const std::string fn = legacyIfaceViewFromFn(ifaceName);
    const std::vector<std::string> impls = allImplementationClasses(ifaceName);

    h_.emitBlank();
    h_.emitLine(ifaceType + " " + fn + "(void* instance, const char* className);");

    c_.emitBlank();
    c_.emitLine("// === 账 #305 #471: 运行期那一枚元素 → " + ifaceName + " 的视图 (For Each 那一站) ===");
    c_.emitLine(ifaceType + " " + fn + "(void* instance, const char* className) {");
    c_.indent();
    c_.emitLine(ifaceType + " v = {0};");
    c_.emitLine("if (!instance || !className) return v;");
    for (const std::string& impl : impls) {
        c_.emitLine("if (_stricmp(className, \"" + impl + "\") == 0)"
                    " return " + legacyIfaceWrapFn(ifaceName, impl) +
                    "((vb6_cls_" + cIdent(impl) + "*)instance);");
    }
    c_.emitLine("return v;  /* 认不出 = 空视图; 在册欠项 = 这一形 VB6 报 13 */");
    c_.dedent();
    c_.emitLine("}");
}

// ============================================================
// 类结构体字段 + _New 初始化
// ============================================================

void CCodeGen::emitIfaceClassFields(Module& module) {
    // 账 #305 §B218 (甲, 第一刀): 计数头不再跟着接口走 —— **每个类模块都带它**.
    // 位置仍钉在 __comObj (第 0 字段, D19 硬约束) 之后、__iv_<I> 槽字段之前 ⇒ "谁握着这枚实例"
    // 从此在类型里有答案; 代价 = 工程类结构体每个字段偏移都动 (两架构 padding 不同形 ⇒ x86 要真验).
    h_.emitLine("    int32_t __refcount;  /* tB B05 / 账 #445 (甲): 1 at New, +1 per owning ref */");
    const std::vector<const IfaceView*> ifaces = ivImplementedIfaces(module);
    for (const IfaceView* v : ifaces) {
        const std::string id = cIdent(v->name);
        h_.emitLine("    vb6_ivref_" + id + " __iv_" + id +
                    ";  /* tB Interface " + v->name + " (B04): iface ptr = &me->__iv_" + id + " */");
    }
}

void CCodeGen::emitIfaceNewInit(Module& module) {
    const std::string clsId = cIdent(moduleName_);
    // 账 #305 §B218 (甲): 初值 1 也无条件发 —— 语义不变 (`_New` 交出的那一次引用归"接住它的那一枚槽"),
    // 变的是那一枚槽从此**要**在覆写/置空/过程出口上放手 (那三条腿见 emitClassRefPair 与 §B218 五腿).
    c_.emitLine("me->__refcount = 1;  /* tB B05 / 账 #445 (甲): 交出的一枚引用归接住它的那一枚槽 */");
    const std::vector<const IfaceView*> ifaces = ivImplementedIfaces(module);
    if (ifaces.empty()) return;
    for (const IfaceView* v : ifaces) {
        const std::string id = cIdent(v->name);
        c_.emitLine("me->__iv_" + id + ".vt = &vb6_ivtbl_" + id + "_for_" + clsId + ";");
    }
}

// ============================================================
// 类侧：适配器 + 槽表实例
// ============================================================

void CCodeGen::emitIfaceImplTables(Module& module) {
    const std::vector<const IfaceView*> ifaces = ivImplementedIfaces(module);
    if (ifaces.empty()) return;

    const std::string clsId = cIdent(moduleName_);
    const std::string clsStruct = "vb6_cls_" + clsId;

    c_.emitBlank();
    c_.emitLine("// === tB Interface contract impl: " + module.moduleName + " (ai/022 B04) ===");

    // IUnknown 三件套一律按 (类, 接口) 各一份：`self` 是某个 `__iv_<I>` 的地址，
    // 回推实例、以及 QI 命中"本类的另一个接口"时要减/加的偏移都依赖具体接口
    // （B05 已为 AddRef/Release 这么做了，QI 同理）。先给本类全部 AddRef 发前向
    // 声明，QI 才能在任意发射顺序下调到兄弟接口的 AddRef。
    for (const IfaceView* v : ifaces) {
        c_.emitLine("static unsigned long __stdcall vb6_iunk_" + clsId + "_" + cIdent(v->name) +
                    "_AddRef(void* self);");
    }

    for (const IfaceView* v : ifaces) {
        const std::string id = cIdent(v->name);
        const std::string ref = "vb6_ivref_" + id;

        // B06a: 真 QueryInterface —— 认 IUnknown、本接口、以及本类实现的其它接口
        const std::string walkQI = "    " + clsStruct + "* me = (" + clsStruct + "*)((char*)self - offsetof(" +
                                   clsStruct + ", __iv_" + id + "));";
        // B13d: 规范 IUnknown —— 不管"从哪个接口问"，返回的都是**本类实现序里第一个接口**
        // 的薄指针（= 类结构体里最靠前的那个 __iv_ 字段）。它是唯一一处"偏移 0 就是 vtable"
        // 的合法 COM 对象指针，所以这个值既能当身份比、也能直接被再次 QI/AddRef/Release，
        // 而且不需要给类结构体加字段（D19：布局不能动）。
        const std::string canonId = cIdent(ifaces.front()->name);
        c_.emitBlank();
        c_.emitLine("static long __stdcall vb6_iunk_" + clsId + "_" + id + "_QueryInterface(void* self, const void* riid, void** ppv) {");
        c_.indent();
        c_.emitLine(walkQI);
        c_.emitLine("if (!ppv) return 0x80070057L;  /* E_POINTER */");
        c_.emitLine("*ppv = NULL;");
        c_.emitLine("if (!riid) return 0x80070057L;  /* E_POINTER */");
        c_.emitLine("if (vb6_IidEqual(riid, vb6_iv_iid_IUnknown)) {");
        c_.indent();
        c_.emitLine("void* canon = &me->__iv_" + canonId +
                    ";  /* 规范指针 (ai/022 B13d): 本类实现的第一个接口 */");
        c_.emitLine("*ppv = canon;");
        c_.emitLine("vb6_iunk_" + clsId + "_" + canonId + "_AddRef(canon);");
        c_.emitLine("return 0L;  /* S_OK */");
        c_.dedent();
        c_.emitLine("}");
        c_.emitLine("if (vb6_IidEqual(riid, vb6_iv_iid_" + id + ")) {");
        c_.indent();
        c_.emitLine("*ppv = self;");
        c_.emitLine("vb6_iunk_" + clsId + "_" + id + "_AddRef(self);");
        c_.emitLine("return 0L;  /* S_OK */");
        c_.dedent();
        c_.emitLine("}");
        for (const IfaceView* s : ifaces) {
            if (s == v) continue;
            const std::string sid = cIdent(s->name);
            c_.emitLine("if (vb6_IidEqual(riid, vb6_iv_iid_" + sid + ")) {");
            c_.indent();
            c_.emitLine("*ppv = &me->__iv_" + sid + ";");
            c_.emitLine("vb6_iunk_" + clsId + "_" + sid + "_AddRef(*ppv);");
            c_.emitLine("return 0L;  /* S_OK */");
            c_.dedent();
            c_.emitLine("}");
        }
        c_.emitLine("return 0x80004002L;  /* E_NOINTERFACE: not implemented by this class */");
        c_.dedent();
        c_.emitLine("}");

        // B05: 真引用计数 (self = &me->__iv_<I>)
        const std::string walk = "    " + clsStruct + "* me = (" + clsStruct + "*)((char*)self - offsetof(" +
                                 clsStruct + ", __iv_" + id + "));";
        c_.emitBlank();
        c_.emitLine("static unsigned long __stdcall vb6_iunk_" + clsId + "_" + id + "_AddRef(void* self) {");
        c_.indent();
        c_.emitLine(walk);
        // 账 #305 §B218 (甲): 公式只住 emitClassRefPair 那一处, 这里转调.
        c_.emitLine("return " + clsStruct + "_AddRef(me);");
        c_.dedent();
        c_.emitLine("}");
        c_.emitLine("static unsigned long __stdcall vb6_iunk_" + clsId + "_" + id + "_Release(void* self) {");
        c_.indent();
        c_.emitLine(walk);
        // 账 #305 §B218 (甲): 退一 + 到 0 才拆 + 「被包装器接管就不拆」那一条规则只住 emitClassRefPair.
        c_.emitLine("return " + clsStruct + "_Release(me);");
        c_.dedent();
        c_.emitLine("}");

        std::vector<std::string> slotFns;
        for (const IfaceSlotView& slot : v->slots) {
            Decl* impl = ivFindImplMember(module, *v, slot);
            if (!impl) {
                // tB 委托式实现 (ai/022 B10): 本类自己没写这一席, 但整份契约转交给了一个
                // 持有字段 → 发一个"转调持有对象同名槽"的适配器。
                // 为什么绕接口表、而不是直调 `vb6_<持有类>_<成员>`: 实现接口的成员按 VB6 惯例
                // 写 Private = C 层 static, 跨翻译单元连不到 (B09 的 MyBase 正是撞上这条才把
                // Private 目标判死的); 表项发在持有类自己的 TU 里, 函数指针从对象里读, 天然可用。
                const ViaView* via = ivViaFor(module, *v);
                if (!via || !slot.sig) {
                    // 契约缺失在语义层已经报过错（B02/B02b），这里只留 NULL 占位避免级联编译错误
                    slotFns.push_back("NULL");
                    continue;
                }
                const std::string fn = "vb6_iimpl_" + clsId + "_" + id + "_" + cIdent(slot.key);
                const std::string ret = ivSlotRetType(slot.sig);
                const std::string fld = cIdent(via->fieldName);
                Decl& sigDecl = *const_cast<Decl*>(slot.sig);
                slotFns.push_back(fn);

                c_.emitBlank();
                c_.emitLine("static " + ret + " __stdcall " + fn + "(" + ref + "* self" +
                            ivParamDecls(sigDecl) + ") {");
                c_.indent();
                c_.emitLine(clsStruct + "* me = (" + clsStruct + "*)((char*)self - offsetof(" +
                            clsStruct + ", __iv_" + id + "));");
                c_.emitLine("vb6_ivref_" + id + "* h = me->" + fld + " ? &me->" + fld +
                            "->__iv_" + id + " : NULL;  /* tB Via " + via->fieldName +
                            " -> " + via->holderModule + " */");
                const std::string call =
                    "h->vt->" + cIdent(slot.key) + "(h" + ivForwardArgs(sigDecl) + ")";
                if (ret == "void") {
                    c_.emitLine("if (h) " + call + ";");
                } else {
                    c_.emitLine("if (h) return " + call + ";");
                    c_.emitLine(ret + " __via0 = {0};  /* 持有对象是 Nothing: 退成该类型的零值 */");
                    c_.emitLine("return __via0;");
                }
                c_.dedent();
                c_.emitLine("}");
                continue;
            }
            const std::string fn = "vb6_iimpl_" + clsId + "_" + id + "_" + cIdent(slot.key);
            slotFns.push_back(fn);

            c_.emitBlank();
            c_.emitLine("static " + ivSlotRetType(slot.sig) + " __stdcall " + fn + "(" + ref + "* self" +
                        ivParamDecls(*impl) + ") {");
            c_.indent();
            c_.emitLine(clsStruct + "* me = (" + clsStruct + "*)((char*)self - offsetof(" +
                        clsStruct + ", __iv_" + id + "));");
            const std::string call = ivImplCName(*impl) + "(me" + ivForwardArgs(*impl) + ")";
            if (ivSlotRetType(slot.sig) == "void") {
                c_.emitLine(call + ";");
            } else {
                c_.emitLine("return " + call + ";");
            }
            c_.dedent();
            c_.emitLine("}");
        }

        c_.emitBlank();
        c_.emitLine("static const vb6_ivtbl_" + id + " vb6_ivtbl_" + id + "_for_" + clsId + " = {");
        c_.indent();
        c_.emitLine("vb6_iunk_" + clsId + "_" + id + "_QueryInterface,");
        c_.emitLine("vb6_iunk_" + clsId + "_" + id + "_AddRef,");
        c_.emitLine("vb6_iunk_" + clsId + "_" + id + "_Release,");
        for (size_t i = 0; i < slotFns.size(); i++) {
            c_.emitLine(slotFns[i] + (i + 1 < slotFns.size() ? "," : ""));
        }
        c_.dedent();
        c_.emitLine("};");
    }

    // B06b: 两个跨模块助手（声明进类的 .h，实现进类的 .c —— 槽表实例是本 TU 的
    // static const 对象，别的 TU 拿不到它的地址，所以"这根薄指针是不是本类的实例"
    // 只能在这里判，不能在使用点写裸强转）。
    const std::string fromIv = "vb6_iv_from_iv_" + clsId;
    const std::string testIid = "vb6_iv_test_iid_" + clsId;
    // Fix <vbeclipse>: 实例 → 薄指针 的反查 (与 fromIv 的方向相反)。
    // fromIv 收的是**薄指针** (&me->__iv_<I>), 拿首字��当槽表比对; 而接口成员桩
    // (`vb6_I_<M>(me)`) 收到的是**实例指针**, 首字段是 __comObj 不是 vt, 直接喂
    // fromIv 永远比不中 => 桩恒返回零值 (m_Scheme.BackColor 读成纯黑, 停靠面板
    // 整片黑底)。这个助手从实例出发走 me->__iv_<I>.vt, 才是桩该用的判据。
    const std::string thinOf = "vb6_iv_thin_of_" + clsId;
    h_.emitBlank();
    h_.emitLine("void* " + fromIv + "(void* self);  /* tB Interface B06b: iface ptr -> instance */");
    h_.emitLine("int32_t " + testIid + "(void* self, const void* riid);  /* tB Interface B06b: TypeOf <cls> Is <iface> */");
    h_.emitLine("void* " + thinOf + "(void* instance);  /* Fix <vbeclipse>: instance -> iface ptr (本类没实现 = NULL) */");

    c_.emitBlank();
    c_.emitLine("void* " + thinOf + "(void* instance) {");
    c_.indent();
    c_.emitLine(clsStruct + "* me = (" + clsStruct + "*)instance;");
    c_.emitLine("if (!me) return NULL;");
    c_.emitLine("/* 按槽表地址认类, 不读 vt 字段 —— 这样对「实例被 COM 包装器握着」和「裸实例」");
    c_.emitLine("   两种形态都成立, 也就不用关心 __comObj 清没清。 */");
    for (const IfaceView* v : ifaces) {
        const std::string id = cIdent(v->name);
        c_.emitLine("if (me->__iv_" + id + ".vt == (const void*)&vb6_ivtbl_" + id + "_for_" + clsId + ")");
        c_.indent();
        c_.emitLine("return &me->__iv_" + id + ";");
        c_.dedent();
    }
    c_.emitLine("return NULL;  /* 不是本类实例 */");
    c_.dedent();
    c_.emitLine("}");

    c_.emitBlank();
    c_.emitLine("void* " + fromIv + "(void* self) {");
    c_.indent();
    c_.emitLine("const void* vt;");
    c_.emitLine("if (!self) return NULL;");
    c_.emitLine("vt = *(const void**)self;");
    for (const IfaceView* v : ifaces) {
        const std::string id = cIdent(v->name);
        c_.emitLine("if (vt == (const void*)&vb6_ivtbl_" + id + "_for_" + clsId + ")");
        c_.indent();
        c_.emitLine("return (char*)self - offsetof(" + clsStruct + ", __iv_" + id + ");");
        c_.dedent();
    }
    c_.emitLine("return NULL;  /* not an instance of this class (Set ... = Nothing 语义) */");
    c_.dedent();
    c_.emitLine("}");

    c_.emitBlank();
    c_.emitLine("int32_t " + testIid + "(void* self, const void* riid) {");
    c_.indent();
    c_.emitLine("if (!self || !riid) return 0;  /* Nothing 不匹配任何类型 */");
    c_.emitLine("if (vb6_IidEqual(riid, vb6_iv_iid_IUnknown)) return 1;");
    for (const IfaceView* v : ifaces) {
        c_.emitLine("if (vb6_IidEqual(riid, vb6_iv_iid_" + cIdent(v->name) + ")) return 1;");
    }
    c_.emitLine("return 0;");
    c_.dedent();
    c_.emitLine("}");

    // B16: 薄面的两个出入口 —— COM 包装器 (vb6comserver) 靠它们把"接口 IID"这一档
    // 交回给调用者，并管住交出去之后的生命期。
    //   vb6_iv_thin_<C>(instance, riid) : IID → 那根薄指针 (本类没实现 = NULL)
    //   vb6_iv_claim_<C>(instance)      : 包装器放手 —— 清掉 __comObj 回指, 再退掉底座引用
    //     (_New / FromInstance 移交给包装器的那一次). 退完仍 > 0 = 还有薄引用在世, 由最后一个
    //     薄引用的 Release 收尾 (那条路的 __comObj 已被清空, 直接走"自己销毁"分支).
    // 为什么必须成对给: `IClassFactory::CreateInstance` 的规范姿势就是
    // `QI(请求的 IID)` 紧接着 `Release(包装器)` —— 只给 thin 不给 claim, 刚交出去的
    // 接口指针会被那次 Release 当场销毁 (悬空). 成对给之后, "谁手里还有引用"由
    // __refcount 一条线说了算 (薄指针的 AddRef/Release 自 B05 起就记在这条线上).
    const std::string thinFn = "vb6_iv_thin_" + clsId;
    const std::string claimFn = "vb6_iv_claim_" + clsId;
    h_.emitBlank();
    h_.emitLine("void* " + thinFn + "(void* instance, const void* riid);  /* tB Interface B16: 接口 IID -> 薄指针 (无则 NULL) */");
    h_.emitLine("void " + claimFn + "(void* instance);  /* tB Interface B16: COM 包装器放手, 退掉底座引用 */");

    c_.emitBlank();
    c_.emitLine("void* " + thinFn + "(void* instance, const void* riid) {");
    c_.indent();
    c_.emitLine(clsStruct + "* me = (" + clsStruct + "*)instance;");
    c_.emitLine("if (!me || !riid) return NULL;");
    for (const IfaceView* v : ifaces) {
        const std::string id = cIdent(v->name);
        c_.emitLine("if (vb6_IidEqual(riid, vb6_iv_iid_" + id + ")) return &me->__iv_" + id + ";");
    }
    c_.emitLine("return NULL;  /* 本类没实现这个接口 */");
    c_.dedent();
    c_.emitLine("}");

    c_.emitBlank();
    c_.emitLine("void " + claimFn + "(void* instance) {");
    c_.indent();
    c_.emitLine(clsStruct + "* me = (" + clsStruct + "*)instance;");
    c_.emitLine("if (!me) return;");
    c_.emitLine("me->__comObj = NULL;  /* B16: 包装器已走, 薄引用的最后一次 Release 自己收尾 */");
    // 账 #305 §B218 (甲): 退数搬到那一处出入口; 清完回指之后 __comObj 为空 ⇒ 到 0 就走"自己销毁"那一支,
    // 与前面手抄的三行同语义 (>0 就直接回, 还有薄引用在世).
    c_.emitLine(clsStruct + "_Release(me);");
    c_.dedent();
    c_.emitLine("}");
}

// ============================================================
// B05: 接口变量的作用域末尾释放
// ============================================================

// 账 #305 §B218 (甲) 腿④: 「这一枚过程级槽在正常出口放手」那张表的唯一登记口。
// 夹 = 返回槽: 那一枚引用是**交出去**的那一枚, 出口再放手 = 调用方刚接手的东西当场读已释放块
// (#445 钉住的那条夹, 原来只有接口一档写着一句注释, 现在两档共用同一句代码)。
void CCodeGen::trackSlotForRelease(const std::string& expr, bool isView,
                                   const std::string& clsName,
                                   const std::string& arrayName, bool freeCarrier) {
    if (!currentReturnVar_.empty() && expr == currentReturnVar_) return;
    for (const SlotRelease& s : slotsToRelease_) {
        if (s.expr == expr) return;
    }
    SlotRelease rec;
    rec.expr = expr;
    rec.isView = isView;
    rec.clsName = clsName;
    rec.arrayName = arrayName;
    rec.isCarrier = !arrayName.empty();
    rec.freeCarrier = freeCarrier;
    slotsToRelease_.push_back(rec);
}

// 账 #305 §B218 (甲) 第一百五十二刀: 「这一枚过程级**数组载体**在正常出口要把元素交回」那张表的
// 唯一登记口。闸一条不加、一条不少:
//  - 认不出元素类 (`classOfArrayCarrier` 空串 = `As Object` 包装器 / `As Long` 裸值 / 外部类 / 点链成员数组)
//    压根不登记 = 维持今天的裸着出作用域;
//  - Static 那一档 (`Static a() As C`, 或整个 `Static Sub`) 由调用方 (cgen_localdecl.cpp 那两条臂)
//    拦在门外 —— 那一格跨调用活着, 出口放手 = 下次进来读已释放块 (与类槽那一档同一句理由)。
// 账 #305 §B218 (甲) 第一百五十六刀: 这一口从「只认工程类元素的载体」放宽到**每一档局部数组**——
// 元素认不出就一声不交回, 但载体那一块该销毁还是要销毁 (Destroy1D 自己会按 elemType 收尾
// bstr / variant 元素)。多出来的那一档只有一条闸: 预扫说这枚载体在本过程里被整体用掉过
// (`Dim v: v = a` / `P(a)` / `For Each x In a` / `UBound(a)` 一律算) ⇒ 不登记 = 维持今天的漏。
void CCodeGen::trackArrayCarrierForRelease(const std::string& cName,
                                           const std::string& arrayName) {
    if (cName.empty() || arrayName.empty()) return;
    std::string lower = arrayName;
    for (char& ch : lower) ch = (char)tolower((unsigned char)ch);
    // 闸① (第一百五十六刀立, 第一百六十四刀改口径): 出口那一发拿的是 C 变量名, 只有声明在那一刻
    // 还在场的才算 —— 名单由 scanLocalArrayEscapes 照提升那一条规则填 (嵌在块里那几档的声明现在
    // 住在序言), 关掉提升就缩回「只有过程体直接子句」那一个答案 (门 #821 那十二发 C2065 归它挡)。
    if (exitVisibleArrayCarriers_.count(lower) == 0) return;
    // 闸②: 预扫说这枚载体在本过程里被整体用掉过 ⇒ 不销毁。
    trackSlotForRelease(cName, false, std::string(), arrayName,
                        escapedLocalArrays_.count(lower) == 0);
}

// 「这枚局部数组有没有被整体用掉」那一问的唯一答案。判据按**节点地址**认, 不按文本:
// traverseExpr 先 visit 父节点再下钻 ⇒ 在父节点里把「这一格不算逃逸」的子节点指针记下,
// 到 IdentifierExpr 那一访按指针放行。认不出的形状默认算逃逸 (fail-safe: 少发一发销毁)。
// 算逃逸的 = 其余每一次裸出现: 右值装箱 (`v = a`, 装的正是这枚载体)、实参 (`P(a)` / `UBound(a)`)、
// 零实参的整枚数组那一形 (`v = a()`, 载体当场交给调用方)、`For Each x In a` 的集合、成员链的宿主 (`a.Field`)。
// 放行三格, 每一格都有一发实测撑着:
//   ① 带实参的元素访问 `a(i)` 里那个被调者 —— 只碰一格, 载体没交出去;
//   ② 第一百五十八刀: **整数组赋值的左值** `a = ...` —— 旧话是「刻意不放行」, 那时那一发是
//      `d = b;` 裸指针赋值 (两边共用同一枚载体, 出口按自己的名义去销毁 = 把别人还活着的东西拆了);
//      现在那一行走 vb6_ArrayAssign1D = 复制, 槽里握的是自己新克隆的那一枚 ⇒ 出口收尾才是正确答案,
//      不放行的实测代价 = 第一百五十六刀那把尺的 B225-take 从 leak=0 涨成 leak=1;
//   ③ 第一百五十九刀: **整数组赋值的右值** —— 复制不吃源的载体 (`vb6_ArrayAssign1D_Cb` calloc 新描述符,
//      元素按档各建副本), 源那枚仍是这一站自己的东西 ⇒ 出口该收尾; 不放行的实测代价 = 同一把页尺的
//      「ReDim 一枚局部数组、只把它复制给另一枚局部数组」那一格 leak=1。放行条件问的是**左值**在
//      exitVisibleArrayCarriers_ 里 —— 左值是 Variant 槽那一发是装箱 (②③都不免), 左值是形参 / 模块级数组
//      那两档压根没有出口收尾, 免不免都不发东西。
// 三格只免**那一格**: 同一个名字在别处任何裸出现照旧算, 所以 `v = a` 之后再 `a = b` 仍然不放手
// (夹具 B263 那一档钉的就是它)。
namespace {
struct EscapeScan : public ASTVisitor {
    std::set<const ASTNode*> exempt;
    std::map<std::string, int> bare;
    // 第一百五十九刀: 右值那一格要不要放行, 问的是「左值是不是本过程的局部数组」——
    // 这张名单由 scanLocalArrayEscapes 先填好, 指针递进来, 不再另起一趟扫描。
    const std::set<std::string>* exitVisibleArrayCarriers = nullptr;
    static std::string low(const std::string& s) {
        std::string r = s;
        for (char& ch : r) ch = (char)tolower((unsigned char)ch);
        return r;
    }
    void visit(IdentifierExpr& node) override {
        if (exempt.count(&node)) return;
        bare[low(node.name)]++;
    }
    // 第一百五十八刀: **整数组赋值的左值那一格不再算逃逸**。
    // 第一百五十六刀把它算进去的理由是「发完 `l = mg` 之后本地槽里握的是别人的载体」——
    // 那是别名时代的事实。这一刀把那一发改成复制 (vb6_ArrayAssign1D 交一枚新载体) ⇒ 左值就是
    // 那枚新载体的唯一主人, 出口销毁它既安全又是唯一不漏的去处。不豁免的实测代价 (本机两台,
    // 第一百五十六刀那把尺): B225-take 从 leak=0 涨成 leak=1 = 每一发整数组赋值净漏一枚载体。
    // 只豁免左值**这一格**: 同一个名字在别处的任何裸出现 (右值装箱 / 实参位 / `A()` 整枚交给
    // 调用方) 照旧算逃逸。标量目标那格豁免是空转 —— 这张名单只被数组那一档消费。
    void visit(AssignmentStmt& node) override {
        if (node.target && node.target->kind == ASTNodeKind::IdentifierExpr) {
            exempt.insert(node.target.get());
            // 第一百五十九刀 (第二格): **整数组赋值的右值那一格也放行**。
            // 复制这一发不吃源的载体: `vb6_ArrayAssign1D_Cb`  calloc 一枚新描述符, 元素按档各自
            // 建副本 (bstr 走 SysAllocString, variant 先 memset 再 VariantCopy, UDT 交生成的
            // vb6_udtcpy_*), 发完源那枚仍然是它自己那一站的东西, 没有第二个持有者 ⇒ 出口该收尾。
            // 不放的代价实测 = 第一百五十九刀那把页尺里「ReDim 一枚局部数组、只把它当右值复制走」
            // 那一格 leak=1 = 每调用净漏一枚 (与 W2/W3 那两格同源)。
            // 只在**左值也是本过程的局部数组**时放行: 左值是 Variant 槽 (`mv = a`) 那一发是装箱,
            // 装的正是这枚载体 ⇒ 照旧算逃逸 (夹具 B263 钉的那一发); 左值是标量同理不走复制。
            if (exitVisibleArrayCarriers && node.value &&
                node.value->kind == ASTNodeKind::IdentifierExpr) {
                auto& tgt = static_cast<const IdentifierExpr&>(*node.target);
                if (exitVisibleArrayCarriers->count(low(tgt.name)))
                    exempt.insert(node.value.get());
            }
        }
    }
    void visit(IndexOrCallExpr& node) override {
        // `a(i)` 那一形: 被调者就是数组本身, 只碰一格, 载体不交出去 ⇒ 放行。
        // `a()` (零实参) 在 VB6 里是**整枚数组**那一形 —— 函数返回数组、整数组赋值的右边都写它,
        // 载体当场交给了调用方, 和 `P(a)` 同一族 ⇒ 这一格不放行, 按裸出现算逃逸。
        if (node.callee &&
            (!node.positional.empty() || !node.named.empty())) {
            exempt.insert(node.callee.get());
        }
    }
};

// 「这条声明是不是数组载体」+ 是就把名字记进名单 (逗号声明要展开成子声明逐条看)。
// 以什么范围调用它 = 闸① 那一条规则, 只住在 scanLocalArrayEscapes; 这一位只管认形状、收名字。
void collectArrayCarrierNames(const Decl* d, std::set<std::string>& out) {
    if (!d) return;
    if (d->kind == ASTNodeKind::MultiDecl) {
        for (const DeclPtr& child : static_cast<const MultiDecl&>(*d).declarations) {
            collectArrayCarrierNames(child.get(), out);
        }
        return;
    }
    if (d->kind != ASTNodeKind::VariableDecl) return;
    const VariableDecl& var = static_cast<const VariableDecl&>(*d);
    // 「这条声明是不是数组」问的正是发码那两条臂各自的条件 (定长/多维看 dimensions,
    // 动态看 isDynamicArray —— `Dim d()` 的 dimensions 是空的), 少一条就把动态那一档漏回今天的漏。
    if (var.dimensions.empty() && !var.isDynamicArray) return;
    std::string lower = var.name;
    for (char& ch : lower) ch = (char)tolower((unsigned char)ch);
    out.insert(lower);
}
} // namespace

void CCodeGen::scanLocalArrayEscapes(const StmtList& body) {
    escapedLocalArrays_.clear();
    exitVisibleArrayCarriers_.clear();
    // 直接子句那几条: 提升开没开, 它们的 C 变量在出口那一刻都在场 (第一百五十六刀的原口径)。
    for (const StmtPtr& st : body) {
        if (st && st->kind == ASTNodeKind::LocalDeclStmt) {
            collectArrayCarrierNames(static_cast<const LocalDeclStmt&>(*st).decl.get(),
                                      exitVisibleArrayCarriers_);
        }
    }
    // 第一百六十四刀: 提升开着 ⇒ 嵌在块里那几档也算在场 —— 动态那一档整条声明就在序言 (Fix 086),
    // 定长与多维那一档由本刀把声明提上来 (hoistLocalDecls 写 preludeArrayCarriers_, 创建留原位)。
    // 走的是提升那一位主人自己的递归 (collectLocalDeclStmts), 不是第二把尺; C3_NO_HOIST 关掉提升
    //  ⇒ 这一问跟着答否, 门 #821 那一形立刻回来。
    if (hoistingLocalDeclsEnabled()) {
        std::vector<LocalDeclStmt*> decls;
        collectLocalDeclStmts(body, decls);
        for (auto* d : decls) {
            if (d && d->decl) {
                collectArrayCarrierNames(d->decl.get(), exitVisibleArrayCarriers_);
            }
        }
    }
    EscapeScan scan;
    scan.exitVisibleArrayCarriers = &exitVisibleArrayCarriers_;
    for (const auto& st : body) {
        if (st) traverseStmt(*st, scan);
    }
    for (const auto& kv : scan.bare) {
        escapedLocalArrays_.insert(kv.first);
    }
}

void CCodeGen::trackClassSlotForRelease(const std::string& cName, const std::string& clsName) {
    trackSlotForRelease(cName, false, clsName);
}

void CCodeGen::trackIvrefLocalForRelease(const std::string& cName) {
    trackSlotForRelease(cName, true, std::string());
}

// 账 #305 §B218 (甲) 腿④ 的第二道夹: 这一格的**地址**交出去过 —— 写它的那一站在计数系统外面
// (一枚外部 Declare 的 `As Any` ByRef 实参, 例如 VB6 运行时的 __vbaObjSetAddref)。本过程从没
// 经自己的主人手里拿过这一枚引用, 出口就不能放手: C3 的那枚桩对**工程内实例**刻意不 AddRef
// (结构体首字段不是 vtable, 加了就 0xC0000005 —— Fix 164x2), 于是"出口退一枚"= 无对价的退。
// 门 #763 实测 = FlexGridX86 起窗期: `Dim This As VBFlexGrid` + `FlexObjSetAddRef This, uId`
// 之后那一句出口 Release 把还活着的控件当场拆掉, 主窗口压根不出来。
// 只撤工程类那一档: 视图那一档的引用是本站 QueryInterface 出来的, 与外部写者无关。
void CCodeGen::dropClassSlotRelease(const std::string& expr) {
    for (auto it = slotsToRelease_.begin(); it != slotsToRelease_.end();) {
        if (!it->isView && it->expr == expr) it = slotsToRelease_.erase(it);
        else ++it;
    }
}

// 账 #305 §B221: `x->vt-><slot>(x, args)` 这一发射形状只许有一位作者 —— com_bind 的
// ivref 臂与「接收者是视图值的调用」那一站 (形②) 都调它。槽解析不出来 = 这个成员不在
// 这张面上，返回空串让调用点继续往下落 (契约检查另有诊断)。
std::string CCodeGen::ivrefVtableCall(const std::string& ifaceName, const std::string& member,
                                      const std::string& recv,
                                      const std::vector<std::string>& args) {
    const IfaceView* v = ivLookupIface(ifaceName);
    if (!v) return std::string();
    const std::string slot = ivSlotKeyForMember(*v, member);
    if (slot.empty()) return std::string();
    std::string call = recv + "->vt->" + cIdent(slot) + "(" + recv;
    for (const std::string& a : args) call += ", " + a;
    call += ")";
    return call;
}

// 账 #305 §B221: B06a 那条视图→视图的 QueryInterface 发射挪成一位主人 —— 原来这段
// 内联在 cgen_setlet_set_prop.inc 的按名臂里，而 RHS 是「返回工程内接口的调用」那一形
// (#448 的形①) 也要同一条腿；留在站点里就会长出第二个答案。
// 计数口径不变: QI 成功即已 AddRef (失败得 NULL = Nothing)，旧值当场 Release。
void CCodeGen::emitIvrefQiSet(const std::string& target, const std::string& ifaceName,
                              const std::string& valueExpr, bool rhsTransferred) {
    const std::string ivCType = ivrefCType(ifaceName);
    // valueExpr 在这个块里出现三次 (判空 / QI 的接收者 / QI 的首参)。裸标识符无所谓，
    // 但形①交进来的是 `MakeDirect()` 那种调用 ⇒ 三次求值 = 三枚实例、三笔引用
    // (实测)。同一 Fix 090q 那条通道：先落一枚临时量，后面三处都用它。
    const bool plainVal = !valueExpr.empty()
        && valueExpr.find_first_not_of(
            "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_")
           == std::string::npos;
    std::string ivVal = valueExpr;
    if (!plainVal) {
        ivVal = "__ivrhs" + std::to_string(ivrefSetSeq_);
        c_.addPending(ivCType + " " + ivVal + " = " + valueExpr + ";");
    }
    const std::string ivOld = "__ivold" + std::to_string(ivrefSetSeq_);
    const std::string ivGot = "__ivgot" + std::to_string(ivrefSetSeq_);
    ivrefSetSeq_++;
    c_.emitLine("{  /* tB Interface B06a: Set " + target + " = QueryInterface */");
    c_.indent();
    c_.emitLine(ivCType + " " + ivOld + " = " + target + ";");
    c_.emitLine("void* " + ivGot + " = NULL;");
    c_.emitLine("if (" + ivVal + ") " + ivVal + "->vt->QueryInterface(" + ivVal +
                ", vb6_iv_iid_" + cIdent(ifaceName) + ", &" + ivGot + ");");
    c_.emitLine(target + " = (" + ivCType + ")" + ivGot + ";");
    c_.emitLine("if (" + ivOld + ") " + ivOld + "->vt->Release(" + ivOld + ");");
    // 账 #305 §B218 第一百三十四刀 (甲) 腿② 视图档: 右值那枚临时量是**接管来的** (被调方 =
    // 返回工程内接口的过程, 它自己的返回槽在 腿④ 不退 ⇒ 交出来就带一枚对价; 契约同 刀3/刀4,
    // 问句也只那一位主人 —— rhsTransfersProjectRef)。本站 QI 出来的那一枚另算 (QI 自己 AddRef),
    // 所以接管来的那一枚当场交回, 不退 = 每调用净 +1 = 永不销毁。
    // 借来的那一档 (裸标识符右值 / New 那一站) 天然 rhsTransferred=false ⇒ 一发都不补。
    if (rhsTransferred) {
        c_.emitLine("if (" + ivVal + ") " + ivVal + "->vt->Release(" + ivVal +
                    ");  /* 账 #305 (甲) 腿②: 交回接管来的那一枚 */");
    }
    c_.dedent();
    c_.emitLine("}");
}

// 出口那一腿与 ansiTempsToFree_ 在同一位置调用（过程统一出口 vb6_proc_exit，
// Exit Sub/Function 自 c97e65bd 起也走那条 goto ⇒ 这一族不漏清理）。
void CCodeGen::emitIvrefScopeRelease() {
    for (const SlotRelease& s : slotsToRelease_) {
        // 第三档 (第一百五十二刀): 数组载体 —— 交给那位数组主人, 出口自己不问臂
        // (`classOfArrayCarrier` + `carrierIsMultidim` 两问都住在主人里)。
        // 刻意**只出厂、不销毁载体**: 载体的形状在出口这一发只有「按代码顺序」那一个答案,
        // 出厂那一支答错的下限是静默不漏 (第一百五十刀那条 elemType 闸), 而 free 那一支答错
        // 就是拿错步长解引用再 free —— 载体级的销毁 (= 内存那半, 全元素档一起) 是另一格机制。
        if (s.isCarrier) {
            const bool nd = carrierIsMultidim(s.arrayName);
            emitClassArrayElemHandback(s.expr, s.arrayName,
                                       s.freeCarrier ? "作用域出口逐枚交回 (载体随后销毁)"
                                                     : "作用域出口逐枚交回 (载体不销毁)");
            // 第一百五十二刀刻意留在册的那一半 (第一百五十五刀量完 = 每次调用实测漏 316/334 字节):
            // 出口销毁载体。臂由同一位主人挑 (`carrierIsMultidim` 读的是当前载体那一档, 与 152 那条
            // 「表在出口这一刻反映的就是当前载体」的实测同前提); 双 free 那两条封死 —— Erase 四支都
            // 要么置 NULL 要么保留载体, 而非 Preserve 的 ReDim 先销毁旧载体再赋新 (读数在台账)。
            if (s.freeCarrier) {
                const std::string kill = nd
                    ? ("vb6_SafeArrayDestroyND((vb6_SafeArrayND*)" + s.expr + ");")
                    : ("vb6_SafeArrayDestroy1D(" + s.expr + ");");
                c_.emitLine(kill +
                            "  /* 账 #305 (甲) 腿⑤: 作用域出口销毁载体 (本过程内它从没整体用掉) */");
            }
            continue;
        }
        if (s.isView) {
            c_.emitLine("if (" + s.expr + ") " + s.expr + "->vt->Release(" + s.expr +
                        ");  /* tB Interface B05: scope-exit release */");
        } else {
            c_.emitLine("if (" + s.expr + ") vb6_cls_" + cIdent(s.clsName) +
                        "_Release(" + s.expr +
                        ");  /* 账 #305 (甲) 腿④: scope-exit release */");
        }
    }
    slotsToRelease_.clear();
}

} // namespace vb6c3
