#include "semantics/symbol_table.hpp"
#include <algorithm>

namespace vb6c3 {

// ============================================================
// Scope
// ============================================================

bool Scope::define(std::unique_ptr<Symbol> sym, const std::string& keyOverride) {
    // 使用storageKey作为键: Property Get/Let/Set加后缀区分同名共存
    // Fix 103: Type/Enum 与过程同名时由 SymbolTable::define 传入 "<name>$ty" 覆盖键.
    std::string key = keyOverride.empty() ? sym->storageKey() : keyOverride;
    auto it = symbols_.find(key);
    if (it != symbols_.end()) {
        return false;  // 重定义
    }
    symbols_[key] = std::move(sym);
    return true;
}

Symbol* Scope::lookupLocal(const std::string& name) const {
    std::string lower = Symbol::toLower(name);
    // 先查找非Property符号 (key == lowerName)
    auto it = symbols_.find(lower);
    if (it != symbols_.end()) {
        return it->second.get();
    }
    // 再查找Property Get (key == lowerName + "$pg") - 默认返回Property Get
    it = symbols_.find(lower + "$pg");
    if (it != symbols_.end()) {
        return it->second.get();
    }
    // 再查找Property Let
    it = symbols_.find(lower + "$pl");
    if (it != symbols_.end()) {
        return it->second.get();
    }
    // 再查找Property Set
    it = symbols_.find(lower + "$ps");
    if (it != symbols_.end()) {
        return it->second.get();
    }
    // Fix 103: Type/Enum 与过程同名时的类型符号 (key == lowerName + "$ty")
    it = symbols_.find(lower + "$ty");
    if (it != symbols_.end()) {
        return it->second.get();
    }
    return nullptr;
}

Symbol* Scope::lookupLocalByKind(const std::string& name, SymbolKind kind) const {
    std::string lower = Symbol::toLower(name);
    if (isPropertyKind(kind)) {
        // Property: 使用带后缀的键精确查找
        std::string key = lower;
        switch (kind) {
            case SymbolKind::PropertyGet:  key += "$pg"; break;
            case SymbolKind::PropertyLet:  key += "$pl"; break;
            case SymbolKind::PropertySet:  key += "$ps"; break;
            default: break;
        }
        auto it = symbols_.find(key);
        if (it != symbols_.end()) {
            return it->second.get();
        }
        return nullptr;
    }
    // 非Property: 直接用lowerName
    auto it = symbols_.find(lower);
    if (it != symbols_.end() && it->second->kind == kind) {
        return it->second.get();
    }
    // Fix 103: 同名过程占用了裸键时, Type/Enum 符号存放在 <name>$ty 下
    auto itTy = symbols_.find(lower + "$ty");
    if (itTy != symbols_.end() && itTy->second->kind == kind) {
        return itTy->second.get();
    }
    return nullptr;
}

Symbol* Scope::lookup(const std::string& name) const {
    // 先在本作用域查找
    if (auto sym = lookupLocal(name)) {
        return sym;
    }
    // 递归向父作用域查找
    if (parent_) {
        return parent_->lookup(name);
    }
    return nullptr;
}

// ============================================================
// SymbolTable
// ============================================================

SymbolTable::SymbolTable(Diagnostics& diag)
    : diag_(diag)
{
    // 创建模块级作用域
    auto modScope = std::make_unique<Scope>(ScopeKind::Module);
    moduleScope_ = modScope.get();
    current_ = moduleScope_;
    scopes_.push_back(std::move(modScope));
}

void SymbolTable::pushScope(ScopeKind kind) {
    auto scope = std::make_unique<Scope>(kind, current_);
    current_ = scope.get();
    scopes_.push_back(std::move(scope));
}

void SymbolTable::popScope() {
    if (current_ && current_->parent()) {
        current_ = current_->parent();
    }
}

bool SymbolTable::define(std::unique_ptr<Symbol> sym) {
    if (!current_) return false;

    std::string lowerName = sym->lowerName;
    SourceLocation loc = sym->location;

    // P6.3 + P24-05: COM内置类型与用户变量可同名 (不同namespace)
    // 如果新符号是builtin的ComClass/ComInterface/ComModule/ComGlobalNs, 且已有同名符号, 跳过不报错
    // 如果已有符号是ComClass/ComInterface/ComModule, 且新符号是用户变量, 也允许覆盖
    // Fix 018: builtin EnumMember 同样跳过 (COM 类型库枚举成员跨 enum 同名时 first-wins, 如 MSHTML)
    if (sym->isBuiltin && (sym->kind == SymbolKind::ComClass || sym->kind == SymbolKind::ComInterface
                         || sym->kind == SymbolKind::ComModule || sym->kind == SymbolKind::ComGlobalNs
                         || sym->kind == SymbolKind::EnumMember)) {
        auto* existing = current_->lookupLocal(lowerName);
        if (existing) {
            // 已有同名符号, 跳过COM类型注册 (避免重复注册)
            return true;  // 不报错
        }
    }
    // 如果已有同名ComClass/ComInterface/ComModule, 新符号是用户变量, 先移除旧的
    {
        auto* existing = current_->lookupLocal(lowerName);
        if (existing && (existing->kind == SymbolKind::ComClass || existing->kind == SymbolKind::ComInterface
                      || existing->kind == SymbolKind::ComModule || existing->kind == SymbolKind::ComGlobalNs)
            && sym->kind != SymbolKind::ComClass && sym->kind != SymbolKind::ComInterface
            && sym->kind != SymbolKind::ComModule && sym->kind != SymbolKind::ComGlobalNs) {
            // 移除COM类型符号, 允许用户变量覆盖
            current_->symbols_.erase(lowerName);
        }
    }

    // 允许用户定义的 Sub/Function/Variable/Constant/Property 覆盖内置符号
    // VB6合法: 类方法/变量名可与内置函数同名 (如 Timer, FormatDateTime, LTrim, App)
    // Fix 018: 用户 EnumMember 也可覆盖内置 COM 枚举成员 (如 cDatabase.cls 自定义 adEmpty
    // 覆盖 ADO DataTypeEnum 的 adEmpty — 用户项目声明优先于引用的类型库)
    // Fix 048: Property Get/Let/Set 同样覆盖内置函数 — 否则类内属性与内置函数同名时
    // (如 cByteBuffer.IsEmpty / cDialog.Filter 对内置 IsEmpty/Filter), 内置 Function 符号
    // (storageKey=无后缀名) 与属性变体 ($pg/$pl) 共存, lookupLocal/lookupModuleByKind
    // 按无后缀键优先返回内置符号, 导致 TypeLib 出现幽灵 Function 变体 (同名同 dispid,
    // SetFuncAndParamNames 返回 TYPE_E_TYPEMISMATCH 0x800280EC)。
    {
        auto* existing = current_->lookupLocal(lowerName);
        if (existing && existing->isBuiltin
            && !sym->isBuiltin
            && (sym->kind == SymbolKind::Sub
                || sym->kind == SymbolKind::Function
                || sym->kind == SymbolKind::Variable
                || sym->kind == SymbolKind::Constant
                || sym->kind == SymbolKind::EnumMember
                || sym->kind == SymbolKind::PropertyGet
                || sym->kind == SymbolKind::PropertyLet
                || sym->kind == SymbolKind::PropertySet)) {
            // 移除内置符号, 允许用户符号替换
            // Fix 216: 但别把形参表一起扔掉 —— 内置函数的形参只在这里, 顶掉之后本模块
            // 再没有第二份来源。`VBA.<成员>` 这种显式限定调用在 VB6 里一定指内置函数,
            // 丢了形参表就会发成 0 实参 (`vb6_Year()` → C2198, 见 .hpp 里 lookupShadowedBuiltin
            // 的注释)。这里把符号整体**转移**到留档表 (move 而非拷贝: 所有权换手,
            // existing 指针仍有效), 再按原键 erase。
            auto itShadowed = current_->symbols_.find(existing->storageKey());
            if (itShadowed != current_->symbols_.end()) {
                if (shadowedBuiltins_.find(existing->storageKey()) == shadowedBuiltins_.end()) {
                    shadowedBuiltins_.emplace(existing->storageKey(), std::move(itShadowed->second));
                }
                current_->symbols_.erase(itShadowed);
            } else {
                current_->symbols_.erase(existing->storageKey());
            }
        }
    }

    // Fix 047: Allow local declarations to overwrite pre-registered external symbols.
    // Cross-module Enum/UDT types are pre-registered before Pass 1 so that parameter
    // types resolve correctly. When the local module's own Pass 1 registers the same
    // Enum/UDT, the external symbol must be replaced by the local definition.
    {
        auto* existing = current_->lookupLocal(lowerName);
        if (existing && existing->isExternal && !sym->isExternal) {
            current_->symbols_.erase(existing->storageKey());
        }
    }

    // Property Get/Let/Set允许同名共存 (VB6合法: Property Get Name + Property Let Name)
    // Scope::define已用storageKey区分, 这里只需确认不报错即可

    // Fix 103: Type/Enum 与过程同名共存 — VB6 合法 (类型名与过程名属不同命名空间).
    // 实例 (Charts 2020 ucProgressCircular/ppProgressCircular.pag):
    //   Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" _
    //       (pChoosecolor As CHOOSECOLOR) As Long
    //   Private Type CHOOSECOLOR ... End Type
    // 二者 lowerName 相同, 原先触发 SemDuplicateDeclaration 直接中止编译.
    // 处理: 恰好一方是类型 (UserDefinedType/EnumType)、另一方是过程时, 类型符号改用
    // "<name>$ty" 存储键; 查找侧见 Scope::lookupLocal / lookupLocalByKind 的 $ty 回退.
    // 仅在该冲突真实存在时改写键, 无冲突的模块行为完全不变.
    std::string keyOverride;
    {
        const bool newIsType = (sym->kind == SymbolKind::UserDefinedType ||
                                sym->kind == SymbolKind::EnumType);
        auto isTypeKind = [](SymbolKind k) {
            return k == SymbolKind::UserDefinedType || k == SymbolKind::EnumType;
        };
        auto* existingBare = current_->lookupLocal(lowerName);
        // lookupLocal 已含 $ty 回退; 这里只关心裸键上的占用者
        auto itBare = current_->symbols_.find(lowerName);
        Symbol* bareOwner = (itBare != current_->symbols_.end())
                                ? itBare->second.get() : nullptr;
        (void)existingBare;
        const std::string tyKey = lowerName + "$ty";
        const bool tyKeyFree = current_->symbols_.find(tyKey) == current_->symbols_.end();
        if (bareOwner && tyKeyFree && isTypeKind(bareOwner->kind) != newIsType) {
            // 裸键被"另一类"符号占用 → 类型符号让位到 $ty
            if (newIsType) {
                keyOverride = tyKey;
            } else {
                // 已有类型占用裸键, 新来的是过程: 把类型搬到 $ty, 过程占裸键
                auto moved = std::move(current_->symbols_[lowerName]);
                current_->symbols_.erase(lowerName);
                current_->symbols_[tyKey] = std::move(moved);
            }
        }
    }

    // 重载分组 (tB 式, O1): 语义层已把"可重载过程"的签名指纹写进 overloadFp.
    // 裸键上已有**同 kind** 且指纹不同的过程 → 本符号挂 "<name>$ov$<fp>" 变体键;
    // 指纹相同 (或 owner 非过程 / owner 未参与分组如类模块成员) → 落到下面原路径,
    // 由 define 冲突报 SemDuplicateDeclaration, 旧行为不变.
    if (keyOverride.empty() && !sym->overloadFp.empty() &&
        (sym->kind == SymbolKind::Sub || sym->kind == SymbolKind::Function)) {
        auto itOwn = current_->symbols_.find(lowerName);
        Symbol* owner = (itOwn != current_->symbols_.end()) ? itOwn->second.get() : nullptr;
        if (owner && owner->kind == sym->kind &&
            !owner->overloadFp.empty() && owner->overloadFp != sym->overloadFp) {
            sym->isOverloadVariant = true;  // 键由 storageKey() 派生 (lower$ov$fp)
            if (current_->define(std::move(sym))) {
                owner->ovlCount++;  // owner 指向堆上 Symbol, map 插入不使其失效
                return true;
            }
            // 变体键撞车: 与非 head 变体同签名 (第三个同指纹声明) → 重复声明
            diag_.error(DiagnosticID::SemDuplicateDeclaration, loc,
                "重复声明: '" + lowerName + "' (重载组内同签名)");
            return false;
        }
    }

    // 类模块允许与类同名的成员 (View.cls 的 `Public Property Get View()` /
    // PopupMenu.cls 的 `Public Sub PopupMenu(...)`) —— VB6 合法。类自符号
    // (SymbolKind::Class, semantic_analyzer 预先 define 进模块作用域) 让位给成员:
    // 裸键换成成员符号。
    //
    // Fix <vbeclipse> rev17: 让位实现从 `erase` 改为**换键到 <name>$ty** (与下面
    // 上面 Fix 103 的 Type/Enum 让位同款)。原先直接 erase 的注释写着"类本身仍可经
    // lookupModule/全局作用域解析", 但那两个出口都建立在"符号还在 moduleScope 里"
    // 这个前提上 —— 而真正消费它的两条路是**直接遍历 moduleScope()->symbols()**:
    //   · getPublicSymbols()  → 跨模块导出 (driver_crossmod 的 globalPublicSyms)
    //   · cgen_util_dllentry_collect.inc 的 coclass 收集
    // 类符号一删, 这个类就再也进不了 coclass 表, 运行期连锁:
    //   vb6_FindCoClassDesc("<类名>") 返 NULL
    //   → vb6_ComObject_FromInstance(NULL, inst) 返 NULL
    //   → vb6_VariantObject(NULL) 产出 vt=VT_DISPATCH / pdispVal=NULL 的**空壳
    //     Variant** (该函数无条件写 vt=dispatch)
    //   → 槽位读回 `xxx Is Nothing` 恒真 / 取回对象即 Nothing
    //   → ucFolder.AddView(ByRef View As View) 收到 NULL
    //   → `View.ViewId` 对 NULL 解引用 (av read 0x4) —— play78 崩点。
    // 实测 play78: coclass 表 14 项, 缺的正是**全工程唯一两个有同名成员的类**
    // View / PopupMenu; Folder/List/Perspective 等另外 14 个类都在表里。
    // 换键后: 裸键仍让给成员 (lookupLocal 的语义不变), 类符号留在
    // <name>$ty, 由既有 $ty 回退 (Scope::lookupLocal / lookupLocalByKind) 找到。
    {
        auto itSelf = current_->symbols_.find(lowerName);
        if (itSelf != current_->symbols_.end() && itSelf->second
            && itSelf->second->kind == SymbolKind::Class
            && (sym->kind == SymbolKind::Sub || sym->kind == SymbolKind::Function
                || isPropertyKind(sym->kind))) {
            diag_.warn(DiagnosticID::SemDuplicateDeclaration, loc,
                "类成员与类同名: '" + lowerName + "' (VB6 允许, 成员让类符号让位)");
            // 让位 = 把类符号挪到类型槽 <name>$ty (不删除), 裸键空出来给成员。
            // 槽已被占 (同名 Type/Enum) 时退回旧行为 (删), 属极端罕见组合。
            const std::string clsKey = lowerName + "$ty";
            if (current_->symbols_.find(clsKey) == current_->symbols_.end()) {
                auto moved = std::move(itSelf->second);
                current_->symbols_.erase(itSelf);
                current_->symbols_[clsKey] = std::move(moved);
            } else {
                current_->symbols_.erase(itSelf);
            }
        }
    }

    if (!current_->define(std::move(sym), keyOverride)) {
        diag_.error(DiagnosticID::SemDuplicateDeclaration, loc,
            "重复声明: \x27" + lowerName + "\x27");
        return false;
    }
    return true;
}

Symbol* SymbolTable::lookup(const std::string& name) const {
    if (!current_) return nullptr;
    return current_->lookup(name);
}

Symbol* SymbolTable::lookupLocal(const std::string& name) const {
    if (!current_) return nullptr;
    return current_->lookupLocal(name);
}

Symbol* SymbolTable::lookupLocalByKind(const std::string& name, SymbolKind kind) const {
    if (!current_) return nullptr;
    return current_->lookupLocalByKind(name, kind);
}

Symbol* SymbolTable::lookupModule(const std::string& name) const {
    if (!moduleScope_) return nullptr;
    return moduleScope_->lookupLocal(name);
}

Symbol* SymbolTable::lookupModuleByKind(const std::string& name, SymbolKind kind) const {
    if (!moduleScope_) return nullptr;
    return moduleScope_->lookupLocalByKind(name, kind);
}

Symbol* SymbolTable::lookupModuleOverloadByLoc(const std::string& name,
                                                const SourceLocation& loc) const {
    if (!moduleScope_) return nullptr;
    std::string lower = Symbol::toLower(name);
    Symbol* head = nullptr;
    for (const auto& [key, sym] : moduleScope_->symbols()) {
        if (key == lower) {
            head = sym.get();
            if (head->location.line == loc.line && head->location.column == loc.column)
                return head;
        } else if (key.size() > lower.size() + 4 &&
                   key.compare(0, lower.size(), lower) == 0 &&
                   key.compare(lower.size(), 4, "$ov$") == 0) {
            if (sym->location.line == loc.line && sym->location.column == loc.column)
                return sym.get();
        }
    }
    return head;  // 无位置命中 → 裸键 head (零重载模块 = 旧 lookupModule 行为)
}

std::vector<Symbol*> SymbolTable::lookupModuleOverloads(const std::string& name) const {
    std::vector<Symbol*> out;
    if (!moduleScope_) return out;
    std::string lower = Symbol::toLower(name);
    Symbol* head = nullptr;
    for (const auto& [key, sym] : moduleScope_->symbols()) {
        if (key == lower) { head = sym.get(); }
        else if (key.size() > lower.size() + 4 &&
                 key.compare(0, lower.size(), lower) == 0 &&
                 key.compare(lower.size(), 4, "$ov$") == 0) {
            out.push_back(sym.get());
        }
    }
    if (head) out.insert(out.begin(), head);
    return out;
}

Symbol* SymbolTable::lookupModuleOverloadBySuffix(const std::string& name,
                                                  const std::string& suffix) const {
    if (!moduleScope_) return nullptr;
    std::string lower = Symbol::toLower(name);
    if (suffix.empty()) return moduleScope_->symbols().count(lower)
                                  ? moduleScope_->symbols().at(lower).get() : nullptr;
    auto it = moduleScope_->symbols().find(lower + suffix);
    return it != moduleScope_->symbols().end() ? it->second.get() : nullptr;
}

int SymbolTable::scopeDepth() const {
    int depth = 0;
    Scope* s = current_;
    while (s && s != moduleScope_) {
        depth++;
        s = s->parent();
    }
    return depth;
}

ScopeKind SymbolTable::currentScopeKind() const {
    return current_ ? current_->kind() : ScopeKind::Module;
}

bool SymbolTable::inProcedure() const {
    Scope* s = current_;
    while (s) {
        if (s->kind() == ScopeKind::Procedure) return true;
        s = s->parent();
    }
    return false;
}

// --- 跨模块符号操作 ---

void SymbolTable::defineExternal(std::unique_ptr<Symbol> sym, bool replaceBuiltinCom) {
    // 在模块级作用域定义外部符号
    // 如果已存在同名符号（本地已有定义），跳过不覆盖
    if (!moduleScope_) return;
    std::string key = sym->storageKey();
    // Fix 103: 同名过程占用了裸键时, 外部 Type/Enum 改注入 "<name>$ty"
    // (与 SymbolTable::define 的冲突处理一致), 避免跨模块 Public Type
    // 与本地同名过程冲突后类型彻底不可见.
    if (sym->kind == SymbolKind::UserDefinedType || sym->kind == SymbolKind::EnumType) {
        if (moduleScope_->symbols_.find(key) != moduleScope_->symbols_.end()) {
            key = sym->lowerName + "$ty";
        }
    }
    auto it = moduleScope_->symbols_.find(key);
    if (it != moduleScope_->symbols_.end()) {
        // Fix <vbeclipse>: 工程类与**类型库内建 coclass/接口**同名时, 工程内定义优先
        // (VB6: 工程类遮蔽引用库同名 coclass). 消费模块里这个名字原本被类型库符号
        // 占着 (如 ScrRun 的 Folder), 其成员表是外部库的 → 工程类的 AddView/
        // ActiveViewId 等成员解析不到, 退回"数据字段"访问 (C2039/C2198).
        // 只有调用方显式请求 (driver_crossmod 的同名冲突分支) 才替换, 其余保持
        // "本地已有定义不注入" 的老语义.
        if (replaceBuiltinCom && sym->kind == SymbolKind::Class
            && !sym->isInterface && it->second && it->second->isBuiltin
            && (it->second->kind == SymbolKind::ComClass
                || it->second->kind == SymbolKind::ComInterface)) {
            it->second = std::move(sym);
            return;
        }
        // 本地已有定义，不注入外部符号
        return;
    }
    moduleScope_->symbols_[key] = std::move(sym);
}

std::vector<const Symbol*> SymbolTable::getPublicSymbols() const {
    std::vector<const Symbol*> result;
    if (!moduleScope_) return result;
    for (const auto& [key, sym] : moduleScope_->symbols()) {
        // O3: 重载变体随组注入跨模块 (storageKey 各自独立, head 占裸键)
        // Fix 084g: Friend 在 VB6 中表示"工程内可见", 与跨模块注入场景一致,
        // 因此 Friend 成员 (如 Friend Property Set fClient) 也应导出供其他模块引用
        if ((sym->access == AccessLevel::Public || sym->access == AccessLevel::Friend) && !sym->isBuiltin) {
            // 仅导出可被其他模块引用的符号类型
            if (sym->kind == SymbolKind::Sub ||
                sym->kind == SymbolKind::Function ||
                sym->kind == SymbolKind::Variable ||
                sym->kind == SymbolKind::Constant ||
                sym->kind == SymbolKind::Class ||
                sym->kind == SymbolKind::PropertyGet ||
                sym->kind == SymbolKind::PropertyLet ||
                sym->kind == SymbolKind::PropertySet ||
                sym->kind == SymbolKind::EnumType ||
                sym->kind == SymbolKind::EnumMember ||
                sym->kind == SymbolKind::UserDefinedType ||
                sym->kind == SymbolKind::DeclareSub ||
                sym->kind == SymbolKind::DeclareFunc) {
                result.push_back(sym.get());
            }
        }
    }
    return result;
}

// 账 §B97: 去重 + 升序。集合内容一处没变（仍是 isExternal 符号的 sourceModule 全集），
// 变的是**顺序的来源**：以前由 unordered_map 的桶布局决定，同一份 .vbp 换一台编译器
// 就换一种顺序，而下游三处码头（.c 的 include 段 / .h 的 crossmod include 段 / 入口点里
// vb6_mod_<X>_init() 的调用序）都按这个顺序逐行发码。排序后它至少是确定的；
// 语义上有意义的顺序（工程声明序）由 driver 用 modules_ 重排。
std::vector<std::string> SymbolTable::getExternalModuleNames() const {
    std::vector<std::string> result;
    if (!moduleScope_) return result;
    std::unordered_set<std::string> seen;
    for (const auto& [key, sym] : moduleScope_->symbols()) {
        if (sym->isExternal && !sym->sourceModule.empty() &&
            seen.insert(sym->sourceModule).second) {
            result.push_back(sym->sourceModule);
        }
    }
    std::sort(result.begin(), result.end());
    return result;
}

} // namespace vb6c3