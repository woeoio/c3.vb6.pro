// driver_crossmod.cpp - C3 编译器驱动: 跨模块符号链接
// 2026-09-17 从 src/driver/driver.cpp 纯搬移（逐行未改）：
//   原第 1192~1195 行
//   原第 1250~1397 行

#include "driver/driver.hpp"
#include "common/diagnostics.hpp"
#include "ast/ast.hpp"
#include "semantics/semantic_analyzer.hpp"
#include "semantics/interface_sig.hpp"   // tB Inherits B09b: ifaceLower (类链键口径)
#include <iostream>
#include <unordered_map>
#include <unordered_set>

namespace vb6c3 {

// === 跨模块符号链接 ===
// 遍历每个模块的符号表，查找未定义的标识符，在其他模块的Public符号中查找匹配
// 为匹配到的符号注入 isExternal=true + sourceModule 的外部符号

namespace {

// Fix <vbeclipse> rev17: 在模块作用域里找**本模块自有**的同名符号, 跳过跨模块注入
// 进来的外部副本 (`isExternal`)。键优先级照抄 Scope::lookupLocal:
// 裸键 → $pg → $pl → $ps → $ty (逐个试, 因为裸键可能被外部副本占着而本地成员在 $pg)。
//
// 用途: driver_crossmod 的**类符号**注入守卫。"本地已有定义"只应指本模块自己的定义:
//   · ucSplitBar.ctl 的 `Public Property Get Folder()` 是本地成员 → 必须挡住
//     `Class Folder` 注入。BASE 之所以正确纯属巧合: SymbolTable::define 在注册该成员时
//     会把内建同名 ComClass 从裸键上删掉 (symbol_table.cpp 的 "用户符号覆盖内置符号"),
//     裸键因此空出来, 名字级 lookupModule 便落到了 $pg 上的本地属性 → 判"已有定义"。
//     裸名 `Folder` 在 ucSplitBar 里必须解析成**本类的同名属性**
//     (`vb6_ucSplitBar_prop_get_Folder((void*)me)`), 一旦让 `Class Folder` 占了裸键,
//     生成的就是静态调用 + Empty 接收者 → C2440 (无法从 vb6_VARIANT 转换为
//     vb6_cls_Folder*)。
//   · 而 View.cls 的 `Property Get View` 经注入铺到**消费**模块的是外部副本,
//     它不能把 `Class View` 挡在门外 —— 否则这个类在任何模块里都不存在
//     (--dump-symbols: `Class View` 0 次, 其它 14 个工程类各 1 次)。
const Symbol* lookupOwnModuleSymbol(const SymbolTable& symTab, const std::string& name) {
    if (!symTab.moduleScope()) return nullptr;
    const std::string lower = Symbol::toLower(name);
    static const char* const kSuffixes[] = { "", "$pg", "$pl", "$ps", "$ty" };
    for (const char* suf : kSuffixes) {
        auto it = symTab.moduleScope()->symbols().find(lower + suf);
        if (it != symTab.moduleScope()->symbols().end() && it->second
            && !it->second->isExternal) {
            return it->second.get();
        }
    }
    return nullptr;
}

}  // namespace


bool Driver::runCrossModuleResolution() {
    if (modules_.size() != analyzers_.size()) return false;

    // Fix 013: 为每个模块计算模块名 — 使用 module.moduleName (VB_Name)
    // 而非文件名stem, 确保与 clsSym->name / mapTypeRef 生成的类型名一致
    std::vector<std::string> moduleBaseNames;
    for (const auto& module : modules_) {
        moduleBaseNames.push_back(module->moduleName);
    }

    // 收集每个模块导出的Public符号: [模块索引] -> vector<Symbol*>
    std::vector<std::vector<const Symbol*>> exportedSymbols(modules_.size());
    for (size_t i = 0; i < analyzers_.size(); i++) {
        exportedSymbols[i] = analyzers_[i]->symbolTable().getPublicSymbols();
    }

    // 构建 "存储键 -> (模块索引, Symbol*)" 的全局查找表
    // Fix 010r-12: 使用storageKey()而非lowerName做去重键
    // 原因: Property Get/Let/Set同名但有不同storageKey ($pg/$pl/$ps)
    // 用lowerName去重会导致只有第一个变体(通常Get)被保留, Let/Set丢失
    // 消费模块的P6.7查找 lookupModuleByKind(name, PropertyLet) 会失败
    std::unordered_map<std::string, std::pair<size_t, const Symbol*>> globalPublicSyms;
    for (size_t i = 0; i < exportedSymbols.size(); i++) {
        for (const Symbol* sym : exportedSymbols[i]) {
            std::string sKey = sym->storageKey();
            // 同一个storageKey只保留第一个 (VB6行为: 先声明的优先)
            auto itExisting = globalPublicSyms.find(sKey);
            if (itExisting == globalPublicSyms.end()) {
                globalPublicSyms[sKey] = {i, sym};
            } else {
                // Fix 095: 标准模块过程优先于类成员 — 撞名时保留 .bas 版.
                // VB6 语义: 类 Public 成员不可裸调 (VB_PredeclaredId=False 时),
                // 裸名引用 (GenerateWebSocketKey()) 只能解析到标准模块的 Public
                // 过程. 先到先得规则在 mWebSocketUtils.bas 与 cWebSocketUtils.cls
                // 同名双定义 (GenerateWebSocketKey/StringToUTF8/UTF8ToString/
                // ComputeAcceptKey/GetHeaderValue/GetCloseCodeDescription) 时
                // 可能保留类版符号 → 裸调生成 vb6_<Class>_<Proc>() 缺 me 首参
                // → C2198 (cWebSocketClient.c 81 等 13 处).
                // 类版成员仍经 memberParams/memberProcKinds 表走显式限定调用
                // (obj.Method), 此处不丢失其任何信息.
                const Symbol* existing = itExisting->second.second;
                bool existingIsSetter = (existing->kind == SymbolKind::PropertyLet
                                         || existing->kind == SymbolKind::PropertySet);
                if (existingIsSetter && sym->kind == SymbolKind::PropertyGet) {
                    globalPublicSyms[sKey] = {i, sym};
                }
                // Fix 095: 过程符号撞名 — 类版已在, 标准 .bas 版到来时切换
                else if ((sym->kind == SymbolKind::Sub || sym->kind == SymbolKind::Function)
                         && (existing->kind == SymbolKind::Sub || existing->kind == SymbolKind::Function)) {
                    size_t curIdx = itExisting->second.first;
                    bool curIsClass = modules_[curIdx]->isClassModule;
                    bool newIsClass = modules_[i]->isClassModule;
                    if (curIsClass && !newIsClass) {
                        globalPublicSyms[sKey] = {i, sym};
                    }
                }
            }
        }
    }

    // 对每个模块，检查其模块级作用域中的所有符号
    // 找到未定义引用（在visit(IdentifierExpr)中可能失败的标识符）
    // 策略：遍历模块级作用域中尚未定义（但被引用的地方找不到）的标识符
    // 实际上更简单的做法：扫描每个模块的AST，找到所有IdentifierExpr引用的名称，
    // 如果在本地符号表中找不到，就在全局Public表中查找并注入外部符号

    // 但为了避免修改AST遍历，采用更简洁的方式：
    // 对每个模块，遍历全局Public表，如果该符号在本模块没有本地定义，且名称匹配
    // 某些被引用但未在本模块定义的标识符，就注入外部符号
    //
    // 更精确的方案：只遍历在visit(IdentifierExpr)中可能需要跨模块的符号类型
    // (Sub/Function/Variable/Constant)

    for (size_t i = 0; i < analyzers_.size(); i++) {
        SymbolTable& symTab = analyzers_[i]->symbolTable();

        for (const auto& [sKey, entry] : globalPublicSyms) {
            auto [srcIdx, srcSym] = entry;
            // 跳过本模块导出的符号
            if (srcIdx == i) continue;

            // === ai/023 S03: 包导出边界 ===
            // 注入是跨模块符号可见的唯一通道, 在此过滤与 Private 的
            // "不导出即不可见" 完全同构。规则:
            //   - 源是包内模块且消费者不在同包 (宿主/另一包):
            //     只放行【导出模块】的 Public 成员 (清单 Friend=True 时含 Friend);
            //     非导出模块整模块不可见。
            //   - 源是宿主模块: 恒放行 (包消费宿主不受限)。
            //   - 同包模块之间: 恒放行 (Friend = 包内可见)。
            if (!modules_[srcIdx]->packageName.empty() &&
                modules_[srcIdx]->packageName != modules_[i]->packageName) {
                const std::string& srcPkg = modules_[srcIdx]->packageName;
                bool allowed = false;
                auto pit = packageExportInfos_.find(srcPkg);
                if (pit != packageExportInfos_.end() &&
                    pit->second.exportedModules.count(Symbol::toLower(moduleBaseNames[srcIdx]))) {
                    allowed = (srcSym->access == AccessLevel::Public) ||
                              (pit->second.friendVisible &&
                               srcSym->access == AccessLevel::Friend);
                }
                if (!allowed) {
                    continue;
                }
            }

            // 检查本模块是否已有此符号的本地定义
            // Fix 010r-12: Property变体需按kind分别检查 (Get/Let/Set各自独立)
            Symbol* localSym = nullptr;
            if (srcSym->kind == SymbolKind::PropertyGet ||
                srcSym->kind == SymbolKind::PropertyLet ||
                srcSym->kind == SymbolKind::PropertySet) {
                localSym = symTab.lookupModuleByKind(srcSym->name, srcSym->kind);
            } else if (srcSym->kind == SymbolKind::Class) {
                // Fix <vbeclipse> rev17: 类符号的"本地已定义"判据只认**本模块自有的**
                // 同名符号 (见 lookupOwnModuleSymbol 注释)。沿用 lookupModule 的名字级
                // 查找会把**刚注入进来的外部成员副本**也当成"本地定义": View.cls 的
                // `Property Get View` 一旦铺到某模块, 该模块的 `Class View` 就被判
                // "已有定义"而 SKIP; 而铺又是全工程性的, 结果这个类在任何模块里都不存在。
                // 同时保留内建 ComClass/ComInterface 占裸键这一档 —— ScrRun 的 coclass
                // Folder 要被工程类 Folder 顶掉 (下面 replaceBuiltinComSym 分支)。
                localSym = const_cast<Symbol*>(lookupOwnModuleSymbol(symTab, srcSym->name));
            } else {
                localSym = symTab.lookupModule(srcSym->name);
            }
            // 重载变体 (O3): 裸键被同源外部 head 占用时**不**拦变体 —
            // 消费模块需要重建 head+variants 组结构. 拦两种情况:
            // 本地(非外部)定义占裸键 → 整组让位; 外部 head 来自别的源模块
            // (先到先得已定主) 或同签名 → 不注入.
            bool ovlVariantAllowed = false;
            if (srcSym->isOverloadVariant && localSym && localSym->isExternal
                && localSym->sourceModule == moduleBaseNames[srcIdx]
                && !localSym->overloadFp.empty()
                && localSym->overloadFp != srcSym->overloadFp) {
                ovlVariantAllowed = true;
            }
            // Fix <vbeclipse>: 工程类与**类型库内建** coclass/接口同名时, 工程内定义优先.
            // 不 continue, 落到下方注入路径, 由 defineExternal(replaceBuiltinCom=true) 用
            // 工程 Class 符号替换内建符号。理由: 消费模块里这个名字原本被类型库符号占着
            // (如 ScrRun 的 coclass Folder / IFolder), 它的成员表是外部库的 —— 工程类
            // Folder.cls 的 AddView / ActiveViewId 等成员解析不到, 退回"数据字段"访问
            // (ucPerspective.c: `l_Folder->AddView` C2039 / `vb6_Folder_prop_let_
            // ActiveViewId((*Folder))` C2198). 替换后成员表/ memberProcKinds 齐备,
            // 且 mapTypeRef 走 Class 分支得到原生 vb6_cls_Folder*。
            // 注意: 仅替换内建 (isBuiltin) 占用者; 本地真实定义仍按老语义让位。
            bool replaceBuiltinComSym = false;
            if (localSym && srcSym->kind == SymbolKind::Class && !srcSym->isInterface
                && localSym->isBuiltin
                && (localSym->kind == SymbolKind::ComClass
                    || localSym->kind == SymbolKind::ComInterface)) {
                replaceBuiltinComSym = true;
            }
            if (localSym && !ovlVariantAllowed && !replaceBuiltinComSym) {
                // Fix 177b: 引用类型库的同名 coclass 被本工程同名类模块遮蔽时,
                // **不替换符号**, 只在 builtin ComClass 上记下工程实现类名.
                // VB6 语义: 工程内定义优先于引用库 (VBMAN.vbp 定义 Class=Dictionary,
                // 同时引用 Microsoft Scripting Runtime 的 Scripting.Dictionary).
                // 为什么不整体替换 (Fix 177 首版): 代码生成期整条 coclass 晚绑定通路
                // (VARIANT 表示 / comDefaultMemberName / ComCall 参数打包 / 返回值
                // 解包) 都建立在类型库符号上, 换成工程类走早绑定通路会大面积改写
                // 生成代码, 87 处 cl 编译错误 (C2063/C2440/C2065...).
                // 现方案: `As Dictionary` 保持 coclass 类型 (晚绑定), 仅创建点
                // (New / Dim As New) 由 CCodeGen::comNewExprFor 改走工程类工厂
                // vb6_ComPack_<类>(vb6_cls_<类>_New()) → 真 IDispatch 包装 (ExeComBridge
                // 03 同通路), 运行期对象即工程类实例, 自有扩展成员 (Count/Exists)
                // 经包装器 GetIDsOfNames 正常分发, DISP_E_MEMBERNOTFOUND 不再出现.
                if (srcSym->kind == SymbolKind::Class
                    && localSym->isBuiltin
                    && (localSym->kind == SymbolKind::ComClass
                        || localSym->kind == SymbolKind::ComInterface)) {
                    localSym->comProjectImplClass = srcSym->name;
                }
                // 泛型 (G2) 暴露的存量洞回填: Fix 047 预注册只建"名字级" UDT
                // 占位符号 (无成员表), 消费模块字段类型推断落空 → String 字段
                // 按 Variant 发 (double)BSTR (普通跨模块 UDT 也复现, C2440).
                // 源模块此时已分析完, 把完整成员表回填进占位符号.
                if (srcSym->kind == SymbolKind::UserDefinedType &&
                    localSym->kind == SymbolKind::UserDefinedType &&
                    localSym->udtMembers.empty() && !srcSym->udtMembers.empty()) {
                    localSym->udtMembers = srcSym->udtMembers;
                }
                continue;  // 已有本地定义，不需要外部符号
            }

            // 注入外部符号
            auto extSym = std::make_unique<Symbol>(
                srcSym->kind, srcSym->name, srcSym->type,
                srcSym->location, srcSym->access
            );
            extSym->isExternal = true;
            extSym->sourceModule = moduleBaseNames[srcIdx];
            extSym->params = srcSym->params;  // 复制参数列表（函数调用需要）
            extSym->isArray = srcSym->isArray;
            // 泛型 (G2) 暴露的存量洞: 跨模块 UDT 未复制成员表 → 消费模块成员
            // 类型全退化 Variant (普通 UDT 复现: g.V As String 被 Debug.Print
            // 按 Variant 发 (double)BSTR 强转 → C2440).
            if (srcSym->kind == SymbolKind::UserDefinedType) {
                extSym->udtMembers = srcSym->udtMembers;
            }
            // O3: 携带重载组身份 — storageKey() 据此把变体注入 "<name>$ov$<fp>"
            // 独立键, 裸键留给 head; 消费模块的 resolveOverload 走 lookupModuleOverloads
            // 收集整组后按实参打分选变体.
            extSym->overloadFp = srcSym->overloadFp;
            extSym->isOverloadVariant = srcSym->isOverloadVariant;
            extSym->ovlCount = srcSym->ovlCount;
            // 类符号: 复制instancing、memberNames、isInterface、implementsNames
            if (srcSym->kind == SymbolKind::Class) {
                extSym->instancing = srcSym->instancing;
                extSym->memberNames = srcSym->memberNames;
                extSym->memberReturnTypes = srcSym->memberReturnTypes;  // Fix 015: 链式调用返回类型表
                extSym->memberProcKinds = srcSym->memberProcKinds;       // Fix 016: 成员过程类型表
                extSym->memberParams = srcSym->memberParams;             // Fix 033: 成员参数表 (calleeParams 跨模块精确查找)
                extSym->memberFieldTypes = srcSym->memberFieldTypes;     // Fix 092m: 字段类型表 (With 字段写 COM 解包)
                extSym->memberFieldNames = srcSym->memberFieldNames;     // Fix 092p: 字段声明原名表 (访问点大小写规范化)
                extSym->publicFieldNames = srcSym->publicFieldNames;     // Fix 099: 显式 Public 字段清单 (COM 暴露)
                extSym->memberFieldDispids = srcSym->memberFieldDispids; // Fix 099: Public 字段 TypeLib DISPID
                extSym->memberLetParams = srcSym->memberLetParams;       // Fix 091a: Let 写方向参数表
                extSym->memberSetParams = srcSym->memberSetParams;       // Fix 091a: Set 写方向参数表
                extSym->memberAccessLevels = srcSym->memberAccessLevels; // tB B08a: 成员访问级别
                extSym->isInterface = srcSym->isInterface;  // P6.4
                extSym->implementsNames = srcSym->implementsNames;  // P6.4
                extSym->interfaceMethodNames = srcSym->interfaceMethodNames;  // P6.4
                extSym->eventNames = srcSym->eventNames;  // P6.5
                extSym->comClsidStr = srcSym->comClsidStr;  // P6.8: CLSID
            }
            // Fix 010r-11 / Fix 015: 复制变量/返回类型名
            // - Variable: 用于跨模块类实例变量识别 (knownClassVars_)
            // - Function / PropertyGet: 用于 method chaining 的返回类型推断
            //   (getClassMethodReturnType 读取该字段判断链式调用能继续到哪一层)
            // 对其它 kind 该字段为空, 无副作用. 原先把此赋值放在 Variable-only 分支内,
            // 导致 Function/PropertyGet 外部符号丢失返回类型名, 链式调用解析失败.
            extSym->variableTypeName = srcSym->variableTypeName;
            // Fix 183: 复制 As New 标志 — 外部模块级 As New 变量在消费模块中
            // 同样需要惰性实例化守卫 (否则全局对象恒为 NULL).
            extSym->isNewVar = srcSym->isNewVar;
            // Fix 017: 复制常量值 (EnumMember / Constant 跨模块注入后需保留值).
            // 外部 EnumMember 若 hasConstValue=false, cgen 会发出裸标识符 (如
            // HASH_ALG_SHA256) 而非数值 → C2065. 此前外部符号构造只复制
            // kind/name/type/location/access, 丢失 hasConstValue/constIntValue.
            // Fix 081d: 也复制 constStringValue/constFloatValue, 否则字符串常量跨模块时丢失.
            extSym->hasConstValue = srcSym->hasConstValue;
            extSym->constIntValue = srcSym->constIntValue;
            extSym->constType = srcSym->constType;
            extSym->constStringValue = srcSym->constStringValue;
            extSym->constFloatValue = srcSym->constFloatValue;
            extSym->constBoolValue = srcSym->constBoolValue;
            if (srcSym->kind == SymbolKind::Variable) {
                extSym->dimCount = srcSym->dimCount;
            }

            symTab.defineExternal(std::move(extSym), replaceBuiltinComSym);
        }
    }

    // tB Inherits (ai/022 B09b, 实测见 022 D39): 祖先模块里 **Private** 的 UDT/Enum 类型符号,
    // 凡被继承字段用到的, 也注进消费者作用域。
    // 为什么必须注: 继承字段是**按值**嵌进派生结构体的, 类型名在消费者模块解不出来时 mapTypeRef
    // 回落到 `void*` —— x64 上 sizeof(void*) 恰好等于小 UDT 的尺寸, 布局"巧合正确"、断言全绿;
    // x86 上短 4 字节 → 该字段之后每个继承字段的偏移全错, 而 _New() 又按短了的 sizeof 分配
    // → 基类那份过程写后面的字段就是越界写 (Inh.vbp 在 x86 段错误就是这个形状)。
    // 只沿 Inherits 链补、只补"被继承字段用到"的名字、本地已有同名符号一律不抢 → 不写 Inherits
    // 的工程一个符号都不会多。类型**定义**的文本早就在基类自己的 .h 里 (visit(TypeDecl) 不分访问
    // 级别, 还带 VB6_TYPE_<X>_DEFINED 守卫), 这里缺的从来只是名字。
    if (!classes_.empty()) {
        std::unordered_map<const Module*, size_t> modIndexOf;
        for (size_t k = 0; k < modules_.size(); k++) modIndexOf[modules_[k].get()] = k;
        for (size_t i = 0; i < analyzers_.size(); i++) {
            auto itSelf = classes_.find(ifaceLower(modules_[i]->moduleName));
            if (itSelf == classes_.end() || itSelf->second.mod != modules_[i].get()) continue;
            const ClassChainView& self = itSelf->second;
            if (self.chainBroken || self.baseKey.empty() || self.inhFields.empty()) continue;
            std::unordered_set<std::string> needNames;   // 只看 SimpleTypeRef (与本洞无关的形状不外扩)
            for (Decl* d : self.inhFields) {
                if (!d || d->kind != ASTNodeKind::VariableDecl) continue;
                VariableDecl& v = static_cast<VariableDecl&>(*d);
                if (!v.asType || v.asType->kind != ASTNodeKind::SimpleTypeRef) continue;
                needNames.insert(ifaceLower(static_cast<SimpleTypeRef&>(*v.asType).name));
            }
            if (needNames.empty()) continue;
            SymbolTable& symTab = analyzers_[i]->symbolTable();
            const std::string selfKey = ifaceLower(modules_[i]->moduleName);
            for (const std::string& ancKey : self.chain) {
                if (ancKey == selfKey) continue;
                auto itAnc = classes_.find(ancKey);
                if (itAnc == classes_.end() || !itAnc->second.mod) continue;
                auto itIdx = modIndexOf.find(itAnc->second.mod);
                if (itIdx == modIndexOf.end() || !analyzers_[itIdx->second]) continue;
                SymbolTable& ancTab = analyzers_[itIdx->second]->symbolTable();
                if (!ancTab.moduleScope()) continue;
                for (const auto& kv : ancTab.moduleScope()->symbols()) {
                    const Symbol* ts = kv.second.get();
                    if (!ts) continue;
                    if (ts->kind != SymbolKind::UserDefinedType && ts->kind != SymbolKind::EnumType)
                        continue;
                    if (!needNames.count(ifaceLower(ts->name))) continue;
                    if (symTab.lookupModule(ts->name)) continue;   // 本地已有/已注入 → 不抢
                    auto extSym = std::make_unique<Symbol>(
                        ts->kind, ts->name, ts->type, ts->location, ts->access);
                    extSym->isExternal = true;
                    extSym->sourceModule = itAnc->second.mod->moduleName;
                    if (ts->kind == SymbolKind::UserDefinedType) extSym->udtMembers = ts->udtMembers;
                    symTab.defineExternal(std::move(extSym));
                }
            }
        }
    }

    // O3: 外部重载组已注入各模块表 — 补跑各模块分析期"当时查无此名"的调用点,
    // 让语义层把变体后缀写回 AST (cgen 只消费 calleeOvlSuffix, 不做选择).
    for (auto& analyzer : analyzers_) {
        analyzer->resolveDeferredCrossModuleOverloads();
    }

    if (diag_->hasErrors()) return false;
    return true;
}

} // namespace vb6c3
