// driver_semantics.cpp - C3 编译器驱动: 语义分析阶段
// 2026-09-17 从 src/driver/driver.cpp 纯搬移（逐行未改）：
//   原第 867~1190 行

#include "driver/driver.hpp"
#include "common/diagnostics.hpp"
#include "semantics/semantic_analyzer.hpp"
#include <iostream>
#include <algorithm>
#include <functional>
#include <unordered_set>

namespace vb6c3 {

namespace {

// get/put 同名冲突保护 (2026-09-18): 类型库里 VB6 的公开字段与读写属性是
// get_X + put_X 一对 FUNCDESC (同 memid、同名字), 而 comMethods 按小写名索引 ——
// 后写入的 put/putref 会覆盖 getter 的签名: returnType 退化成 Empty (VT_VOID),
// isPropertyGet 变 false。于是 obj.Prop.Method() 这类链式访问被判成"标量取值",
// 再对 int16_t 拼 C 结构体成员访问 → C2224 (vbman-demo 的 ctx.Request.QueryString)。
// 读取路径只认 getter; 属性写入走名字化的晚绑定, 不查 comMethods, 故让 getter 优先。
void insertComMethod(Symbol& sym, const std::string& key, Symbol::ComMethodSig sig,
                     ComMemberKind kind) {
    bool incomingSetter = (kind == ComMemberKind::PropertyPut
                        || kind == ComMemberKind::PropertyPutRef);
    auto it = sym.comMethods.find(key);
    if (incomingSetter && it != sym.comMethods.end() && it->second.isPropertyGet) {
        return;  // 已有 getter, setter 不得覆盖
    }
    sym.comMethods[key] = std::move(sig);
}

} // namespace

bool Driver::runSemanticAnalysis(const CompileOptions& options) {
    // ai/084a M1: 类成员访问级别预计算表 (只读, 随后随各分析器下发)
    buildMemberAccessTable();

    analyzers_.clear();
    // <vbeclipse>: 先把"工程级已有的名字"收齐，随各分析器下发。两份都必须从**已解析的
    // AST** 取，不能取符号表那份 getExternalModuleNames()：跨模块注入跑在逐模块分析之后，
    // 而隐式变量声明发生在分析当刻 —— 晚了就把 Module1.ShowForm2 这种名字挤成局部变量
    // (实测 ext_show_test C2063、test_modulemethod 恒返 0)。见 namesProjectLevel 注释。
    std::unordered_set<std::string> projModNames;
    std::unordered_set<std::string> projPubProcs;
    for (auto& module : modules_) {
        if (!module) continue;
        projModNames.insert(Symbol::toLower(module->moduleName));
        // 只有标准模块的 Public 过程能被别处按**裸名**点到 (类/窗体的公开成员必须经
        // 实例名), 所以第二份只收 .bas 模块。
        if (module->isClassModule || module->isFormModule) continue;
        for (auto& d : module->declarations) {
            if (!d) continue;
            std::string nm;
            AccessLevel acc = AccessLevel::Private;
            switch (d->kind) {
                case ASTNodeKind::SubDecl: {
                    auto& s = static_cast<SubDecl&>(*d); nm = s.name; acc = s.access; break;
                }
                case ASTNodeKind::FunctionDecl: {
                    auto& s = static_cast<FunctionDecl&>(*d); nm = s.name; acc = s.access; break;
                }
                case ASTNodeKind::PropertyDecl: {
                    auto& s = static_cast<PropertyDecl&>(*d); nm = s.name; acc = s.access; break;
                }
                default: continue;
            }
            if (acc == AccessLevel::Public && !nm.empty()) {
                projPubProcs.insert(Symbol::toLower(nm));
            }
        }
    }
    for (auto& module : modules_) {
        auto analyzer = std::make_unique<SemanticAnalyzer>(*diag_, options.verbose);
        // <vbeclipse>: 工程级名字表 (见上面两份的注释)
        analyzer->setProjectModuleNames(projModNames);
        analyzer->setProjectPublicProcNames(projPubProcs);
        // 泛型 (tB, G3): 调用点推断需要模板只读视图 (runGenericsPrepass 已构建)
        analyzer->setGenericRegistry(&genView_);
        // Interface 契约 (tB, B02): stage 2.7 建好的只读登记表
        analyzer->setInterfaceRegistry(&ifaces_);
        // 委托式实现 (tB, B10): stage 2.7 Pass D 建好的裁决表 (整份契约转交给持有字段)
        analyzer->setViaRegistry(&vias_);
        // 类继承 (tB, B07b): stage 2.8 建好的只读链登记表 (裸名继承成员判定)
        analyzer->setClassChainRegistry(&classes_);
        // ai/084a M1/M2: 成员访问级别表 (Private 越界裁决的数据源)
        analyzer->setMemberAccessTable(&memberAccessTable_);
        // ai/084c: Class_Initialize 形参个数表 (New Cls(args) 元数校验)
        analyzer->setCtorParamCounts(ctorParamCounts_);

        // ai/023 S03: 包导出边界屏蔽表 (预计算, driver_compile 填)。
        // 本模块消费"其他包"时, 那些包的被屏蔽成员名在本模块查无此名的瞬间
        // 报 VB7006 —— 必须在分析器创建时就位 (无实参裸调用不走 deferred 通路)。
        {
            const std::string& ownPkg = module->packageName;
            std::map<std::string, std::string> blocked;
            for (const auto& [pkg, names] : packageBlockedNames_) {
                if (pkg == ownPkg) continue;
                blocked.insert(names.begin(), names.end());
            }
            if (!blocked.empty()) analyzer->setBlockedPackageNames(std::move(blocked));

            // ai/023 S04: 非导出包类名 (同样排除自己的包)
            std::map<std::string, std::string> blockedCls;
            for (const auto& [pkg, names] : packageBlockedClasses_) {
                if (pkg == ownPkg) continue;
                blockedCls.insert(names.begin(), names.end());
            }
            if (!blockedCls.empty()) analyzer->setBlockedPackageClasses(std::move(blockedCls));
        }

        // P6.3: 如果有TypeLib解析结果, 注入COM类型信息到符号表
        if (typelibParser_) {
            for (auto& tl : typelibParser_->cachedResults()) {
                for (auto& cc : tl->coclasses) {
                    // 注册ComClass符号
                    auto sym = std::make_unique<Symbol>(
                        SymbolKind::ComClass, cc->name, Vb6Type::Object,
                        SourceLocation{}, AccessLevel::Public);
                    sym->isBuiltin = true;
                    sym->comClsidStr = cc->clsidStr;
                    sym->comProgId = cc->progId;
                    sym->comDefaultIfaceName = cc->defaultIfaceName;

                    // 复制默认接口的方法签名到ComClass
                    if (cc->defaultIface) {
                        sym->comIidStr = cc->defaultIface->iidStr;
                        sym->comIsDual = cc->defaultIface->isDual;
                        sym->comVtblBase = cc->defaultIface->isDispatch ? 7 : 3;

                        // P24-10: 传播默认成员名 (DISPID_VALUE=0)
                        sym->comDefaultMemberName = cc->defaultIface->defaultMemberName;
                        sym->comDefaultMemberRealName = cc->defaultIface->defaultMemberRealName;

                        for (auto& member : cc->defaultIface->members) {
                            Symbol::ComMethodSig sig;
                            sig.realName = member.realName;
                            sig.memid = member.memid;
                            sig.vtableIndex = member.vtableIndex;
                            sig.returnType = member.returnType;
                            sig.isPropertyGet = (member.kind == ComMemberKind::PropertyGet);
                            sig.isPropertyPut = (member.kind == ComMemberKind::PropertyPut);
                            sig.isPropertyPutRef = (member.kind == ComMemberKind::PropertyPutRef);
                            for (auto& param : member.params) {
                                ParameterInfo pi;
                                pi.name = param.name;
                                pi.type = param.type;
                                pi.isByVal = (param.direction == ComParamDir::In);
                                pi.isOptional = param.isOptional;
                                sig.params.push_back(std::move(pi));
                            }
                            insertComMethod(*sym, member.name, std::move(sig), member.kind);
                        }
                        // 填充memberNames (类成员名列表, 兼容现有逻辑)
                        for (auto& member : cc->defaultIface->members) {
                            sym->memberNames.push_back(member.realName);
                        }
                    }

                    // P13.23: 填充事件源接口信息 (用于外部COM WithEvents)
                    if (!cc->defaultSourceIfaceName.empty()) {
                        sym->comHasSourceIface = true;
                        sym->comSourceIfaceName = cc->defaultSourceIfaceName;
                        if (cc->defaultSourceIface) {
                            sym->comSourceIfaceIid = cc->defaultSourceIface->iidStr;
                            sym->comSourceIfaceIsDispOnly = cc->defaultSourceIface->isDispatch;
                            for (auto& member : cc->defaultSourceIface->members) {
                                sym->eventNames.push_back(member.realName);
                                sym->comEventDispids[Symbol::toLower(member.name)] = member.memid;
                                Symbol::ComMethodSig sig;
                                sig.realName = member.realName;
                                sig.memid = member.memid;
                                sig.vtableIndex = member.vtableIndex;
                                sig.returnType = member.returnType;
                                for (auto& param : member.params) {
                                    ParameterInfo pi;
                                    pi.name = param.name;
                                    pi.type = param.type;
                                    pi.isByVal = (param.direction == ComParamDir::In);
                                    pi.isOptional = param.isOptional;
                                    sig.params.push_back(pi);
                                }
                                sym->comSourceMethods[Symbol::toLower(member.name)] = std::move(sig);
                            }
                        }
                    }
                    analyzer->symbolTable().define(std::move(sym));
                }

                // 也注册ComInterface符号
                for (auto& iface : tl->interfaces) {
                    auto sym = std::make_unique<Symbol>(
                        SymbolKind::ComInterface, iface->name, Vb6Type::Object,
                        SourceLocation{}, AccessLevel::Public);
                    sym->isBuiltin = true;
                    sym->comIidStr = iface->iidStr;
                    sym->comIsDual = iface->isDual;
                    sym->comVtblBase = iface->isDispatch ? 7 : 3;

                    for (auto& member : iface->members) {
                        Symbol::ComMethodSig sig;
                        sig.realName = member.realName;
                        sig.memid = member.memid;
                        sig.vtableIndex = member.vtableIndex;
                        sig.returnType = member.returnType;
                        sig.isPropertyGet = (member.kind == ComMemberKind::PropertyGet);
                        sig.isPropertyPut = (member.kind == ComMemberKind::PropertyPut);
                        sig.isPropertyPutRef = (member.kind == ComMemberKind::PropertyPutRef);
                        for (auto& param : member.params) {
                            ParameterInfo pi;
                            pi.name = param.name;
                            pi.type = param.type;
                            pi.isByVal = (param.direction == ComParamDir::In);
                            pi.isOptional = param.isOptional;
                            sig.params.push_back(std::move(pi));
                        }
                        insertComMethod(*sym, member.name, std::move(sig), member.kind);
                    }

                    analyzer->symbolTable().define(std::move(sym));
                }

                // Fix 018: 注册COM枚举成员为EnumMember符号 (TextCompare/adStateClosed 等)
                // COM 类型库的 TKIND_ENUM 解析后, 成员作为全局可见命名常量注入符号表.
                // hasConstValue=true 使 cgen 发出数值 (经 Fix 017-P1/P3 路径), 避免裸名 C2065.
                // Volume guard: 跳过枚举成员过多的类型库 (如 MSHTML 数千成员), 避免命名空间
                // 污染、内存膨胀 (125 模块 × N 成员) 和跨 enum 同名碰撞. 小型库 (Scripting ~30,
                // ADO ~200) 正常注册, 覆盖 TextCompare / adXXX 等已知 C2065.
                size_t totalEnumMembers = 0;
                for (auto& en : tl->enums) totalEnumMembers += en->members.size();
                if (totalEnumMembers <= 1000) {
                    for (auto& en : tl->enums) {
                        for (auto& m : en->members) {
                            auto enumSym = std::make_unique<Symbol>(
                                SymbolKind::EnumMember, m.name, Vb6Type::Long,
                                SourceLocation{}, AccessLevel::Public);
                            enumSym->isBuiltin = true;
                            enumSym->hasConstValue = true;
                            enumSym->constIntValue = m.value;
                            enumSym->constType = Vb6Type::Long;
                            analyzer->symbolTable().define(std::move(enumSym));
                        }
                    }
                }
                // Fix 084m: ADODB 非标准别名。类型库 ExecuteOptionEnum 只有 adAsyncExecute(=16),
                // 但源码 (cDatabase.Exec) 使用 adExecuteAsync; 未注册时生成裸名 → C2065。
                {
                    std::string tlPathLower = tl->tlbPath;
                    std::transform(tlPathLower.begin(), tlPathLower.end(), tlPathLower.begin(), ::tolower);
                    bool isAdodb = tlPathLower.find("msado") != std::string::npos
                        || tl->typeLibProjectName.find("ADODB") != std::string::npos;
                    if (isAdodb) {
                        static const std::vector<std::pair<const char*, long>> adodbAliases = {
                            {"adExecuteAsync", 16},
                        };
                        for (auto& al : adodbAliases) {
                            std::string alLower = al.first;
                            std::transform(alLower.begin(), alLower.end(), alLower.begin(), ::tolower);
                            if (!analyzer->symbolTable().lookupModule(alLower)) {
                                auto enumSym = std::make_unique<Symbol>(
                                    SymbolKind::EnumMember, al.first, Vb6Type::Long,
                                    SourceLocation{}, AccessLevel::Public);
                                enumSym->isBuiltin = true;
                                enumSym->hasConstValue = true;
                                enumSym->constIntValue = al.second;
                                enumSym->constType = Vb6Type::Long;
                                analyzer->symbolTable().define(std::move(enumSym));
                            }
                        }
                    }
                }
            }
        }
            // P24-04: 注册ComModule符号 (TKIND_MODULE → ActiveX DLL全局函数命名空间)
            // VBMAN.Version() 这种调用: VBMAN是工程名/模块名, Version是全局函数
            for (auto& tl : typelibParser_->cachedResults()) {
                for (auto& mod : tl->modules) {
                    auto sym = std::make_unique<Symbol>(
                        SymbolKind::ComModule, mod->name, Vb6Type::Object,
                        SourceLocation{}, AccessLevel::Public);
                    sym->isBuiltin = true;
                    sym->comModuleDllPath = mod->dllPath;
                    for (auto& func : mod->functions) {
                        Symbol::ComMethodSig sig;
                        sig.realName = func.realName;
                        sig.memid = func.memid;
                        sig.vtableIndex = -1;  // 模块函数无vtable
                        sig.returnType = func.returnType;
                        sig.isPropertyGet = false;
                        sig.isPropertyPut = false;
                        sig.isPropertyPutRef = false;
                        for (auto& param : func.params) {
                            ParameterInfo pi;
                            pi.name = param.name;
                            pi.type = param.type;
                            pi.isByVal = (param.direction == ComParamDir::In);
                            pi.isOptional = param.isOptional;
                            sig.params.push_back(std::move(pi));
                        }
                        sym->comModuleFunctions[func.name] = std::move(sig);
                    }
                    analyzer->symbolTable().define(std::move(sym));
                }
            }

            // P24-04: 注册VB_GlobalNameSpace promoted函数为ComGlobalNs符号
            // GlobalNameSpace coclass (如sGlobal) 的默认接口Public方法提升为全局符号
            // 示例: VBMAN库的sGlobal._sGlobal.VBMAN() → 全局"VBMAN"函数 (SymbolKind::ComGlobalNs)
            // 代码生成: VBMAN() → vb6_ComCallObject(vb6_CreateObject(L"VBMANLIB.sGlobal"), L"VBMAN", NULL, 0)
            for (auto& tl : typelibParser_->cachedResults()) {
                for (auto& cc : tl->coclasses) {
                    if (!cc->isGlobalNamespace || !cc->defaultIface) continue;
                    // 只提升非标准方法 (排除 IUnknown/IDispatch 的标准方法)
                    // Fix 094: 改为按名过滤, 不再"跳过前 N 个". 接口 members 是否包含
                    // 标准方法取决于 TypeLib 的生成方 —— 真 VB6 生成的含(7 个),
                    // C3 自产的只含业务方法. 原先假设"前 7 个是标准方法"会让 C3 自产
                    // TypeLib 在 members.size() < 7 时整个循环不执行, 一个全局方法都
                    // 提升不了 (实测 sGlobal._sGlobal members=1 只有 VBMAN, 于是
                    // VBMAN.Version() 退化成 vb6_VBMAN_Version() → LNK2019).
                    static const std::unordered_set<std::string> stdMethodNames = {
                        "QueryInterface", "AddRef", "Release",
                        "GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke"
                    };
                    for (size_t mi = 0; mi < cc->defaultIface->members.size(); mi++) {
                        auto& member = cc->defaultIface->members[mi];
                        if (stdMethodNames.count(member.realName)) continue;
                        // 每个promoted方法注册为独立的ComGlobalNs符号
                        auto sym = std::make_unique<Symbol>(
                            SymbolKind::ComGlobalNs, member.realName, member.returnType,
                            SourceLocation{}, AccessLevel::Public);
                        sym->isBuiltin = true;
                        sym->comClsidStr = cc->clsidStr;
                        sym->comProgId = cc->progId;
                        sym->comDefaultIfaceName = cc->defaultIfaceName;
                        sym->comDefaultIfaceIid = cc->defaultIface->iidStr;
                        sym->comDefaultMemberName = cc->defaultIface->defaultMemberName;
                        sym->comDefaultMemberRealName = cc->defaultIface->defaultMemberRealName;
                        sym->comGlobalNsMethodName = member.realName;
                        // 复制方法签名
                        Symbol::ComMethodSig sig;
                        sig.realName = member.realName;
                        sig.memid = member.memid;
                        sig.vtableIndex = member.vtableIndex;
                        sig.returnType = member.returnType;
                        sig.isPropertyGet = (member.kind == ComMemberKind::PropertyGet);
                        sig.isPropertyPut = (member.kind == ComMemberKind::PropertyPut);
                        sig.isPropertyPutRef = (member.kind == ComMemberKind::PropertyPutRef);
                        for (auto& param : member.params) {
                            ParameterInfo pi;
                            pi.name = param.name;
                            pi.type = param.type;
                            pi.isByVal = (param.direction == ComParamDir::In);
                            pi.isOptional = param.isOptional;
                            sig.params.push_back(std::move(pi));
                        }
                        insertComMethod(*sym, member.name, std::move(sig), member.kind);
                        analyzer->symbolTable().define(std::move(sym));
                    }
                }
            }
        // P7.5: 注册窗体控件名为符号 (否则语义分析报"未声明的标识符")
        auto frmIt = frmFiles_.find(module->moduleName);
        if (frmIt != frmFiles_.end()) {
            const auto& frmDesc = frmIt->second.form;
            // 注册窗体名本身 (Form1.Caption 访问)
            // Fix 110: 仅 .frm 需要 — .ctl/.pag 的模块名已经是 Class 符号,
            // 再定义为 Object 变量会把类符号覆盖掉 (Dim x As New ucChartArea
            // 随之退化为 Object/void*, 方法调用全部变成 COM 晚绑定).
            if (module->isFormModule) {
                auto sym = std::make_unique<Symbol>(
                    SymbolKind::Variable, module->moduleName, Vb6Type::Object,
                    SourceLocation{}, AccessLevel::Public);
                sym->isBuiltin = true;
                analyzer->symbolTable().define(std::move(sym));
            }
            // 注册控件名 (去重: 控件数组只注册一次)
            // Fix 110: 递归注册嵌套控件 (Frame 内的子控件也要注册).
            std::unordered_set<std::string> registeredCtrls;
            std::function<void(const FrmControl&)> registerCtrlsRec =
                [&](const FrmControl& ctrl) {
                    std::string ctrlLower = ctrl.controlName;
                    for (auto& c : ctrlLower) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
                    if (!ctrlLower.empty() && registeredCtrls.insert(ctrlLower).second) {
                        // Fix 110: 已存在的同名类型符号 (工程类/UDT/COM 类, 如
                        // PropertyPage 上的 LabelPlus1 与工程类同名) 不覆盖 —
                        // 否则类符号被 Object 变量顶掉, 方法调用退化为 COM 晚绑定.
                        Symbol* exist = analyzer->symbolTable().lookupModule(ctrl.controlName);
                        const bool isTypeSym = exist &&
                            (exist->kind == SymbolKind::Class ||
                             exist->kind == SymbolKind::UserDefinedType ||
                             exist->kind == SymbolKind::EnumType ||
                             exist->kind == SymbolKind::ComClass ||
                             exist->kind == SymbolKind::ComInterface);
                        if (!isTypeSym) {
                            auto sym = std::make_unique<Symbol>(
                                SymbolKind::Variable, ctrl.controlName, Vb6Type::Object,
                                SourceLocation{}, AccessLevel::Public);
                            sym->isBuiltin = true;
                            analyzer->symbolTable().define(std::move(sym));
                        }
                    }
                    for (const auto& child : ctrl.children) registerCtrlsRec(child);
                };
            for (const auto& ctrl : frmDesc.formControl.children) {
                registerCtrlsRec(ctrl);
            }
        }

        // Fix 047: Pre-register Public Enum and UDT types from all other modules
        // so that parameter types like "As UcsAsyncSocketEventMaskEnum" resolve
        // to Long/UserDefinedType during Pass 1 (before runCrossModuleResolution).
        // Without this, cross-module Enum/UDT parameter types are resolved as Variant,
        // causing false-positive vb6_VariantFromValue() wrapping at call sites → C2440.
        for (auto& otherModule : modules_) {
            if (otherModule.get() == module.get()) continue;
            // 泛型模板类 (G4): 本体声明不外泄 (特化副本另有扁名, 各自正常预注册)
            if (!otherModule->classTypeParams.empty()) continue;
            for (auto& decl : otherModule->declarations) {
                if (decl->kind == ASTNodeKind::EnumDecl) {
                    auto& enumDecl = static_cast<EnumDecl&>(*decl);
                    // AccessLevel::Default == Public, so all non-Private enums are public
                    if (enumDecl.access != AccessLevel::Private) {
                        std::string lower = Symbol::toLower(enumDecl.name);
                        if (!analyzer->symbolTable().lookupModule(lower)) {
                            auto sym = std::make_unique<Symbol>(
                                SymbolKind::EnumType, enumDecl.name,
                                Vb6Type::Long, enumDecl.loc, AccessLevel::Public
                            );
                            sym->isExternal = true;
                            sym->sourceModule = otherModule->moduleName;
                            analyzer->symbolTable().defineExternal(std::move(sym));
                        }
                        // Fix 197: 预注册枚举成员常量 — Optional 参数默认值 `= ucsSfdAll`
                        // 经 evalOptionalDefault 查符号表发生在 Pass 1 期间, 而
                        // runCrossModuleResolution 在其后才注入跨模块枚举成员, 不预注册
                        // 则默认值回退类型零值 0 (vbman cTlsSocket.Create 的
                        // EventMask=0 → WSAAsyncSelect 注销通知 → 永不 accept 实证).
                        int64_t nextValue = 0;
                        for (auto& member : enumDecl.members) {
                            std::string mLower = Symbol::toLower(member->name);
                            if (!analyzer->symbolTable().lookupModule(mLower)) {
                                auto msym = std::make_unique<Symbol>(
                                    SymbolKind::EnumMember, member->name,
                                    Vb6Type::Long, member->loc, AccessLevel::Public
                                );
                                msym->isExternal = true;
                                msym->sourceModule = otherModule->moduleName;
                                msym->hasConstValue = true;
                                msym->constType = Vb6Type::Long;
                                int64_t evaluated = 0;
                                if (member->value
                                    && SemanticAnalyzer::evalEnumConstIntForDriver(member->value.get(), evaluated)) {
                                    nextValue = evaluated;
                                }
                                msym->constIntValue = nextValue;
                                nextValue++;
                                analyzer->symbolTable().defineExternal(std::move(msym));
                            }
                        }
                    }
                } else if (decl->kind == ASTNodeKind::TypeDecl) {
                    auto& typeDecl = static_cast<TypeDecl&>(*decl);
                    // 泛型模板 (tB, G2): 不预注册 (泛型器已把特化副本注入声明表,
                    // 预注册走的就是这条 AST 扫描, 特化副本 typeParams 为空自然通过)
                    if (typeDecl.typeParams.empty() &&
                        typeDecl.access != AccessLevel::Private) {
                        std::string lower = Symbol::toLower(typeDecl.name);
                        if (!analyzer->symbolTable().lookupModule(lower)) {
                            auto sym = std::make_unique<Symbol>(
                                SymbolKind::UserDefinedType, typeDecl.name,
                                Vb6Type::UserDefinedType, typeDecl.loc, AccessLevel::Public
                            );
                            sym->isExternal = true;
                            sym->sourceModule = otherModule->moduleName;
                            analyzer->symbolTable().defineExternal(std::move(sym));
                        }
                    }
                }
            }
        }

        bool ok = analyzer->analyze(*module);

        if (options.dumpSymbols) {
            analyzer->dumpSymbols(std::cout);
        }

        if (!ok) return false;
        analyzers_.push_back(std::move(analyzer));
    }
    return !diag_->hasErrors();
}

// ai/084a M1: 扫全部类模块的 AST 声明建成员访问级别表。与语义层 pass-1 的
// memberAccessLevels 回填 (semantic_analyzer.cpp:42-163) 同一套 precedence:
// Property Get 总是覆盖; Let/Set 仅在键不存在或现有也是 Let/Set 时写入。
// 泛型模板类与语义层同口径跳过 (本体声明不外泄, 特化副本走完整管线自建)。
void Driver::buildMemberAccessTable() {
    memberAccessTable_.clear();
    ctorParamCounts_.clear();
    for (const auto& module : modules_) {
        if (!module || !module->isClassModule) continue;
        if (!module->classTypeParams.empty()) continue;  // 泛型模板 (G4): 与 analyze() 同口径
        const std::string clsLower = Symbol::toLower(module->moduleName);
        auto& members = memberAccessTable_[clsLower];
        // ai/084c: Class_Initialize 形参个数 (带参构造 New Cls(args) 的元数校验源)
        int ctorN = 0;
        for (const auto& decl : module->declarations) {
            if (decl && decl->kind == ASTNodeKind::SubDecl) {
                auto& s = static_cast<SubDecl&>(*decl);
                if (s.name == "Class_Initialize") { ctorN = (int)s.params.size(); break; }
            }
        }
        ctorParamCounts_[clsLower] = ctorN;
        // ai/084a M3: 定义包归属 + Friend 豁免位 (packageExportInfos_ 在 stage 0 已填)
        const std::string pkgLower = Symbol::toLower(module->packageName);
        bool pkgFriendOpen = false;
        if (!pkgLower.empty()) {
            auto pit = packageExportInfos_.find(pkgLower);
            if (pit != packageExportInfos_.end()) pkgFriendOpen = pit->second.friendVisible;
        }
        std::map<std::string, ProcKind> procKindOf;  // 复刻 memberProcKinds 的 wins 判定
        for (const auto& decl : module->declarations) {
            if (!decl) continue;
            switch (decl->kind) {
                case ASTNodeKind::VariableDecl: {
                    auto& v = static_cast<VariableDecl&>(*decl);
                    members[Symbol::toLower(v.name)] =
                        { v.access, clsLower, true, pkgLower, pkgFriendOpen };
                    break;
                }
                case ASTNodeKind::SubDecl: {
                    auto& s = static_cast<SubDecl&>(*decl);
                    members[Symbol::toLower(s.name)] = { s.access, clsLower, false, pkgLower, pkgFriendOpen };
                    procKindOf[Symbol::toLower(s.name)] = ProcKind::Sub;
                    break;
                }
                case ASTNodeKind::FunctionDecl: {
                    auto& f = static_cast<FunctionDecl&>(*decl);
                    members[Symbol::toLower(f.name)] = { f.access, clsLower, false, pkgLower, pkgFriendOpen };
                    procKindOf[Symbol::toLower(f.name)] = ProcKind::Function;
                    break;
                }
                case ASTNodeKind::PropertyDecl: {
                    auto& p = static_cast<PropertyDecl&>(*decl);
                    std::string lower = Symbol::toLower(p.name);
                    auto it = procKindOf.find(lower);
                    bool wins = false;
                    if (p.propKind == ProcKind::PropertyGet) {
                        wins = true;
                    } else if (p.propKind == ProcKind::PropertyLet) {
                        wins = (it == procKindOf.end()
                                || it->second == ProcKind::PropertyLet
                                || it->second == ProcKind::PropertySet);
                    } else if (p.propKind == ProcKind::PropertySet) {
                        wins = (it == procKindOf.end()
                                || it->second == ProcKind::PropertySet);
                    }
                    if (wins) {
                        members[lower] = { p.access, clsLower, false, pkgLower, pkgFriendOpen };
                        procKindOf[lower] = p.propKind;
                    }
                    break;
                }
                case ASTNodeKind::EventDecl: {
                    auto& e = static_cast<EventDecl&>(*decl);
                    members[Symbol::toLower(e.name)] = { e.access, clsLower, false, pkgLower, pkgFriendOpen };
                    break;
                }
                default:
                    break;
            }
        }
    }
}

} // namespace vb6c3
