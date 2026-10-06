// driver_codegen_dll_typelib.inc - runCodeGeneration 函数体片段：P6.6/P9 ActiveX DLL TypeLib 构建（位于 if (options.isDll) 块内）
// 由 src/driver/driver.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 463~698 行，逐行未改。


    // P6.6: ActiveX DLL模式, 额外生成 dll_entry.c (包含DLL导出函数)
    // 这个文件始终由driver生成, 而不是在单个模块的.c中生成
    // 原因: DLL工程可能只有类模块而没有标准模块, emitActiveXDll()在类模块中不会被执行
    if (options.isDll && !analyzers_.empty()) {
        // 使用最后一个analyzer的符号表 (已包含跨模块符号)
        auto& lastAnalyzer = analyzers_.back();
        // Fix 023: 传入 void* 字段表 (与主 codegen 循环一致)
        CCodeGen dllCgen(*diag_, lastAnalyzer->symbolTable(), lastAnalyzer->typeSystem(),
                         &classVoidFieldMap, &classTypedFieldMap, &variantReturnFuncs, options.verbose);
        dllCgen.setFormModuleNames(formModuleNames);  // Fix 086
        dllCgen.setModulePublicConsts(&modulePublicConsts);  // Fix 086
        dllCgen.setDesignerFiles(&frmFiles_);  // Fix 112
        dllCgen.setStaticModuleVarNames(&dupModuleVarNames);  // Fix 093b: 撞名模块级变量转 static
        dllCgen.setInterfaceImplementers(&ifaceImplementersMap);  // P6.4+: 接口唯一实现类映射
        dllCgen.setInterfaceAllImplementers(&ifaceAllImplementersMap);  // Fix <vbeclipse>: 接口全部实现类 (桩分派用)
        dllCgen.setDefaultInstanceClasses(&defaultInstanceClasses);  // P6.4+: 类模块默认实例
        // ai/022 B13b: dll_entry 的身份一律取 stage 2.7 的唯一出口 coclassIds_。
        // 只接 DLL 这条 cgen —— EXE 的 com_entry.c 本批刻意不接线（理由见 D57）。
        dllCgen.setCoClassIdentityRegistry(&coclassIds_);
        dllCgen.setInterfaceIidMap(&ifaceIds_);  // ai/022 B13c: 非默认接口的 IID 也读同一张表
        // ai/022 B16: 接口登记表 —— dll_entry 靠它判断"这个类的某个 Implements 是不是
        // 新式接口"（是则引用 vb6_iv_thin_<类>/vb6_iv_claim_<类>），判据与类模块自身发码
        // 时逐字相同（同一张 ivreg_）。
        dllCgen.setInterfaceRegistry(&ifaces_);
        // Collect all symbol tables for cross-module Property lookup
        std::vector<SymbolTable*> allSymTabs;
        for (auto& analyzer : analyzers_) {
            allSymTabs.push_back(&analyzer->symbolTable());
        }
        // P6.8: 注入VBP指定的CLSID到类符号
        for (auto& [key, sym] : lastAnalyzer->symbolTable().moduleScope()->symbols()) {
            if (sym->kind == SymbolKind::Class && !sym->isInterface) {
                std::string lowerName = sym->name;
                std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::tolower);
                auto it = classClsidMap_.find(lowerName);
                if (it != classClsidMap_.end() && sym->comClsidStr.empty()) {
                    sym->comClsidStr = it->second;
                }
            }
        }
        // P6.6.6: TypeLib building MUST run before generateDllEntry()
        // Reason: TypeLib builder writes back comDefaultIfaceIid and comClsidStr
        // to the symbol table, which cgen's generateDllEntry() reads

        // P9: Build TypeLib using CreateTypeLib2 API (replaces MIDL)
        {
            TypeLibBuilder tlbBuilder;
            std::string tlbPath = pathToUtf8(std::filesystem::absolute(utf8ToPath(outputDir + "/" + options.dllProgId + ".tlb")));
            std::string libId = options.libidStr.empty()
                ? TypeLibBuilder::generateUuid(options.dllProgId + ".TypeLib")
                : options.libidStr;

            // 类型库内部名 = 工程名 (= ProgID 前缀)。真 VB6 的 ActiveX DLL 就是这样: 实测
            // VBMAN.dll 的库名是 "VBMANLIB" (工程 Name=), 不是 "VBMANLIB.TypeLib"。
            // 客户端导入类型库后, coclass 的 ProgID 靠注册表反查 (ProgIDFromCLSID), 反查
            // 不到时退化成 "<库名>.<类名>" —— 库名多一段后缀就会拼出 CLSIDFromProgID 查不
            // 到的 ProgID, 运行期 New 报 429 (0x800401F3)。
            // LIBID 种子保持不变: 改它会给已构建产物换一套 GUID, 与本次修复无关。
            if (tlbBuilder.beginLib(tlbPath, libId, options.dllProgId + " TypeLib", options.dllProgId,
                                   /*is64=*/options.arch != "x86")) {
                int dispidCounter = 1;
                std::set<std::string> processedClasses;
                // 新式接口的 IID：与 vtable QI、dll_entry 的 IID 表读同一张 Pass E 表
                // （`ifaceIds_`）。查不到 = 不是新式接口 → 保持类型库自己那枚 legacy mint。
                auto ifaceIidFromMap = [&](const std::string& name) -> std::string {
                    auto it = ifaceIds_.find(Symbol::toLower(name));
                    return it == ifaceIds_.end() ? std::string() : it->second;
                };
                // 类 → 手写块的身份（与 cgen_util_dllentry_collect.inc 的匹配口径一致：
                // 折算记录一律跳过，存量 DLL 的 .tlb 逐字节不变）
                auto identityForClass = [&](const std::string& className) -> const CoClassIdentity* {
                    const std::string lower = Symbol::toLower(className);
                    for (const auto& kv : coclassIds_) {
                        const CoClassIdentity& id = kv.second;
                        if (id.name.empty() || id.legacyFolded) continue;
                        if (!id.implName.empty() && Symbol::toLower(id.implName) == lower) return &id;
                    }
                    return nullptr;
                };
                for (auto* symTab : allSymTabs) {
                    // 遍历模块级作用域中的符号, 找出 Public Class (ActiveX DLL 的 coclass)
                    auto* modScope = symTab->moduleScope();
                    if (!modScope) continue;
                    for (auto& [key, symPtr] : modScope->symbols()) {
                        if (!symPtr || symPtr->kind != SymbolKind::Class) continue;
                        // 跳过跨模块注入的外部类副本 (isExternal=true), 只处理宿主表中的本地类,
                        // 否则其成员查找会命中本模块/其他模块的同名方法导致参数污染
                        if (symPtr->isExternal) continue;
                        if (symPtr->instancing == VBInstancing::Private) continue;
                        auto& clsSym = *symPtr;
                        // Deduplicate: each class only once across all symbol tables
                        if (processedClasses.count(clsSym.name)) continue;
                        processedClasses.insert(clsSym.name);

                        // ai/022 B15 (D62-4/D62-5): **新式接口的宿主模块**在库里发成真接口
                        // (`TKIND_INTERFACE`)，并到此为止 —— 不再发 `_<名>` 那张空 dispinterface，
                        // 也不再发那一行 coclass。今天这条循环是唯一没有 `isInterface` 过滤的 COM
                        // 消费者（对照 `cgen_util_dllentry_collect.inc:26`、`driver_codegen_dll_sync.inc:8`），
                        // 于是接口模块在 COM 表里不存在、却在库里占了一整行**可创建**的 coclass，
                        // CLSID 还是另一枚谁也不认的 mint（`generateUuid(类名)`）。
                        // B16 起成员面如实发进去（见 addVtableInterface）：生成码自 B16 起整条薄面
                        // 都是 `__stdcall`，与库里的 `CC_STDCALL` + `oVft=(3+槽号)*指针宽` 一一对应。
                        // 判据沿用 B13c 那一条（Pass E 命中 = 这个模块名就是一个新式接口），
                        // legacy 那条（被 `Implements` 点名但没有 Interface 块的类）一行不动。
                        {
                            const std::string ifaceRowIid = ifaceIidFromMap(clsSym.name);
                            auto ifaceIt = ifaces_.find(Symbol::toLower(clsSym.name));
                            if (!ifaceRowIid.empty() && ifaceIt != ifaces_.end()) {
                                // 成员面 (B16)。三处口径都从 `IfaceView::slots` 来 —— 与生成码
                                // `vb6_ivtbl_<I>` 读的是同一份、同一个顺序，所以槽号对得上：
                                //   名字: COM 惯例 (Property 是 get_/put_/putref_ 前缀, 过程用声明里的
                                //         大小写); 同名 Property 的 Get/Let/Set 共享一个 memid。
                                //   类型: SimpleTypeRef 走类型表 (Long→VT_I4/String→VT_BSTR), 其余
                                //         (无 As / 数组 / LongPtr / UDT / 工程类) 落 VT_VARIANT 兜底
                                //         —— 这一档"类型保真"的缺口属 B17, 见 022 D63。
                                std::vector<TypeLibBuilder::MethodInfo> vtableMembers;
                                {
                                    TypeSystem& ts = lastAnalyzer->typeSystem();
                                    auto tlTypeOf = [&](const ASTNode* tr) -> Vb6Type {
                                        if (!tr || tr->kind != ASTNodeKind::SimpleTypeRef) {
                                            return Vb6Type::Variant;
                                        }
                                        return ts.resolveTypeName(
                                            static_cast<const SimpleTypeRef*>(tr)->name);
                                    };
                                    std::map<std::string, int32_t> propMemid;
                                    int32_t nextMemid = 1;
                                    for (const IfaceSlotView& slot : ifaceIt->second.slots) {
                                        const Decl* sig = slot.sig;
                                        if (!sig) {
                                            std::cerr << "C3: TypeLib vtable member skipped (no sig): "
                                                      << ifaceIt->second.name << "." << slot.key
                                                      << std::endl;
                                            continue;
                                        }
                                        TypeLibBuilder::MethodInfo mi;
                                        std::string base;
                                        bool isGet = false, isPut = false, isPutRef = false;
                                        const std::vector<std::unique_ptr<ParameterDecl>>* ps = nullptr;
                                        switch (sig->kind) {
                                            case ASTNodeKind::SubDecl: {
                                                const SubDecl& d = static_cast<const SubDecl&>(*sig);
                                                base = d.name;
                                                mi.returnType = Vb6Type::Void;
                                                ps = &d.params;
                                                break;
                                            }
                                            case ASTNodeKind::FunctionDecl: {
                                                const FunctionDecl& d = static_cast<const FunctionDecl&>(*sig);
                                                base = d.name;
                                                mi.returnType = tlTypeOf(d.returnType.get());
                                                ps = &d.params;
                                                break;
                                            }
                                            case ASTNodeKind::PropertyDecl: {
                                                const PropertyDecl& d = static_cast<const PropertyDecl&>(*sig);
                                                base = d.name;
                                                isGet = d.propKind == ProcKind::PropertyGet;
                                                isPut = d.propKind == ProcKind::PropertyLet;
                                                isPutRef = d.propKind == ProcKind::PropertySet;
                                                mi.returnType = isGet ? tlTypeOf(d.returnType.get())
                                                                      : Vb6Type::Void;
                                                ps = &d.params;
                                                break;
                                            }
                                            default:
                                                break;
                                        }
                                        mi.name = (isGet ? "get_" : (isPut ? "put_"
                                                   : (isPutRef ? "putref_" : std::string()))) + base;
                                        mi.isPropertyGet = isGet;
                                        mi.isPropertyPut = isPut;
                                        mi.isPropertyPutRef = isPutRef;
                                        if (isGet || isPut || isPutRef) {
                                            const std::string pkey = Symbol::toLower(base);
                                            auto pm = propMemid.find(pkey);
                                            if (pm != propMemid.end()) {
                                                mi.dispid = pm->second;
                                            } else {
                                                mi.dispid = nextMemid++;
                                                propMemid[pkey] = mi.dispid;
                                            }
                                        } else {
                                            mi.dispid = nextMemid++;
                                        }
                                        if (ps) {
                                            for (const auto& p : *ps) {
                                                ParameterInfo pi;
                                                pi.name = p->name;
                                                pi.type = tlTypeOf(p->asType.get());
                                                pi.isByVal = p->isByVal;
                                                pi.isOptional = p->isOptional;
                                                pi.isParamArray = p->isParamArray;
                                                mi.params.push_back(pi);
                                            }
                                        }
                                        vtableMembers.push_back(std::move(mi));
                                    }
                                }
                                if (tlbBuilder.addVtableInterface(ifaceIt->second.name, ifaceRowIid,
                                                                  vtableMembers)) {
                                    clsSym.comDefaultIfaceIid = ifaceRowIid;
                                    clsSym.comDefaultIfaceName = ifaceIt->second.name;
                                } else {
                                    std::cerr << "C3: TypeLib addVtableInterface failed for "
                                              << clsSym.name << ": " << tlbBuilder.lastError() << std::endl;
                                }
                                continue;
                            }
                        }

                        std::vector<TypeLibBuilder::MethodInfo> methods;
                        // 类→宿主符号表映射 (与 cgen_util.cpp Fix 016b 一致):
                        // 只记录 isExternal=false 的本地 Class 所在表, 成员查找优先用宿主表,
                        // 彻底避免跨表 lookupModuleByKind 命中其他类同名方法 (参数污染,
                        // 如 cCollection.Add 被查成 cHttpServerRouterAfter.Add 变成
                        // RouteName/Handler/MethodLimit, 导致 SetFuncAndParamNames 因参数
                        // 个数不匹配返回 TYPE_E_TYPEMISMATCH 0x800280EC)。
                        std::unordered_map<std::string, SymbolTable*> classOwningSymTab;
                        for (auto* st_ : allSymTabs) {
                            if (!st_ || !st_->moduleScope()) continue;
                            for (auto& [k, cs] : st_->moduleScope()->symbols()) {
                                if (cs && cs->kind == SymbolKind::Class && !cs->isExternal) {
                                    classOwningSymTab.emplace(Symbol::toLower(cs->name), st_);
                                }
                            }
                        }
                        auto findMemberInClass = [&](const std::string& name, SymbolKind kind) -> Symbol* {
                            // 1) 类宿主表精确查找 (无跨类碰撞)。
                            //    类宿主表 (classOwningSymTab) 由语义分析注册, 包含该类全部成员变体
                            //    (Sub/Function/Property Get/Let/Set), 跨表 fallback 反而会命中
                            //    其他类的同名方法 (如 cCollection 的 Sub Add(Item,Key) 之外, 又
                            //    从其他类找到 Function Add(RouteName,Handler,MethodLimit), 导致
                            //    SetFuncAndParamNames 因参数个数不匹配返回 TYPE_E_TYPEMISMATCH)。
                            auto itOwn = classOwningSymTab.find(Symbol::toLower(clsSym.name));
                            if (itOwn != classOwningSymTab.end() && itOwn->second && itOwn->second->moduleScope()) {
                                // Fix 048: 排除 isBuiltin — 内置函数 (IsEmpty/Filter/Timer 等)
                                // 注册于每个模块符号表 (key=无后缀名), 若类成员查找命中它们,
                                // TypeLib 会出现幽灵 Function 变体 (与真实属性 $pg/$pl 同名同
                                // dispid, SetFuncAndParamNames 返回 TYPE_E_TYPEMISMATCH)。
                                Symbol* s = itOwn->second->lookupModuleByKind(name, kind);
                                if (s && !s->isExternal && !s->isBuiltin) {
                                    return s;
                                }
                            }
                            // 2) 当前表兜底 (只接受本地符号)
                            Symbol* s2 = symTab->lookupModuleByKind(name, kind);
                            if (s2 && !s2->isExternal && !s2->isBuiltin) {
                                return s2;
                            }
                            return nullptr;
                        };
                        // memberNames 可能含大小写不同的重复名 (VB6 标识符不区分大小写),
                        // 重复处理会导致同一方法被添加多次, 且第二次查找可能命中其他类的同名
                        // 方法符号 (参数污染)。按大小写不敏感去重。
                        std::unordered_set<std::string> seenMemberNames;
                        for (auto& memberName : clsSym.memberNames) {
                            if (!seenMemberNames.insert(Symbol::toLower(memberName)).second) continue;
                            // 同名属性可能有 PropertyGet/Let/Set 多个符号
                            // memberNames 中同名属性只存一次, 需要分别查找各变体
                            // M29: cross-symTab fallback - the class may be found in a "merged" analyzer
                            // (e.g., lastAnalyzer), but its Property Let/Set variants might only exist in
                            // the original per-module analyzer's table. Fall back to allSymTabs to find them.
                            auto* subSym = findMemberInClass(memberName, SymbolKind::Sub);
                            auto* fnSym = findMemberInClass(memberName, SymbolKind::Function);
                            auto* propGetSym = findMemberInClass(memberName, SymbolKind::PropertyGet);
                            auto* propLetSym = findMemberInClass(memberName, SymbolKind::PropertyLet);
                            auto* propSetSym = findMemberInClass(memberName, SymbolKind::PropertySet);

                            // 收集所有 Public 成员变体 (PropertyGet 必须在 PropertyLet 前面)
                            struct MemberRef { Symbol* sym; bool isGet; bool isPut; bool isPutRef; };
                            std::vector<MemberRef> refs;
                            if (subSym && subSym->access == AccessLevel::Public)
                                refs.push_back({subSym, false, false, false});
                            if (fnSym && fnSym->access == AccessLevel::Public)
                                refs.push_back({fnSym, false, false, false});
                            if (propGetSym && propGetSym->access == AccessLevel::Public)
                                refs.push_back({propGetSym, true, false, false});
                            if (propLetSym && propLetSym->access == AccessLevel::Public)
                                refs.push_back({propLetSym, false, true, false});
                            if (propSetSym && propSetSym->access == AccessLevel::Public)
                                refs.push_back({propSetSym, false, false, true});

                            for (auto& mr : refs) {
                                TypeLibBuilder::MethodInfo mi;
                                mi.name = mr.sym->name;
                                mi.returnType = mr.sym->type;
                                mi.isPropertyGet = mr.isGet;
                                mi.isPropertyPut = mr.isPut;
                                mi.isPropertyPutRef = mr.isPutRef;
                                // Property Get/Let/Set with same name share DISPID
                                bool foundExistingDispid = false;
                                for (const auto& existing : methods) {
                                    if (existing.name == mi.name) {
                                        mi.dispid = existing.dispid;
                                        foundExistingDispid = true;
                                        break;
                                    }
                                }
                                if (!foundExistingDispid) {
                                    mi.dispid = dispidCounter++;
                                }
                                // M29: 回写dispid到method symbol, 供cgen生成dll_entry.c方法表使用, 确保TypeLib与dll_entry.c dispid一致
                                mr.sym->comDispid = mi.dispid;
                                for (auto& param : mr.sym->params) {
                                    mi.params.push_back(param);
                                }
                                methods.push_back(mi);
                            }
                        }

                        // Fix 099: 显式 Public 字段 —— 真 VB6 把类模块的 `Public X As T`
                        // 在 TypeLib 里暴露为 Property Get/Let 对 (共享同一 dispid).
                        // 不生成会有两处后果: ① 早绑定客户端在类型库里找不到该属性;
                        // ② dll_entry.c 的字段表取不到 dispid, 只能走防御分支, 与 TypeLib
                        // 错位 (Invoke 按 dispid 线性匹配会命中错误表项).
                        // 这里只分配 dispid 并回写到 Class 符号; 运行期读写由 cgen 生成的
                        // vb6_cls_<cls>_field_get_/_let_<fld> 提供 (见 cgen_com.cpp).
                        // 字段类型统一按 Variant 声明 —— 不在这里做类型解析 (避免解析失败
                        // 导致整个 TypeLib 构建出错), 早绑定客户端可隐式转换 Variant.
                        {
                            std::unordered_set<std::string> seenFields;
                            for (auto& fname : clsSym.publicFieldNames) {
                                if (!seenFields.insert(Symbol::toLower(fname)).second) continue;
                                int32_t fieldDispid = dispidCounter++;
                                TypeLibBuilder::MethodInfo fGet;
                                fGet.name = fname;
                                fGet.dispid = fieldDispid;
                                fGet.returnType = Vb6Type::Variant;
                                fGet.isPropertyGet = true;
                                methods.push_back(fGet);

                                TypeLibBuilder::MethodInfo fPut;
                                fPut.name = fname;
                                fPut.dispid = fieldDispid;
                                fPut.returnType = Vb6Type::Void;
                                fPut.isPropertyPut = true;
                                ParameterInfo fieldValueParam;
                                fieldValueParam.name = "Value";
                                fieldValueParam.type = Vb6Type::Variant;
                                fieldValueParam.isByVal = true;
                                fPut.params.push_back(fieldValueParam);
                                methods.push_back(fPut);

                                clsSym.memberFieldDispids[Symbol::toLower(fname)] = fieldDispid;
                            }
                        }

                        std::string ifaceName = "_" + clsSym.name;
                        // ai/022 B13c: 这个"类"其实是个**新式接口模块**（头行 `Interface IProbe`）
                        // → 它的类型库 GUID 就是那个接口的 IID，取 Pass E 的表，不再 mint 一枚
                        // 只有类型库知道的值（D58 读数：IProbe 在一次编译里曾有 4 枚 GUID）。
                        const std::string mapIid = ifaceIidFromMap(clsSym.name);
                        std::string iid = mapIid.empty()
                            ? TypeLibBuilder::generateUuid("_" + clsSym.name) : mapIid;
                        bool ifaceOk = tlbBuilder.addDispInterface(ifaceName, iid, methods);
                        if (!ifaceOk) {
                            std::cerr << "C3: TypeLib addDispInterface failed for " << ifaceName << ": " << tlbBuilder.lastError() << std::endl;
                            for (size_t mi_ = 0; mi_ < methods.size(); mi_++) {
                                auto& mm = methods[mi_];
                                fprintf(stderr, "  [%zu] %s get=%d put=%d putref=%d dispid=%d nparams=%zu\n",
                                        mi_, mm.name.c_str(), (int)mm.isPropertyGet, (int)mm.isPropertyPut,
                                        (int)mm.isPropertyPutRef, mm.dispid, mm.params.size());
                            }
                        }

                        // Only write back IID if addDispInterface succeeded
                        if (ifaceOk) {
                            clsSym.comDefaultIfaceIid = iid;
                            clsSym.comDefaultIfaceName = ifaceName;
                        }

                        // If class has events, create source dispinterface (_ClassNameEvents)
                        std::string sourceIfaceName;
                        bool sourceIfaceOk = false;
                        if (!clsSym.eventNames.empty()) {
                            sourceIfaceName = "_" + clsSym.name + "Events";
                            std::string sourceIid = TypeLibBuilder::generateUuid(sourceIfaceName);
                            std::vector<TypeLibBuilder::MethodInfo> eventMethods;
                            for (size_t ei = 0; ei < clsSym.eventNames.size(); ei++) {
                                TypeLibBuilder::MethodInfo emi;
                                emi.name = clsSym.eventNames[ei];
                                emi.dispid = (int32_t)(ei + 1);
                                emi.returnType = Vb6Type::Void;
                                eventMethods.push_back(emi);
                            }
                            sourceIfaceOk = tlbBuilder.addDispInterface(sourceIfaceName, sourceIid, eventMethods);
                            if (!sourceIfaceOk) {
                                std::cerr << "C3: TypeLib addDispInterface failed for " << sourceIfaceName << ": " << tlbBuilder.lastError() << std::endl;
                            }

                            // Only write back sourceIfaceIid if addDispInterface succeeded
                            if (sourceIfaceOk) {
                                clsSym.comSourceIfaceIid = sourceIid;
                                clsSym.comSourceIfaceName = sourceIfaceName;
                            }

                            // Write event DISPIDs to symbol table
                            for (size_t ei = 0; ei < clsSym.eventNames.size(); ei++) {
                                std::string evtLower = Symbol::toLower(clsSym.eventNames[ei]);
                                clsSym.comEventDispids[evtLower] = eventMethods[ei].dispid;
                            }
                        }

                        // Only add CoClass if primary interface was added successfully
                        if (ifaceOk) {
                            const CoClassIdentity* ccid = identityForClass(clsSym.name);
                            std::string clsid;
                            if (ccid && !ccid->clsid.empty()) {
                                // ai/022 B13c: 手写块的 CLSID 取唯一出口（与 dll_entry 同值）
                                clsid = ccid->clsid;
                            } else if (!clsSym.comClsidStr.empty()) {
                                clsid = clsSym.comClsidStr;
                            } else {
                                auto it = classClsidMap_.find(clsSym.lowerName);
                                if (it != classClsidMap_.end()) clsid = it->second;
                            }
                            if (clsid.empty()) clsid = TypeLibBuilder::generateUuid(clsSym.name);
                            if (clsSym.comClsidStr.empty()) clsSym.comClsidStr = clsid;
                            // ai/022 B13e: 默认接口**一律**是 `<_类名>`，不跟块里的 [Default]。
                            // B13c 曾把它改指向块的 `[Default]` 接口（`_IProbe`），实测是把"广告"
                            // 与"应答"分开：类型库说默认视图是 `_IProbe`，而服务器的成员面
                            // （`GetIDsOfNames`/`Invoke` 查 `desc->methods`）只有类的 Public 成员，
                            // 两头都点不到。谁应答就广告谁。
                            if (!tlbBuilder.addCoClass(clsSym.name, clsid, ifaceName,
                                                       sourceIfaceOk ? sourceIfaceName : std::string())) {
                                std::cerr << "C3: TypeLib addCoClass failed for " << clsSym.name << ": "
                                          << tlbBuilder.lastError() << std::endl;
                            }
                        }
                    }
                }

                if (!tlbBuilder.endLib(tlbPath)) {
                    std::cerr << "C3: TypeLib generation failed: " << tlbBuilder.lastError() << std::endl;
                } else if (options.verbose) {
                    std::cout << "C3: TypeLib generated: " << tlbPath << std::endl;
                }
            } else {
                std::cerr << "C3: TypeLib init failed: " << tlbBuilder.lastError() << std::endl;
            }
        }

