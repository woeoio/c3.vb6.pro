// cgen_util_dllentry_collect.inc - generateDllEntry 函数体片段：coclass 收集、空类桩、前向声明与 IDispatch 桥接
// 由 src/backend/cgen_util.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 73~397 行，逐行未改。

    // Fix 016b: 构建 className(lower) -> 定义该类的 SymbolTable* 映射.
    // COM 分派表收集每个类的 Public 方法时, 必须只搜索该类自身的符号表,
    // 不能跨 allSymTabs 搜索同名方法 — 否则跨类同名冲突 (如 6 个类的 Mode:
    // cCryptoHMAC/cCryptoHash 是 Function, cDelay/cModbusTransportTCP/... 是
    // PropertyGet/Let) 会导致 cDelay 的 PropertyGet "mode$pg" 泄漏进 cCryptoHMAC
    // 的分派表, 生成不存在的 vb6_cCryptoHMAC_prop_get_Mode → LNK2019.
    // 每个类的本地方法符号只存在于定义该类的模块符号表中, 搜索 ownTab 即可
    // 精确获取该类自己的方法, 完全避免跨类碰撞.
    std::unordered_map<std::string, SymbolTable*> classOwningSymTab;
    for (auto* st : allSymTabs) {
        if (!st->moduleScope()) continue;
        for (const auto& [ck, cs] : st->moduleScope()->symbols()) {
            if (cs->kind == SymbolKind::Class && !cs->isExternal) {
                classOwningSymTab.emplace(Symbol::toLower(cs->name), st);
            }
        }
    }

    // Fix <vbeclipse> rev17: coclass 收集必须遍历**全部模块**的符号表, 不能只看
    // `symTab_` (= 最后一个模块) 的那一张。原先只看单表的理由是"类符号会经跨模块
    // 注入把副本铺进每个模块" —— 但那条路有两处会漏 (实测 tests/ve_list, 4 个类只
    // 收进 2 个):
    //   ① 类名与成员同名时 (View.cls `Property Get View` / PopupMenu.cls
    //      `Sub PopupMenu`), 成员的 storageKey 与类符号的裸键相同或经名字级回落被
    //      命中 → driver_crossmod 的注入守卫判"本地已定义"直接 SKIP;
    //   ② 注入是"先到先得"的单键存放, 一个名字只能活一个符号。
    // 而**类在自己模块的表里一定是权威副本** (semantic_analyzer 预先 define, 且那里
    // 带完整 memberNames/publicFieldNames/comClsidStr…)。改成遍历 allSymTabs 后,
    // 收集与注入彻底解耦, 上面两条漏路都不再影响 coclass 表。
    // 实测 play78: 修前表里 14 项, 缺 View/PopupMenu (恰是全工程唯一两个有同名成员的
    // 类) → 运行期 vb6_FindCoClassDesc("View")=NULL → 装箱出 vt=9/pdisp=NULL 空壳
    // Variant → `Is Nothing` 恒真 → ucFolder.AddView 收到 NULL → av read 0x4。
    // Fix 053: 去重 — 同一类会在多个模块的表里各出现一次 (自身副本 + 注入副本)
    std::set<std::string> seenCoClassNames;
    for (auto* st : allSymTabs) {
    if (!st->moduleScope()) continue;
    for (auto& [key, sym] : st->moduleScope()->symbols()) {
        if (sym->kind == SymbolKind::Class && !sym->isInterface) {
            // BUG-4 fix: PublicNotCreatable(2) is visible but NOT externally creatable
            // Only MultiUse(5)/SingleUse(3) should have ClassFactory entries
            if (sym->instancing == VBInstancing::MultiUse || sym->instancing == VBInstancing::SingleUse) {
                std::string lowerName = Symbol::toLower(sym->name);
                if (seenCoClassNames.count(lowerName)) continue;  // 已处理, 跳过
                seenCoClassNames.insert(lowerName);
                // Fix <vbeclipse> rev17c: rev17 让本收集锚定 allSymTabs 的**首个**命中 ——
                // 对被消费的类, 首个命中可能是别的模块里的**注入副本** (isExternal)。
                // 成员面 (memberNames/instancing/implementsNames) 注入时逐字段拷贝过,
                // 副本读来无碍; 但 driver 的回写三样 (P6.6.6 typelib 的 comDefaultIfaceIid/
                // comSourceIfaceIid、P6.8 的 comClsidStr) 只落**宿主表**的本地符号
                // (typelib 扫描跳过 isExternal), dll_sync 又只灌最后一张表 —— 副本上
                // 永远是空。实测 CoDll: IProbe.cls 排在 CImpl.cls 之前, collect 在
                // IProbe 的表里先撞见 CImpl 的注入副本 ⇒ defaultIfaceIid=NULL ⇒
                // dll_entry.c 不发 IID_vb6def_<类> ⇒ cc_dll_identity_single_source 的
                // 0x7CA8CD81 针与 cc_dll_tlb_matches_table 同时红 (GA vbp#3/#4)。
                // 口径与上面 rev17 注释一致 ("类在自己模块的表里一定是权威副本"):
                // 回写类字段一律锚定宿主符号 authCls, 副本只供成员面; 查不到宿主时
                // 原样回落 sym (单模块工程两者本就是同一只)。
                SymbolTable* ownTab = nullptr;
                {
                    auto itOwnH = classOwningSymTab.find(lowerName);
                    if (itOwnH != classOwningSymTab.end()) ownTab = itOwnH->second;
                }
                Symbol* hostClsSym = nullptr;
                if (ownTab) {
                    Symbol* hs = ownTab->lookupModule(sym->name);
                    if (hs && hs->kind == SymbolKind::Class && !hs->isExternal) hostClsSym = hs;
                }
                Symbol& authCls = hostClsSym ? *hostClsSym : *sym;
                CoClassInfo info;
                info.moduleName = sym->name;
                info.clsidStr = authCls.comClsidStr.empty() ? generateClsid(progId + "." + sym->name) : authCls.comClsidStr;  // P6.8: 优先使用VBP指定的CLSID
                info.progId = progId + "." + sym->name;

                // ai/022 B13b: 身份从唯一出口 coclassIds_ 取（026 三节"不可让步" + D47"不许
                // 第二条通道"；D56-3 实测：登记表在后端一个消费者都没有，产品注册的从来不是
                // 语义层那一份）。登记表按**块名**索引，这里按 [Implementation] 反查实现类。
                // **只接手写块**：折算记录（C04）的 clsid/iid 一旦覆盖上去，存量 DLL 工程的
                // 注册身份就会换 mint（两族种子不同），那是"把改码前编得过的工程产物改掉"，
                // 与逐字节护栏冲突 —— legacyFolded 一律照旧走 comClsidStr / legacy lambda。
                const CoClassIdentity* ccid = nullptr;
                if (ccreg_) {
                    for (const auto& kv : *ccreg_) {
                        const CoClassIdentity& id = kv.second;
                        if (id.name.empty() || id.legacyFolded) continue;
                        if (Symbol::toLower(id.implName) == lowerName) { ccid = &id; break; }
                    }
                }
                if (ccid) {
                    if (!ccid->clsid.empty()) info.clsidStr = ccid->clsid;
                    // ai/022 B13e：**默认接口的 IID 不再从出口覆盖**。B13b 那样做之后，产品里的
                    // "默认接口"变成块指定的那个新式接口，可它对外既没有成员（契约成员是 `Private`）、
                    // 类型库里的渲染又是 0 成员的 dispinterface ⇒ 服务器 `GetIDsOfNames` 认的成员面
                    // （= 类的公有成员那一档）与库广告出来的那一档不再是同一档。新式接口在这里只留
                    // **身份**（下面 `iidMap` 走 `iidreg_` 那条，与 vtable QI/类型库同值），
                    // 门面交回 `<_类名>`；`defaultIfaceIid` 沿用 typelib 的回写值（见本文件末尾），
                    // 于是"广告的那一枚 == 服务器应答的那一枚"重新成立。
                    // 组名档 ProgID：块名与类模块名不同名才有意义（VB6 惯用的同名写法不产生第二行），
                    // 且必须 [ComCreatable(True)] —— 这个位今天第一次在产品侧有后果。
                    if (ccid->comCreatable && Symbol::toLower(ccid->name) != lowerName) {
                        info.groupProgId = progId + "." + ccid->name;
                    }
                }

                // 使用类的memberNames查找Public方法符号
                // (不能按前缀"ClassName_"搜索, 因为单模块工程中方法名是"SetValue"而非"Calc_SetValue")
                // Fix 016b: 优先搜索该类自身的符号表 (classOwningSymTab), 避免跨类同名碰撞.
                // rev17c: ownTab/hostClsSym 已在上面随 authCls 一起锚定, 这里直接复用.
                std::string classNameLower = lowerName;

                // Fix 053: 去重 memberNames — Event 和 Sub/Function 可能同名,
                // memberNames 中会有重复条目, 导致 dispatch 函数被生成两次 (C2084)
                std::set<std::string> seenMemberNames;
                // Fix 093a: 事件名集合 (小写) — Event 无实现函数. 若被当作 Public
                // 方法收集 (如 cTimer/cTimers 的 `Event Timer()` 与内置 Timer() 同名,
                // lookupModule 命中内置符号) 会生成 vb6_disp_<cls>_<ev>_invoke 并
                // extern 引用不存在的 vb6_<cls>_<ev> → LNK2019. 事件不进方法表.
                std::set<std::string> eventNamesLower;
                for (const auto& en : authCls.eventNames) eventNamesLower.insert(Symbol::toLower(en));
                for (auto& memberName : authCls.memberNames) {
                    if (!seenMemberNames.insert(Symbol::toLower(memberName)).second) continue;
                    if (eventNamesLower.count(Symbol::toLower(memberName))) continue;
                    // Try Sub/Function: 优先在类自身符号表搜索 (本类方法, 无碰撞)
                    Symbol* memSym = nullptr;
                    if (ownTab) memSym = ownTab->lookupModule(memberName);
                    if (!memSym) memSym = symTab_.lookupModule(memberName);
                    if (!memSym) memSym = symTab_.lookup(memberName);
                    if (!memSym) {
                        for (auto* st : allSymTabs) {
                            memSym = st->lookupModule(memberName);
                            if (memSym) break;
                        }
                    }
                    // Fix 016b: 验证找到的符号属于当前类 (外部符号查 sourceModule;
                    // 本地符号只需 ownTab 命中即可, 因 ownTab 只含本类定义)
                    if (memSym && (memSym->kind == SymbolKind::Sub || memSym->kind == SymbolKind::Function)
                        && memSym->access == AccessLevel::Public && !memSym->isBuiltin) {
                        bool belongs = false;
                        if (memSym->isExternal) {
                            belongs = (Symbol::toLower(memSym->sourceModule) == classNameLower);
                        } else {
                            belongs = (ownTab && ownTab->lookupModule(memberName) == memSym);
                        }
                        if (belongs) {
                            info.publicMethodNames.push_back(memSym->name);
                            info.publicMethodSyms.push_back(memSym);
                            memSym = nullptr;  // Don't double-count
                        }
                    }
                    // Lookup Property Get/Let/Set: 优先在类自身符号表搜索
                    for (auto kind : {SymbolKind::PropertyGet, SymbolKind::PropertyLet, SymbolKind::PropertySet}) {
                        Symbol* propSym = nullptr;
                        if (ownTab) propSym = ownTab->lookupModuleByKind(memberName, kind);
                        if (!propSym) propSym = symTab_.lookupModuleByKind(memberName, kind);
                        if (!propSym) {
                            for (auto* st : allSymTabs) {
                                propSym = st->lookupModuleByKind(memberName, kind);
                                if (propSym) break;
                            }
                        }
                        if (propSym && propSym->access == AccessLevel::Public) {
                            // Fix 016b: 验证归属, 阻止其他类的同名 Property 泄漏
                            bool belongs = false;
                            if (propSym->isExternal) {
                                belongs = (Symbol::toLower(propSym->sourceModule) == classNameLower);
                            } else {
                                belongs = (ownTab && ownTab->lookupModuleByKind(memberName, kind) == propSym);
                            }
                            if (belongs) {
                                info.publicMethodNames.push_back(propSym->name);
                                info.publicMethodSyms.push_back(propSym);
                            }
                        }
                    }
                }
                // P12.1: 收集 Implements 接口列表
                info.implementsNames = authCls.implementsNames;
                // ai/022 B16: 本类是否实现**新式接口** —— 判据与 cgen_iface_vtbl 发码
                // 时用的那条逐字相同（ivLookupIface 命中：登记表里有、链没断、有槽），
                // 因为这一位决定 dll_entry.c 要不要引用 vb6_iv_thin_<类> /
                // vb6_iv_claim_<类> 这两个符号（引用一处就必须发一处）。
                // 注: dllCgen 的 ivreg_ 由 driver 在调用 generateDllEntry 前注入
                // （driver_codegen_dll_typelib.inc），没注时这里恒 false = 退回胖路。
                for (const auto& ifaceName : info.implementsNames) {
                    if (ivLookupIface(ifaceName)) { info.hasNewStyleIface = true; break; }
                }
                // Fix 099: 收集显式 Public 字段 (semantic_analyzer 填的 publicFieldNames).
                // 真 VB6 把类的 Public 字段暴露为 Property Get/Let; 此前 C3 完全不进
                // IDispatch 表 → 客户端 GetIDsOfNames("Router") 失败 → vb6_ComGetProp
                // 返回 NULL → 调用方按对象解引用即崩.
                // dispid 必须从该类**宿主符号表**里的符号读: driver 在 TypeLib 阶段把
                // dispid 回写到宿主符号, 而本函数用的 symTab_ 是最后一个模块的表, 其中
                // 的类可能是 driver_crossmod 在 TypeLib **之前**注入的 external 副本 —
                // 那份副本的 memberFieldDispids 永远是空的 (实测 cHttpServer 等落到
                // 防御段 9000+), 而 TypeLib 里是真的连续 dispid, 两边错位会让早绑定
                // 客户端按 TypeLib 的 dispid 调 Invoke 命中错误表项.
                // rev17c: hostClsSym 已随 authCls 在上面锚定 (同一判据: 宿主表 + 本地
                // Class 符号), dispid 与 IID/CLSID 同源同锚.
                const Symbol* dispidSym = (hostClsSym && hostClsSym->kind == SymbolKind::Class)
                    ? hostClsSym : sym.get();
                for (auto& fname : authCls.publicFieldNames) {
                    info.publicFields.push_back(fname);
                    auto itFd = dispidSym->memberFieldDispids.find(Symbol::toLower(fname));
                    info.publicFieldDispids.push_back(
                        (itFd != dispidSym->memberFieldDispids.end()) ? itFd->second : 0);
                }
                // P6.6.4: 收集事件信息
                if (!authCls.eventNames.empty()) {
                    info.eventNames = authCls.eventNames;
                    // P6.6.6: 优先使用TypeLib builder回写的IID, 保证与TypeLib中一致
                    info.sourceIfaceIid = authCls.comSourceIfaceIid.empty()
                        ? generateIid("source:" + progId + "." + sym->name + "Events")
                        : authCls.comSourceIfaceIid;
                }
                // P6.6.6: 默认接口IID (由driver TypeLib builder回写)
                if (info.defaultIfaceIid.empty()) info.defaultIfaceIid = authCls.comDefaultIfaceIid;
                coClasses.push_back(std::move(info));
            }
        }
    }
    }

    if (coClasses.empty()) {
        // BUG-2 fix: Generate minimal stubs (prevents linker errors for empty ActiveX DLL)
        entry.emitLine("// Generated by vb6c3 (C3.exe) - ActiveX DLL entry point (no public classes)");
        entry.emitLine("#include \"vb6comserver.h\"");
        entry.emitBlank();
        entry.emitLine("const vb6_CoClassDesc g_vb6_coclasses[] = {{0}};");
        entry.emitLine("const int g_vb6_coclassCount = 0;");
        entry.emitBlank();
        entry.emitLine("HRESULT WINAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv) {");
        entry.indent();
        entry.emitLine("return CLASS_E_CLASSNOTAVAILABLE;");
        entry.dedent();
        entry.emitLine("}");
        entry.emitBlank();
        entry.emitLine("HRESULT WINAPI DllCanUnloadNow(void) { return S_OK; }");
        entry.emitLine("HRESULT WINAPI DllRegisterServer(void) { return S_OK; }");
        entry.emitLine("HRESULT WINAPI DllUnregisterServer(void) { return S_OK; }");
        entry.emitBlank();
        return entry.str();
    }

    // 生成 dll_entry.c 内容

    // 头文件注释
    entry.emitLine("// Generated by vb6c3 (C3.exe) - ActiveX DLL entry point");
    entry.emitLine("// Contains COM server exports: DllGetClassObject, DllRegisterServer, etc.");
    entry.emitBlank();

    // 包含必要头文件
    // Fix 010b: 添加vb6rtl.h以提供vb6_VARIANT等类型定义
    // (vb6comserver.h只include windows.h/oleauto.h, 不含vb6_VARIANT)
    // vb6rtl.h使用vb6_前缀避免与Windows VARIANT冲突, 可安全与oleauto.h共存
    entry.emitLine("#include \"vb6rtl.h\"");
    entry.emitLine("#include \"vb6comserver.h\"");
    entry.emitBlank();

    // 前向声明: 类结构体和工厂/销毁函数 (C语言需要struct关键字)
    // Fix 186: 必须先集中发完所有类结构体的 typedef, 再发任何函数声明.
    // 本文件既有 "struct X*" 用法, 也有裸 "X*" 用法 (主要是工程类实例返回值的
    // extern 声明与其 _p 局部变量); typedef 若与其使用点交错, 声明靠前的使用点
    // 仍会撞上未声明的裸名 → C2143/C2065, 整个 dll_entry.c 编译失败.
    for (auto& cc : coClasses) {
        std::string clsStruct = "vb6_cls_" + cIdent(cc.moduleName);
        entry.emitLine("typedef struct " + clsStruct + " " + clsStruct + ";  /* forward decl */");
    }
    entry.emitBlank();
    for (auto& cc : coClasses) {
        std::string clsId = cIdent(cc.moduleName);
        std::string clsStruct = "vb6_cls_" + clsId;
        entry.emitLine("extern struct " + clsStruct + "* vb6_cls_" + clsId + "_New(void);");
        entry.emitLine("extern void vb6_cls_" + clsId + "_Destroy(struct " + clsStruct + "*);");
        // ai/022 B16: 新式接口那一面（薄指针出入口）由类模块自己的 TU 提供，见
        // cgen_iface_vtbl.cpp 的 emitIfaceImplTables 末尾. 只在真的发了这两个符号的
        // 类上声明（判据同上），否则整行都不出现 —— 无新式接口的工程逐字节不变.
        if (cc.hasNewStyleIface) {
            entry.emitLine("extern void* vb6_iv_thin_" + clsId +
                           "(void* instance, const void* riid);  /* B16: 接口 IID -> 薄指针 */");
            entry.emitLine("extern void vb6_iv_claim_" + clsId +
                           "(void* instance);  /* B16: COM 包装器放手, 退底座引用 */");
        }
        // 类的方法函数前向声明
        for (size_t mi = 0; mi < cc.publicMethodNames.size(); mi++) {
            Symbol* methodSym = cc.publicMethodSyms[mi];
            // Property需要加前缀: prop_get_/prop_let_/prop_set_
            std::string methodCName = cc.publicMethodNames[mi];
            if (methodSym->kind == SymbolKind::PropertyGet) {
                methodCName = "prop_get_" + methodCName;
            } else if (methodSym->kind == SymbolKind::PropertyLet) {
                methodCName = "prop_let_" + methodCName;
            } else if (methodSym->kind == SymbolKind::PropertySet) {
                methodCName = "prop_set_" + methodCName;
            }
            std::string procName = cProcName(methodCName, methodSym->access, cc.moduleName);
            // Fix 052: 前向声明必须与函数定义签名完全一致
            // - ByRef 参数需要添加 * (指针)
            // - Optional 参数需要添加 _has_ 标志
            {
                std::string retType = (methodSym->kind == SymbolKind::Function || methodSym->kind == SymbolKind::PropertyGet)
                    ? mapType(methodSym->type) : "void";
                // Fix 184: 返回工程类实例时真实 C 类型是 vb6_cls_<类>*, 须与函数定义一致.
                // 否则 extern 与定义 ABI 不符 (结构体返回走隐藏指针参数, 指针返回走 EAX),
                // x86 下返回值与实参全部错位 → 0xC0000005.
                if (retType == "vb6_VARIANT" && !methodSym->variableTypeName.empty()) {
                    auto* rcInstSym = lookupModuleDotted(methodSym->variableTypeName);
                    if (rcInstSym && rcInstSym->kind == SymbolKind::Class && !rcInstSym->isInterface) {
                        // ai/022 B09-0: 带 `struct` 关键字。本文件的 typedef 是**按类分块**发的
                        // (每个类一行 "typedef struct vb6_cls_X vb6_cls_X;"), 而这里引用的可能是
                        // **另一个还没轮到**的类 (派生类块排在基类块之前就是这样): 裸 `vb6_cls_X*`
                        // 在该行还是未声明名 → error C2143。`struct vb6_cls_X*` 自己就把标签带上了,
                        // 与上面的 _New/_Destroy extern 同一口径, 不再有顺序依赖。
                        retType = "struct vb6_cls_" + cIdent(rcInstSym->name) + "*";
                    }
                }
                std::string params = "struct " + clsStruct + "*";
                for (auto& p : methodSym->params) {
                    std::string pType = mapType(p.type);
                    // <vbeclipse>: 上面 Fix 184 只把**返回类型**的工程类补对了, 形参是
                    // 同一处毛病: `Sub Wire(ByVal h As DemoShape)` 的 p.type 是 Variant
                    // (类名在符号表里就丢了), mapType 只能给 vb6_VARIANT, 于是本文件的
                    // extern 与类模块发的定义 (vb6_cls_DemoShape*) ABI 不符; 调用点照着
                    // extern 那份把类指针装箱成 VARIANT → C2440 (实测 cc_demo/itf_via/
                    // cls_inh/modulemethod 一系)。判定和调用点共用 cParamClassPtrType
                    // 这一处权威（它内部就是 mapTypeRef，不是第二份表），并同 Fix 184
                    // 带 `struct` 关键字免声明顺序依赖 —— 原型与实参打包不会再次分叉。
                    const std::string pClsPtr = cParamClassPtrType(p);
                    if (pType == "vb6_VARIANT" && !pClsPtr.empty()) {
                        pType = "struct " + pClsPtr;
                    }
                    if (!p.isByVal) {
                        // ByRef → C指针 (与 cgen_decl.cpp 函数定义一致)
                        pType += "*";
                    }
                    params += ", " + pType;
                }
                // Optional 参数的 _has_ 标志
                for (auto& p : methodSym->params) {
                    if (p.isOptional && !p.isParamArray) {
                        params += ", int _has_" + cIdent(p.name);
                    }
                }
                entry.emitLine("extern " + retType + " " + procName + "(" + params + ");");
            }
        }
    }
    // vb6_Init/vb6_Exit 前向声明 (定义在vb6rtl.c)
    entry.emitLine("extern void vb6_Init(void);");
    entry.emitLine("extern void vb6_Exit(void);");
    entry.emitBlank();

    // 1. IDispatch方法桥接函数
    for (auto& cc : coClasses) {
        std::string clsId = cIdent(cc.moduleName);
        std::string clsStruct = "vb6_cls_" + clsId;

        for (size_t mi = 0; mi < cc.publicMethodNames.size(); mi++) {
            std::string& methodName = cc.publicMethodNames[mi];
            Symbol* methodSym = cc.publicMethodSyms[mi];
            std::string prefix = cc.moduleName + "_";
            std::string bareName = methodName;
            if (Symbol::toLower(bareName).substr(0, prefix.size()) == Symbol::toLower(prefix)) {
                bareName = bareName.substr(prefix.size());
            }
            std::string invokeSuffixMD;
            if (methodSym->kind == SymbolKind::PropertyGet) invokeSuffixMD = "_get";
            else if (methodSym->kind == SymbolKind::PropertyLet) invokeSuffixMD = "_let";
            else if (methodSym->kind == SymbolKind::PropertySet) invokeSuffixMD = "_set";
            std::string invokeName = "vb6_disp_" + clsId + "_" + cIdent(bareName) + invokeSuffixMD + "_invoke";

            entry.emitLine("static void " + invokeName + "(void* instance, void** args, int32_t argc, void* result) {");
            entry.indent();
            entry.emitLine("struct " + clsStruct + "* me = (struct " + clsStruct + "*)instance;");

            // Fix 052: 构建实参列表 (不含me的类型, 只有表达式)
            // 注意: args[i] 是 VARIANT* (来自DISPPARAMS.rgvarg), 需用VARIANT字段提取值
            // Fix 093: 参数缺失守卫 —— 调用方可能少传参数 (VB6 Optional 未提供, 或
            //   脚本引擎只传前 N 个), 此时 argc < 声明参数个数, args 为 NULL (argc==0)
            //   或 args[i] 越界, 直接解引用即 0xC0000005. 实测 cVBMAN.Version
            //   (Optional HostApp As Object, ByRef) 被零参 vm.Version() 晚绑定调用时
            //   必崩于 &((VARIANT*)NULL)->lVal. 现为每个参数生成 (argc > i) 三元:
            //   已提供取 VARIANT 字段, 未提供取零值哑元 (复合字面量, 同 Fix 089h 手法).
            std::string callArgs = "me";
            for (int i = 0; i < (int)methodSym->params.size(); i++) {
                const ParameterInfo& prm = methodSym->params[i];
                Vb6Type pType = prm.type;
                const std::string idx = std::to_string(i);
                const std::string guard = "(argc > " + idx + ")";
                std::string provided;
                std::string missing;
                if (prm.isByVal) {
                    // 从VARIANT中提取ByVal值
                    // <vbeclipse>: 工程类形参的槽 —— 与上面 extern 那处同一个权威
                    // (cParamClassPtrType)。DISPPARAMS 送来的是 VT_DISPATCH, 旧代码按
                    // "Variant 槽"整份 VARIANT 结构透传, 而定义侧形参是 vb6_cls_X* ⇒
                    // C2440 (实测 com_entry.c 的 vb6_DemoHolder_Wire)。这里剥出实例指针
                    // 再交出去, 与**调用点**那条直传路径同一个口径。
                    const std::string cpCls = cParamClassPtrType(prm);
                    if (!cpCls.empty() && cpCls.compare(0, 8, "vb6_cls_") == 0) {
                        provided = "(" + cpCls + ")vb6_ComObject_GetInstance("
                                   "((VARIANT*)args[" + idx + "])->pdispVal)";
                        missing = "NULL";
                    } else if (pType == Vb6Type::Long || pType == Vb6Type::Integer ||
                        pType == Vb6Type::Byte) {
                        provided = "((VARIANT*)args[" + idx + "])->lVal";
                        missing = "0";
                    } else if (pType == Vb6Type::Boolean) {
                        provided = "(((VARIANT*)args[" + idx + "])->vt == VT_BOOL ? (((VARIANT*)args[" + idx + "])->boolVal == VARIANT_TRUE ? -1 : 0) : ((VARIANT*)args[" + idx + "])->lVal)";
                        missing = "0";
                    } else if (pType == Vb6Type::Double || pType == Vb6Type::Single) {
                        provided = "((VARIANT*)args[" + idx + "])->dblVal";
                        missing = "0.0";
                    } else if (pType == Vb6Type::String) {
                        provided = "((VARIANT*)args[" + idx + "])->bstrVal";
                        missing = "(BSTR)NULL";
                    } else if (pType == Vb6Type::Variant) {
                        // Fix 051: ByVal Variant — 传递整个 VARIANT 结构 (vb6_VARIANT 与 Windows VARIANT 布局兼容)
                        provided = "*(vb6_VARIANT*)args[" + idx + "]";
                        missing = "(vb6_VARIANT){0}";
                    } else if (pType == Vb6Type::Object) {
                        // Fix 093: 对象 ByVal — 直接传 IDispatch* 本身, 不能取 lVal.
                        // lVal 是 LONG (4 字节), x86 下与 void* 同宽恰好可用,
                        // x64 下指针被截断成低 32 位 → 无效指针.
                        provided = "((VARIANT*)args[" + idx + "])->pdispVal";
                        missing = "NULL";
                    } else {
                        provided = "((VARIANT*)args[" + idx + "])->lVal";
                        missing = "0";
                    }
                } else {
                    // ByRef: 传递VARIANT中的值指针
                    if (pType == Vb6Type::Long || pType == Vb6Type::Integer ||
                        pType == Vb6Type::Byte) {
                        provided = "&((VARIANT*)args[" + idx + "])->lVal";
                        missing = "&(int32_t){0}";
                    } else if (pType == Vb6Type::Boolean) {
                        provided = "&((VARIANT*)args[" + idx + "])->lVal";  /* ByRef Boolean: same as Long, VB6 uses int32_t */
                        missing = "&(int32_t){0}";
                    } else if (pType == Vb6Type::Double || pType == Vb6Type::Single) {
                        provided = "&((VARIANT*)args[" + idx + "])->dblVal";
                        missing = "&(double){0.0}";
                    } else if (pType == Vb6Type::Variant) {
                        // Fix 051: ByRef Variant — 传递 VARIANT 指针
                        provided = "(vb6_VARIANT*)args[" + idx + "]";
                        missing = "&(vb6_VARIANT){0}";
                    } else if (pType == Vb6Type::String) {
                        provided = "&((VARIANT*)args[" + idx + "])->bstrVal";
                        missing = "&(BSTR){NULL}";
                    } else if (pType == Vb6Type::Object) {
                        // Fix 093: 对象 ByRef — 必须传"指针字段的地址"(void**),
                        // 不能取 &lVal (LONG*, 4 字节): x86 下与 void* 同宽巧合可用,
                        // x64 下被当 void** 解引用会读 8 字节越过该字段 → 读到栈垃圾
                        // 当 IDispatch* → 0xC0000005 (实测 cVBMAN.Version 的
                        // Optional HostApp As Object 在 x64 崩、x86 侥幸通过).
                        provided = "(void*)&((VARIANT*)args[" + idx + "])->pdispVal";
                        missing = "(void*)&(void*){0}";
                    } else {
                        provided = "(void*)&((VARIANT*)args[" + idx + "])->lVal";
                        missing = "(void*)&(int32_t){0}";
                    }
                }
                callArgs += ", (" + guard + " ? (" + provided + ") : (" + missing + "))";
            }
            // Fix 052: Optional 参数的 _has_ 标志
            // Fix 093: 改按实际实参个数判定 (此前恒为 1, 少传参数时被调方误以为收到了实参)
            for (int i = 0; i < (int)methodSym->params.size(); i++) {
                const ParameterInfo& p = methodSym->params[i];
                if (p.isOptional && !p.isParamArray) {
                    callArgs += ", ((argc > " + std::to_string(i) + ") ? 1 : 0)";
                }
            }

            // Property需要加前缀: prop_get_/prop_let_/prop_set_
            std::string methodCName = methodName;
            if (methodSym->kind == SymbolKind::PropertyGet) {
                methodCName = "prop_get_" + methodCName;
            } else if (methodSym->kind == SymbolKind::PropertyLet) {
                methodCName = "prop_let_" + methodCName;
            } else if (methodSym->kind == SymbolKind::PropertySet) {
                methodCName = "prop_set_" + methodCName;
            }
            std::string procName = cProcName(methodCName, methodSym->access, cc.moduleName);

            if (methodSym->kind == SymbolKind::Function || methodSym->kind == SymbolKind::PropertyGet) {
                std::string retType = mapType(methodSym->type);
                // Fix 184: 工程类实例返回值先包装成 IDispatch 再放入 VARIANT
                std::string retClsInst;
                if (retType == "vb6_VARIANT" && !methodSym->variableTypeName.empty()) {
                    auto* rcInstSym2 = lookupModuleDotted(methodSym->variableTypeName);
                    if (rcInstSym2 && rcInstSym2->kind == SymbolKind::Class && !rcInstSym2->isInterface) {
                        retClsInst = cIdent(rcInstSym2->name);
                    }
                }
                entry.emitLine("if (result) {");
                entry.indent();
                if (!retClsInst.empty()) {
                    entry.emitLine("vb6_cls_" + retClsInst + "* _p = " + procName + "(" + callArgs + ");");
                    entry.emitLine("vb6_VARIANT _r = vb6_VariantObject(vb6_ComObject_FromInstance("
                                   "vb6_FindCoClassDesc(\"" + retClsInst + "\"), (void*)_p));");
                } else {
                    entry.emitLine(retType + " _r = " + procName + "(" + callArgs + ");");
                }
                if (!retClsInst.empty()) {
                    entry.emitLine("memcpy(result, &_r, sizeof(vb6_VARIANT));");
                } else if (methodSym->type == Vb6Type::Long || methodSym->type == Vb6Type::Integer ||
                    methodSym->type == Vb6Type::Byte) {
                    entry.emitLine("((VARIANT*)result)->vt = VT_I4; ((VARIANT*)result)->lVal = (int32_t)_r;");
                } else if (methodSym->type == Vb6Type::Boolean) {
                    entry.emitLine("((VARIANT*)result)->vt = VT_BOOL; ((VARIANT*)result)->boolVal = _r ? VARIANT_TRUE : VARIANT_FALSE;");
                } else if (methodSym->type == Vb6Type::Double || methodSym->type == Vb6Type::Single) {
                    entry.emitLine("((VARIANT*)result)->vt = VT_R8; ((VARIANT*)result)->dblVal = (double)_r;");
                } else if (methodSym->type == Vb6Type::String) {
                    entry.emitLine("((VARIANT*)result)->vt = VT_BSTR; ((VARIANT*)result)->bstrVal = _r;");
                } else if (retType == "vb6_VARIANT") {
                    // Fix 052: 所有映射为 vb6_VARIANT 的返回类型 (Variant/Decimal/UDT等) 都用 memcpy
                    // vb6_VARIANT 结构体不能 cast 为 intptr_t,
                    // 用 memcpy 复制到 Windows VARIANT (布局兼容: vt + padding + union).
                    entry.emitLine("memcpy(result, &_r, sizeof(vb6_VARIANT));");
                } else {
                    entry.emitLine("((VARIANT*)result)->vt = VT_I4; ((VARIANT*)result)->lVal = (int32_t)(intptr_t)_r;");
                }
                entry.dedent();
                entry.emitLine("}");
            } else {
                entry.emitLine(procName + "(" + callArgs + ");");
            }
            entry.dedent();
            entry.emitLine("}");
            entry.emitBlank();
        }
    }

    // Fix 099: Public 字段的 IDispatch 桥接 (Property Get/Let)
    // getter/setter 实体 (vb6_cls_<cls>_field_get_/_let_<fld>) 定义在**类模块自身**
    // 的 .c 里 (cgen_com.cpp emitClassFieldAccessors) —— dll_entry.c 只有结构体的
    // 前向声明 (不完整类型), 无法直接 `me->Fld`, 只能 extern 调用.
    // Get/Let 共享同一 dispid (COM 约定, 与 TypeLib 里的同名 Property 一致).
    {
        bool anyFieldBridges = false;
        for (auto& cc : coClasses) {
            if (!cc.publicFields.empty()) { anyFieldBridges = true; break; }
        }
        if (anyFieldBridges) {
            for (auto& cc : coClasses) {
                if (cc.publicFields.empty()) continue;
                std::string fClsId = cIdent(cc.moduleName);
                std::string fClsStruct = "vb6_cls_" + fClsId;
                for (auto& fname : cc.publicFields) {
                    std::string fId = cIdent(fname);
                    entry.emitLine("extern vb6_VARIANT " + fClsStruct + "_field_get_" + fId
                                   + "(struct " + fClsStruct + "*);");
                    entry.emitLine("extern void " + fClsStruct + "_field_let_" + fId
                                   + "(struct " + fClsStruct + "*, vb6_VARIANT);");
                }
                entry.emitBlank();
                for (auto& fname : cc.publicFields) {
                    std::string fId = cIdent(fname);
                    entry.emitLine("static void vb6_disp_" + fClsId + "_" + fId
                                   + "_get_invoke(void* instance, void** args, int32_t argc, void* result) {");
                    entry.indent();
                    entry.emitLine("(void)args; (void)argc;");
                    entry.emitLine("if (!result) return;");
                    entry.emitLine("{");
                    entry.indent();
                    entry.emitLine("vb6_VARIANT _r = " + fClsStruct + "_field_get_" + fId
                                   + "((struct " + fClsStruct + "*)instance);");
                    entry.emitLine("memcpy(result, &_r, sizeof(vb6_VARIANT));");
                    entry.dedent();
                    entry.emitLine("}");
                    entry.dedent();
                    entry.emitLine("}");
                    entry.emitBlank();

                    entry.emitLine("static void vb6_disp_" + fClsId + "_" + fId
                                   + "_let_invoke(void* instance, void** args, int32_t argc, void* result) {");
                    entry.indent();
                    entry.emitLine("(void)result;");
                    entry.emitLine("vb6_VARIANT _v = (argc > 0 && args) ? *(vb6_VARIANT*)args[0] : vb6_VariantEmpty();");
                    entry.emitLine(fClsStruct + "_field_let_" + fId
                                   + "((struct " + fClsStruct + "*)instance, _v);");
                    entry.dedent();
                    entry.emitLine("}");
                    entry.emitBlank();
                }
            }
        }
    }
