// driver_codegen_dup_module_vars.inc - runCodeGeneration 函数体片段：Fix 093b 跨模块同名模块级变量
// 由 src/driver/driver.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 342~370 行，逐行未改。

    // P6.4+: 预扫描 — 构建 接口 → 唯一实现类名 映射.
    // 遍历所有类模块的 Class 符号的 implementsNames, 反向记录每个接口被哪些
    // 类实现. 恰好一个实现类 → 记录其规范名; 多于一个 → 置空 (cgen 依赖唯一
    // 实现类的场景报错而非静默误编). 供 Set 接口引用赋值类型转换与 ByVal
    // 接口形参回调重建 (SUBCLASSPROC ABI) 使用.
    std::map<std::string, std::string> ifaceImplementersMap;
    // Fix <vbeclipse>: 接口 → **全部**实现类名 (同名去重, 保持首见序)。
    // 与上面那张"唯一实现类"表并存, 两者用途不同、不能互相替代:
    //   · ifaceImplementersMap 用于**推断类型** (Set 接口引用赋值 / ByVal 接口形参
    //     回调重建)。多实现类时置空, 让 cgen 报错而不是猜。
    //   · 本表用于**接口成员桩的运行期分派**: 桩必须挨个试"me 是不是这些实现类
    //     之一的实例", 命中就经它的槽表转调。IScheme 在 VbEclipse 里有 7 个实现类
    //     (SchemeWinXP / SchemeEclipse3 / SchemeLonghorn / SchemeOffice2003 /
    //     SchemeVS2005 / SchemeCustom / ...), 上面那张表对它是空的 —— 桩若因此
    //     退成"恒返回零值", `m_Scheme.BackColor` 之类一律读到 0 (纯黑), 左侧停靠
    //     面板整片黑底 (实测 24 次 SetBackColor 里 17 次 color=000000, 而
    //     SchemeWinXP 定义的配色里根本没有黑色)。
    std::map<std::string, std::vector<std::string>> ifaceAllImplementersMap;
    // 账 #305 新账 #464: 接口名的**书写**拼写 (小写键 → 首见原样)。legacy 形状表要用它拼
    // C 类型名 (`vb6_vtbl_<I>`), 而消费侧 `mapTypeRef` 用的是符号表里的规范名 —— 宿主在
    // 工程内时下面直接取宿主模块名, 只有工程外宿主才退回收书写形式。
    std::map<std::string, std::string> ifaceWrittenNames;
    for (size_t i = 0; i < modules_.size(); i++) {
        if (!modules_[i]->isClassModule) continue;
        Symbol* clsSym = analyzers_[i]->symbolTable().lookupModule(modules_[i]->moduleName);
        if (!clsSym || clsSym->kind != SymbolKind::Class) continue;
        if (clsSym->implementsNames.empty()) continue;
        for (const auto& iface : clsSym->implementsNames) {
            std::string ifLower = iface;
            std::transform(ifLower.begin(), ifLower.end(), ifLower.begin(), ::tolower);
            auto it = ifaceImplementersMap.find(ifLower);
            if (it == ifaceImplementersMap.end()) {
                ifaceImplementersMap[ifLower] = clsSym->name;
            } else if (it->second != clsSym->name) {
                it->second.clear();  // 多实现类 → 不可推断
            }
            // 同一实现类对同一接口只登记一次 (implementsNames 自身已去重, 但两个
            // 大小写不同的接口名会归一到同一个 key, 这里再挡一次)。
            std::vector<std::string>& all = ifaceAllImplementersMap[ifLower];
            if (std::find(all.begin(), all.end(), clsSym->name) == all.end()) {
                all.push_back(clsSym->name);
            }
            if (ifaceWrittenNames.count(ifLower) == 0) ifaceWrittenNames[ifLower] = iface;
        }
    }

    // 账 #305 新账 #464: legacy 接口（.cls 宿主、**没**进 tB 登记表）那张表的**形状归接口自己定**。
    // 从前发码层 `emitInterfaceVtable` 在每枚实现类里按「自家成员声明顺序 + 自家 `me` 类型」harvest，
    // 于是同一枚接口被 N 枚类实现就发 N 份同名 `vb6_vtbl_<I>` / `vb6_iface_<I>` —— 真语料
    // `图像/光线追踪/Ray.vbp`（四枚类 Implements `IRayTraceable`）实测 cl **54 发 C2011**，
    // 且四份表里**槽序有两套**（`_for_Sphere` 与另三份不同）⇒ 「留第一份、丢其余」不是修法，
    // 那是把三通槽按两套含义发出去。这里把形状建成一张工程级只读表：
    //   键 = 小写接口名；值 = **接口模块自己的**可实现成员声明，顺序 = 接口里的声明顺序。
    // 存 AST 指针而不存渲染文本：C 类型仍由各模块的 cgen 现拼（与 stage 2.7 `IfaceRegistry` 同一口径，
    // 那里也只存 `Decl*`）。进了登记表的接口由 `emitIfaceContractTypedefs` 那条 canonical 路接管，这里不收。
    //
    // 表的来源分两段，顺序固定：
    //   1) 接口宿主模块**自己**的可实现成员（声明序 = 槽序）—— 工程内有 `.cls` 宿主时才有；
    //   2) 各实现类写的 `<接口名>_成员` 里，第 1 段没有的槽**补到表尾**（按实现类首见序、类内声明序）。
    // 第 2 段两头都有用：宿主在工程外（引用库里的接口）时它是唯一的形状来源；宿主在工程内时
    // 它保证「今天某枚实现类的自家表里有的槽，改后仍在」——否则调用点会从「有字段」退成 C2039。
    std::map<std::string, LegacyIfaceShape> legacyIfaceShapeMap;
    {
        std::map<std::string, Module*> modByLowerName;
        for (auto& mod : modules_) {
            std::string k = mod->moduleName;
            std::transform(k.begin(), k.end(), k.begin(), ::tolower);
            modByLowerName.emplace(k, mod.get());
        }
        // 声明成员名 → 槽键: 摘掉 VB6 实现成员的 `<接口名>_` 前缀再小写。接口宿主写 `Bar`、
        // 实现类写 `IFoo_Bar`, 两种拼法必须归到同一个键, 否则第 2 段会把第 1 段已有的槽再补一份。
        auto ifaceSlotKey = [](const std::string& ifLower, const std::string& written) {
            std::string lower = written;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            const std::string prefix = ifLower + "_";
            if (lower.size() > prefix.size() && lower.compare(0, prefix.size(), prefix) == 0) {
                return lower.substr(prefix.size());
            }
            return lower;
        };
        // 可实现成员 = Sub/Function/Property 三类过程声明 (Type/Enum/Declare 不算)。
        // 名字由调用方取, 这里只认 kind。
        auto isMemberProc = [](const std::unique_ptr<Decl>& d) {
            return d && (d->kind == ASTNodeKind::SubDecl ||
                         d->kind == ASTNodeKind::FunctionDecl ||
                         d->kind == ASTNodeKind::PropertyDecl);
        };
        for (const auto& kv : ifaceAllImplementersMap) {
            const std::string& ifLower = kv.first;
            if (ifaces_.count(ifLower) != 0) continue;  // 新式接口: 归 canonical 那条路
            LegacyIfaceShape& shape = legacyIfaceShapeMap[ifLower];
            std::vector<const Decl*>& slots = shape.slots;
            std::set<std::string> seen;
            // 同一槽键只认第一份 —— 但键现在**带档** (D2: Sub/Function 裸名, 属性按
            // get_/put_/putref_ 分向; 账 #305 新账 #489)。从前这把键是裸成员名, 于是同名
            // `Property Get Label` 与 `Property Let Label` 在表里共用一席、由**声明序**决定
            // 是哪一档 (实测 Let 先写就铸出 setter 的形状 = 本账的病灶)。字段名那边只给
            // 第二档改口 (`Label` / `Label_put` / `Label_putref`) => 两份同名结构体成员
            // = C2086 这条仍然躲得开, 而 Get 那一席与所有现有产物逐字节不动。
            auto addSlot = [&](const std::unique_ptr<Decl>& d, bool wantPrefix) {
                if (!isMemberProc(d)) return;
                std::string written;
                if (d->kind == ASTNodeKind::SubDecl) {
                    written = static_cast<SubDecl&>(*d).name;
                } else if (d->kind == ASTNodeKind::FunctionDecl) {
                    written = static_cast<FunctionDecl&>(*d).name;
                } else {
                    written = static_cast<PropertyDecl&>(*d).name;
                }
                const std::string prefix = ifLower + "_";
                std::string lower = written;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                const bool hasPrefix =
                    lower.size() > prefix.size() && lower.compare(0, prefix.size(), prefix) == 0;
                if (wantPrefix && !hasPrefix) return;  // 实现类里不是本接口那一份成员
                if (!seen.insert(ifaceSlotKeyNamed(ifaceSlotKey(ifLower, written),
                                                   d->kind, ifaceDeclProcKind(*d))).second) return;
                slots.push_back(d.get());
            };
            auto itMod = modByLowerName.find(ifLower);
            if (itMod != modByLowerName.end() && itMod->second->isClassModule) {
                for (const auto& d : itMod->second->declarations) addSlot(d, false);
            }
            for (const std::string& implName : kv.second) {
                std::string k = implName;
                std::transform(k.begin(), k.end(), k.begin(), ::tolower);
                auto itImpl = modByLowerName.find(k);
                if (itImpl == modByLowerName.end()) continue;
                for (const auto& d : itImpl->second->declarations) addSlot(d, true);
            }
            // 类型名的规范拼写: 宿主在工程内 → 宿主模块名 (与它的 Class 符号同名, 消费侧
            // `mapTypeRef` 发的 `vb6_iface_<I>` 用的正是这个); 宿主在工程外 → 取首个
            // `Implements <I>` 的书写形式 (去掉 `Project.` 限定段)。
            if (itMod != modByLowerName.end() && itMod->second->isClassModule) {
                shape.writtenName = itMod->second->moduleName;
            } else {
                auto itW = ifaceWrittenNames.find(ifLower);
                if (itW != ifaceWrittenNames.end()) {
                    const std::string& w = itW->second;
                    const size_t dot = w.rfind('.');
                    shape.writtenName = dot == std::string::npos ? w : w.substr(dot + 1);
                }
            }
            if (slots.empty() || shape.writtenName.empty()) {
                legacyIfaceShapeMap.erase(ifLower);  // 既没形状也没名字: 这张表里不留空壳
            }
        }
    }

    // P6.4+: 预扫描 — 类模块默认实例 (Attribute VB_PredeclaredId = True).
    // VB6 会为这类类模块生成隐藏的全局默认实例, frm 里裸写 cTT.CreateToolTip(...)
    // 就是调用它. 记录 小写类名 → 类规范名, 供 cgen 解析裸类名成员访问并生成
    // vb6_cls_X_Default() 惰性单例访问器.
    std::map<std::string, std::string> defaultInstanceClasses;
    for (size_t i = 0; i < modules_.size(); i++) {
        if (!modules_[i]->isClassModule) continue;
        for (const auto& attr : modules_[i]->attributes) {
            std::string aName = attr->attrName;
            std::transform(aName.begin(), aName.end(), aName.begin(), ::tolower);
            if (aName != "vb_predeclaredid") continue;
            bool isTrue = false;
            if (attr->value && attr->value->kind == ASTNodeKind::LiteralExpr) {
                auto& lit = static_cast<const LiteralExpr&>(*attr->value);
                if (lit.literalKind == LiteralKind::Boolean) isTrue = lit.boolValue;
                else if (lit.literalKind == LiteralKind::Integer) isTrue = (lit.intValue != 0);
            }
            if (isTrue) {
                std::string lower = modules_[i]->moduleName;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                defaultInstanceClasses[lower] = modules_[i]->moduleName;
            }
            break;
        }
    }

    // Fix 093b: 预扫描 — 找出"跨模块同名模块级变量".
    // VB6 允许不同模块各自声明同名模块级变量 (模块内引用绑定自身副本, 跨模块引用
    // 必须写 "模块名.变量名"), 但 C 语言下两个模块会生成同名全局符号 → LNK2005.
    // 本工程实例:
    //   Tools.bas:  Public Delay As New cDelay   /  mDelay.bas: Public Delay As New cDelay
    //   FToastDrawer.frm: Dim Sad/m_Theme/m_State/PosVal/TopVal/m_ParentToast/m_TagName
    //   FToastCenter.frm: 同名 Dim 一组
    // 命中的名字在生成端强制 static 内部链接 (见 CCodeGen::isPublicModuleDecl):
    // 各模块保留自己的副本, 语义与 VB6 一致, 且不再产生重复符号.
    std::unordered_map<std::string, std::unordered_set<std::string>> moduleVarOwners;
    for (size_t i = 0; i < modules_.size(); i++) {
        // 类模块的成员变量是结构体字段 (me->x), 不产生全局符号; 不能参与撞名判定
        // (否则 cVBMAN.cls 的 Public ToolsUtf8 会让 Tools.bas 的 Public ToolsUtf8
        //  被误判为撞名 → 强制 static → 跨模块引用 cHttpServer.c 等 C2065).
        if (modules_[i]->isClassModule) continue;
        std::string modLower = modules_[i]->moduleName;
        std::transform(modLower.begin(), modLower.end(), modLower.begin(), ::tolower);
        for (const auto& decl : modules_[i]->declarations) {
            if (decl->kind != ASTNodeKind::VariableDecl) continue;
            auto& vd = static_cast<const VariableDecl&>(*decl);
            std::string lower = vd.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            moduleVarOwners[lower].insert(modLower);
        }
    }
    std::unordered_set<std::string> dupModuleVarNames;
    for (const auto& [varName, owners] : moduleVarOwners) {
        if (owners.size() > 1) dupModuleVarNames.insert(varName);
    }
