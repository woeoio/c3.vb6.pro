// cgen_api.inc - CCodeGen 的 public 接口段（类体片段）
// 由 src/backend/cgen.hpp 拆出（2026-09-17），在类体内被 #include，非独立头文件。
// 内容 = 原 cgen.hpp 第 77~275 行，逐行未改；`public:` 由伞头给出。

    CCodeGen(Diagnostics& diag, const SymbolTable& symTab,
             const TypeSystem& typeSys,
             const std::unordered_map<std::string, std::set<std::string>>* classVoidFieldMap = nullptr,
              const std::unordered_map<std::string, std::unordered_map<std::string, std::string>>* classTypedFieldMap = nullptr,
              const std::unordered_set<std::string>* variantReturnFuncs = nullptr,
              bool verbose = false);

    // 主入口: 生成C代码，返回是否成功
    // externalModules: 当前模块引用的外部模块基名列表 (用于生成 #include)
    // isDll: P6.6 ActiveX DLL模式, 生成COM服务端代码
    // dllProgId: P6.6 DLL的ProgID前缀
    // frmDesc: P7 窗体描述 (仅.frm有效, nullptr=非窗体模块)
    bool generate(Module& module, const std::string& baseName,
                  const std::unordered_set<std::string>& externalModules = {},
                  bool isDll = false, const std::string& dllProgId = "",
                  const FrmFormDesc* frmDesc = nullptr);

    // opt4: 裁剪标准模块.c中未实际引用的跨模块#include (降低预处理量, 加速cl编译)
    void setTrimIncludes(bool b) { trimIncludes_ = b; }
    // opt4: 设置"当前模块实际引用的外部模块"集合(小写), 由driver通过AST扫描提供
    void setTrimModules(const std::unordered_set<std::string>& m) { trimModules_ = m; }

    // Fix 142: 设置 VBP 的 Startup= 启动对象 ("Sub Main" 或窗体模块名).
    // 多模块工程下仅该模块生成进程入口点 (WinMain/main), 其余模块不生成,
    // 否则每个窗体模块各出一个 WinMain → LNK2005.
    void setStartupObject(const std::string& s) { startupObject_ = s; }

    // ai/vb-asm-extension-spec: 目标架构 ("x64"/"x86") — Asm 块 v1 仅 x64 (x86 报 3030)
    void setTargetArch(const std::string& a) { targetArch_ = a; }
    // ai/vb-asm-extension-spec: 已收集的 Asm 过程 (函数体=单个 Asm 块),
    // driver 侧据此生成 .asm 并调 ml64 → obj → 链接 (见 driver_link.cpp)
    const std::vector<AsmProcInfo>& asmProcs() const { return asmProcs_; }

    // Fix 086: 设置工程内全部窗体模块名集合(小写), 由driver预扫描提供.
    // 跨模块窗体默认实例引用 (cLogs 里 FLogs.Visible / Unload FLogs) 需要知道
    // FLogs 是窗体, 才能生成 vb6_form_hwnd_FLogs() 访问器调用而非裸标识符.
    void setFormModuleNames(const std::unordered_set<std::string>& names) { knownFormModuleNames_ = names; }
    // Fix 086: 各模块 Public 常量表 (小写模块名 → 小写常量名 → 整数值).
    // 本地同名过程符号会阻止跨模块常量注入 (如 cSerialPort.Sub SetDTR 与
    // modSerialPortAPI.Const SETDTR), 此表兜底内联数值.
    void setModulePublicConsts(const std::unordered_map<std::string,
        std::unordered_map<std::string, long long>>* consts) { modulePublicConsts_ = consts; }

    // Fix 093b: 跨模块同名模块级变量集合 (小写名), 由 driver 预扫描提供.
    // 两个不同模块各自声明同名模块级变量时 (如 Tools.bas 与 mDelay.bas 都有
    // Public Delay As New cDelay; FToastDrawer/FToastCenter 都有 Dim m_Theme /
    // PosVal / m_ParentToast ...), C 语言下会生成同名全局符号 → LNK2005.
    // 命中的名字改为 static 内部链接: VB6 语义下模块内引用绑定本模块副本,
    // 跨模块引用必须写 "模块名.变量名" (不会与裸名撞), 故 static 不改变行为.
    void setStaticModuleVarNames(const std::unordered_set<std::string>* names) {
        staticModuleVarNames_ = names;
    }
    // 名字(大小写不敏感)是否命中"跨模块同名模块级变量"
    bool isStaticForcedVarName(const std::string& vb6Name) const {
        if (!staticModuleVarNames_) return false;
        std::string lower;
        lower.reserve(vb6Name.size());
        for (char ch : vb6Name) {
            lower += (ch >= 'A' && ch <= 'Z') ? (char)(ch - 'A' + 'a') : ch;
        }
        return staticModuleVarNames_->count(lower) != 0;
    }
    // 模块级变量: Public 且撞名 → 强制 static (内部链接)
    bool isPublicModuleDecl(const VariableDecl& n) const {
        if (n.access != AccessLevel::Public) return false;
        return !isStaticForcedVarName(n.name);
    }
    // 模块级常量: 生成 #define 宏, 不产生符号, 不受撞名影响
    bool isPublicModuleDecl(const ConstDecl& n) const {
        return n.access == AccessLevel::Public;
    }

    // P6.6: 单独生成ActiveX DLL入口文件 (dll_entry.c)
    // 当DLL工程只有类模块(无标准模块)时, 由Driver调用此方法生成DLL导出代码
    // progId: DLL的ProgID前缀
    // ExeComBridge 01: includeDllExports=false 时只生成 coclass 表 (EXE 工程的
    //   com_entry.c), 跳过 DllGetClassObject/DllMain 等 DLL 导出段.
    // 返回: 生成的dll_entry.c文件内容
    std::string generateDllEntry(const std::string& progId,
                                  const std::vector<SymbolTable*>& allSymTabs = {},
                                  bool includeDllExports = true);

    // 获取生成的代码
    const std::string& headerCode() const { return header_; }
    const std::string& sourceCode() const { return source_; }

    // --- 声明 ---
    void visit(SubDecl& node) override;
    void visit(FunctionDecl& node) override;
    void visit(PropertyDecl& node) override;
    void visit(TypeDecl& node) override;
    void visit(TypeMember& node) override;
    void visit(EnumDecl& node) override;
    void visit(EnumMember& node) override;
    void visit(DeclareDecl& node) override;
    void visit(EventDecl& node) override;
    void visit(ConstDecl& node) override;
    void visit(VariableDecl& node) override;
    void visit(ParameterDecl& node) override;

    // --- 语句 ---
    void visit(Block& node) override;
    void visit(AssignmentStmt& node) override;
    void visit(SetStmt& node) override;
    void visit(LetStmt& node) override;
    void visit(IfStmt& node) override;
    void visit(ElseIfClause& node) override;
    void visit(ForStmt& node) override;
    void visit(ForEachStmt& node) override;
    void visit(DoLoopStmt& node) override;
    void visit(WhileWendStmt& node) override;
    void visit(SelectCaseStmt& node) override;
    void visit(CaseClause& node) override;
    void visit(WithStmt& node) override;
    void visit(GoToStmt& node) override;
    void visit(GoSubStmt& node) override;
    void visit(OnGoSubStmt& node) override;  // P17.4
    void visit(OnGoToStmt& node) override;  // P18-A
    void visit(MidStmt& node) override;  // P18-A
    void visit(OnErrorStmt& node) override;
    void visit(ResumeStmt& node) override;
    void visit(ErrorStmt& node) override;
    void visit(ExitStmt& node) override;
    void visit(CallStmt& node) override;
    void visit(ReDimStmt& node) override;
    void visit(EraseStmt& node) override;
    // Fix 084y-5: ReDim/Erase 目标的 C 标识符解析 — 处理成员访问形态
    // (.Field With块成员 → _vb6_with_N->Field; obj.Field UDT成员 →
    // (*obj).Field 或 obj.Field), 避免 cIdent 把 '.' 替换成 '_' 生成
    // 未声明的单标识符 (uOutput.Buffer → uOutput_Buffer → C2065)
    std::string resolveArrayTargetIdent(const std::string& varName);
    // Fix 180: ReDim/Erase 的"owner 不带下标的点号链"目标 (h1.Items) 的元素 UDT C 类型。
    // 入参是小写后的 node.varName; 解析不到返回空串, 由调用方保持原有 Variant 回落。
    std::string resolveUdtMemberArrayCType(const std::string& lowerVar) const;
    // Fix 084aa: 判断标识符是否引用 String 类型常量 (#define 宏, 不可取址)
    bool isStringConstIdent(const std::string& name) const;
    // Fix 084aa: 查找常量符号 (Constant kind), 无则返回 nullptr
    Symbol* lookupConstSym(const std::string& name) const;
    // Fix 084aa: 判断标识符是否引用任何类型的常量 (#define 宏)
    bool isConstIdent(const std::string& name) const;
    // Fix 084aa: 获取常量标识符的 VB6 类型
    Vb6Type constIdentType(const std::string& name) const;
    // Fix 084aa: 常量宏作为 ByRef 实参 → 生成可寻址复合字面量包装
    std::string wrapConstArgForByRef(const std::string& argVal, Vb6Type paramVb6Type) const;
    // Fix 189: 当前函数的 ByRef 具体类型形参作为 ByRef Variant 实参 → 字段式 VARIANT 包装
    std::string wrapByRefVariantParamArg(const std::string& paramName, Vb6Type curParamType) const;
    void visit(OpenStmt& node) override;
    void visit(CloseStmt& node) override;
    void visit(PrintStmt& node) override;
    void visit(WriteStmt& node) override;
    void visit(LineInputStmt& node) override;
    void visit(InputStmt& node) override;
    void visit(GetStmt& node) override;
    void visit(PutStmt& node) override;
    void visit(SeekStmt& node) override;
    void visit(LockStmt& node) override;
    void visit(UnlockStmt& node) override;
    void visit(ResetStmt& node) override;
    void visit(WidthStmt& node) override;
    void visit(KillStmt& node) override;
    void visit(NameStmt& node) override;
    void visit(MkDirStmt& node) override;
    void visit(RmDirStmt& node) override;
    void visit(ChDirStmt& node) override;
    void visit(ChDriveStmt& node) override;
    void visit(FileCopyStmt& node) override;
    void visit(LabelStmt& node) override;
    void visit(OptionStmt& node) override;
    void visit(LocalDeclStmt& node) override;
    // Fix 086: LocalDeclStmt 的实际声明发射代码 (visit 判重后调用; 提升阶段也直接调用)
    void emitLocalDeclCode(LocalDeclStmt& node);
    void visit(RaiseEventStmt& node) override;
    void visit(BeepStmt& node) override;
    void visit(DoEventsStmt& node) override;

    // --- 表达式 (返回C表达式字符串) ---
    void visit(BinaryExpr& node) override;
    void visit(UnaryExpr& node) override;
    void visit(LiteralExpr& node) override;
    void visit(IdentifierExpr& node) override;
    void visit(MemberAccessExpr& node) override;
    void visit(DictionaryAccessExpr& node) override;
    void visit(IndexOrCallExpr& node) override;
    void visit(NewExpr& node) override;
    void visit(TypeOfExpr& node) override;
    void visit(AddressOfExpr& node) override;
    void visit(MeExpr& node) override;
    void visit(WithMemberExpr& node) override;

    // --- 类型引用 ---
    void visit(SimpleTypeRef& node) override;
    void visit(ArrayTypeRef& node) override;
    void visit(FixedStringTypeRef& node) override;

    // --- 模块 ---
    void visit(Module& node) override;

    // 限定类型名符号查找: "Scripting.Dictionary" → 先查全名，失败则查短名 "Dictionary"
    Symbol* lookupModuleDotted(const std::string& name) const {
        auto* sym = symTab_.lookupModule(name);
        if (!sym) {
            size_t dot = name.find('.');
            if (dot != std::string::npos) {
                sym = symTab_.lookupModule(name.substr(dot + 1));
            }
        }
        return sym;
    }
    Symbol* lookupDotted(const std::string& name) const {
        auto* sym = symTab_.lookup(name);
        if (!sym) {
            size_t dot = name.find('.');
            if (dot != std::string::npos) {
                sym = symTab_.lookup(name.substr(dot + 1));
            }
        }
        return sym;
    }

    // P6.4+: 接口 → 唯一实现类名 映射 (小写接口名 → 实现类规范名), 由 driver 预扫描提供.
    void setInterfaceImplementers(const std::map<std::string, std::string>* m) {
        interfaceImplementers_ = m;
    }
    // Interface 契约 (tB, ai/022 B04): stage 2.7 建好的工程级只读登记表, 发码期取槽表.
    void setInterfaceRegistry(const IfaceRegistry* reg) { ivreg_ = reg; }
    // 委托式实现 (tB, ai/022 B10): stage 2.7 Pass D 的裁决表, 发码期据此把无自家实现的
    // 槽转调持有对象的同名槽.
    void setViaRegistry(const ViaRegistry* reg) { ivviareg_ = reg; }
    // 类继承 (tB, ai/022 B07b): stage 2.8 链登记表, 发码期取继承字段与转发桩清单.
    void setClassChainRegistry(const ClassChainRegistry* reg) { clsreg_ = reg; }
    // 工程级类名表 (小写类名 → 原大小写类名), driver 侧 runCodeGeneration 建.
    //
    // 为什么不能只靠本模块符号表判"X 是工程类": 每模块一张 SymbolTable, 同名成员
    // (字段自身 Variable / 同名 Sub / 同名 PropertyGet) 会遮蔽跨模块类名. vbeclipse:
    //   ucSplitBar.ctl `Private WithEvents SplitBar As SplitBar` → 类型名被字段自身占位;
    //   ucCaption.ctl `Private WithEvents m_PopupMenu As PopupMenu` → 被同名 Sub 占位.
    // (也不能用 ClassChainRegistry: 工程无 Inherits 时 driver_classchain.cpp:355 的
    //  D24 护栏让整表早退为空.)
    void setProjectClassNames(const std::unordered_map<std::string, std::string>* m) {
        projClassNames_ = m;
    }
    // 该名字是不是工程类 (空串 = 不是)
    std::string projectClassNameOf(const std::string& name) const {
        if (!projClassNames_ || name.empty()) return std::string();
        auto it = projClassNames_->find(Symbol::toLower(name));
        return it == projClassNames_->end() ? std::string() : it->second;
    }
    // CoClass 身份出口 (tB, ai/022 B13b): 只喂 generateDllEntry。EXE 的 com_entry.c 本批
    // **不接线** —— 它今天整个跑在 legacy mint 上，接上就是把改码前编得过的存量工程换掉。
    void setCoClassIdentityRegistry(const std::unordered_map<std::string, CoClassIdentity>* reg) {
        ccreg_ = reg;
    }
    // 接口 IID 唯一出口 (ai/022 B13c): 同上，只读注入
    void setInterfaceIidMap(const std::unordered_map<std::string, std::string>* map) {
        iidreg_ = map;
    }
    // 查询接口唯一实现类名 (空串 = 无映射/多实现类, 不可用)
    std::string singleImplementationClass(const std::string& ifaceName) const {
        if (!interfaceImplementers_) return "";
        std::string lower = ifaceName;
        for (auto& ch : lower) {
            if (ch >= 'A' && ch <= 'Z') ch = (char)(ch - 'A' + 'a');
        }
        auto it = interfaceImplementers_->find(lower);
        if (it == interfaceImplementers_->end()) return "";
        return it->second;
    }
    // Fix <vbeclipse>: 注入"接口 → 全部实现类"表 (小写接口名 → 规范类名列表).
    void setInterfaceAllImplementers(const std::map<std::string, std::vector<std::string>>* m) {
        interfaceAllImplementers_ = m;
    }
    // 查询接口的全部实现类 (保持 driver 预扫描的首见序)。空 vector = 无实现类
    // 或表未注入 => 接口桩退回零值。
    //
    // 为什么不能用上面的 singleImplementationClass: 它在**多实现类**接口上刻意
    // 返回空串 (让依赖唯一实现类的场景报错而非猜), 而 VbEclipse 的 IScheme 有 7 个
    // 实现类。接口桩需要的是"挨个试", 不是"唯一那个" —— 否则桩恒返回零值,
    // `m_Scheme.BackColor` 一律读到 0 (纯黑), 停靠面板整片黑底。
    std::vector<std::string> allImplementationClasses(const std::string& ifaceName) const {
        std::vector<std::string> out;
        if (!interfaceAllImplementers_) return out;
        std::string lower = ifaceName;
        for (auto& ch : lower) {
            if (ch >= 'A' && ch <= 'Z') ch = (char)(ch - 'A' + 'a');
        }
        auto it = interfaceAllImplementers_->find(lower);
        if (it == interfaceAllImplementers_->end()) return out;
        return it->second;
    }

    // P6.4+: 类模块默认实例 (VB_PredeclaredId=True), 小写类名 → 类规范名.
    void setDefaultInstanceClasses(const std::map<std::string, std::string>* m) {
        defaultInstanceClasses_ = m;
    }
    // 查询默认实例类名 (大小写不敏感); 非默认实例类返回空串
    std::string defaultInstanceClassName(const std::string& name) const {
        if (!defaultInstanceClasses_) return "";
        std::string lower = name;
        for (auto& ch : lower) {
            if (ch >= 'A' && ch <= 'Z') ch = (char)(ch - 'A' + 'a');
        }
        auto it = defaultInstanceClasses_->find(lower);
        if (it == defaultInstanceClasses_->end()) return "";
        return it->second;
    }
    // 若 name 是默认实例类, 注册到 knownClassVars_, 返回类规范名; 否则返回空串.
    std::string registerDefaultInstanceClass(const std::string& name);

    // Fix 112: 工程内 UserControl (.ctl) 模块的设计器描述 (模块名小写 → FrmFile).
    void setDesignerFiles(const std::map<std::string, FrmFile>* m) { designerFiles_ = m; }

    // Fix 143: vbp Object= 的 OCX 文件表 (CLSID 小写去花括号 → ocx 绝对路径).
    // 供 vb6_OcxHost_Create 免注册 LoadLibrary 实例化第三方 OCX 控件.
    void setOcxFiles(const std::map<std::string, std::string>* m) { ocxFiles_ = m; }

    // Fix 143b: vbp Object= 的原始引用列表 (CLSID 小写去花括号, ocx 绝对路径).
    // 控件的实例化 CLSID 是 vbp Object= 里的 GUID (OCX 的 DllGetClassObject 只认它),
    // 而 typelib 的 coclass GUID 是类型信息 ID, 两者在 VB6 UserControl 上**不同** —
    // 用 typelib GUID 会实例化出错误/失败的对象. 命中规则:
    //   typelib CLSID 在列表里 → 用它; 否则列表恰有一条 → 直接用 (单 OCX 工程).
    void setOcxRefs(const std::vector<std::pair<std::string, std::string>>* v) { ocxRefs_ = v; }

    // Fix 160: vbp ComLib= 免注册 COM 组件表 {ProgID, CLSID, coclass名, 相对exe路径}.
    // 入口点生成时烘焙成 static 表 + vb6_ComLibRegister 调用; 表空则一行不发.
    void setComLibRefs(const std::vector<std::array<std::string, 4>>* v) { comLibRefs_ = v; }

    // Fix <vbeclipse> 2026-10-06: vbp Object= OCX 免注册表 {ProgID, CLSID, coclass名, 相对exe路径}.
    // 入口点生成时烘焙成 static 表 + vb6_OcxRefRegister 调用; 表空则一行不发.
    void setOcxLibRefs(const std::vector<std::array<std::string, 4>>* v) { ocxLibRefs_ = v; }
