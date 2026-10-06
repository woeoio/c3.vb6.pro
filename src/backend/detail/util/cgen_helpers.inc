// cgen_helpers.inc - CCodeGen 的 private 方法声明段（类体片段）
// 由 src/backend/cgen.hpp 拆出（2026-09-17），在类体内被 #include，非独立头文件。
// 内容 = 原 cgen.hpp 第 616~993 行，逐行未改；类体收尾 `};` 由伞头给出。

    // ---- 类型映射 ----

    // Vb6Type → C类型字符串
    std::string mapType(Vb6Type type) const;

    // Vb6Type → COM vtable 方法参数 C类型字符串 (数组→SAFEARRAY*等)
    std::string mapComType(Vb6Type type) const;

    // IID字符串 {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} → GUID结构体初始化字符串
    std::string emitGuidInitializer(const std::string& iidStr) const;

    // TypeRefPtr → C类型字符串 (非const: P6.3收集COM接口类型名)
    std::string mapTypeRef(ASTNode* typeRef);

    // Fix 081e: Declare函数返回类型映射 - Long/LongPtr → intptr_t (x64指针安全)
    std::string mapDeclareType(ASTNode* typeRef);

    // Fix 010b: 尝试将AST表达式常量折叠为int64_t (用于enum成员值)
    // 成功返回true并设置result, 失败返回false
    bool tryEvalConstInt(ASTNode* expr, int64_t& result);

    // VB6默认值 → C表达式
    std::string defaultValue(Vb6Type type) const;

    // Fix 190: Optional 形参省略时, 按形参传值方式给出缺省实参表达式.
    // ByVal 形参直接给值; ByRef 形参 (C 形参是 T*) 必须给 &(T){缺省值} —— 否则把
    // T 值当 T* 交给被调方, 被调方解引用 (如 String 的 (*p)) 会把 T 的数据首几字节
    // 当指针用: Optional ByRef String 交 vb6_BSTR_Empty() 的 BSTR 值 → 被调方读到
    // BSTR 数据首 4 字节 (L"" 之后的未初始化堆字节) → 在 OLEAUT32 里 AV.
    // explicitDefault 非空时优先用它 (形参自带 `= 常量` 缺省).
    std::string defaultArgForParam(Vb6Type type, bool isByVal,
                                   const std::string& explicitDefault) const;

    // ---- 标识符命名 ----

    // VB6标识符 → 安全C标识符 (处理关键字冲突、特殊字符)
    std::string cIdent(const std::string& vb6Name) const;

    // P8.4: 根据表达式类型推断Variant构造函数
    // 例如: 整数字面量→vb6_VariantLong, 浮点→vb6_VariantDouble, BSTR→vb6_VariantString
    std::string wrapVariantValue(ASTNode* valueNode, const std::string& cExpr) const;

    // Fix 198: 装箱点上的布尔口径修正. `As Boolean` 在 C 层就是 int16_t (与 Integer
    // 同型), 布尔字面量又是裸 (-1) (int) → 交给 _Generic vb6_VariantFromValue 会分别
    // 选到 short:/int: 分支, 装箱成 VT_I2 / VT_I4: 实测 VarType(b)=2、
    // TypeName(True)="Long"、经 Variant 的 CStr/Format 打成 -1/0.
    // 本助手只在 inferExprType 判为 Boolean 时把包装换成 vb6_VariantBool, 其余表达式
    // 逐字节回退到 vb6_VariantFromValue(原样) —— 非布尔路径的生成码不变.
    // expr 允许为 nullptr (该位置没有 AST 实参, 如命名实参) → 按原样装箱.
    std::string boxToVariant(Expr* expr, const std::string& cExpr) const;

    // VB6标识符 → C全局函数名
    // 本模块: vb6_<ProcName>
    // 跨模块(外部): vb6_<ModuleName>_<ProcName>
    std::string cProcName(const std::string& procName, AccessLevel access,
                          const std::string& sourceModule = "") const;

    // ---- 表达式求值 ----

    // 生成C表达式字符串, 存入lastExpr_
    void emitExpr(Expr& expr);

    // ---- 语句生成 ----

    // 生成语句列表
    void emitStmtList(StmtList& stmts, bool emitResumePoints = false);

    // ---- 声明生成 ----

    // 生成模块级声明
    void emitModuleDecl(Decl& decl);

    // 生成函数签名 (不含函数体)
    std::string makeProcSignature(SubDecl& node);
    // 账 #222: 事件处理器的 ABI 与"这枚处理器的首形参是不是元素号"两处判断的唯一实现。
    void applyEventHandlerAbi(SubDecl& node);
    bool prepareEventHandlerProc(Module& mod, const std::string& loweredName);
    std::string makeProcSignature(FunctionDecl& node);

    // ai/vb-asm-extension-spec: 过程体 = 单个 Asm 块时降级为独立 MASM 过程 (v1: x64/ml64)。
    // 返回 true = 已接管该过程 (调用方应立即 return, 不再走常规 C 体发射)。
    //   x64: 只发 extern 原型 + 收集元数据到 asmProcs_ (driver 生成 MASM → ml64 → 链接);
    //   x86: 直接发 C 函数体, 内含 MSVC `__asm { }` 内联块;
    //   混排/类方法/Optional/ParamArray/(x64)>4参/浮点参 → 3037; 参数或返回类型越界 → 3037。
    //   isNaked = 过程上的 `<Naked>` (spec §7): 不生成保存/对齐/ret。
    bool tryEmitAsmProc(const std::string& procName, AccessLevel access,
                        std::vector<std::unique_ptr<ParameterDecl>>& params,
                        ASTNode* returnType, const StmtList& body, SourceLocation loc,
                        bool isNaked);

    // 混排 (项2/项3): 在语句流里就地降级一条 AsmStmt。由 emitStmtList 的
    // AsmStmt 分支调用。
    //   x86: 发 MSVC `__asm { }` 内联块 —— [X] → C 变量名 (MSVC 按名解析), 天然
    //        指向 VB 局部变量/参数; 非 Naked 时自动 push/pop 踩到的 callee-saved。
    //   x64: 没有内联汇编, 片段降级为**独立** MASM 过程, 被引用的 VB 变量以**地址**
    //        作参数传入 ([X] → [rcx] 等); 调用点发 `vb6_asm_<proc>_<n>(&X, ...)`。
    //   细节与限制见 ai/vb-asm-extension-spec §12.5。
    void emitAsmStmtMixed(AsmStmt& node);

    // 混排辅助: 扫描汇编块里出现的 [名字] 引用, 分类成
    //   regs  —— 寄存器名 (原样保留)
    //   vbVars —— 解析到可见 VB 变量/参数的名字 (原名 + 是否 ByRef 参数)
    //   function —— 是否出现 [Function]
    // 解析不到 (既非寄存器也非可见变量) 的名字收进 unknown。
    struct AsmRefScan {
        std::vector<std::string> vbVars;        // 出现顺序, 去重
        std::vector<std::string> byRefVars;     // vbVars 里属于 ByRef 参数的
        std::vector<std::string> unknown;       // 无法解析的名字
        bool function = false;
    };
    AsmRefScan scanAsmRefs(const AsmStmt& node);

    // --- Delegate (tB 扩展): 桩类型名/桩名/声明趟发射/收尾定义/直调 ---
    std::string delegateTyName(const std::string& delName);
    std::string delegateThunkName(const std::string& delName, const std::string& procName);
    // 重载 C 名后缀 (O2): 变体符号/AST 键后缀 → "_ov<fp消毒>", 非变体 = 空串
    std::string ovlCSuffix(const Symbol* sym) const;
    std::string ovlCSuffixFromKey(const std::string& suffixKey) const;
    void emitDelegateDecls(Module& module);
    void emitDelegateThunks();
    // --- AddressOf 回调桩 (账 #184) ---
    // VB6 的 AddressOf 交出去的是 __stdcall 调用桩, 不是本体地址。x86 下本体是
    // __cdecl ⇒ 直接把本体给 OS/COM = 每次回调少弹 N*4 字节 (VBFlexGrid 起窗即
    // 0xC0000374 的根因)。被标记的过程 (Symbol::addressOfCallback) 在**定义模块**
    // 里另发一枚 __stdcall 桩, AddressOf 站点取桩地址; 本体约定一字不动 ⇒
    // 现有直接调用与 RTL 的 cdecl 登记面完全不受影响。x64 忽略该约定。
    std::string addressOfThunkName(const std::string& procCName);
    std::string addressOfTargetCName(const Symbol* sym, const std::string& procCName);
    void emitAddressOfThunkDecls(Module& module);   // decl pass: 桩原型
    void emitAddressOfThunkDefs();                  // epilogue: 桩定义
    void emitDelegateCall(IndexOrCallExpr& node);
    // 委托桩目标过程声明查找 (O2: procFp 非空时按重载变体指纹精确匹配)
    Decl* findDelegateTargetDecl(Module& module, const std::string& name,
                                 const std::string& procFp);

    // 生成参数列表
    // staticEntry: ai/024 T02 —— true 表示这条声明要直连**静态归档**里的真实导出
    // 符号 (不是 Win32 DLL 导入库)。此时 `ByVal <x> As String` 声明为 `char*`,
    // 与调用点实际发出的 ANSI `char*` 口径对齐 (见 makeParamCType 的注释)。
    std::string makeParamList(std::vector<std::unique_ptr<ParameterDecl>>& params,
                              bool isDeclare = false, bool staticEntry = false);

    // Fix 084k: 生成单个参数的C类型+名字 ("int32_t x" / "vb6_cls_cWinsock** o"),
    // 与makeParamList逐参数逻辑完全一致, 供事件包装器等复用
    std::string makeParamCType(ParameterDecl* p, bool isDeclare = false, bool staticEntry = false);

    // 生成Property签名 (不含函数体)
    std::string makePropertySignature(PropertyDecl& node);

    // Fix <vbeclipse>: 接口类成员的**默认实现** (VB6 `Public x As New <Interface>`
    // 实例化接口本身 → 接口成员声明即默认实现, 非抽象). 有返回值的发 `return 零值;`,
    // 其余发空操作体. 实现见 cgen_decl_prop.cpp.
    void emitIfaceMemberStub(const SubDecl& node, const std::string& sig, bool returnsValue);
    void emitIfaceMemberStub(const FunctionDecl& node, const std::string& sig, bool returnsValue);
    void emitIfaceMemberStub(const PropertyDecl& node, const std::string& sig, bool returnsValue);
    void emitIfaceStubBody(const std::string& sig, const std::string& retType,
                           const std::string& slotKey, bool takesArgs);

    // 生成类工厂函数 (New/Destroy)
    void emitClassFactory(Module& module);
    // Fix 099: 生成 Public 字段的 COM 读写访问器 (ActiveX DLL 专用, 见 cgen_com.cpp)
    void emitClassFieldAccessors(Module& module);
    void emitInterfaceVtable(Module& module);
    // tB Interface 契约发码 (ai/022 D2/D3, 批次 B04, 见 cgen_iface_vtbl.cpp)
    // ---- Inherits 发码 (tB, ai/022 B07b; 实现在 cgen_inherit.cpp) ----
    const ClassChainView* classChainOf(const Module& module) const;  // 空 = 本模块无继承链
    std::vector<VariableDecl*> structFieldDecls(Module& module);      // 继承字段 (根→叶) + 自有字段
    std::string inheritedRetType(Decl& d);
    std::string inheritedStubSig(Decl& d);
    void emitInheritedProcDecls(Module& module);   // .h: 桩声明
    void emitInheritedProcDefs(Module& module);    // .c: 桩定义

    // ---- 类虚表发码 (tB, ai/022 B08d; 实现在 cgen_inherit.cpp) ----
    // 表类型/字段/实例三处同源: 槽的 C 字段名与签名都取自 stage 3.4b 的 virtSlots。
    const ClassChainView* classViewOf(const Module& module) const;   // 按模块身份取视图 (含链根)
    const ClassChainView* classViewByName(const std::string& cls) const;  // 按类名取 (调用点用)
    void emitClassVirtField(Module& module);       // .h: const void* __cvtbl 字段
    void emitClassVirtType(Module& module);       // .h: vb6_cvtbl_<X> 类型 + extern 表实例声明
    void emitClassVirtNewInit(Module& module);    // .c: _New() 里装载本类的表
    void emitClassVirtImpl(Module& module);       // .c 末尾: 表实例定义
    std::string cvtblImplName(const ClassChainView& v,
                              const ClassChainView::VirtSlot& s) const;  // 表项目标函数 C 名
    std::string virtDispatchCallee(const std::string& cls, const std::string& member,
                                   const std::string& objExpr, bool mustDispatch,
                                   const SourceLocation& loc);  // 空 = 本类不派发这个成员

    // ---- MyBase 显式基调用 + 构造链 (tB, ai/022 B09; 实现在 cgen_inherit.cpp) ----
    // 去虚化: `MyBase.M` 直调基类会跑的那一份实现, 不查 __cvtbl。
    bool tryEmitMyBaseMember(MemberAccessExpr& node);  // true = 已接管 (含判死)
    bool tryEmitMyBaseAssign(MemberAccessExpr& ma, const std::string& rhsC,
                             bool forSet = false);  // MyBase.X = v / Set MyBase.X = obj
    const ClassChainView* myBaseClassOf(MemberAccessExpr& node);  // 空 = 已判死 VB3028
    std::string chainInitCName(const std::string& clsName) const;  // vb6_<X>_chain_init
    bool classIsBaseOfSomething(const std::string& clsLower) const;
    void emitChainInitDecl(Module& module);   // .h: 桥接声明 (仅"有 Class_Initialize 且被继承")
    void emitChainInitDef(Module& module);    // .c 末尾: 桥接定义
    void emitClassInitChain(Module& module);  // .c: 派生类 _New() 里根→叶跑祖先初始化

    void emitIfaceContractTypedefs();                 // .h: 工程级 vb6_ivtbl_/vb6_ivref_ 类型
    void emitIfaceClassFields(Module& module);        // .h: 类结构体里的 __iv_<I> 字段
    void emitIfaceImplTables(Module& module);         // .c: 适配器 + 类侧槽表实例
    void emitIfaceNewInit(Module& module);            // .c: _New() 里回填 vt
    void trackIvrefLocalForRelease(const std::string& cName);  // B05: 登记局部接口变量
    void emitIvrefScopeRelease();                     // B05: 过程出口发 Release
    std::string ivrefCType(const std::string& typeName);   // 空串 = 不是新式接口
    const IfaceView* ivLookupIface(const std::string& written);
    std::vector<const IfaceView*> ivImplementedIfaces(Module& module);
    Decl* ivFindImplMember(Module& module, const IfaceView& v, const IfaceSlotView& slot);
    // tB 委托式实现 (ai/022 B10): 本类把接口 v 整份转交给哪个持有字段 (无则 nullptr)
    const ViaView* ivViaFor(const Module& module, const IfaceView& v);
    std::string ivSlotRetType(const Decl* sig);
    std::string ivParamDecls(Decl& d);
    std::string ivParamDeclsRef(const Decl* sig);
    std::string ivForwardArgs(Decl& d);
    std::string ivImplCName(Decl& d);
    std::string ivSlotKeyForMember(const IfaceView& v, const std::string& member);
    void emitEventSink(Module& module);
    void emitComVtableSinkDecls();  // .h 前向声明
    void emitComVtableSinks();      // .c 实现



    // P7: 生成Win32窗体框架代码 (WndProc + 控件创建 + 消息映射)
    void emitFormFramework(const FrmFormDesc& frmDesc, Module& module);

    // P7.8: 递归生成菜单项 (VB.Menu子项)
    void emitMenuItem(const std::string& parentVar, const FrmControl& menuCtrl, int& menuId);

    // P7.8: 递归生成菜单点击事件派发 (WM_COMMAND中)
    void emitMenuClickDispatch(const FrmControl& menuCtrl, int& menuId);

    // M22: 转义宽C字符串（UTF-8→\xNNNN），用于AppendMenuW等宽字符API
    static std::string escapeWideCString(const std::string& s);

    // M22: 将非BSTR表达式包装为BSTR (用于字符串连接运算)
    std::string wrapToBSTR(const std::string& expr, Expr& node);

    // Fix 194: 字符串型控件属性写入的**唯一**取值入口。RTL 的
    // vb6_SetControlText(hwnd, void* bstr) / vb6_SetMenuCaption 第二参是 BSTR 指针,
    // RHS 若是数值就会被当成指针 —— 必须在此统一转 BSTR, 否则每个属性写入点都得
    // 各贴一遍守卫, 漏一处就静默 UB (见实现处注释)。
    std::string controlPropValueExpr(const std::string& writeFn,
                                     const std::string& valExpr, Expr& valueNode);

    // Fix 100a: 判断表达式本身是否为"返回 BSTR 的调用" (顶层前缀判定, 非子串判定).
    // 用于 wrapToBSTR 的 BSTR 直通分支 — 实参中含 vb6_BSTR_FromStr 的标量函数
    // (vb6_Len("abc") 等) 不得因未锚定子串检查被判为已是 BSTR.
    static bool isBstrReturningCall(const std::string& expr);

    // P7.8: 转义C字符串中的特殊字符
    static std::string escapeCString(const std::string& s);

    // P7.5: 判断控件属性访问 → 返回RTL读取函数名 (如"vb6_GetControlText")
    // 空字符串表示不是已知控件属性
    std::string getControlPropReadFn(FrmControlType ctrlType, const std::string& propName) const;

    // P7.5: 判断控件属性写入 → 返回RTL写入函数名 (如"vb6_SetControlText")
    std::string getControlPropWriteFn(FrmControlType ctrlType, const std::string& propName) const;
    // P20-36: 生成控件属性访问的HWND参数 (Menu控件用GetMenu+menuId)
    std::string makeCtrlHwndArg(const std::string& ctrlNameLower, FrmControlType ctrlType) const;
    // Fix <vbeclipse>: 控件数组属性写的 callee 实参串。普通控件 = CtrlArr_GetAt(...);
    // 菜单 (vb6_SetMenu* 三参: HMENU, menuId, 值) = "(GetMenu((HWND)窗体), 基址+下标)"。
    std::string ctrlArrWriteCalleeArgs(const std::string& writeFn,
                                       const std::string& arrOrigName,
                                       const std::string& idxArg) const;
    // Fix <vbeclipse>: 在其他窗体的设计器描述里找控件 (frmViewViews.tvwViews)。
    // 命中返回该控件描述, 否则 nullptr。designerFiles_ 未设 / 非窗体名都返回 nullptr。
    const FrmControl* findExternalFormControl(const std::string& formName,
                                              const std::string& ctrlName) const;

    // 应用 .frm 设计期样式属性 (BorderStyle / Alignment)。顶层控件与容器子控件共用，
    // hwndExpr 为控件句柄表达式 (如 "vb6_hwnd_Foo")。
    void emitDesignerStyleProps(const FrmControl& ctrl, const std::string& hwndExpr);

    // 账 #125: 应用 .frm 设计期的**状态**三件 (Enabled / Visible / CheckBox·OptionButton 的 Value)。
    // 顶层控件与容器子控件两条创建路共用; hwndExpr 同 emitDesignerStyleProps。
    void emitDesignerStateProps(const FrmControl& ctrl, const std::string& hwndExpr);
    // 账 #142: 设计期那两条**字符串**属性（ToolTipText / Tag）的发码。与上面那条一样被
    // 顶层控件与容器子控件两条创建路共用；只在 .frm 写了非空值时发。
    void emitDesignerStringProps(const FrmControl& ctrl, const std::string& hwndExpr);

    // 账 #154（C29-SL-q）: 设计期 `FontName` / `FontSize` 打到窗口上。
    // 与 #125 / #142 同一形状 —— **两条创建路都要调**（顶层那趟与容器子控件那趟），只接一头就是本线踩过几次的那声不响。
    void emitDesignerFontProps(const FrmControl& ctrl, const std::string& hwndExpr);

    // 账 #160: 设计期 `TabIndex` 下发到窗口（存窗口属性 `VB6_TabIndex`）。
    // `.frm` 写了照发的，没写用 `declarationIndex`（同一父窗内的声明序号）兜底 ——
    // 兜成 0 会让多枚控件同时声称 TabIndex=0。两条创建路都要调。
    void emitDesignerTabIndexProp(const FrmControl& ctrl, const std::string& hwndExpr,
                                  long declarationIndex);

    // 账 #197: 设计期 `ScaleMode` 下发到窗口属性 `VB6_ScaleMode`（写了才发）。
    // **三条落点共用这一处**：顶层创建路、容器子控件创建路、窗体自己的 WM_CREATE。
    void emitDesignerScaleModeProp(const FrmControl& ctrl, const std::string& hwndExpr);

    // 返回控件类型相关的 Win32 样式位 (BS_*/SS_*/ES_*/CBS_*/LBS_*/WS_BORDER...)。
    // 顶层控件与容器(Frame/PictureBox)子控件必须使用同一套逻辑，否则子控件缺
    // SS_NOTIFY/BS_GROUPBOX 等样式会导致外观/交互 (含 tooltip) 异常。
    long controlTypeStyleBits(const FrmControl& ctrl) const;

    // 账 #83(a)（C29-SL-r，读数见 029 的 `C29-SL-r-0`）: 创建时要不要立 `WS_TABSTOP`。
    // VB6 的 `TabStop` 默认是 True，而这个位在**两条创建路上以前都没立过**（backend 里这个常量
    // 一次都没出现过）⇒ 所有控件读回 False、模态窗体里 Tab 不走。两处调用（顶层 + 容器子控件），
    // 拿不到焦点的那几类不立，`.frm` 显式写了 `TabStop` 的按写的来。
    long controlTabStopStyleBit(const FrmControl& ctrl) const;

    // 账 #83(b2): 容器（Frame / PictureBox / SSTab）要不要挂 `WS_EX_CONTROLPARENT`。
    // `IsDialogMessage`（账 #83(b1) 刚给主泵接上）默认只看**活动窗口的直接子窗口**，
    // 不走进没挂这一位的容器 ⇒ 容器里的控件压根进不了 tab 序（实测站点序列
    // `cmdIn1 → cmdTop2 → cmdTop1`，Frame 里的 cmdIn2/cmdDeep 一步都不到）。
    // 容器清单与 `cgen_form_frame_menu.inc` 里那条递归清单同一族 —— **两条创建路都要挂**
    // （容器自己也可以是子控件：嵌在 Frame 里的 Frame 走第二条路）。
    // 账 #167：这类型算不算**容器**（会发子控件、要挂 CONTROLPARENT、子控件要吃派发）。
    // 创建侧与两趟派发侧共用这一处 —— 派发侧以前少认 SSTab ⇒ 子控件 id 整体错位。
    bool controlIsContainerType(const FrmControl& ctrl) const;

    long controlContainerExStyleBit(const FrmControl& ctrl) const;

    // 账 #158（按钮那一半）: BUTTON 类的焦点通知（`BN_SETFOCUS=6` / `BN_KILLFOCUS=7`）只有创建时
    // 挂了 `BS_NOTIFY` 才会往父窗发（裸码实测 `.build/bnnotify/bnnotify.c`，两枚按钮只差这一位）。
    // 两条创建路共用这一处判断（顶层 + 容器子控件）。`BN_CLICKED` 不需要这一位，
    // 所以按钮的 `_Click` 一直是通的——但挂上这一位之后，通知里多了 6/7，
    // 所以 Click 那条 arm 必须开始按 `code == 0` 过滤（见 create 那一趟）。
    long controlButtonNotifyStyleBit(const FrmControl& ctrl) const;

    // 账 #157: VB6 把窗体显示出来时，焦点交给**这枚窗体里 TabIndex 最小的那枚拿得到焦点的控件**
    // （不是创建顺序 —— `.frm` 里控件的书写顺序与 TabIndex 常常相反，实测探针里 cmdB 先创建、
    //  cmdA 才是 TabIndex=0）。这一份**留在发码期算**（账 #160 之后运行期已经有了 `VB6_TabIndex`，
    //  可选改运行时；挪过去要连带把排除表与"每父窗自己编号"的口径一起搬，不值当为省一行发码去冒这个险）。
    // 算出来的目标一次性存到窗体句柄上，由 RTL 的 `vb6_ShowForm` 在激活之后应用。
    void emitFormInitialFocus(const FrmControl& formNode);

    // C29-SL-h: Slider 的创建样式位（TBS_AUTOTICKS / TBS_VERT / TickStyle 那三位）。
    // 顶层与容器子控件两条创建路各自都要它，所以单独一处折算、两处调用（账 #83）。
    long sliderStyleBits(const FrmControl& ctrl) const;

    // C29-SL-l（账 #143）: 控件的**零实参方法**（`SetFocus` / Slider 的 `ClearSel` 这一族）。
    // 一形带括号（走表达式路）、一形不带（走语句路），两条码头都调这两个函数，
    // 免得又只接一头、留下编得过而运行期一声不响的那条静默空转。
    // objExpr 两种形态都认：`vb6_hwnd_List1`（表达式路）与裸小写名 `list1`（语句路）。
    bool formCtrlSlot(const std::string& objExpr, FrmControlType& outType,
                      std::string& outHwndExpr) const;
    std::string controlZeroArgMethod(FrmControlType ctrlType,
                                     const std::string& memberLower) const;
    // 账 #196: **一个实参**的控件方法表（TextHeight / TextWidth）。只交名字，实参由码头拼 ——
    // 与上面那张表同一套规矩，With 形与带括号裸形两条码头都要接。
    std::string controlOneArgMethod(FrmControlType ctrlType,
                                    const std::string& memberLower) const;

    // 账 #196 第三条: 控件的单位换算方法（ScaleX/ScaleY，三个实参）名字表 —— 同一套
    // 「表只交名字、实参由码头拼」的规矩，且刻意不给通用行（理由见定义处）。
    std::string controlScaleMethod(FrmControlType ctrlType,
                                   const std::string& memberLower) const;

    // 账 #221 = C29-PL-a: 控件的**画表面方法**名表（目前只有 Line，六个实参）。
    // 与上面三张表同一套规矩: 表只交名字、实参由码头拼，且**刻意不给通用行** ——
    // VB6 只有画得上去的那几枚（Form / PictureBox）有 Line，给成通用行就等于允许
    // `List1.Line (0,0)-(1,1)` 也答一次成功（同 #192/#196 那条"别伪造成功"的口径）。
    std::string controlCanvasMethod(FrmControlType ctrlType,
                                    const std::string& memberLower) const;

    // C29-SL-n（账 #141）: 「这枚控件要不要子类化」的**唯一一份**判据。
    // 这一趟原本被抄成三份：发 arm 的 `cgen_form_wndproc_subclass.inc`（按 info.hasXxx 汇总）、
    // 装的 `cgen_form_frame_menu.inc`、拆的 `cgen_form_wndproc_dispatch.inc` —— 后两份各写一遍，
    // 实测两份都**少了 `_DblClick` 与 `_Paint`**：只挂 `xxx_DblClick` 的控件，子类过程与
    // `WM_LBUTTONDBLCLK` 那一条 arm 在产物里生成得好好的，却**一次都没被 install**
    // ⇒ 处理器编得过、永不触发（`Slider` 首当其冲，本页 §3 就写着 DblClick）。
    // ⚠ 内容与发 arm 那一趟的 info.hasXxx 逐条对应，**两边一起改**才有意义；
    //   焦点与 Click 那两条必须带排除表，否则同一次操作会从两条路各发一次（C29-SL-g / SL-d）。
    // 账 #189: 控件数组整体成员 (Count/LBound/UBound) 的唯一出口; 不是这三条返回空串。
    std::string ctrlArrayMetaMemberExpr(const std::string& arrName,
                                        const std::string& memberLower) const;
    // 账 #190: 事件处理器的 C 函数名 = 符号自己的拼写 (与 Sub 的写法一致); 没处理器返空串。
    std::string eventHandlerFn(const std::string& ctrlName,
                               const std::string& suffix) const;
    // 账 #191: 工程内 UC 的实例成员访问 (Fix 112 的出口)。recvExpr = 该实例的宿主句柄表达式,
    // 单枚控件与控件数组元素两条路只差这一个串。命中返回 true (已填 lastExpr_/pendingChainObj_)。
    bool emitUcInstanceMemberExpr(const std::string& ucClass, const std::string& member,
                                  const std::string& recvExpr);
    bool controlNeedsSubclass(const FrmControl& ctrl) const;

    // C29-SL-o（账 #83 的 Slider 半边）: Slider 的**设计期值面**下发。
    // 这段原本内联在 cgen_form_ctrl_style_apply.inc（只有顶层控件走得到），容器
    // （Frame / PictureBox / SSTab）里的子滑杆拿不到 ⇒ `.frm` 写的 Min/Max/Value/
    // Small·LargeChange/TickFrequency/Sel* 整条丢（实测顶层读回 10/100/42/2/9/5、
    // Frame 内读回 0/100/0/1/20/0 = 原生默认档）。抽出来两条创建路共用，hwndExpr 由调用方给
    // （顶层那趟自己算槽、子控件那趟用 childHwndRead）。
    void emitSliderDesignTimeInit(const FrmControl& ctrl, const std::string& hwndExpr);

    // VB6 PictureBox/Image 的 STATIC 控件不显示 Caption (标题在创建时应清空)。
    bool controlTypeClearsCaption(const FrmControl& ctrl) const;

    // C29-1a: Shape / Line 的两条创建路 (顶层控件 与 Frame/PictureBox 子控件) 共用一份
    // 口径 —— 之前这两类在 controlTypeToWin32Class 里返回 nullptr, 控件被当"不可见控件"
    // 跳过 (句柄永远 NULL), 而属性表/自绘窗口类其实早就写好了, 缺的就是创建那一刀。
    //  1) Line 没有 Left/Top/Width/Height: 窗口矩形 = X1/Y1/X2/Y2 (容器缇值) 的包围盒;
    //  2) 设计期的 Shape/FillStyle/BorderStyle/BorderWidth/两个颜色 要落到控件属性上。
    static void lineRectFromEndpoints(const FrmControl& ctrl,
                                      int& l, int& t, int& w, int& h);
    void emitShapeLineProps(const FrmControl& ctrl, const std::string& hwndExpr);
    // 控件初始化时的句柄表达式: 控件数组走 CtrlArr_GetAt, 其余走 vb6_hwnd_<名>。
    // (Shape/Line 的设计期初始化以前硬写 vb6_hwnd_*, 控件数组那条路会打到空句柄上)
    std::string ctrlHwndExprForInit(const FrmControl& ctrl) const;

    // C29-SL-d: 该控件的 Click 是否已由原生通知（BN_CLICKED / *_SELCHANGE /
    // TCN_SELCHANGE …）送出。子类化那一路补 Click 档时必须按它排除，否则双发。
    // 定义与逐条来源清单在 cgen_util_ctrl.cpp；两个使用点分别是
    // cgen_form_wndproc_subclass.inc（发不发这条 arm）与 cgen_form_frame_menu.inc
    // （要不要装子类化）。
    static bool controlClickFromNativeNotify(FrmControlType ctrlType);
    // 账 #222: 控件数组元素的键 (见 cgen_util_ctrl.cpp 的定义处)。
    static std::string ctrlElemKey(const std::string& ctrlName, int index);
    // 账 #229: `对象.成员` 的对象位是不是窗体控件（单枚 / 数组元素两形同归），
    // 定义与实测症状在 cgen_util_ctrl.cpp。
    bool ctrlTypeOfMemberObject(const Expr* obj, FrmControlType& outType) const;
    // C29-SL-g: 同理，但排除的是**焦点**通知（TextBox/ComboBox/ListBox/命令按钮那六类的
    // GotFocus/LostFocus 由父窗的 WM_COMMAND 码供，见 cgen_form_wndproc_dispatch.inc 开头）。
    // 三处使用点：发 arm（subclass）、装子类化（frame_menu）、拆子类化（dispatch）——
    // 这一句判据在仓里本来就被抄成三份，任何一处不同步都会留下"装了不拆"或"发了不装"。
    static bool controlFocusFromNativeNotify(FrmControlType ctrlType);

    // P11.7: 返回控件类型的默认属性名 (如 TextBox->"Text", Label->"Caption")
    // 空字符串表示该控件类型无默认属性
    static const char* getDefaultPropertyName(FrmControlType ctrlType);

    // 生成类方法函数体中的Me引用名
    std::string classMeParam() const;

    // Fix 019: 在事件包装函数体内把 void* handler 转换为类指针类型
    // 用于替代历史上误用的 classMeParam() (那是参数声明, 不能作为函数调用实参)
    std::string classHandlerCast() const;

    // ---- 二元运算符映射 ----
    std::string mapBinaryOp(BinaryOp op) const;

    // ---- 表达式类型推断 ----
    // 推断表达式的Vb6Type（简化版，用于Select Case等需要类型判断的场景）
    Vb6Type inferExprType(Expr& expr) const;

    // ai/009 §5.10 (P3): 给**窄整型目标**的赋值右值套上溢出检查 (vb6_ChkByte/
    // ChkInt/ChkLong), 越界即 run-time error 6。原先这类赋值是裸 `t = v;`, 由 C
    // 静默截断 (`a As Byte: a = -1` ⇒ a=255)。
    // 只在右值类型**可能**超出目标范围时套; 同宽/更窄的右值套上去是恒真的噪声。
    // 目标限 IdentifierExpr —— 成员访问/数组下标的目标类型另有一整套解析, 不在
    // 这里猜 (返回原值 = 不套检查, 与改动前的行为一致)。
    std::string narrowCheckAssign(Expr* target, Expr* value, const std::string& cValue) const;

    // P20-42: 窗体控件的**属性返回类型**。见 cgen_util_type.cpp 的 MemberAccessExpr
    // 分支: 那里对 `对象.成员` 的兜底是 lookupModule(memberName), 只按成员**裸名**
    // 查模块符号 —— 于是 `SSTab1.Tab` 会命中返回 BSTR 的内置函数 vb6_Tab 而被判成
    // String, 生成码不给 vb6_CStr(vb6_VariantFromValue(...)) 包装, 而 RTL 的
    // vb6_SSTab_GetTab 返回的是 int32_t → 整数被当 BSTR 指针解引用 → 0xC0000005。
    // 返回 Unknown 表示"这张表不管", 调用方照旧走原有兜底。
    Vb6Type controlPropType(FrmControlType ctrlType, const std::string& propName) const;

    // Fix 029: 严格 Variant 推断 — 仅当表达式"明确"为 Variant/Variant 数组时返回 true.
    // 与 inferExprType 不同: inferExprType 对内置函数/Udt 字段访问等无法识别的情况
    // 会回退到 Variant (默认值), 此处不认为那些是 Variant, 避免误用 vb6_VariantToXxx 包装.
    // isArrOut: 若不为 nullptr, 用于返回是否为 Variant() 数组类型.
    bool isDefinitelyVariantExpr(Expr& expr, bool* isArrOut = nullptr) const;

    // 账 #238: 「这个操作数能不能当 vb6_VARIANT* 取址」只有一处判据。vb6_VarCmp*(&A,&B)
    // 那四个取址点以前只看名字形状 (裸标识符就 & 取址), 于是 `Dim d As Double` 的地址被当
    // vb6_VARIANT* 递进 RTL —— 按 VARIANT 的布局读一个 8 字节标量: 答案错, 且越界读。
    // 裸名字那一形交回 isDefinitelyVariantExpr (声明那几张表 + 符号表); 其余形状不在这里回答。
    bool cmpOperandMayTakeAddr(const std::string& c, Expr* ast) const;

    // Fix 038b: 基于 C 表达式字符串的 Variant 检测 — 补充 isDefinitelyVariantExpr
    // 的 AST 级检测. 当 codegen 生成的 C 表达式包含已知返回 vb6_VARIANT 的函数调用
    // (如 vb6_VariantArrayGet, vb6_VariantFromComResult 等) 时, 判定为 Variant.
    // 仅检查顶层表达式 (去除前导括号后), 避免对子表达式误判.
    bool cExprIsVariant(const std::string& cExpr) const;

    // <vbeclipse>: 与 cExprIsVariant 相对的一侧 —— C 表达式是否**已经是
    // vb6_SafeArray1D\* 载体** (Split/Filter/Array 等内置函数的 C 返回, 以及 _arr_N
    // 匿名临时量)。数组槽实参的 vb6_VariantToSafeArray1D 提取对这类表达式必须跳过,
    // 否则 C2440 (载体指针 → vb6_VARIANT)。见 cgen_util_type.cpp 的定义。
    bool cExprIsSafeArrayCarrier(const std::string& cExpr) const;

    // <vbeclipse>: 装箱权威用它判断"这个 C 表达式顶层已经是 vb6_VARIANT 了"。
    // 必须结构化判 (顶层 vb6_Variant* 且不是 vb6_VariantTo*), 不能靠 ctor 名字清单 ——
    // 清单漏一枚 (vb6_VariantByte / 新增的 VariantSingle / VariantDate) 就会把
    // 已是 VARIANT 的表达式再按 VB 类型装一层, 换来 C2440 + C2198。
    bool cExprIsVariantCarrier(const std::string& cExpr) const;

    // <vbeclipse>: 只有声明名（没有 AST 类型引用）时的 C 类型入口 —— 合成一次
    // SimpleTypeRef 走 mapTypeRef, 不复制那份映射。dll/COM 入口的 extern 原型用它，
    // 以便工程类形参与类模块的定义给出同一个类型。详见 cgen_base_type.cpp 定义处注释。
    std::string cTypeForDeclaredTypeName(const std::string& name);
    // <vbeclipse> <VBFlexGridDemo>: 与 cTypeForDeclaredTypeName 平行的一条判据 ——
    // 专问"这个形参槽是不是类型化 COM 接口 (vb6_ComIface_X*)"，走全局 symTab_.lookup
    // 绕开当前模块作用域 (mapTypeRef 的截短名只看当前模块，TLB 注入的接口挂在
    // 全局，caller 侧就查不到 → 兜底 void*)。空串 = 不是接口槽。
    std::string cParamTypedComIfaceCType(const ParameterInfo& p);

    // <vbeclipse>: 该形参槽在 **C 侧**是否收工程类/接口的**实例指针**。是 → 返回那个
    // 指针类型 (mapTypeRef 的原样结果, 不含 ByRef 的再取址)；否 → 空串。
    // 为什么单独立一条：`Sub Wire(ByVal h As DemoShape)` 这种形参在符号表里 p.type 就是
    // Vb6Type::Variant（类名只留在 p.typeRefName），于是全代码库那句
    // `calleeParams[i].type == Vb6Type::Variant` 会把它当 Variant 槽 —— 调用点装箱成
    // vb6_VARIANT、类模块的定义却是 vb6_cls_DemoShape* ⇒ C2440。所有"这个槽要不要装箱"
    // 的判断都必须再过这一道，否则原型和调用点各按各的理解写（本条就是那次分叉的收口）。
    // 只认 vb6_cls_ / vb6_ivref_ 两档：legacy 接口的 vb6_iface_<I>（非指针）和 COM 的
    // vb6_ComIface_X* 一律返回空串 = 维持本改动之前的行为，不在这里顺手扩大。
    std::string cParamClassPtrType(const ParameterInfo& p);

    // <vbeclipse>: 该名字是否为本工程的某个类/窗体模块 (它的 prop_let_/成员函数真会被发出)。
    // 与 cParamClassPtrType 同一口径, 见 cgen_base_type.cpp 定义处注释。凡是要"按类实例
    // 发码"的地方都该先过这一道 —— RTL/typelib 里同名的"类"没有那份定义。
    bool isProjectClassName(const std::string& name);

    // <vbeclipse>: 配套 cParamClassPtrType 的**交付**侧：实参要放进那个类槽时该怎么发。
    // clsPtr 为空 (= 不是类槽) 时原样返回，所以调用点写成 argVal = deliverToClassSlot(...)
    // 是不分支的一句，不必每个站点各写一遍"是类槽就跳过装箱"。
    std::string deliverToClassSlot(const std::string& clsPtr, const std::string& argVal) const;

    // <vbeclipse> rev7: 配套 deliverToClassSlot 的**落临时**侧名生成器。ByRef 类形参
    // 收到右值实参时必须先落局部再取址 (&(右值) = C2102)。名里只保留标识符字符
    // (`vb6_ivref_IScheme*` 截断会带 '*'), 详见 cgen_base_type.cpp 定义处。
    std::string classSlotTempName(const std::string& clsPtr, int idx) const;

    // Fix 038b: 运行时函数参数 C 类型查找表 — 当 calleeParams 为空 (运行时/内置函数)
    // 时, 通过函数名和参数索引查找期望的 C 类型. 返回空字符串表示未知.
    // 用于在函数调用参数生成时插入类型转换 (VARIANT→具体, 具体→VARIANT).
    static std::string getRuntimeParamCType(const std::string& funcName, size_t paramIdx);

    // ---- AST辅助 ----
    // 检测语句列表中是否包含GoSubStmt
    bool hasGoSubInStmts(StmtList& stmts) const;
    // P12.3: 检测语句列表中是否包含OnErrorStmt
    bool hasOnErrorInStmts(StmtList& stmts) const;
    // P14.1.2: 检测语句列表中是否包含Resume/Resume Next
    bool hasResumeInStmts(StmtList& stmts) const;
    // Bug #1 fix (082h): 预扫描语句中的UBound/LBound(arr,N>1)收集ND数组名
    void scanNDArraysInStmts(StmtList& stmts);
    void scanNDArraysInExpr(Expr& expr);

    // Fix 086: 局部声明过程级作用域提升
    // VB6 的 Dim/Const 是过程级作用域 (块内 Dim 在块外仍可见), 而 C 块作用域
    // 会造成 "分支内 Dim, 另一分支使用" 时 C2065 未声明标识符. 过程体发射前
    // 预扫描收集所有可提升的 LocalDeclStmt (非数组标量/动态数组指针/Const),
    // 在函数序言处统一声明, 原位置的 visit 跳过.
    void hoistLocalDecls(StmtList& body);
    void collectLocalDeclStmts(StmtList& stmts, std::vector<LocalDeclStmt*>& out);
    // 已提升的声明节点集合 (按指针, 每过程清空)
    std::set<const ASTNode*> hoistedLocalDeclSet_;
    // Fix 086: 工程内窗体模块名集合 (小写), driver预扫描填充
    std::unordered_set<std::string> knownFormModuleNames_;
    // Fix 086: For方向拆分的循环体副本索引 (0=首份, 1=第二份); Label/GoTo配对加后缀
    int labelCopyIdx_ = 0;
    // Fix 090o: For方向拆分期间各被拆 body 内定义的标签名栈 (嵌套层)。
    // GoTo 目标若定义在(某个被拆的) For body 内 → 第二份副本 goto 须加 _dN 后缀与
    // 副本内 LabelStmt 配对; 若目标在 body 外 (函数级出口标签如 QH/EH/ErrHandler,
    // VB6: GoTo 跳出循环到过程尾) 只在过程尾定义一份 → 盲目加后缀会 C2094 标签未定义。
    std::vector<std::unordered_set<std::string>> forSplitLabelStack_;
    void collectForBodyLabels(const StmtList& stmts, std::unordered_set<std::string>& out);
    // Fix 090m/090p: ReDim/Erase 目标是 Variant 数组判定 (顶层 As Variant 变量 /
    // UDT 的 As Variant 字段如 (*uFile).BufferArray)
    bool isVariantArrayTarget(const std::string& name);
    // Fix 086: 各模块 Public 常量表 (driver预扫描填充, 本体归driver所有)
    const std::unordered_map<std::string,
        std::unordered_map<std::string, long long>>* modulePublicConsts_ = nullptr;

    // Fix 093b: 跨模块同名模块级变量集合(小写), 本体归 driver 所有
    const std::unordered_set<std::string>* staticModuleVarNames_ = nullptr;

    // Fix 093b: WithEvents 事件包装函数名 (带模块前缀).
    // 不同模块里的同名 WithEvents 变量 (如各处的 m_ListenSocket / m_socket) 会生成
    // 同名包装函数 → LNK2005; 包装函数只在本模块内使用 (sink 是 Set 处的函数内
    // static 局部变量), 带模块前缀即可彻底消除撞名, 且不改变链接属性.
    std::string evtWrapperName(const std::string& varLower, const std::string& evtName) const {
        return "vb6_evt_wrap_" + cIdent(moduleName_) + "_" + varLower + "_" + cIdent(evtName);
    }
    // Fix 093b: COM 源接口事件包装函数名 (带模块前缀, 同上)
    std::string comEvtWrapperName(const std::string& varLower, const std::string& evtName) const {
        return "vb6_com_evt_" + cIdent(moduleName_) + "_" + varLower + "_" + cIdent(evtName);
    }

    // ---- COM辅助 (P6.2) ----
    // 推断COM参数的封装函数: 根据表达式类型选择vb6_ComPackBSTR/Int/Double/Object
    std::string comPackExpr(Expr& expr);

    // Fix 177b: 创建 coclass 实例的表达式 (New / Dim As New 自动实例化共用).
    // 默认 vb6_NewObject(L"<ProgID>") — 运行期按 ProgID 走 CoCreateInstance.
    // 但若该 coclass 被本工程的同名类模块遮蔽 (Symbol::comProjectImplClass 非空,
    // 如 VBMAN.vbp 的 Class=Dictionary vs Microsoft Scripting Runtime 的
    // Scripting.Dictionary), ProgID 仍解析到引用库的实现 → 拿到的对象缺少工程类
    // 自有成员 → DISP_E_MEMBERNOTFOUND. 此时改为直接创建工程类实例, 再经
    // vb6_ComPackVB6InstanceRaw 包装成真 IDispatch (与 ExeComBridge 03 同一通路).
    std::string comNewExprFor(const Symbol* comSym);

    // P25: 解析COM标记为类型化属性取值, 用于COM调用参数打包
    // 当isComMarker_为true时调用, 根据packFnHint选择ComGetObjectProp/GetIntProp/GetStringProp等
    // packFnHint: comPackExpr返回的封装函数名, 用于推断所需属性类型
    std::string resolveComMarkerForPack(const std::string& packFnHint);

    // 解析COM标记为C值表达式 (属性读取语义)
    // 当isComMarker_为true时调用, 生成vb6_ComGetProp+Unpack, 并清除标记
    // unresolvedType: 期望的解封类型, 默认为BSTR (最通用)
    std::string resolveComValue(const std::string& unresolvedType = "BSTR");

    // Fix 092m: With 块类字段写的 COM 解包 hint — 按目标类字段的 VB 类型名
    // (Symbol::memberFieldTypes, 语义分析阶段填充 + driver 跨模块拷贝) 映射到
    // resolveComValue 期望的解封类型:
    //   String → "BSTR" / Long|Integer|Boolean|Byte → "Long" /
    //   Single|Double|Date|Currency → "Double" / Object → "Object" /
    //   Variant → "Variant" / 项目类|UDT → "Object" / 未登记 → "BSTR"(默认).
    std::string classFieldComUnpackHint(const std::string& className,
                                        const std::string& memberName) const;

    // Fix 092p: 类数据字段名规范化 — VB6 大小写不敏感, 源码 `.socket` 与声明
    // `Socket` 是同一字段, 但 C 结构体成员名按声明原样生成. 访问点先经本函数换回
    // 声明名 (Symbol::memberFieldNames), 避免 `me->Client->socket` C2039.
    // 表中无此项 (方法/属性/非本类成员/未知类) 时原样返回.
    std::string canonicalClassFieldName(const std::string& className,
                                        const std::string& memberName) const;

    // Fix 093a: 类成员(方法/属性)名规范化 — VB6 大小写不敏感, 调用点拼写 (count)
    // 可能与类内声明 (Count) 不同; C 符号大小写敏感 → 生成的 vb6_<cls>_prop_get_<m>
    // 与类定义不一致 → LNK2019. 以 Class 符号 memberNames 的声明拼写为准; 无此项
    // (非本类成员/未知类) 时原样返回.
    std::string canonicalClassMemberName(const std::string& className,
                                         const std::string& memberName) const;

    // Fix 093a: 当前类是否声明了同名成员字段 (裸标识符赋值时字段优先于外部同名
    // Property Let/Set).
    bool isOwnClassField(const std::string& memberName) const;

    // Fix <vbeclipse>: 宿主伪对象 (UserControl/Ambient/Extender/PropertyPage) 成员名
    // 规范化 — RTL 的符号名是手写约定 (vb6_UserControl_hDC), VB6 源码拼写可以不
    // 同 (`UserControl.hDc`), C 符号大小写敏感 → C2065 未声明的标识符.
    // 例: ucButton.ctl:64 / ucCaption.ctl:119 / ucTab.ctl:92 写 hDc, RTL 只有 hDC.
    // 非宿主伪对象 / 表中无此项时原样返回.
    std::string canonicalHostPseudoMember(const std::string& pseudoObj,
                                           const std::string& memberName) const;

    // 账 #159 = C29-CH-d: 上面那条规范化与下面三个问题, 全部由 cgen_util_com.cpp
    // 里 kHostPseudoRows 这**一张表**回答 (成员名清单只此一处, 别处再抄就是回归):
    //   hostPseudoBareName   — .ctl/.pag 里裸写的成员要不要接、接成哪个 RTL 符号
    //   hostPseudoValueType  — 按值读时的 VB6 类型 (决定比较走不走装箱那一路)
    //   hostPseudoIsNumeric  — 赋值时 RHS 是 Variant 要不要先拆成数值
    // 每行的类型都必须与 src/rtl/core/vb6rtl/vb6rtl_userctl.h 的 extern 对上,
    // 由 scripts/check_host_pseudo_table.ps1 逐行钉 (含负控)。
    bool hostPseudoBareName(const std::string& pseudoObj, const std::string& memberName,
                            std::string& outRtlName) const;
    bool hostPseudoValueType(const std::string& pseudoObj, const std::string& memberName,
                             Vb6Type& outType, bool requireBare) const;
    bool hostPseudoIsNumeric(const std::string& pseudoObj,
                             const std::string& memberName) const;

    // Fix 010r-16: COM/Property-Get 左值重写辅助
    // 当赋值语句(Let/Set/Assignment fallback)的 LHS 是非常量 C 表达式(非左值)时,
    // 尝试重写为 COM SetProp/SetPropArg 或 Property Let/Set 调用.
    // 参数:
    //   target  - LHS 的 C 表达式(已被 emitExpr 产生)
    //   value   - RHS 的 C 表达式(已被 emitExpr 产生, 已完成 COM 解封)
    //   valueExpr - RHS 的 AST 节点(用于 comPackExpr 推断封装函数)
    //   isSet   - true 表示 Set 语句(对象引用语义), false 表示 Let/Assignment
    // 返回值: 若成功重写并已 emit, 返回 true; 否则返回 false(让调用者继续 fallback)
    bool tryRewriteCOMLvalue(const std::string& target, const std::string& value,
                             Expr* valueExpr, bool isSet);

    // Fix 090w/090x: Property Let/Set 值实参打包 — Let 末形参 As Variant 时,
    // Pattern C/D2 (prop_get_ LHS rewrite) 与 With 块 prop_let_ 调用追加的值实参
    // (double/int/BSTR 等标量 C 表达式) 需按方向转换:
    //   - ByVal Variant  → vb6_VariantFromValue(v)
    //   - ByRef Variant  → 可寻址字段式复合字面量 (&(vb6_VARIANT){.vt=..., .xxx=v})
    //     (按 valueExpr 推断的 VB 类型选字段; 与 cgen_expr M22 ByRef Variant 分支一致)
    // 非法形态 (&(vb6_VARIANT){vb6_VariantFromValue(v)}) 会触发 C2440
    // "vb6_VARIANT→vb6_vartype". lastP.type 非 Variant 时原样返回.
    // Fix 185: 非 const — String 分支复用 wrapToBSTR (非 const 成员).
    std::string packLetValueArg(const ParameterInfo& lastP, Expr* valueExpr,
                                const std::string& val);

    // Fix 159-B: VB6 隐式 CStr 的实参类型清单 — 日期/数值标量赋给 ByVal String 形参
    // 时按 CStr 转换 (vb6_CStrDate/CStrDbl/CStrLong), 见 wrapToBSTR.
    // String/Variant/Object/未知 返回 false (不介入, 维持既有发射).
    static bool isScalarImplicitStr159(Vb6Type t);

    // Fix 090ae: 链式 COM 默认属性索引赋值 (P25b 提取为可复用 helper)
    // VB: dic(a)(b) = v / Set dic(a)(b) = v — 多层默认属性/Item 索引写.
    // AST 形态: 多层 IndexOrCallExpr 嵌套, 最内层 callee 为 COM 变量/类 void*
    // COM 字段 (IdentifierExpr) 或模块默认成员 PropertyGet. 中间层对象用
    // vb6_ComCallObject 逐层解包, 最外层走 vb6_ComSetPropArg.
    // AssignmentStmt/SetStmt 均调用 (090ae: SetStmt 原先缺失该分支 → LHS 被
    // emitExpr 生成值语义 vb6_VariantFromComResult(...) = ... → C2440).
    // 返回值: 若已识别并 emit, 返回 true; 否则 false (让调用者继续 fallback).
    bool tryEmitChainedComWrite(Expr* targetNode, Expr* valueNode);

    // Fix 011r-1: 类实例成员调用解析辅助
    // 给定类名与成员名, 在当前作用域的模块级符号表中查找属于该类的方法/属性符号,
    // 返回构造的 C 函数名:
    //   - Sub/Function     → vb6_<className>_<memberName>
    //   - Property Get     → vb6_<className>_prop_get_<memberName>  (读上下文优先)
    //   - Property Let     → vb6_<className>_prop_let_<memberName>  (仅当没有 Get 时返回)
    //   - Property Set     → vb6_<className>_prop_set_<memberName>  (仅当没有 Get/Let 时返回)
    // 匹配条件 (跨模块): sym->isExternal && sym->sourceModule == className
    // 匹配条件 (同模块类): !sym->isExternal && isClassModule_ && moduleName_ == className  (Fix 013: moduleName_ = VB_Name)
    // 返回空串表示该类中无对应方法/属性, 调用者应视为数据字段访问 (obj->member)
    std::string resolveClassMemberCall(const std::string& className,
                                       const std::string& memberName) const;

    // Fix 033: 类感知方法/属性参数查找 (类符号 + memberParams 表的精确解析).
    // IndexOrCallExpr calleeParams 解析: 按 className 查找指定类的方法/属性参数表,
    // 避免跨模块 storageKey 冲突命中的首个注册者导致 Optional _has_ 标志数错 (→ C2197).
    // Phase A: 与 resolveClassMemberCall 同迭代规则 (匹配 sourceModule 或 同模块
    //          moduleName_), 在 moduleScope 找到真实 Sub/Function/Property 符号,
    //          返回其 params + isBuiltin. 命中即返回 true.
    // Phase B: Phase A 失败时 (storageKey 冲突让消费模块的 external Create 符号的
    //          sourceModule 不匹配 className), 退到 Class 符号自身的 memberParams[lower]
    //          (语义分析阶段已按 Get > Function > Sub > Let > Set 优先级填充), 返回该
    //          参数表 (isBuiltin=false). 命中即返回 true.
    // 都未命中返回 false, 调用者回退到 class-unaware lookupModule.
    bool findClassMemberCallParams(const std::string& className,
                                   const std::string& memberName,
                                   std::vector<ParameterInfo>& outParams,
                                   bool& outIsBuiltin) const;

    // ---- Fix 091a: 属性写方向 (Property Let/Set) 参数查找 ----
    // Pattern C/D2 (prop_get_ LHS → prop_let_/prop_set_) 与 With 块属性写需要
    // **写方向** 参数表来判定末参是否 As Variant (值实参打包). 读方向优先级的
    // memberParams 对「Get 有参 + Let 末参才是 value」的属性 (cJson.Item(key) /
    // Let Item(key, Dat As Variant)) 会取到 Get 的 [key] → 不打包 → C2440.
    // Phase A: 模块作用域中 sourceModule 匹配 className 的 PropertyLet/PropertySet
    //          符号 (跨模块 storageKey 冲突致其缺失时转 Phase B).
    // Phase B: Class 符号自身的 memberLetParams/memberSetParams (driver.cpp 已跨
    //          模块拷贝; 同类内 Let/Set 各自唯一, 不受 Get 优先级遮蔽).
    bool findClassMemberWriteParams(const std::string& className,
                                    const std::string& memberName,
                                    bool isSet,
                                    std::vector<ParameterInfo>& outParams) const;

    // ---- Fix 015: Method chaining 解析辅助 ----
    // 给定一个表达式 AST 节点, 推断其在运行时返回的类名 (如果它返回类实例)
    //  - IdentifierExpr: 查 knownClassVars_, 找到则返回该变量的声明类名
    //  - IndexOrCallExpr: 递归推断 callee.object 的类名, 再用 getClassMethodReturnType 找方法的返回类名
    //  - 其他: 返回空串 (不可推断为类实例)
    // 用于链式调用 db.Sql(s).Exec(...) 中 .Exec 的对象表达式 (db.Sql(s)) 类型推断
    std::string inferClassTypeOfExpr(const ASTNode& expr) const;

    // Fix <vbeclipse> rev8: 该类成员的返回类型是不是 `As Object` (即 C 侧 void*)。
    // `Property Get View() As Object` (View.cls:66) 就是这一类。
    // 用途: 链式成员访问的**接收者是这种成员**时, 后面那一段必须走 COM 晚绑定
    // (`vb6_ComGetProp(obj, L"Member")`), 不能拼 `obj.Member` / `obj->Member` ——
    // obj 实际是 void*, 拼出来必是 C2223 "左侧必须具有结构/联合类型" (实测
    // ucPerspective.c 36 处 .hWnd / ucFolder.c 11 处 .Caption / modSubClass.c 6 处)。
    // 判据 = 类符号的 memberReturnTypes[member] 归一化后是 Object/Variant/Empty,
    // 或该成员的 PropertyGet/Function 符号 type==Object 且 variableTypeName 不是
    // 某个 Class 符号 (与 getClassMethodReturnType 的收紧口径同源)。
    bool classMemberReturnsAsObject(const std::string& className,
                                    const std::string& memberName) const;

    // ---- Fix 037: UDT 类型推断辅助 ----
    // 给定一个表达式 AST 节点, 递归推断其 UDT C 类型标识符 (如 "vb6_type_UcsBuffer").
    //  - IdentifierExpr: 查 knownUdtVars_ 得到 UDT C 类型
    //  - MemberAccessExpr: 递归推断 object 的 UDT 类型, 再在 udtMembers 中找
    //    memberName 对应成员. 若该成员本身是 UDT (typeRefName 非空且符号表中有
    //    UserDefinedType), 返回 "vb6_type_<memberUdtName>"; 否则返回空串.
    //    支持嵌套 UDT: _vb6_with_3.DecrBuffer → 推断 DecrBuffer 的 UDT 类型.
    //  - 其他: 返回空串
    // 用于 IndexOrCallExpr 中 obj.udtArrayField(idx) 模式识别, 避免把 UDT 数组
    // 字段误当函数调用 (C2064).
    std::string inferUdtTypeOfExpr(const ASTNode& expr) const;

    // ---- Fix 085: UDT 对象字段类型推断 ----
    // 给定 UDT C 类型标识符 (如 "vb6_type_tZipFileItem") 与字段名(小写), 返回该字段
    // 的对象类别 C 类型:
    //   - ""                          → 非对象字段 (标量/数组等)
    //   - "vb6_type_Y"                → 嵌套 UDT 字段 (非对象, 供链式推断)
    //   - "vb6_cls_X*"                → 项目类对象字段 (X=类名, 走类方法调用)
    //   - "void*"                     → Collection/COM/接口 对象字段 (走 COM dispatch)
    //   - "BSTR"                      → String 字段 (Fix 177, 赋值须深拷贝)
    // 依据语义阶段 TypeDecl 在 udtMembers 中登记的 type/typeRefName.
    std::string udtFieldObjCType(const std::string& udtCType,
                                 const std::string& memberLower) const;

    // Fix 177: 目标已是**生成的 C 串**时, 判断它是否为 UDT 动态数组元素上的 String
    // 字段 (`VB6_SA_AT(vb6_type_TCELL, ...).Text`) → 赋值必须走 vb6_BSTR_Assign。
    bool udtFieldIsBstrInCTarget(const std::string& target) const;

    // Fix 085: 在 UDT 字段访问拼接处追加对象字段注释标记 (供外层 MemberAccessExpr
    // 消费): obj.field 追加 "  /* udt objfield <CType> */"; 非对象字段原样返回.
    // objExpr 为已生成的左值表达式, udtCType/member 用于判定字段类别.
    std::string appendUdtObjFieldMarker(const std::string& objExpr,
                                        const std::string& udtCType,
                                        const std::string& member,
                                        const std::string& accessOp = ".") const;

    // Fix 084n: 推断 target 是否为 UDT 字段链, 是则返回该字段的 Vb6Type (含 Array 标志), 否则 Unknown.
    // 供赋值语句判断 Variant RHS 需转换的目标类型 (如 cZipArchive 的 .FileName As String
    // ← vb6_VariantArrayGet → vb6_VariantToString; uBuf.MaxMatch As Long ← At() → vb6_VariantToLong)
    Vb6Type inferUdtFieldVb6Type(const ASTNode* target) const;

    // ai/022 B08f-1 (D37): 走同一条 UDT 解析路, 但返回字段在 C 里的**对象类型**
    // (`vb6_cls_X*` / `void*` / "" = 不是对象字段或推不出)。判"这个左值到底是不是 Variant
    // 容器"用它 —— `mi.type` 那条口径对跨模块的 `As <项目类>` 会说是 Variant。
    std::string udtFieldCTypeOfTarget(const ASTNode* target) const;
    // <vbeclipse> <VBFlexGridDemo>: UDT 左值字段的**声明类型名** (`mi.typeRefName`),
    // 空 = 真 Variant 槽, 非空 = 声明过具体类型 (Object/接口/类/UDT)。
    std::string udtFieldTypeRefNameOfTarget(const ASTNode* target) const;

    // Fix 084o: 若 C 表达式为 Variant (字符串级检测或已知 Variant 变量), 返回
    // vb6_VariantToLong(...) 包装, 否则原样返回. 用于需要 int32_t 的上下文:
    // 整除运算符 \ 的操作数、vb6_VariantArrayGet/GetVal 的索引、For 循环 Variant
    // 控制变量的比较/步进等. astExpr 可为 nullptr.
    std::string toLongIfVariant(const std::string& cExpr, const Expr* astExpr);

    // Fix 156: toLongIfVariant 的 double 版 — 用于浮点除法 `/` 的
    // `(double)(x) / (double)(y)`. x/y 是 vb6_VARIANT 结构体时 C 强转非法
    // (C2440 "无法从 vb6_VARIANT 转换为 double"), 改走 vb6_VariantToDouble.
    std::string toDoubleIfVariant(const std::string& cExpr, const Expr* astExpr);

    // 给定类名与成员名, 在当前模块作用域符号表中查找属于该类的 Function/PropertyGet 符号,
    // 若其返回类型为 Object 且记录了 variableTypeName (Fix 015), 返回经 canonicalClassName
    // 规范化后的类名; 否则返回空串. 用于推断方法返回值的类类型.
    std::string getClassMethodReturnType(const std::string& className,
                                         const std::string& memberName) const;

    // 将可能存在大小写差异的类型名 (来自源码 variableTypeName) 规范化为符号表中
    // Class 符号记录的标准名称 (clsSym->name 或 sourceModule), 与 struct 定义
    // vb6_cls_<canonicalName> 大小写一致. 找不到时原样返回.
    std::string canonicalClassName(const std::string& typeName) const;

    // ---- 数组辅助 ----
    // VB6类型 → SAFEARRAY元素类型C枚举名
    std::string mapSaElemType(Vb6Type type) const;
    // VB6类型 → SAFEARRAY元素C类型 (如int32_t)
    std::string mapSaElemCType(Vb6Type type) const;
    // 从ArrayTypeRef或asType获取元素Vb6Type
    Vb6Type resolveArrayElemType(ASTNode* typeRef) const;
    // Fix 055: 从ArrayTypeRef获取UDT元素C类型名 (如"vb6_type_RECT"), 非UDT返回空串
    std::string resolveArrayUdtElemCType(ASTNode* typeRef) const;
    // Fix 192: 从ArrayTypeRef/asType获取元素**类名** (如"ShapeAct"), 非项目类返回空串
    std::string resolveArrayClassElemType(ASTNode* typeRef) const;

    // Fix 100: ReDim complex target (member chain with subscript)
    void emitReDimComplexTarget(ReDimStmt& node);
    // Fix 100: resolve element type of a complex ReDim target
    void resolveReDimComplexElemType(const ReDimStmt& node, Vb6Type& outType,
                                     std::string& outUdtCType) const;
    // P6.4+: class-module default instance singleton
    void emitClassDefaultInstance(Module& module);
    // Fix 110: emit symbol support for .ctl/.pag designer sub-controls
    void emitDesignerControlDecls(const FrmFormDesc& frmDesc, DocumentKind kind);
    // Fix 110: recursively emit designer control handle/array/COM declarations
    void emitControlHandleDecls(const FrmFormDesc& frmDesc);
    // Fix 107: lookup type-category symbol (Class/UDT/Enum/COM), ignoring proc/var
    Symbol* lookupTypeSymbol(const std::string& name) const;
    // Fix 112: find .ctl designer spec by control type name
    const FrmFile* findUserControlSpec(const std::string& controlTypeName) const;
    // Fix 092w: rewrite RHS assigned to a Byte array into byte-array semantics
    std::string rewriteByteArrayValue(const std::string& value) const;
    // Fix 170: `A()` 空括号整体数组引用判定 (无实参 + callee 是已知/模块级数组变量)
    bool isWholeArrayRef(const Expr* e) const;
    // Fix 178: callee 是 UDT 的动态数组成员 (`.Cols()`) 时为 true, 顺带取元素 UDT
    bool isDynamicArrayMemberCallee(const Expr* callee, std::string* elemUdt) const;
    // Fix 170: 整体数组赋值右侧收口 —— 新建载体直接赋, 其余深拷贝避免别名
    std::string wrapWholeArrayAssign(const std::string& target,
                                     const std::string& rhs,
                                     const Expr* targetNode = nullptr) const;

    // ---- Fix 178: 含所有权成员 (String/动态数组/Variant/嵌套 UDT) 的 UDT 深拷贝 ----
    // VB6 运行时按类型描述符对这类 UDT 的赋值与 LSet 做深拷贝; C 端 `d = s` /
    // `memcpy` 是浅拷贝 → 成员指针被多行共享 (VBFlexGridDemo 行 0 与行 >=2 同格)。
    // 命中所有权的 UDT 拷贝点调用 requestUdtCopy, 模块出口 (generate epilogue) 统一
    // 生成 static vb6_udtcpy_<T>() 并插到 .c 的 #include 之后。
    bool udtHasOwnedMembers(const std::string& udtCType) const;
    void requestUdtCopy(const std::string& udtCType, bool needVoidPtrWrapper) const;
    std::string emitUdtCopyBlock() const;
    // 赋值两端都是同一"含所有权成员 UDT" 时返回 vb6_udtcpy_X(&t, &v); 否则空串。
    std::string udtDeepCopyAssign(const Expr& targetNode, const std::string& target,
                                  const Expr* valueNode, const std::string& value) const;
    // 递归版 (stack 防自引用); needVoidPtrWrapper 的传递闭包在 emitUdtCopyBlock 里重算。
    bool udtHasOwnedMembersR(const std::string& udtCType,
                             std::vector<std::string>& stack) const;
