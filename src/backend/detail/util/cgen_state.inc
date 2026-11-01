// cgen_state.inc - CCodeGen 的 private 状态段（嵌套类型 + 成员变量，类体片段）
// 由 src/backend/cgen.hpp 拆出（2026-09-17），在类体内被 #include，非独立头文件。
// 内容 = 原 cgen.hpp 第 278~615 行，逐行未改；`private:` 由伞头给出。

    Diagnostics& diag_;    const SymbolTable& symTab_;
    const TypeSystem& typeSys_;
    bool verbose_;

    // 枚举成员值求值时的**同枚举兄弟成员表** (VbEclipse 批次): 小写成员名 → 常量值。
    // 语义层 visit(EnumDecl) 把兄弟成员定成 SymbolKind::EnumMember, 但那是**按模块**建的表,
    // 而 cgen 这一层拿到的是当前模块的 symTab_ —— 私有枚举的兄弟名在这里查不到,
    // tryEvalConstInt 于是走"回退 emitExpr", 吐出裸 `MSG_AFTER` (C2065/C2057 满屏)。
    // 这里按枚举自己的成员表就地折叠, 不依赖符号表可见性; 非空时优先于 symTab_ 查找。
    std::map<std::string, int64_t> enumSiblingConsts_;

    // Fix 056b: 清理过程内局部数组注册, 保留模块级/类成员数组 (跨过程需要)
    // 过程开始时清空knownArrays_会丢失模块级UDT数组的arrayUdtElemTypes_注册,
    // 导致访问时元素类型回退vb6_VARIANT (MSVC C2440).
    void clearProcArrayTracking();

    // 代码输出
    CodeEmitter h_;   // .h 文件内容
    CodeEmitter c_;   // .c 文件内容
    // C3 bloat fix: 窗体 frx 字节数组 (Icon/Picture/控件 Picture/ImageList/SSTab) 的**定义**。
    // 这些定义在 emitFormFramework 期间从多个调用点产生, 但那些调用点大多落在生成代码的
    // **函数体**里 (CreateControls 等) —— `extern ... = {...}` 在块作用域非法, 故统一攒到
    // 这里, 等 emitFormFramework 返回、c_ 回到文件作用域时一次性发射 (见
    // cgen_base_generate_decl_pass.inc)。.h 侧只发 extern 声明 (bytesToHexArrayDecl)。
    std::vector<std::string> frxArrayDefs_;
    std::string header_;
    std::string source_;
    std::string baseName_;        // 输出文件基名 (如 "hello")
    std::string moduleName_;      // 当前模块名 (用于跨模块函数名前缀)

    // 状态
    Module* currentModule_ = nullptr;
    // <vbeclipse>: 当前模块的 `Option Compare Text` (VB6 是**按模块**的编译期属性)。
    // 由 generate 序言从 module.options 读出; 决定字符串比较发 vb6_StrCmp 还是
    // vb6_StrCmpT、Like 发 vb6_Like 还是 vb6_LikeC, 以及 InStr/StrComp/Replace/
    // Split/Filter/InStrRev 省略 compare 时的缺省值。
    bool optionCompareText_ = false;
    Symbol* currentProc_ = nullptr;  // 当前过程符号
    std::string currentReturnVar_;   // Function返回值变量名 (如 "vb6_ret_CalculateSum")
    std::string currentReturnCType_; // Fix 054: 当前函数返回值的C类型名 (如 "vb6_type_QRCodegenSegment")
    int labelCounter_ = 0;           // 标签计数器 (避免C标签冲突)
    int tempCounter_ = 0;            // 临时变量计数器
    // 账 #305 §B156 (第四十四刀): 多维 ReDim 的界数组临时量以前**按数组名**拼
    // (_redim_bounds_<名>), 同一过程里第二条多维 ReDim 就在同一个 C 函数里把定义发两遍
    // ⇒ cl C2374「重定义；多次初始化」。这一枚计数器给的是「同一模块内第几条多维 ReDim」,
    // 与 tempCounter_ 同生同灭 (同处复位), 使名字在同一过程内必然唯一。
    int ndBoundsSeq_ = 0;             // 多维 ReDim 界数组临时量的序号

    // 表达式求值结果 (累加器模式: 每个emitExpr调用设置lastExpr_)
    std::string lastExpr_;           // 最近一次表达式生成的C代码字符串

    // 已生成的C标识符 (避免重复)
    std::unordered_set<std::string> emittedSymbols_;

    // With语句名称栈
    std::vector<std::string> withObjectVars_;

    // P17.1: With对象元信息栈 — 与withObjectVars_平行，记录对象类别
    enum class WithObjKind {
        Unknown,        // UDT/未知: 生成 struct.field 访问
        FormControl,    // 窗体控件HWND: 使用 getControlPropReadFn/WriteFn
        WithEventsCtrl, // WithEvents控件变量HWND: 使用WE属性路径
        COMObject,      // COM IDispatch*: 使用 vb6_ComGetProp/ComSetProp
        ClassInstance,  // 类实例指针: 使用类方法调用
        BuiltinObject   // 内置全局对象(Err/App等): 使用函数调用
    };
    struct WithObjInfo {
        WithObjKind kind = WithObjKind::Unknown;
        FrmControlType ctrlType = FrmControlType::Unknown;
        std::string ctrlOrigName;    // 控件原始大小写名称(用于HWND变量名)
        // Fix 011r-1: 类实例的类名(如"cJson"), 仅 ClassInstance kind 有意义
        // 已知时优先通过 resolveClassMemberCall 精确解析该类的成员方法/属性,
        // 并把 tempType 设为 vb6_cls_<className>* (避免 void* 上的 .member/->member 错误)
        std::string className;
        // 账 #257 第十九层: With 目标是**宿主伪对象** (UserControl / PropertyPage) 时,
        // 记下源码里那份拼法 (PascalCase 由 hostPseudoRtlSymbol 归一); 空串 = 不是。
        // 体内 `.成员` 是 VB6 的**内建**成员, 与控件自己声明的同名公开属性是两个东西
        // (见 CCodeGen::withHostPseudoIntrinsic)。
        std::string hostPseudoObj;
    };
    std::vector<WithObjInfo> withObjectInfoStack_;
    bool suppressDefaultProp_ = false;  // P17.1: With对象表达式时抑制默认属性读取

    // P14.3.1: Dim As New自动实例化变量集合 (小写key → 类名C标识)
    std::unordered_map<std::string, std::string> knownNewVars_;
    // Fix <vbeclipse> rev27: `Dim X As New <窗体名>` 的窗体实例化集合
    // (小写**变量名** → 窗体名的 C 标识)。
    // 与 knownNewVars_ 分开是因为**发的东西不同**: 窗体没有 vb6_cls_<F>_New(),
    // 它的运行时入口是 cgen_form_ctrl_registry.inc 发的 vb6_form_hwnd_<F>()
    // (首次引用即隐式 Load = VB6 `New <窗体>` 语义)。判据 = knownFormModuleNames_。
    // 之所以要存**窗体名**而不是从变量名推: VB6 惯例变量名与窗体同名但**不保证**
    // (`Dim frm As New frmEditorText` 完全合法), 推错就会发成
    // vb6_form_hwnd_frm() → C2065 未声明。
    std::unordered_map<std::string, std::string> knownNewFormVars_;
    // Fix 090v: 模块/类级 As New 变量持久注册 (knownNewVars_ 在函数入口清理,
    // 需从本表恢复; 局部 As New 由 LocalDeclStmt 每函数重注册) — 否则跨函数
    // 同名局部变量 (如 cColl 的 Dim A As New cCollection vs JsonStr 的
    // Const A) 残留 knownNewVars_ → const 变量被注入 auto-instantiate (C2166).
    std::unordered_map<std::string, std::string> moduleNewVars_;
    // Fix 090v-com: 前期绑定COM变量C类型 (小写key → 如 "vb6_ComIface_IChilkatJsonObject*"),
    // 供 Dim As New 守卫把 vb6_NewObject 的 void* 正确转型到类型化接口指针.
    std::unordered_map<std::string, std::string> knownTypedComVarCType_;

    // Fix 054: 模块级变量延迟初始化语句 (C2099: 文件作用域变量不能用运行时函数调用初始化)
    std::vector<std::string> moduleInitStmts_;

    // P14.3.2: 循环栈 - 支持嵌套Exit For/Exit Do跳转到正确层
    struct LoopInfo {
        ExitKind kind;          // For 或 Do
        std::string exitLabel;  // 跳出标签 (如 "vb6_loop_exit_0")
    };
    std::vector<LoopInfo> loopStack_;

    // 当前过程的已知数组变量名集合 (小写)
    // 用于IndexOrCallExpr中区分数组访问(vs函数调用)
    std::unordered_set<std::string> knownArrays_;
    // 数组名 → 元素Vb6Type (小写key)
    std::unordered_map<std::string, Vb6Type> arrayElemTypes_;
    // P8.1: 数组名 → 维度数 (小写key, 1=一维1D, 2+=多维ND)
    std::unordered_map<std::string, int> arrayDimCounts_;
    // Fix 055: 数组名 → UDT元素C类型 (小写key, e.g. "uvectors" → "vb6_type_RECT")
    // 当数组元素类型为 UserDefinedType 时, 记录实际 UDT C 类型名,
    // 让 VB6_SA_AT / VB6_SA_ND_AT2 使用正确的 struct 类型而非 vb6_VARIANT
    std::unordered_map<std::string, std::string> arrayUdtElemTypes_;
    // Fix 192: 数组名 → 元素**类名** (小写key, e.g. "shapes" → "ShapeAct").
    // 与 arrayUdtElemTypes_ 对称: resolveArrayElemType 把 `As 某类` 压成
    // Vb6Type::Object (枚举装不下"哪个类"), 于是 arr(i).Method 的接收者推断不出类
    // → 发 VB6_SA_AT(void*, arr, i).Method(...) (void* 取成员 → C2224).
    // 静态数组 `Dim s(1) As C` / 动态数组 `Dim a() As C` / `ReDim a(1) As C` 都要登记。
    std::unordered_map<std::string, std::string> arrayClassElemTypes_;

    // Fix 170: 赋值目标 `A() = expr` (空括号整体数组) 上下文标记。
    // 只在 visit(AssignmentStmt) 里 emit node.target 的那一瞬为 true,
    // 让 IndexOrCallExpr 发数组载体 (vb6_SafeArray1D*) 而不是 VB6_SA_AT(elem, A, 0)
    // —— 未 ReDim 的动态数组 data==NULL, 写 0 号元素即解引用 NULL → 0xC0000005。
    bool wholeArrayLvalue_ = false;

    // Fix 178: 本模块内被请求过的「含所有权成员 UDT」深拷贝函数 (vb6_type_X → 是否
    // 需要 void* 版薄封装, 用于 vb6_ArrayAssign1D_Cb 的元素回调)。模块出口处一次性
    // 生成 static vb6_udtcpy_<T>() 插到 .c 的 #include 之后 (见 generate epilogue)。
    mutable std::map<std::string, bool> udtCopyRequested_;

    // Bug #1 fix (082h): 过程内已确认的ND数组名集合 (小写)
    // 当 UBound(arr, N>1) 触发ND版本时将arr注册到此集合,
    // 后续 UBound(arr, 1) 检查此集合也使用ND版本
    std::unordered_set<std::string> knownNDArraysInProc_;

    // 已知BSTR变量名集合 (小写) - 用于Debug.Print等场景判断表达式类型
    std::unordered_set<std::string> knownBstrVars_;

    // 已知double变量名集合 (小写) - 用于Debug.Print区分整数/浮点输出
    std::unordered_set<std::string> knownDoubleVars_;
    // Fix 117c: Single(float) 变量独立登记 — CStr 需要 VT_R4 语义 (7 位有效数字),
    // 否则 CSng(21.1) 会按 Double 打印成 21.1000003814697。
    std::unordered_set<std::string> knownSingleVars_;
    // Fix 175: `As Date` 在 C 层与 Double 同为 double, 只按 C 类型串登记会让
    // inferExprType 永远看不见 Date —— `Print "x=" & (d + 1)` / String 形参收 Date
    // 实参都按 double 走 vb6_CStrDbl → 打出序列号 46023 而不是 2026/1/2。
    // 与 knownDoubleVars_ **并存**登记 (仿 Fix 117c 的 Single 双登记口径),
    // 消费点一律先判 Date, 以免既有 double 路径失去变量。
    std::unordered_set<std::string> knownDateVars_;

    // ai/022 W1: `As Boolean` 在 C 层是 int16_t / VBABOOL, 与 Integer/Long 同一登记串
    // ⇒ 只按 C 类型登记会让 inferExprType 永远看不见 Boolean: `CStr(b)`、`"x=" & b`、
    // String 形参收 Boolean 实参都走 vb6_CStrLong → 打出 -1 而 VB6 打 "True"。
    // 与 knownLongVars_ **并存**登记 (口径同 Fix 117c 的 Single / Fix 175 的 Date),
    // 消费点一律先判 Boolean, 以免既有整数路径失去变量。
    // 账 #123: 局部/形参 As Byte 的专属登记表 (口径同上 knownBoolVars_ 那条 ai/022 W1)。
    // As Byte 的 C 型是 uint8_t, 不匹配既有的 int32_t|int16_t|VBABOOL 那一支, 所以以前
    // 它进不了任何 knownXxxVars_ ⇒ inferExprType 答 Unknown、isDefinitelyVariantExpr 又回退
    // 去查符号表 ⇒ 比较被当 VARIANT 取地址 (实测 `(bt = 65)` 今天返回 False)。
    std::unordered_set<std::string> knownByteVars_;
    std::unordered_set<std::string> knownBoolVars_;

    // Integer 与 Long 共用 knownLongVars_ (C 侧都是整数, 见下面那一份), 收窄检查
    // 却要分出 16 位的 Integer 范围, 否则 `i As Integer: i = 40000` 会被当成 Long
    // 放行 (实测不报错)。口径同上面 knownBoolVars_ 的并存登记: 与 knownLongVars_
    // **并存**, 消费点先判它, 以免既有整数路径失去变量。
    std::unordered_set<std::string> knownIntVars_;

    // ai/032: 无符号/窄整型 (SByte/UInteger/ULong/ULongLong) 的局部/形参/返回槽专属登记表
    // (小写名 → 真 Vb6Type)。口径同上面 knownByteVars_ / knownBoolVars_:
    // 这四档的 C 型 (int8_t / uint16_t / uint32_t / uint64_t) 都不在既有的
    // `int32_t | int16_t | VBABOOL` 那一支里, 而**局部变量在 symTab_ 里不可达**
    // (symTab_ 只有模块作用域) ⇒ 以前 inferExprType 一律答 Variant。后果:
    //   · `CStr(u)` (u As ULong) 走 vb6_CStr(vb6_VariantFromValue(u)), _Generic 的
    //     `unsigned int: vb6_VariantLong` 把 &HFFFFFFFF 打成 "-1";
    //   · `Debug.Print u` 落到 vb6_DebugWriteLong((int32_t)(u)) → 同样 -1。
    // 用一张 name→Vb6Type 的表 (而不是四个 set) 是因为消费点需要知道**哪一档**,
    // 才能选 vb6_CStrULong / vb6_CStrULongLong。模块级变量不进此表 —— 它们在
    // symTab_ 里可达, inferExprType 已能答出, 且本表随过程入口 clear()。
    std::unordered_map<std::string, Vb6Type> knownNarrowIntVars_;

    // 已知Long/Integer/Boolean变量名集合 (小写) - 用于COM值解封类型推断
    std::unordered_set<std::string> knownLongVars_;

    // Bug #2 fix: 已知LongPtr/intptr_t变量名集合 (小写) - 用于直接比较而非VarCmpLong
    std::unordered_set<std::string> knownLongPtrVars_;


    // P8.4: 已知Variant变量名集合 (小写) - 用于赋值时包装值
    std::unordered_set<std::string> knownVariantVars_;
    // P20-44: 处理器形参里 `As DataObject` 的名字 (小写)。这些形参在 C 里是
    // void** (指向 DataObject), 其成员访问 (Data.GetText 等) 必须改道 RTL。
    // **为什么不用 symTab_**: codegen 生成处理器体时形参不在符号表 (实测 lookup
    // found=0), 只能在注册参数的同一循环里顺手收进集合。
    std::unordered_set<std::string> dataObjectParams_;

    // Fix 091m: 模块级 Variant 变量名集合 (小写). knownVariantVars_ 在每个过程
    // 开始时 clear(), 模块级变量 (如 ToolsLogs: Dim EnumLevelNames As Variant)
    // 因而在过程体内不可见 → 赋值时漏包装 (091l 的 Variant 目标判定失效).
    // 模块级声明时登记, 过程开始时回灌.
    std::unordered_set<std::string> moduleVariantVars_;

    // Fix 161f: 无 As Type 但**整数可折叠**的 Const (小写名 → 折叠值)。
    // 这类 Const 在符号表里 type = Variant (未指定类型)，而 C 侧已 emit 成
    // `#define NAME (4127)` —— 过程入口 known*Vars_ 被 clear 后重新按符号表
    // 判定会得到 Variant，于是调用实参被包成 vb6_VariantToLong(LVM_GETHEADER)
    // → C2440 "无法从 int 转换为 vb6_VARIANT" (extlist MListView SendMessage)。
    // 登记后标识符直接内联数值，与已有 `As Long` Const 的内联口径一致。
    std::unordered_map<std::string, int64_t> moduleIntConstValues_;

    // Fix 091p: 类模块/窗体字段中的 Variant 名字集合 (小写). 类字段同样在过程
    // 开始时不可见 (knownVariantVars_ 被 clear), 但字段访问形如 me->field,
    // 前缀明确, 故单独成集合, 只在 me-> 形态下参与判定, 避免污染裸名判定.
    std::unordered_set<std::string> classVariantFields_;

    // Fix 179: 持有 IDispatch 包装器 (vb6_ComObject*) 的变量/字段名 (小写).
    // 179a 创建点登记; 179b 该值再赋给 typed 工程类目标时解包. 每模块新实例, 模块级作用域.
    std::unordered_set<std::string> wrappedComVars_;

    // Fix 062: 已知Byte数组变量名集合 (小写) - 用于Variant→SafeArray1D*转换
    std::unordered_set<std::string> knownByteArrayVars_;

    // Fix 140: 类模块/窗体字段中的 Byte() 数组成员集合 (小写, 含 m_ 前缀).
    // 与 knownByteArrayVars_ (过程局部变量) 对称: 类字段在过程开头 knownByteArrayVars_
    // 被 clear 后不可见, 故单独持久化, 每个过程入口再 copy 回 knownByteArrayVars_.
    // 用途: LabelPlus.ctl 的 `Dim m_Caption() As Byte` — `Property Let Caption` 内
    // `m_Caption = New_Caption` 需按"字符串→字节数组"改写 (vb6_StringToByteArray),
    // 否则直接把 BSTR 当 SafeArray1D* 赋给字段 → UBound/StrPtr 读到的 caption 为空,
    // 右侧 KPI 卡片 (Venta diaria / USD $532.00) 空白 (Charts 2020 实测).
    std::unordered_set<std::string> classByteArrayMembers_;


    // P6.11: 类模块成员变量类型集合 (小写, 含m_前缀格式)
    // 在generate()开头从模块声明填充, 每个方法/属性入口的clear()后从此恢复
    // 防止类方法内对me->m_Xxx的BSTR赋值无法识别类型
    std::unordered_set<std::string> classBstrMembers_;
    std::unordered_set<std::string> classLongMembers_;

    // 账 #305 §B183: Integer 类字段的窄档名单（两种拼法都登，同上面那张表）。
    // 唯一消费者是 narrowTargetTypeOf，它只问「这枚左值有多窄」。
    std::unordered_set<std::string> classIntMembers_;
    std::unordered_set<std::string> classDoubleMembers_;
    // Fix 037: 类 Variant 成员变量集合 (As Variant 字段, 含 m_ 前缀格式).
    // 用于 IndexOrCallExpr 处理 `me->VarField(idx)` 或 `obj.VarField(args)` 模式:
    // Variant 字段可能持有 Collection/SafeArray/Object, 通过 idx 访问时应走
    // vb6_VariantArrayGet(&me->VarField, idx) 而非误把 Variant 当函数调用 (C2064).
    std::unordered_set<std::string> classVariantMembers_;
    // Fix 010r: ALL class member variable names (lowercase, both with/without m_ prefix)
    // Used for me-> prefix detection in Erase/ReDim/Assignment statements
    std::unordered_set<std::string> classMemberVars_;
    // Fix 023: 全局类 void* (外部 COM 类型) 字段表 — 跨 CCodeGen 实例共享.
    // 由 driver.cpp 在 runCodeGeneration 主循环前一次性预扫描所有模块 AST 计算.
    // key: 类名 (如 "cDataBase"), value: 该类结构体中 void* 字段的 lowercase 名集合
    // (包含原名小写 + "m_" 前缀小写两种形式).
    // 在 cgen_expr.cpp knownClassVars_ fallback 路径查询以判断 obj->member 是否
    // 返回 void* COM 指针. 若是, 在 emit 的注释中加入 "voidptr" 标记, 由外层
    // MemberAccessExpr 的链式 COM 检测2 (~line 1477) 识别并切换为 COM dispatch
    // (vb6_ComCall/ComGet*Prop). 例: Db.Rs.EOF (Rs 是 ADODB.Recordset 字段)
    //   内层 Db.Rs → emit "Db->Rs  /* class var .Rs field voidptr */"
    //   外层 .EOF  → 识别 voidptr → vb6_ComGet*Prop(Db->Rs..., L"EOF")
    const std::unordered_map<std::string, std::set<std::string>>* classVoidFieldMap_ = nullptr;
    // Fix 037b: 类模块/窗体模块中 typed (非 void*) 对象字段的类型信息表.
    // key: 类名 (如 "cDataBase"), value: (字段名小写 → VB6类型名)
    //   - 项目类字段 (如 Rows As cCollection) → value="cCollection" (项目类名)
    //   - COM 接口字段 (如 Header As Dictionary) → value="COM:Dictionary" (COM标记)
    // 用于 IndexOrCallExpr 中 obj.typedField(idx) 模式识别 — typedField 不是函数
    // 而是对象数据字段, VB6 语义为调用其默认 Item 属性.
    // 项目类 → 直接调用 vb6_<Type>_prop_get_Item(obj.field, arg);
    // COM 接口 → vb6_ComCall((void*)obj.field, L"Item", args, argc).
    const std::unordered_map<std::string, std::unordered_map<std::string, std::string>>* classTypedFieldMap_ = nullptr;
    // 账 #305 §B168 (第五十九刀) + §B169 (第六十刀): 工程级「类 → (标量数据字段 → 声明档)」表
    // — driver 预扫描。key: 小写类名, value: 小写字段名 (原名 + "m_" 前缀两种拼法都在) → Vb6Type。
    // 用途: 回答「`<对象>.<字段>` 这个左值到底属于哪一档」—— 只有三档用得着:
    //   Variant → 必须装箱 (vb6_VARIANT 槽);  String → 必须深拷贝/接管 (BSTR 槽);
    //   其它标量 → 直发字段赋值;  数组/定长串/对象名/UDT **不入表** (各有自己的账与表)。
    // 赋值站的四条既有判据 (裸名 / 本模块 classVariantMembers_ / VB6_SA_* 文本 / UDT 字段类型)
    // 都够不着**别类**的字段, 于是 §B168 发的是不装箱裸赋值 (cl C2440, `h->PV = 7;`)、
    // §B169 形② 发的是指针别名裸赋值 (`h->PS = s;` — 旧值不释放, 析构再 free 一遍),
    // 而 With 块那一站 (§B169 形①) 在查不到 Property Let/Set 之后**凭空拼一个函数名**
    // `vb6_<Cls>_prop_let_<F>` (实测 DllPrint 281 句) ⇒ C4013 + LNK2019。
    // 查询走 classFieldKind() 一处, 不要在调用点再抄一遍 find; 三个布尔投影
    // (classFieldIsVariantSlot / classFieldIsBstrSlot / classFieldIsScalarData) 也都问它。
    const std::unordered_map<std::string, std::unordered_map<std::string, Vb6Type>>*
        classFieldKindMap_ = nullptr;
    // 账 #476 §B190: 「类 → 动态数组成员 → 声明语法」表 (driver 同一次扫描另运一份, 只运
    // 语法不运档位)。上面那张表刻意不收数组字段, `clsreg_` 在无继承工程里不建表 ⇒ 类字段
    // 那一腿只有这一份来源; 档位到 C 元素类型的映射在 resolveMemberArrayDecl 里做。
    const ClassMemberArrayMap* classMemberArrayMap_ = nullptr;
    // Fix 045: 跨模块 Variant 返回值函数 C 名集合 — 由 driver.cpp 预扫描所有模块构建.
    // 包含所有返回 vb6_VARIANT 的项目函数的 C 名称 (如 "vb6_cAsyncSocket_Znl").
    // 在 cExprIsVariant() 中检查 C 表达式是否调用这些函数, 以正确识别 Variant 值.
    // 解决 cExprIsVariant() 无法检测项目类函数返回 Variant 的核心缺口.
    const std::unordered_set<std::string>* variantReturnFuncs_ = nullptr;
    // Fix 010n: 类模块UDT成员变量 (小写var名 → UDT类型C标识符)
    // 用于在过程开始时恢复 knownUdtVars_ (因clear()会丢失类成员UDT变量)
    std::unordered_map<std::string, std::string> classUdtMembers_;
    // Fix 010n (扩展): 普通模块(.bas/.frm)模块级UDT变量 (小写var名 → UDT类型C标识符)
    // classUdtMembers_ 仅覆盖类模块; 标准模块的模块级UDT变量 (如 Private m_uData As UcsCryptoData)
    // 在过程开头 knownUdtVars_.clear() 后同样会丢失, 导致 With m_uData 被误分类为 COM 对象.
    // 本集合在 generate() 扫描模块声明时填充, 过程开头 clear() 后恢复.
    std::unordered_map<std::string, std::string> moduleUdtMembers_;
    // opt4: 裁剪标准模块.c中未实际引用的跨模块#include
    bool trimIncludes_ = false;
    // opt4: 当前模块实际引用的外部模块(小写), 由 driver 的 AST 收集器填充
    std::unordered_set<std::string> trimModules_;
    // 已知类实例变量名 → 类名映射 (小写var名 → 类名, 如 "me" → "cDialog")
    // 用于方法调用翻译 c.Method → vb6_cls_ClassName_Method(c)
    // Fix 010r-10: 从 unordered_set 改为 unordered_map 以支持类名查找
    std::unordered_map<std::string, std::string> knownClassVars_;

    // 账 #305 #408: 本过程形参表里声明成 As Object 的名字 (小写)。
    // 每张过程入口清空, 只描述当前过程的作用域 —— 用来挡住上面那张表跨过程漏答案。
    std::unordered_set<std::string> procObjectFormals_;
    // 账 #305 §B218 (甲): 本过程**工程类形参**的名字名册 (同上面那张的生命周期, 按过程重开)。
    // 引用计数那两条腿 (覆写 claim / 置空 release) 要它落空 —— ByVal 形参槽里那一枚引用
    // 今天没有任何一站 claim 过它 (实参那一站是腿⑤, 还没接), 让它放手就是拆调用方还在用的实例。
    std::unordered_set<std::string> procClassFormals_;

    // 账 #305 #408: knownClassVars_ 里模块级那一档登记的持久快照 (模块级 Dim As 类、类字段、窗体上的
    // UserControl 子控件、跨模块实例变量)。每张过程入口恢复一份 —— 与 moduleNewVars_ / moduleFixedStringLen_
    // 同一条形状; 形参那一档不进快照, 所以形参遮蔽只影响本过程。
    std::unordered_map<std::string, std::string> moduleClassVars_;

    // Fix 037b: 标记最近一次 IndexOrCallExpr 生成的表达式是否为返回 VARIANT 的
    // 项目类 Item 属性调用 — Set 语句需要将其转换为 void* 对象引用.
    bool lastExprNeedsObjectUnpack_ = false;

    // Fix 010o: 过程局部变量名集合 (小写) — Dim声明的局部变量 + For/ForEach循环变量
    // 用于在IdentifierExpr中避免对局部变量错误添加 me-> 前缀
    std::unordered_set<std::string> knownLocalVars_;
    // 账 #305 第二百四十九刀: 「**本过程**里这一枚名字声明成**非数组**的局部」—— 与
    // knownLocalVars_ 同一条生命周期 (每过程开头清一次, 见 cgen_decl_proc/func/prop 三处)。
    // 为什么需要单列一张: knownArrays_ **从不 clear** (cgen_decl.cpp 那条注释: 清了会丢模块级
    // UDT 数组的元素类型注册) ⇒ 它是跨过程累积的, 于是「同一模块另一过程里 `Dim Count(0 To 1)
    // As Long`」会把本过程那枚 `Dim Count As Long` 也答成数组。多一个读者就多一处 AV。
    std::unordered_set<std::string> procNonArrayLocalVars_;
    std::unordered_set<std::string> knownByRefParams_;  // Fix 081g: ByRef params (lowercase)

    // UDT变量名集合 (小写var名 → UDT类型C标识符, 如 "p" → "vb6_type_Point")
    // 用于成员访问时区分"p.X"(结构体字段) vs "Module1.X"(模块变量)
    std::unordered_map<std::string, std::string> knownUdtVars_;
    // Fix 090q: 已 emit 声明、返回 UDT 的函数/方法 (小写函数名 → C 返回类型,
    // 如 "pvtofiletime" → "vb6_type_FILETIME"). emitFunctionDecl 在函数体
    // (含调用点) 之前注册; 供 Declare As Any ByRef 打包判断实参是否为
    // "返回 UDT 的函数调用" — 需用临时 UDT 复合字面量取址传参, 不能强转
    // struct 值 (cZipArchive pvVfsSetEof: SetFileTime ..., pvToFileTime(...)).
    std::unordered_map<std::string, std::string> funcUdtRetCType_;

    // 定长字符串变量 (小写var名 → 长度表达式, 如 "a" → "10")
    // 用于LSet/RSet使用固定长度而非SysStringLen
    std::unordered_map<std::string, std::string> knownFixedStringLen_;

    // 账 #118: 定长字符串 **UDT 字段** 的长度表 (UDT 的 C 类型名 → 成员小写名 → 长度表达式)。
    // `Type Rec: f As String * 8` 的 `r.f = ...` / `Len(r.f)` 靠它收口到 8。
    // 在 visit(TypeDecl&) 里按声明填 (那一步的 C 类型名就是 key), 全模块有效 —— **不随
    // 过程入口 clear** (与 knownFixedStringLen_ 不同), 所以别把它加进任何一个 clear 列表。
    // 跨模块 UDT 查不到 (符号表按模块隔离, 同 udtFieldObjCType 的既有边界)。
    std::unordered_map<std::string, std::unordered_map<std::string, std::string>> udtFixedStrFieldLen_;

    // 账 #118: **模块级**定长串的持久副本 (小写变量名 → 长度表达式)。knownFixedStringLen_
    // 在每个过程入口 clear (cgen_decl_func.cpp / cgen_decl_proc.cpp / cgen_decl_prop.cpp),
    // 而模块级 `Private gT As String * 5` 是在 declaration 阶段登记的 —— 于是第一个过程
    // 一入口那条登记就没了, `gT = "ab"` 与 `Len(gT)` 双双退回"普通动态串"口径 (实测
    // A-len=2 / A=[ab])。与 classBstrMembers_ / moduleUdtMembers_ 同一套办法:
    // 在 cgen_base_generate_state_scan.inc 里扫一遍留底, 每个过程入口再插回来。
    std::unordered_map<std::string, std::string> moduleFixedStringLen_;

    // M22: Declare A版API函数名集合 (小写) - 需要BSTR->ANSI转换的ByVal String参数
    // 判断: 函数名或Alias以'A'结尾 (如 GetWindowTextA, Alias "WritePrivateProfileStringA")
    std::unordered_set<std::string> knownDeclareAnsi_;

    // Fix 161b-decl-out: Declare VB名(小写) → Alias 导出名 (保留原大小写)。
    // 由来: VB6 `Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA"`
    // 里 VB 名 GetUserName 恰好被 SDK WinBase.h `#define GetUserName GetUserNameW`
    // 抢占 —— 生成的 `#ifndef GetUserName` 守卫为假 ⇒ 调用点 `GetUserName(x)` 被
    // SDK 宏改写成 **GetUserNameW**(宽版), 而我们按 VB6 语义已生成 ANSI 编组的
    // char* 缓冲 → 窄缓冲交给宽 API, 写入 UTF-16 (每字符隔一个 NUL), 回读只得到
    // 首字符 (t1 实测 name=[A])。修法: 有 Alias 时调用点用 Alias 名 (GetUserNameA),
    // 天然绕开 SDK 的 A/W 宏。
    std::unordered_map<std::string, std::string> declareAliasMap_;

    // M22: ANSI临时变量待释放列表 (变量名) + 计数器
    // 在Declare ANSI函数调用前声明临时char*变量, 调用后立即FreeANSI
    std::vector<std::string> ansiTempsToFree_;
    // Fix <vbeclipse>: 本过程被 push_macro/#undef 隔离的隐式变量名 (出过程 pop 复原)
    std::vector<std::string> implicitMacroNames_;
    // Fix 161b-decl-out: Declare A 版 ByVal String **出参回写**登记表。
    // VB6 语义: ByVal String 当缓冲时, 被调用方写入的字节必须按 NUL 截断回写成
    // 新 String (GetUserName/GetModuleFileName 正是靠这条工作)。C3 原先只生成
    // `char* _ansi_N = vb6_BSTR_ToANSI(buf); Call(_ansi_N,...); FreeANSI(_ansi_N);`
    // —— API 写的是临时副本, 原 BSTR 一字未动 ⇒ 出参读回恒空 (t1/t2 实测)。
    // 每条: { 临时char*变量名, 回写目标左值表达式, 缓冲容量表达式(可空) }。
    struct AnsiOutParam { std::string tempVar; std::string targetExpr; std::string capExpr; };
    std::vector<AnsiOutParam> ansiOutParams_;
    // 账 #305 §B217 (第一百二十一刀): 本站为 ByRef Variant 实参 mint 的那枚具名临时槽
    // (`_vb6_vararg_<n>`) 里住着一枚**装箱自己取的引用** —— 语句末收回, 与上面
    // `ansiOutParams_`/`ansiTempsToFree_` 同一条"编译器自己 mint 的临时量, 范围就是这条语句"
    // 的口径。连着记下**发声明时那层的缩进**: 语句的收回行只在同层落地, 层不同就说明声明发在
    // 一对已经合上的花括号里 (Debug.Print 那族 `{ ... }` 自带块), 那时宁可不收也不许发成
    // 引用越界的 C。
    std::vector<std::pair<std::string, int>> boxTempsToClear_;
    int ansiCounter_ = 0;
    int vcmpCounter_ = 0;  // P25: Variant比较临时变量计数器

    // 账 #464 §B164 回收刀: 一条语句里现造的带权临时量（**整枚实参 = 一个 mint 调用**）
    // 提成具名临时量之后, 语句末按**档**收回。
    // 档: 1 = BSTR（vb6_BSTR_Free）· 2 = VARIANT（vb6_VariantClear(&t)）。
    // 档号**只由 mint 出口的返回档定**（owned_temp_keeps.hpp 的 [own] 节）——
    // 源表达式的 VB 档不作数, 消费者的 VB 形参档也不作数。
    // 连着记下**发声明时那层的缩进**: 语句末的收回行只在同层落地, 层不同就说明声明发在
    // 一对已经合上的花括号里 ⇒ 那时**宁可不收**也不许发成引用越界的 C（同 boxTempsToClear_）。
    struct OwnedTemp { std::string name; int level; int kind; };
    std::vector<OwnedTemp> ownedTempsToFree_;
    int ownedTempCounter_ = 0;
    // 拼接站左操作数守卫: 左操作数的**发码文本**被拼接站那条「含 vb6_BSTR_Concat 子串」的
    // 判据读过 ⇒ 左操作数里的一切**不许提名**（提了名那条判据会翻面 ⇒ 行为变了）。
    // 左操作数发码期间把计数抬起来, 右操作数发码期间落回去（递归地护住每一层的左操作数）。
    int concatLeftDepth_ = 0;
    // 当前**语句**所在的那一层（emitStmtList 每进一条语句记一次）——
    // 回收刀只在「这一枚临时量的声明与它的收回落在同一层」时才提名：
    // 声明发在语句这一层 ⇒ 语句末的收回行也发在语句这一层 ⇒ 同作用域，引用不越界。
    // 声明若发在更深的块里（那一层的块已经合上了）⇒ 宁可不收。
    int stmtLevel_ = 0;
    // 账 #464 §B164: **块内条件**的深度 —— 回收刀的门 ⑨。
    // 有一类表达式是在**块已经开起来之后**才求值的：ElseIf 的条件
    //   （生成器发成 `} else {` + 条件 + `if (...) {`，见 visit(IfStmt) 的 Fix 063b）
    //   —— 那一枚临时量的**声明**因此落在 else 块**内**，而语句末的收回行落在 else 块**外**
    //   ⇒ 引用越界（cl C2065: '_own_N': 未声明的标识符）。
    // 这一格只许在「块内条件」里挡下提名（宁可不收 ⇒ 与基线同漏，不新增漏）。
    int blockCondDepth_ = 0;

    // 已知COM对象变量名集合 (小写) - 用于后期绑定 obj.Method → vb6_ComCall(obj, L"Method", ...)
    std::unordered_set<std::string> knownObjectVars_;

    // P6.3: 前期绑定COM变量 (小写变量名 → ComClass符号指针)
    // Dim fso As FileSystemObject → knownTypedComVars_["fso"] = &FileSystemObject符号
    // 生成vtable直接调用而非IDispatch后期绑定
    std::unordered_map<std::string, const Symbol*> knownTypedComVars_;

    // P6.3: 已使用的COM接口类型名 (如 "IFileSystem3")
    // 用于在.h文件中生成typedef前向声明, 使 vb6_ComIface_<Name>* 类型可用
    std::unordered_set<std::string> usedComIfaceTypes_;

    // P6.4: 已使用的VB6接口类型名 (如 "IFoo")
    // 用于在.h文件中生成typedef前向声明, 使 vb6_iface_<Name> 类型可用
    std::unordered_set<std::string> usedVb6IfaceTypes_;

    // Fix 010: 已使用的VB6类类型名 (如 "cHttpServerContext")
    // 用于在.h文件中生成typedef前向声明, 使 vb6_cls_<Name>* 类型可用
    // 解决循环#include导致的类型未定义问题 (C2081错误)
    std::unordered_set<std::string> usedClassTypes_;

    // Fix 010: 已使用的VB6 UDT类型名 (如 "TypeLang")
    // 用于在.h文件中生成typedef前向声明, 使 vb6_type_<Name> 类型可用
    std::unordered_set<std::string> usedUdtTypes_;

    // COM后期绑定中间状态 (P6.2)
    // MemberAccessExpr为COM对象设置此字段, IndexOrCallExpr/AssignmentStmt/SetStmt读取后清除
    // 当此字段非空时, lastExpr_中的"值"是对象表达式, comMemberName_是成员名
    std::string comObjExpr_;        // COM对象C表达式 (如 "fso")
    std::string comMemberName_;     // COM成员名 (如 "CreateTextFile")
    bool isComMarker_ = false;      // lastExpr_是否为COM标记

    // Bug #3 fix: 标记lastExpr_是vb6_ComIface_Picture*类型的函数返回值
    // 当函数返回类型为StdPicture/IPictureDisp时设置, 用于Picture属性赋值时
    // 选择vb6_SetControlPictureFromCom而非vb6_SetControlPicture
    bool lastExprIsComPicture_ = false;

    // P6.3: 前期绑定中间状态 (在isComMarker_基础上额外标记)
    bool isEarlyBoundCom_ = false;  // 当前COM标记是否为前期绑定 (vtable直接调用)
    const Symbol* earlyBoundSym_ = nullptr;  // 前期绑定的ComClass符号 (含方法签名)

    // 是否需要 setjmp.h (On Error GoTo label)
    bool needSetjmp_ = false;

    // 多模块项目标志 (影响Public函数命名: vb6_<Module>_<Proc> vs vb6_<Proc>)
    bool isMultiModule_ = false;

    // Fix 142: VBP 的 Startup= 启动对象 ("Sub Main" 或窗体模块名).
    // 多模块工程里每个窗体模块此前都生成 WinMain → LNK2005 多重定义.
    // 只有启动对象所在模块可以生成进程入口点.
    std::string startupObject_;

    // M22: 外部模块名称 (用于跨模块变量解析时判断模块是否已#include)。
    // 账 §B97: vector 不是 unordered_set —— 三处会出顺序的码头（.c/.h 的 #include 段、
    // 入口点里的模块 init 调用序）直接迭代它，顺序必须是 driver 定下来的工程声明序。
    std::vector<std::string> externalModules_;

    // M22: 当前IdentifierExpr是否在IndexOrCallExpr的callee位置
    // VB6语义: 同名函数引用 — callee上下文返回函数名(供调用), 其他上下文返回返回值变量
    bool asCallCallee_ = false;

    // Fix 015: 类方法链式调用对象参数传递管道
    // visit(MemberAccessExpr) 在 Fix 015 路径中, 当 asCallCallee_=true (外层是
    // IndexOrCallExpr 或 CallStmt 的 callee context) 时, 不直接 emit "func(wrappedObj)"
    // (那样会让 IndexOrCallExpr 的空参数shortcut 或 CallStmt 的 bare-call 分支错误地
    // 把 callee 当作已完成调用, 跳过 Optional 参数默认值填充), 而是把 wrappedObj 存储
    // 到此字段, lastExpr_ 只返回裸函数名. visit(IndexOrCallExpr) / visit(CallStmt) 在
    // 完成参数处理后从此字段取出 wrappedObj, 作为首个 (this指针) 参数前置.
    // 一旦消费即清空, 防止跨调用泄漏.
    std::string pendingChainObj_;

    // 类虚表 (ai/022 B08e-4 站点⑭): 赋值/Set/Let 的**左值**文本发出来之后还要过一遍
    // `cgen_util_comwrite.cpp` 的 Pattern C/D2 —— 它把 `prop_get_X` 整体换成 `prop_let_X`/
    // `prop_set_X` (写方向在 3.4b 没有槽, 只能直调)。callee 若是派发表达式
    // `((const vb6_cvtbl_C*)(...)->__cvtbl)->prop_get_x`, 那个盲替换会得到表里不存在的
    // `->prop_let_x` (error C2039)。所以左值求值期间置 true, 让 ①⑤⑪ 三处都不改派。
    // 读上下文 (普通表达式里的 `Me.f.Level`) 不受影响, 照旧派发。
    bool suppressVirtDispatch_ = false;

    // 链中对象求值旗标: visit(MemberAccessExpr) 的 generic_access 在 emitExpr(对象)
    // 期间置 true —— 自类型过程作链中接收者 (frmMain.RibbonBar.FindTab) 的那几处
    // 发码分支据此把 comMemberName_ 留空 (对象值已就绪, 成员名其实是被调函数名,
    // 不能再交给 resolveComValue 解一层); 左值/成员读上下文保持原名。
    bool inMemberObjectEmit_ = false;

    // 类模块标志
    bool isClassModule_ = false;
    bool isFormModule_ = false;

    // Fix 010: 类模块变量注册模式 — 只填充tracking set, 不生成变量声明(已在结构体中)
    bool trackOnly_ = false;
    std::string formName_;  // M22-Issue6: 当前窗体模块名 (用于Form Print)

    // P6.4: Implements 接口引用变量 (小写变量名 → 接口名)
    // Dim x As IFoo → knownIfaceVars_["x"] = "IFoo"
    std::unordered_map<std::string, std::string> knownIfaceVars_;
    // Interface 契约 (tB, ai/022 B04): `Dim x As <新式接口>` → 变量名小写 → 接口名.
    // 与 legacy 的 knownIfaceVars_ 分开: 新式接口值是**薄指针** (vb6_ivref_<I>*),
    // 派发/Set 都走 ivref 分支, 不共用胖对文本.
    std::unordered_map<std::string, std::string> knownIvrefVars_;
    // 成员访问片段 -> 调用片段之间传递"这次是接口槽调用"的接口名 (B04)
    std::string ivrefIfaceName_;
    // tB Interface B05: 本过程内需要"作用域末尾 Release"的接口变量 (C 名)
    // 账 #305 §B218 (甲) 腿④: 「过程出口这一格该不该放手」收成一张**带档**的表
    // (原来只有 ivref 一档有答案, 具体类那一档压根没有登记面 = 今天每条 `Dim x As C` 漏一枚)。
    // isView = true 走 `->vt->Release` (接口薄指针), false 走本类那对 mint 出入口 (工程类实例)。
    // isCarrier = true = 第三档「数组载体」(第一百五十二刀): 这一格躺的不是引用而是一枚载体,
    // 出口发的不是 `_Release` 而是那位数组主人 (出厂逐枚)。arrayName = 登记那一刻的**数组名**,
    // 出口照它问「哪一类的元素 / 几维」—— 两问与语句站点共用同一位主人, 不另开第二问。
    struct SlotRelease {
        std::string expr;
        bool isView;
        std::string clsName;   // 只有工程类那一档带类名
        bool isCarrier = false;
        bool freeCarrier = false;  // 只有载体档用: 出口除逐枚交回之外还销毁载体
        std::string arrayName; // 只有载体档带带数组名
    };
    std::vector<SlotRelease> slotsToRelease_;
    // 账 #305 §B218 (甲) 第一百五十六刀: 「这一档局部数组的载体有没有在本过程里被整体用掉」
    // 由过程入口那一趟预扫 (scanLocalArrayEscapes) 一次性回答, 存小写裸名。登记口只在这张名单
    // 不含它时才允许出口销毁载体 —— 认不出的形状一律算逃逸 (漏判的方向是继续漏, 不是悬空读)。
    std::set<std::string> escapedLocalArrays_;
    // 第一百五十六刀立、第一百六十四刀改口径的第二道闸: 「出口那一发拿这个名字去销毁, 那一刻它
    // 在不在场」。从前这一问等于「Dim 是不是过程体的直接子句」—— 嵌在 If / For / With 括号里的那
    // 几档只在那对括号内存在, 拿它们去销毁 = cl C2065 (门 #821 的 FlexGridX86 十二发)。现在声明
    // 住在序言里就算在场: 动态那一档从 Fix 086 起就是, 定长与多维那一档由本刀补上 ⇒ 提升开着 ⇒
    // 每一档局部数组都进名单; C3_NO_HOIST 关掉提升 ⇒ 缩回「只有直接子句」那一个答案。
    std::set<std::string> exitVisibleArrayCarriers_;
    // 第一百六十四刀: 提升那一趟发过序言声明的那几枚载体名 (小写)。名单由主人一处写、原位那一发
    // 一处读 —— 原位不许自己再推一遍「我这枚到底在不在序言」。
    std::set<std::string> preludeArrayCarriers_;
    // Set 语句里"旧值"临时变量的序号 (只有新语法路径消耗, 不影响其他工程)
    int ivrefSetSeq_ = 0;

    // VB6 Static Sub/Function标志: 过程内所有局部变量都应生成C static
    bool inStaticProc_ = false;

    // GoSub返回地址计数器和标志 (每个过程独立)
    bool hasGoSub_ = false;
    int gosubReturnCounter_ = 0;

    // P12.3: On Error嵌套支持标志 (每个过程独立)
    bool hasOnError_ = false;

    // P14.1.2: Resume dispatch switch支持 (每个过程独立)
    bool hasResume_ = false;            // 当前过程使用了Resume/Resume Next
    bool inProtectedBlock_ = false;     // 在On Error GoTo和错误处理器标签之间
    int resumePointCounter_ = 0;        // 当前resume点索引
    std::vector<int> dispatchPoints_;   // 已生成的resume点索引列表
    std::string currentErrorHandlerLabel_; // 当前On Error GoTo的错误处理器标签名

    // Fix <vbeclipse>: 过程统一出口 (vb6_proc_exit) 是否被引用。
    // VB6 语义: 过程退出时错误处理状态**自动恢复**。原先 Exit Sub/Function/Property
    // 发裸 `return;`, 直接越过过程尾部的 vb6_RestoreErrState() —— vb6_err_stack_top
    // 不弹栈且 vb6_error_jmp_ptr 仍指向**已失效的栈帧**; 之后任何 VB 错误
    // (vb6_ErrRaise) 都会 longjmp 进死帧 → 野读崩。
    // 现在 Exit* 改为 `goto vb6_proc_exit;`, 尾部统一做 restore/release/return。
    bool procExitLabelUsed_ = false;

    // P6.5: WithEvents变量 (小写变量名 → 源类名)
    // Dim WithEvents obj As ClassName → knownWithEventsVars_["obj"] = "ClassName"
    std::unordered_map<std::string, std::string> knownWithEventsVars_;
    // P16: WithEvents控件变量 (小写变量名 → FrmControlType)
    // Dim WithEvents cmd As CommandButton → knownWithEventsCtrlVars_["cmd"] = CommandButton
    std::unordered_map<std::string, FrmControlType> knownWithEventsCtrlVars_;
    // P16: WithEvents控件变量原始名 (小写 → cIdent原名, 如"cmd" → "cmd")
    std::unordered_map<std::string, std::string> knownWithEventsCtrlOrigNames_;

    // P7.5: 窗体控件名映射 (小写控件名 → FrmControlType)
    // 由emitFormFramework从FrmFormDesc填充，用于识别 ctrl.Property 的控件属性访问
    std::unordered_map<std::string, FrmControlType> knownFormControls_;

    // P7.9: Window control name mapping (lowercase -> original casing for HWND vars)
    std::unordered_map<std::string, std::string> knownFormControlOriginalNames_;

    // 控件名 (小写) → 设计期类型名 (如 "DockingPaneManager" → "XtremeDockingPane.DockingPane")。
    // 控件变量名与 typelib coclass 名不一定同名 (CommandBars 同名, DockingPaneManager 不同名),
    // 晚绑定"成员是方法还是属性"的判据必须先按类型名找回 coclass。
    std::unordered_map<std::string, std::string> knownFormControlTypeNames_;

    // P20-36: 菜单项ID映射 (小写菜单名 → menuId, 与emitMenuItem/emitMenuClickDispatch一致)
    std::unordered_map<std::string, int> knownMenuIds_;
    std::string knownMenuFormHwnd_;  // 当前窗体HWND变量名 (如 "vb6_hwnd_Form1")
    // Fix <vbeclipse>: 菜单控件数组基址 (小写菜单名 → 首元素 menuId)
    std::unordered_map<std::string, int> knownMenuArrayBaseIds_;

    // P7.5: 窗体名 (小写)，用于识别 Form.Caption 等窗体自身属性
    std::string knownFormName_;


    // P7.6: 控件数组名集合 (小写控件名 → 是否为数组)
    // 同名控件出现多次时标记为数组, 生成 vb6_CtrlArr 而非 void* HWND
    std::unordered_map<std::string, bool> knownControlArrays_;

    // P7.6: 控件数组名 → 控件ID映射 (控件ID → 数组Index, 用于事件分发)
    // key=小写控件名, value=vector<int>按控件ID顺序对应的Index
    std::unordered_map<std::string, std::vector<int>> controlIdToIndexMap_;

    // P20-39: 原生复刻的 ImageList 控件槽 (形如 "vb6_com_ImageList1")。
    // ImageList 无窗口, 复刻后它就是个 HIMAGELIST, 存在沿用下来的 vb6_com_<Name> 槽里。
    // 用于把 `ImageList1.ListImages.Add/Remove/Clear`、`ListImages.Count`、
    // `ListImages(i).Key`、`ImageList1.ImageWidth` 这些**字符串形态**的 COM 表达式
    // 改道到 RTL 原生函数上。Toolbar/StatusBar/CommonDialog 这三个还是 OCX 占位, 不登记。
    std::unordered_set<std::string> imageListVars_;
    // C29-3: 当前右值是否落在 **Let 赋值**（`n = ...`）上。Let 与 Set 对成员对象的
    // 语义不同: `n = ListImages.Add(..)` 在 VB6 里取的是 ListImage 的**默认属性**
    // (`Index`), 而 `Set img = ListImages.Add(..)` 拿对象本身。两者走的是两个不同的
    // visit (AssignmentStmt vs SetStmt), 天然分开 —— 所以这个标志只在 Let 期间置真。
    bool memObjLetScalar_ = false;

    // C29-7: 原生复刻的 ListView 控件 (登记裸控件名, 如 "ListView1")。
    // ⚠ ListView 是**真窗口**, 所以 RTL 各函数第一参要的是 **HWND 槽 vb6_hwnd_X** ——
    // 与 ImageList 那族的 vb6_com_X (无窗口, 槽里是复刻实例指针) 不同, 别照抄。
    std::unordered_set<std::string> listViewVars_;    // objExpr 是否落在一个已登记的 ListView 上 (返回裸控件名, 否则空串)。
    // 形态: ① 裸槽 vb6_hwnd_X
    //       ② 集合对象 vb6_ListView_ListItems((void*)vb6_hwnd_X) / ...ColumnHeaders(...)
    //       ③ Item 结果 vb6_ComCallObject(②, L"Item", …) —— ③ 靠 ② 的子串命中
    // Fix 161f: 形参/局部变量持有 ListView 的名集合 (`Sub RefillList(lv As ListView)`)。
    // 这类槽在 C 侧**本身就是 HWND** (调用点传的实参就是 vb6_hwnd_ListView1),
    // 与控件名集合 listViewVars_ (那是窗体设计期控件, 要拼 vb6_hwnd_<名>) 分开存,
    // 因为拼法不同 —— 形参直接就是槽变量名, 不能再加 vb6_hwnd_ 前缀。
    std::unordered_set<std::string> listViewSlotVars_;

    // 返回该 objExpr 对应的 **ListView HWND 表达式** (可直接当 (void*) 实参用)。
    // 形态: ① 裸槽 vb6_hwnd_X (设计期控件)
    //       ② 集合对象 vb6_ListView_ListItems((void*)vb6_hwnd_X) / ...ColumnHeaders(...)
    //       ③ Item 结果 vb6_ComCallObject(②, L"Item", …) —— ③ 靠 ② 的子串命中
    //       ④ Fix 161f: 裸形参/局部变量名 (lv / (*lv)) —— 本身就是 HWND
    std::string listViewHwndExprOf(const std::string& objExpr) const {
        static const char* kHeads[] = {
            "vb6_ListView_ListItems((void*)vb6_hwnd_",
            "vb6_ListView_ColumnHeaders((void*)vb6_hwnd_" };
        for (const char* h : kHeads) {
            size_t p = objExpr.find(h);
            if (p == std::string::npos) continue;
            size_t s = p + strlen(h);
            size_t e = s;
            while (e < objExpr.size()
                   && (isalnum((unsigned char)objExpr[e]) || objExpr[e] == '_')) e++;
            std::string n = objExpr.substr(s, e - s);
            if (listViewVars_.count(n)) return "vb6_hwnd_" + n;
            if (listViewSlotVars_.count(n)) return n;   // 形参本身就是 HWND
        }
        if (objExpr.size() > 9 && objExpr.compare(0, 9, "vb6_hwnd_") == 0) {
            std::string n = objExpr.substr(9);
            if (listViewVars_.count(n)) return "vb6_hwnd_" + n;
        }
        // Fix 161f: 裸形参/局部名 (lv) 或 ByRef 解引用 ((*lv)) —— C 侧已是 HWND。
        {
            std::string b = objExpr;
            if (b.size() > 4 && b[0] == '(' && b[1] == '*' && b.back() == ')')
                b = b.substr(2, b.size() - 3);
            size_t sp = b.find_first_of(" \t");
            if (sp != std::string::npos) b = b.substr(0, sp);
            if (!b.empty() && listViewSlotVars_.count(b)) return b;
        }
        return "";
    }
    // 兼容旧调用点: 只要一个"是 ListView 宿主"的真假判据。
    bool isListViewHostExpr(const std::string& objExpr) const {
        return !listViewHwndExprOf(objExpr).empty();
    }
    // Fix 161f: 保留旧名做别名, 老调用点 (frmLVStyles 系列的 C29-7 路径) 语义不变。
    std::string listViewHwndVarOfExpr(const std::string& objExpr) const {
        return listViewHwndExprOf(objExpr);
    }
    // C29-Data: Data 控件 (不可见类 VB6_DATA 属性宿主, 登记裸控件名如 "Data1")。
    // 账 #278 §B124: Recordset 链今天**只有一条路** —— cgen_util_ctrl.cpp 在 recordset
    // 那一档交出 `vb6_Data_RecordsetObj`，RTL 那边是 memberobj 的真 IDispatch（成员名表
    // vb6forms_memberobj.c:136-137）。从前那套"透传形态直译"认的是 `vb6_Data_Self(` 前缀，
    // 而这个前缀今天没有任何发射者 ⇒ 判据恒假，直译分支一条都进不去（九形探针量过，
    // 见台账 §B124）。⇒ 撤掉，别再给死码立表。dataVars_ 留着：控件自己那一档
    // (dataNameOfExpr) 还在用。
    std::unordered_set<std::string> dataVars_;
    std::string dataNameOfExpr(const std::string& objExpr) const {
        if (objExpr.size() > 9 && objExpr.compare(0, 9, "vb6_hwnd_") == 0) {
            std::string n = objExpr.substr(9);
            if (dataVars_.count(n)) return n;
        }
        return "";
    }
    // C29-OLE: 真窗口的 OLE 容器控件 (登记裸控件名, 如 "OLE1")。
    // 方法 (CreateEmbed/CreateLink/DoVerb/Close/InsertObjDlg/SaveToFile/ReadFromFile/
    // Copy/Paste) 与属性读写全走 vb6_OleCon_*, 第一参是 HWND 槽 vb6_hwnd_X。
    std::unordered_set<std::string> oleConVars_;
    // objExpr 是否落在一个已登记的 OLE 容器上 (返回裸控件名, 否则空串)。
    std::string oleConNameOfExpr(const std::string& objExpr) const {
        if (objExpr.size() > 9 && objExpr.compare(0, 9, "vb6_hwnd_") == 0) {
            std::string n = objExpr.substr(9);
            if (oleConVars_.count(n)) return n;
        }
        return "";
    }
    // objExpr 是否落在一个已登记的 ImageList 上。命中返回**裸控件名** (如 "ImageList1"),
    // 否则返回空串。objExpr 有三种形态, 三种都要认:
    //   ① 控件槽本身                       → vb6_com_ImageList1
    //   ② .ListImages 集合                 → vb6_ComGetObjectProp(vb6_com_X, L"ListImages")
    //   ③ ListImages(i) 的结果              → vb6_ComCallObject(②…, L"Item", (void*[]){…}, 1)
    std::string imageListNameOfExpr(const std::string& objExpr) const {
        const std::string head = "vb6_ComGetObjectProp(vb6_com_";
        size_t p = objExpr.find(head);
        if (p != std::string::npos) {
            size_t q = objExpr.find(", L\"ListImages\")", p);
            if (q != std::string::npos) {
                std::string n = objExpr.substr(p + head.size(), q - (p + head.size()));
                if (imageListVars_.count(n)) return n;
            }
        }
        // C29-3: 两个新形态 ——
        //   ④ 裸槽 `vb6_com_X`: MemberAccess 的 object 就是控件标识符时的形态
        //      (cgen_expr_member_form_builtin.inc:350 设的就是它)。上面 ② 是"已经被
        //      包成 ComGetObjectProp"之后的形态, 两条路都会来, 所以两个都要认。
        //   ⑤ 真集合对象 `vb6_ImageList_ListImages((void*)vb6_com_X)` —— 本批新增。
        // ⑤ 的串里**含** ④ 的前缀, 所以先判 ⑤ 再判 ④。
        {
            static const std::string kHead = "vb6_ImageList_ListImages((void*)vb6_com_";
            size_t p2 = objExpr.find(kHead);
            if (p2 != std::string::npos) {
                size_t s = p2 + kHead.size();
                size_t e = s;
                while (e < objExpr.size()
                       && (isalnum((unsigned char)objExpr[e]) || objExpr[e] == '_')) e++;
                std::string n = objExpr.substr(s, e - s);
                if (imageListVars_.count(n)) return n;
            }
        }
        if (objExpr.size() > 8 && objExpr.compare(0, 8, "vb6_com_") == 0) {
            std::string n = objExpr.substr(8);
            if (imageListVars_.count(n)) return n;
        }
        if (imageListVars_.count(objExpr)) return objExpr;
        return "";
    }
    // 同一件事的"槽变量"形态 (RTL 各函数第一参要的是这个)。
    std::string imageListSlotVarOfExpr(const std::string& objExpr) const {
        std::string n = imageListNameOfExpr(objExpr);
        return n.empty() ? "" : "vb6_com_" + n;
    }

    // P20-40: 原生复刻的 StatusBar 控件 (形如 StatusBar1)。与 ImageList 同构, 但
    // **RTL 函数第一参要的是 HWND 槽 vb6_hwnd_X** (面板状态挂在窗口属性上), 不是
    // vb6_com_X。集合表达式的三种形态:
    //   ① 控件槽本身                 → vb6_hwnd_StatusBar1
    //   ② .Panels 集合               → vb6_ComGetObjectProp(vb6_com_X, L"Panels")
    //   ③ Panels(i) 的结果           → vb6_ComCallObject(②…, L"Item", (void*[]){…}, 1)
    // ②③ 只要抠出**裸控件名**就够了 —— 面板函数都只认"HWND + 1 基下标",
    // 下标另从 Item 的实参里抠 (与 ImageList 分支同一套手法)。
    std::unordered_set<std::string> statusBarVars_;
    std::string statusBarNameOfExpr(const std::string& objExpr) const {
        // 集合形态有三种写法, 槽前缀各不相同但结构一致:
        //   vb6_ComGetObjectProp(vb6_com_X, L"Panels")     OCX 占位期 / ImageList
        //   vb6_ComCallObject(vb6_hwnd_X, L"Panels", …)     真窗口期链式下标
        // 都按 `vb6_<X>` 抠, 再用 statusBarVars_ 核对是不是真 StatusBar
        // (vb6_ComCallObject(vb6_hwnd_Form1, L"Controls") 会抠出 Form1, 核对刷掉)。
        // 账 #299 起还有第三种写法: 集合本体由 memberobj 的真 IDispatch 答
        //   `vb6_StatusBar_Panels((void*)vb6_hwnd_X)` (以及它外面套的 Item 那一层)。
        // 这一条**必须**跟着拦子一起加: 不认新头, 链式那 8 条成员与 `Panels.Add/Remove`
        // 的直译就全都判不出"接收者是 StatusBar", 一起打回晚绑定 —— 那是行为改动, 不是收口。
        static const char* kHeads[] = { "vb6_ComGetObjectProp(vb6_", "vb6_ComCallObject(vb6_",
                                        "vb6_StatusBar_Panels((void*)vb6_" };
        for (const char* h : kHeads) {
            size_t p = objExpr.find(h);
            while (p != std::string::npos) {
                size_t s = p + strlen(h);
                size_t e = s;
                while (e < objExpr.size()
                       && (isalnum((unsigned char)objExpr[e]) || objExpr[e] == '_')) e++;
                // 到标识符边界为止 —— 中间那层可能是 vb6_hwnd_X / vb6_com_X, 前缀
                // 与裸控件名都吃进来了, 按空白切会连 "hwnd_" 一起留下 (比对必败)。
                std::string n = objExpr.substr(s, e - s);
                if (statusBarVars_.count(n)) return n;
                // 槽前缀不止一种: 真窗口期是 `vb6_hwnd_X`, OCX 占位期是 `vb6_com_X`。
                // 上面只吃掉了中间的 vb6_, 这里再剥一层 hwnd_/com_ 还原裸控件名。
                static const char* kPrefixes[] = { "hwnd_", "com_", "arr_" };
                for (const char* pf : kPrefixes) {
                    size_t pl = strlen(pf);
                    if (n.size() > pl && n.compare(0, pl, pf) == 0) {
                        n = n.substr(pl);
                        break;
                    }
                }
                if (statusBarVars_.count(n)) return n;
                p = objExpr.find(h, s);
            }
        }
        if (statusBarVars_.count(objExpr)) return objExpr;
        // 2026-09-30 (门 #233 / c29sbevt): **宿主槽本身**也是这个函数的输入之一。
        // 集合形态 (`Panels.Add` / `Panels.Count`) 一直走上面那两段; 而 `StatusBar1.SimClick`
        // 这类**直接落在宿主槽上**的成员调用送进来的是 `vb6_hwnd_StatusBar1` —— 以前这里只有
        // `statusBarVars_.count(objExpr)`(拿带前缀的槽名去比裸控件名, 恒假) ⇒ 函数答空
        // ⇒ 宿主被当成"不是 StatusBar", 成员被当 COM 占位属性读 (`SimClick` 于是编成
        // `ComCallObject(ComGetObjectProp(hwnd, L"SimClick"), L"Item", …)`: 编得过、跑起来
        // 一条事件都不发)。姊妹函数 `treeViewNameOfExpr` / `toolbarNameOfExpr` 是拿
        // `vb6_hwnd_<裸名>` 扫子串的, 所以它们认得这一形 —— 只有 StatusBar 这一路漏了。
        // 这里补的正是"裸槽"那一形, 判据仍是登记表 (不是按前缀猜): 剥掉 hwnd_/com_/arr_
        // 前缀后必须真的在 `statusBarVars_` 里。
        {
            static const char* kSlotPrefixes[] = { "vb6_hwnd_", "vb6_com_", "vb6_arr_" };
            for (const char* pf : kSlotPrefixes) {
                size_t pl = strlen(pf);
                if (objExpr.size() > pl && objExpr.compare(0, pl, pf) == 0) {
                    std::string bare = objExpr.substr(pl);
                    // 只有"整串就是那个槽"才算宿主 (带成员/下标的走集合那两段)
                    if (bare.find_first_of("[(. ") == std::string::npos
                        && statusBarVars_.count(bare)) {
                        return bare;
                    }
                }
            }
        }
        return "";
    }
    // C29-5a: Toolbar 的裸名登记 (设计期发射那侧 insert)。认法比 StatusBar 那套窄得多 ——
    // Toolbar 已经有真窗口，集合表达式的槽一定是 `vb6_hwnd_<裸名>`，不必再抠 vb6_com_ 前缀。
    // 名字比对要过一道**标识符边界**检查: 只 find 子串的话 `vb6_hwnd_tb1` 会命中
    // `vb6_hwnd_tb10` 的表达式 (两枚工具栏同窗体就翻车)。
    std::unordered_set<std::string> toolbarVars_;
    std::string toolbarNameOfExpr(const std::string& objExpr) const {
        for (const auto& n : toolbarVars_) {
            const std::string hwnd = "vb6_hwnd_" + n;
            size_t p0 = objExpr.find(hwnd);
            while (p0 != std::string::npos) {
                size_t e = p0 + hwnd.size();
                if (e >= objExpr.size() || !(isalnum((unsigned char)objExpr[e]) || objExpr[e] == '_'))
                    return n;
                p0 = objExpr.find(hwnd, e);
            }
        }
        if (toolbarVars_.count(objExpr)) return objExpr;
        return "";
    }
    // C29-8b: TreeView 的裸名登记 (emitFormFramework 开头统一插, 含容器子控件)。
    // 认法照 toolbarNameOfExpr: 真窗口 ⇒ 集合表达式的槽只会是 `vb6_hwnd_<裸名>`,
    // 并且比对要过一道标识符边界 (tv1 不能命中 tv10)。
    std::unordered_set<std::string> treeViewVars_;
    std::string treeViewNameOfExpr(const std::string& objExpr) const {
        for (const auto& n : treeViewVars_) {
            const std::string hwnd = "vb6_hwnd_" + n;
            size_t p0 = objExpr.find(hwnd);
            while (p0 != std::string::npos) {
                size_t e = p0 + hwnd.size();
                if (e >= objExpr.size() || !(isalnum((unsigned char)objExpr[e]) || objExpr[e] == '_'))
                    return n;
                p0 = objExpr.find(hwnd, e);
            }
        }
        if (treeViewVars_.count(objExpr)) return objExpr;
        return "";
    }
    // Fix <vbeclipse>: 跨窗体 TreeView。`frmViewViews.tvwViews` 解析成控件句柄
    // 访问器 `vb6_form_ctrl_<Form>_<Ctrl>()` 时在这里登记 (键=值=访问器文本),
    // 供 Nodes 集合两条路 (resolveComMember / resolveComValue) 认领。
    mutable std::unordered_map<std::string, std::string> externalTreeViewAccs_;
    // <vbeclipse>: 控件的**集合对象表达式** —— 它的值已经是那个集合本身了 (真 IDispatch
    // 或原生复刻的集合)，对它点成员是**方法调用** (`Nodes.Add(…)` 调的就是 Add)，不是
    // "再取一次集合属性 + Item 下标"。
    // 判据一律按**开头**认，绝不按子串：集合表达式里嵌着宿主槽 `vb6_hwnd_tv1`，而
    // treeViewNameOfExpr / toolbarNameOfExpr / isListViewHostExpr 都是子串扫的 —— 按子串
    // 判"这是不是宿主槽"会把集合对象也算成宿主槽，于是 `Nodes.Add(…)` 编成
    // vb6_ComCallObject(vb6_ComGetObjectProp(vb6_TreeView_Nodes(…), L"Add"), L"Item", …)：
    // 节点根本没被加进去，Index 全部读空。实测 12 条控件回读红
    // (ctrltreeview / ctrltoolbar / c29listview / c29sbevt 各 x64+x86)。
    // dev 那边是靠 `comObjExpr_.find("vb6_hwnd_") == 0` 这个**前缀锚**挡住的；本条把
    // 这件事写成一处有名字的话，免得下一个控件再加一路子串判据时又把它冲掉。
    bool isControlCollectionExpr(const std::string& expr) const {
        static const std::vector<std::string> heads = {
            "vb6_TreeView_Nodes(", "vb6_Toolbar_Buttons(", "vb6_ImageList_ListImages(",
            "vb6_ListView_ListItems(", "vb6_ListView_ColumnHeaders(",
            // 账 #299: StatusBar 的 Panels 集合对象 (与上面几族同一个 memberobj 机制)。
            "vb6_StatusBar_Panels(",
        };
        for (const auto& h : heads) {
            if (expr.size() >= h.size() && expr.compare(0, h.size(), h) == 0) return true;
        }
        return false;
    }

    // <vbeclipse> 账 #299 第二十一刀: "这枚成员集合的创建式是什么" —— 唯一一处答案。
    // 这张表以前在**三处**各写一遍 (resolveComValue / resolveComMarkerForPack /
    // cgen_expr_call_com_bind 的默认成员下标那一段)，而且每一处各缺几族: Panels 三处全缺
    // ⇒ `Set po = SB.Panels(1)` 与 `With SB.Panels(1)` 的头一直发成
    // `vb6_ComGetObjectProp(<HWND>, L"Panels")` = 拿 HWND 当 IDispatch 用 (一条读数都交不出)，
    // 第二十刀只补了 resolveComValue 那一处 ⇒ 四形里只有 `For Each` 活了过来。
    // 三条发射路的**返回语义**各不相同 (改 lastExpr_+isComMarker_ / 直接 return /
    // 拼完还要再套一层 Item)，所以这里只交**串**，语义留在调用点。空串 = 不是已知的成员集合。
    // 宿主槽前缀按族走各自的事实: ImageList 无窗口 ⇒ vb6_com_X 是实例指针；
    // TreeView/Toolbar/StatusBar 是真窗口 ⇒ vb6_hwnd_X；ListView 用 listViewHwndExprOf
    // (它还认 With 槽里那个裸 HWND 变量)；跨窗体 TreeView 查 externalTreeViewAccs_。
    std::string memberCollectionObjectExpr(const std::string& hostExpr,
                                           const std::string& memberName) const {
        std::string mem = memberName;
        for (char& ch : mem) ch = (char)tolower((unsigned char)ch);
        if (mem == "nodes") {                                   // C29-8b
            std::string bare = treeViewNameOfExpr(hostExpr);
            if (!bare.empty()) return "vb6_TreeView_Nodes((void*)vb6_hwnd_" + bare + ")";
            auto itExt = externalTreeViewAccs_.find(hostExpr);   // 跨窗体: 槽是访问器文本
            if (itExt != externalTreeViewAccs_.end())
                return "vb6_TreeView_Nodes((void*)" + itExt->second + ")";
            return "";
        }
        if (mem == "buttons") {                                 // C29-5b
            std::string bare = toolbarNameOfExpr(hostExpr);
            return bare.empty() ? "" : "vb6_Toolbar_Buttons((void*)vb6_hwnd_" + bare + ")";
        }
        if (mem == "panels") {                                  // 账 #299
            std::string bare = statusBarNameOfExpr(hostExpr);
            return bare.empty() ? "" : "vb6_StatusBar_Panels((void*)vb6_hwnd_" + bare + ")";
        }
        if (mem == "listimages") {                               // C29-3
            std::string bare = imageListNameOfExpr(hostExpr);
            return bare.empty() ? "" : "vb6_ImageList_ListImages((void*)vb6_com_" + bare + ")";
        }
        if (mem == "listitems" || mem == "columnheaders") {      // C29-7
            std::string lvHwnd = listViewHwndExprOf(hostExpr);
            if (lvHwnd.empty()) return "";
            return (mem == "listitems") ? ("vb6_ListView_ListItems((void*)" + lvHwnd + ")")
                                        : ("vb6_ListView_ColumnHeaders((void*)" + lvHwnd + ")");
        }
        return "";
    }
    // C29-DT-c: DTPicker 的裸名登记。认法照 treeViewNameOfExpr —— 它也是真窗口，
    // 判据方法 (SimChange/SimDropDown/SimCloseUp) 的宿主槽只会是 `vb6_hwnd_<裸名>`，
    // 名字比对同样要过标识符边界 (dt1 不能命中 dt10)。
    std::unordered_set<std::string> dtpickerVars_;
    std::string dtpickerNameOfExpr(const std::string& objExpr) const {
        for (const auto& n : dtpickerVars_) {
            const std::string hwnd = "vb6_hwnd_" + n;
            size_t p0 = objExpr.find(hwnd);
            while (p0 != std::string::npos) {
                size_t e = p0 + hwnd.size();
                if (e >= objExpr.size() || !(isalnum((unsigned char)objExpr[e]) || objExpr[e] == '_'))
                    return n;
                p0 = objExpr.find(hwnd, e);
            }
        }
        if (dtpickerVars_.count(objExpr)) return objExpr;
        return "";
    }
    // C29-MV-c: MonthView 的裸名登记（判据方法 SimDateClick 靠它认宿主是真窗口槽）。
    std::unordered_set<std::string> monthviewVars_;
    std::string monthviewNameOfExpr(const std::string& objExpr) const {
        for (const auto& n : monthviewVars_) {
            const std::string hwnd = "vb6_hwnd_" + n;
            size_t p0 = objExpr.find(hwnd);
            while (p0 != std::string::npos) {
                size_t e = p0 + hwnd.size();
                if (e >= objExpr.size() || !(isalnum((unsigned char)objExpr[e]) || objExpr[e] == '_'))
                    return n;
                p0 = objExpr.find(hwnd, e);
            }
        }
        if (monthviewVars_.count(objExpr)) return objExpr;
        return "";
    }
    // C29-SL-c: Slider 的裸名登记（判据方法 SimNotify 靠它认宿主是真窗口槽，同 DT/MV 那两格）。
    std::unordered_set<std::string> sliderVars_;
    std::string sliderNameOfExpr(const std::string& objExpr) const {
        for (const auto& n : sliderVars_) {
            const std::string hwnd = "vb6_hwnd_" + n;
            size_t p0 = objExpr.find(hwnd);
            while (p0 != std::string::npos) {
                size_t e = p0 + hwnd.size();
                if (e >= objExpr.size() || !(isalnum((unsigned char)objExpr[e]) || objExpr[e] == '_'))
                    return n;
                p0 = objExpr.find(hwnd, e);
            }
        }
        if (sliderVars_.count(objExpr)) return objExpr;
        return "";
    }
    std::string statusBarHwndVarOfExpr(const std::string& objExpr) const {
        std::string n = statusBarNameOfExpr(objExpr);
        return n.empty() ? "" : "vb6_hwnd_" + n;
    }
    // P20-40: `StatusBar1.Panels(i).Prop = v` 的**写**目标。读那边在 resolveComValue
    // 里改道了, 写这边不拦就会落到 vb6_ComSetProp(<HWND>, L"Text", …) —— 对假
    // IDispatch (HWND) 调 PutProp, 链接都过、运行静默无效。返回 RTL 写入函数名,
    // hwndOut/idxOut 出参拿槽与 1 基下标; 空串 = 不是 Panel 链式写。
    std::string statusBarPanelWriteFn(const std::string& targetExpr,
                                      const std::string& memLower,
                                      std::string* hwndOut, std::string* idxOut) const {
        static const struct { const char* vb; const char* fn; } kSbWriters[] = {
            { "text",       "vb6_StatusBar_SetPanelText"     },
            { "key",        "vb6_StatusBar_SetPanelKey"      },
            { "width",      "vb6_StatusBar_SetPanelWidth"    },
            { "minwidth",   "vb6_StatusBar_SetPanelMinWidth" },
            { "autosize",   "vb6_StatusBar_SetPanelAutoSize" },
            { "style",      "vb6_StatusBar_SetPanelStyle"    },
            { "tooltiptext","vb6_StatusBar_SetPanelToolTip"  },
        };
        std::string hwnd = statusBarHwndVarOfExpr(targetExpr);
        if (hwnd.empty()) return "";
        std::string fn;
        for (const auto& w : kSbWriters) {
            if (memLower == w.vb) { fn = w.fn; break; }
        }
        if (fn.empty()) return "";
        // 抠出面板下标。这里没有 `L"Item"` 那一层 —— 生成器把 Panels 当**默认成员**,
        // 索引直接跟着它那一层的 `(void*[]){…}` 走 (形如
        // `vb6_ComCallObject(vb6_hwnd_StatusBar1, L"Panels", (void*[]){vb6_ComPackInt(1)}, 1)`),
        // 所以按参数包抠, 剥掉 vb6_ComPackInt(...) 只留数字。抠不到就返回空:
        // 宁可什么都别发, 也别发个下标为空的坏调用。
        std::string idx;
        const std::string kArr = "(void*[]){";
        size_t ia = targetExpr.find(kArr);
        if (ia != std::string::npos) {
            size_t open = ia + kArr.size();
            size_t close = targetExpr.find('}', open);
            if (close != std::string::npos) {
                std::string argTxt = targetExpr.substr(open, close - open);
                size_t lp = argTxt.find('(');
                if (lp != std::string::npos) {
                    size_t end = argTxt.find_first_of("),", lp + 1);
                    idx = (end == std::string::npos) ? argTxt.substr(lp + 1)
                                                     : argTxt.substr(lp + 1, end - lp - 1);
                } else {
                    idx = argTxt;
                }
            }
        }
        if (idx.empty()) return "";
        if (hwndOut) *hwndOut = hwnd;
        if (idxOut)  *idxOut  = idx;
        return fn;
    }

    // P6.6: ActiveX DLL模式
    bool isDll_ = false;                    // 编译为ActiveX DLL
    std::string dllProgId_;

    // Fix 110: true = .ctl/.pag designer sub-control module (has designer desc, not .frm)
    bool isDesignerModule_ = false;
    // Fix 110f: true = PropertyPage (.pag), false = UserControl (.ctl)
    bool isPropertyPageDesigner_ = false;
    // 账 #179: 只有 UserControl 的实例方法要成对换入宿主上下文 (体首 Push / 统一出口尾 Pop)。
    // 谓词写在这一处, 三个发射点 (func/proc/prop) 共用 —— 别再各抄一遍条件。
    bool ucCtxScoped() const { return isDesignerModule_ && !isPropertyPageDesigner_; }

    // P6.4+: interface -> single implementer map (injected by driver)
    const std::map<std::string, std::string>* interfaceImplementers_ = nullptr;

    // Fix <vbeclipse>: interface -> **全部**实现类 (小写接口名 → 实现类规范名列表),
    // driver 侧预扫描提供。接口成员桩靠它把调用转进实现类的槽表 —— 单实现映射
    // (interfaceImplementers_) 在多实现接口上是空的, 不能用于分派。
    const std::map<std::string, std::vector<std::string>>* interfaceAllImplementers_ = nullptr;

    // 账 #305 新账 #464: legacy 接口（没进 tB `IfaceRegistry` 的那批 .cls 宿主）的**形状表**，
    // driver 预扫描建、发码期只读。键 = 小写接口名；值 = 规范拼写 + 按序的可实现成员声明。
    const LegacyIfaceShapeMap* legacyIfaceShapes_ = nullptr;

    // Interface 契约 (tB, ai/022 B04): stage 2.7 建好的工程级只读登记表 (driver 注入).
    // 槽表/适配器/薄指针类型都由它驱动, 与 legacy 的 implementsNames 口径互不影响.
    const IfaceRegistry* ivreg_ = nullptr;

    // 委托式实现 `Implements I Via m_h` (tB, ai/022 B10): stage 2.7 Pass D 的只读裁决表.
    const ViaRegistry* ivviareg_ = nullptr;

    // 类继承 (tB, ai/022 B07b): stage 2.8 建好的工程级只读链登记表 (driver 注入).
    // 继承字段与前缀布局、转发桩都由 view 的 inhFields/inhProcs 驱动 (stage 3.4 回填).
    const ClassChainRegistry* clsreg_ = nullptr;

    // 工程级类名表 (小写类名 → 原大小写类名, driver 注入). 本模块符号表会被同名成员
    // 遮蔽, 判"某类型名是不是工程类"只能问这张表 (详见 setProjectClassNames 注释).
    const std::unordered_map<std::string, std::string>* projClassNames_ = nullptr;

    // 账 #305 §B187 第二腿: 工程级「在运行期 coclass 描述表里占一行」的类名名单
    // (小写, driver 与建表那侧同一条判据扫出 —— 见 driver_codegen_voidfield_scan.inc)。
    // 装箱腿发每一句 vb6_FindCoClassDesc("<类>") 之前先问它: 表里没有这行 ⇒ desc=NULL
    // ⇒ 包装器=NULL ⇒ 槽里成了 Nothing, 比裸指针还差, 所以那种形状一律不装箱。
    const std::unordered_set<std::string>* projCoclassNames_ = nullptr;

    // 账 #305 §B189 第二档: 工程级「被谁 Implements 过」的名单 (小写, driver 从语义层那份
    // Symbol::implementsNames 注入)。 TypeName 按声明档折叠那一刀用它挡掉接口视图 (注释见
    // driver_codegen_voidfield_scan.inc)。
    const std::unordered_set<std::string>* projImplementedNames_ = nullptr;

    // 账 #305 第320刀 (账 #592): .vbp 的 Title（空则 Name）—— 生成的入口点把它播种进
    // RTL 那一格运行期真值（`vb6_SetAppTitle(L"...")`），MsgBox 的默认题头与
    // `App.Title` 读取都只问那一格。单文件编译（无 .vbp）时为空 ⇒ 入口点一行不发。
    std::string projectTitle_;

    // CoClass 身份 (tB, ai/022 B13b): stage 2.7 Pass E 求出的唯一出口 (coclassIds_)。
    // 只有 generateDllEntry 读它 —— 产品里的 CLSID / 默认接口 IID / 组名 ProgID 一律
    // 从这里取，prelude 里那两枚 lambda 退成"读不到块的类"的兜底路径。
    const std::unordered_map<std::string, CoClassIdentity>* ccreg_ = nullptr;

    // 接口 IID 表 (tB, ai/022 B13c): stage 2.7 Pass E 由 `buildIfaceIdMap` 建的只读表，
    // key = 接口名小写。新式接口的 IID 三处消费者（vtable QI / dll_entry 的 IID 表 /
    // 类型库）一律读它；表里没有的接口保持各自 legacy 派生（存量产物不变）。
    const std::unordered_map<std::string, std::string>* iidreg_ = nullptr;

    // P6.4+: class-module default instance (VB_PredeclaredId=True) lowercase name -> canonical
    const std::map<std::string, std::string>* defaultInstanceClasses_ = nullptr;

    // Fix 112: in-project UserControl (.ctl) designer spec table (lowercase module -> FrmFile)
    // czUI fix: UC 事件 attach 行 (文件作用域 thunk/sink 在 prelude 生成)
    // 账 #222: 键从"控件名"改成"控件名 + 元素号" —— 控件数组的 24 枚元素共用一枚处理器,
    // 但每枚要自己的 thunk 与自己的 sink (元素号得交得回去), 按名字存就只剩最后一枚。
    // sink 变量名一起存: 挂接那侧以前把 "_sink_" + 名字**再拼一遍**, 那是第二处拼名。
    struct UcEventAttach {
        std::vector<std::string> lines;
        std::string sinkVar;
    };
    std::map<std::string, UcEventAttach> ucEventAttach_;
    // 账 #222: 同一趟登记顺手记下"sink 供了哪一枚事件" (键 = ctrlElemKey + "." + 事件小写名)。
    // UC 的 Click 走 RaiseEvent 到 sink 到 thunk, 子类化那条 WM_LBUTTONUP 臂再发一遍就是
    // 一次点击两声 (实测 ucProgressCircular 的 Form1.c 两路都有)。
    std::set<std::string> ucSinkEvents_;

    // Fix <vbeclipse> rev17: 本窗体里**充当 WithEvents 事件处理器**的过程名 (小写)。
    // 由 cgen_form_prelude 在扫 <ctrl>_<Event> 命中时登记。用途: 事件处理器是"跨对象
    // 边界"的回调, 形参一律按**事件 ABI = ByVal** 发 (与 emitEventSink 的
    // `pi.isByVal = true` 同一条规矩), 否则处理器声明出 `void**` 而回调 typedef /
    // RaiseEvent 发送侧给的是 `void*` —— 处理器多解一层引用 → 拿到 NULL。
    // 实证 play78: `Private Sub ucPerspective1_OpenEditor(Editor As Object)` 发成
    // `void** Editor` 且体里 `(*Editor)`, 而 typedef 是 `void* Editor`、
    // 发送侧是 `onOpenEditor(handler, (*Editor))` ⇒ Editor.Caption 取到 NULL
    // ⇒ av read 0x0 (_vb6_frmMain_ucPerspective1_OpenEditor+0x19)。
    std::set<std::string> ucEventHandlers_;

    // 账 #305 第二百五十六刀: 处理器名(小写) -> **每一位形参**要不要保住 ByRef。
    // 上面那条"一律 ByVal"对**没写 ByRef** 的形参是对的 (play78 那条实证就在这一档)，
    // 但事件声明里明写 `ByRef Text As String` 的那几位是**出参**，翻成 ByVal 就把写回丢了。
    // 由 cgen_form_prelude 经 prepareEventHandlerProc 登记 (那里才看得见事件的形参表)，
    // applyEventHandlerAbi 读 —— 声明侧与定义侧同读这一张表，与上面那条同源。
    std::map<std::string, std::vector<char>> ucEventHandlerByAddr_;

    // Fix <vbeclipse> rev18: 本模块发出去的 UserControl 自有属性桥条数。
    // emitFormFramework 末尾 (cgen_form_uc_props.inc) 算表时写入, emitDesignerControlDecls
    // 填 vb6_UserControlDesc.props/propCount 时读 (同一次 decl-pass 内, 前者先跑 ——
    // 见 cgen_base_generate_decl_pass.inc:40 与 :46)。
    int32_t ucHostPropCount_ = 0;

    // Fix <vbeclipse> rev21: 本模块发出去的 UserControl 自有 **Public Sub 方法桥**条数。
    // 与 ucHostPropCount_ 同一对生命周期: cgen_form_uc_methods.inc 算表时写入,
    // emitDesignerControlDecls 填 vb6_UserControlDesc.methods/methodCount 时读。
    int32_t ucHostMethodCount_ = 0;
    const std::map<std::string, FrmFile>* designerFiles_ = nullptr;

    // Fix 143: vbp Object= 的 OCX 文件表 (CLSID 小写去花括号 → ocx 绝对路径)
    const std::map<std::string, std::string>* ocxFiles_ = nullptr;

    // Fix 143b: vbp Object= 原始引用列表 (CLSID 小写去花括号, ocx 绝对路径)
    const std::vector<std::pair<std::string, std::string>>* ocxRefs_ = nullptr;

    // Fix 160: vbp ComLib= 免注册 COM 组件表 {ProgID, CLSID, coclass名, 相对exe路径}
    const std::vector<std::array<std::string, 4>>* comLibRefs_ = nullptr;

    // Fix <vbeclipse> 2026-10-06: vbp Object= OCX 免注册表 {ProgID, CLSID, coclass名, 相对exe路径}
    const std::vector<std::array<std::string, 4>>* ocxLibRefs_ = nullptr;

    // Fix 112: sub-controls on a form whose type is an in-project UserControl
    std::unordered_map<std::string, std::string> knownUserControlCtrlVars_;
                  // DLL的ProgID前缀


    // Delegate (tB 扩展): decl pass 登记的桩定义文本, epilogue 统一发射
    std::vector<std::string> delegateThunkDefs_;

    // AddressOf 回调桩 (账 #184): 同上, 与委托桩同批后置发射
    std::vector<std::string> addressOfThunkDefs_;

    // 重载调用变体 (O2): IndexOrCallExpr 发射 callee ident 前置入的符号键后缀,
    // 由 ident 发射路径消费并清空 ("" = 裸键 head, 产出不受影响)
    std::string pendingCalleeOvl_;

    // ai/vb-asm-extension-spec: Asm 块过程收集 + 目标架构
    //   函数体 = 单个 AsmStmt 时, 不发 C 体, 改发 extern 原型并把元数据记在这里,
    //   由 driver 生成独立 MASM 过程 (ml64) 参与链接。
    //   混排 (项2): 逐条 AsmStmt 就地降级, 也往这里追加 (每片段一个独立过程)。
    std::vector<AsmProcInfo> asmProcs_;
    std::string targetArch_ = "x64";
    // 混排 (项2): 每个过程体清零; 给降级出来的 Asm 辅助过程/标签起唯一序号。
    int asmMixedBlockCounter_ = 0;
