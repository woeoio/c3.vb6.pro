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

    // Fix 037b: 标记最近一次 IndexOrCallExpr 生成的表达式是否为返回 VARIANT 的
    // 项目类 Item 属性调用 — Set 语句需要将其转换为 void* 对象引用.
    bool lastExprNeedsObjectUnpack_ = false;

    // Fix 010o: 过程局部变量名集合 (小写) — Dim声明的局部变量 + For/ForEach循环变量
    // 用于在IdentifierExpr中避免对局部变量错误添加 me-> 前缀
    std::unordered_set<std::string> knownLocalVars_;
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
    int ansiCounter_ = 0;
    int vcmpCounter_ = 0;  // P25: Variant比较临时变量计数器

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

    // M22: 外部模块名称集合 (用于跨模块变量解析时判断模块是否已#include)
    std::unordered_set<std::string> externalModules_;

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
    std::vector<std::string> ivrefLocalsToRelease_;
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
    // Recordset 链 (Data1.Recordset.Fields("x").Value / .MoveNext / .BOF) 全直译 RTL。
    std::unordered_set<std::string> dataVars_;
    // 从 vb6_Data_Self(vb6_hwnd_X) / vb6_Data_Self((void*)vb6_hwnd_X) 抠出裸槽名。
    // ⚠ 控件属性分支发的形态**不带** (void*) 强转, 两种都要认。
    std::string dataSelfHwndExpr(const std::string& objExpr) const {
        const char* kHeads[] = { "vb6_Data_Self((void*)", "vb6_Data_Self(" };
        for (const char* kHead : kHeads) {
            size_t p = objExpr.find(kHead);
            if (p == std::string::npos) continue;
            size_t s = p + strlen(kHead);
            if (s > 0 && objExpr.compare(s, 9, "vb6_hwnd_") != 0 && s > 0) {
                // 允许中间有空白
                while (s < objExpr.size() && objExpr[s] == ' ') s++;
            }
            size_t e = s;
            while (e < objExpr.size() && (isalnum((unsigned char)objExpr[e]) || objExpr[e] == '_')) e++;
            std::string n = objExpr.substr(s, e - s);
            // ⚠ 扫出来的是完整槽名 "vb6_hwnd_Data1" —— 剥掉前缀再查表
            if (n.size() > 9 && n.compare(0, 9, "vb6_hwnd_") == 0) n = n.substr(9);
            if (dataVars_.count(n)) return "vb6_hwnd_" + n;
        }
        return "";
    }
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
        static const char* kHeads[] = { "vb6_ComGetObjectProp(vb6_", "vb6_ComCallObject(vb6_" };
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
        };
        for (const auto& h : heads) {
            if (expr.size() >= h.size() && expr.compare(0, h.size(), h) == 0) return true;
        }
        return false;
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
