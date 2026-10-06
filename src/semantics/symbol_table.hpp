#pragma once
// VB6符号表 - 作用域层次 + 符号查找
// 支持: 模块级/过程级/块级作用域, VB6不区分大小写

#include "common/types.hpp"
#include "common/diagnostics.hpp"
#include <algorithm>
#include <map>
#include <string>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <memory>
#include <optional>

namespace vb6c3 {

// ============================================================
// 084a: 类成员访问级别预计算表 (ai/084a M1)
// ============================================================
// 为什么预计算而不是 visit 期查符号表: 跨模块类符号 (extSym) 的成员表要到
// stage 3.5 注入才有, 当前模块分析时 lookup 查无成员 —— 与 023 S03 预计算
// 屏蔽表同因同解。driver 在 runSemanticAnalysis 入口处扫一遍全部类模块的
// AST 声明建表, 之后只读, 经 SemanticAnalyzer::setMemberAccessTable 下发。
struct MemberAccessEntry {
    AccessLevel level = AccessLevel::Public;
    std::string definingModuleLower;  // 声明成员的类模块名 (小写)
    // 字段与过程在继承合并下不对称 (tB B07b): 祖先 Private **字段**仍并入派生类
    // (inhFields 无 Private 剔除), Private **过程**被剔除 (转发桩只收非 Private)。
    // 故家族内放行只对字段成立; Private 过程在派生类里必须语义期拦下 —— 否则
    // 语义放行、发码期 LNK2019, 比报错糟一档。
    bool isField = false;
    // ai/084a M3: 定义类所属包 (小写; 空 = 宿主工程模块)。Friend 成员的跨包
    // 裁决依据: 消费方包 != 定义包 且 !pkgFriendOpen → 7008。
    std::string definingPkg;
    bool pkgFriendOpen = false;  // 定义包清单 [Export] Friend=True
};
// key1 = 类名小写, key2 = 成员名小写。只含工程内类模块声明的成员 (白名单):
// 查不到 = 不裁决, 维持旧行为, 避免屏蔽表过宽误杀。
using MemberAccessTable = std::map<std::string, std::map<std::string, MemberAccessEntry>>;

class TypeSystem;  // 前向声明

// ============================================================
// 符号类别
// ============================================================

enum class SymbolKind : uint8_t {
    Variable,       // Dim/Private/Public/Static 变量
    Constant,       // Const 常量
    Sub,            // Sub 过程
    Function,       // Function 过程
    PropertyGet,    // Property Get
    PropertyLet,    // Property Let
    PropertySet,    // Property Set
    Parameter,      // 过程参数
    UserDefinedType,// Type ... End Type
    EnumType,       // Enum ... End Enum
    EnumMember,     // Enum 成员
    DeclareSub,     // Declare Sub (外部)
    DeclareFunc,    // Declare Function (外部)
    Event,          // Event 声明
    Delegate,       // Delegate 声明 (tB 扩展: 具名函数指针类型)
    Class,          // 类模块 (.cls)
    Label,          // 行标签
    ComClass,       // COM coclass (来自TypeLib, 前期绑定)
    ComInterface,   // COM 接口 (来自TypeLib, 前期绑定)
    ComModule,      // COM 模块 (P24-04: TKIND_MODULE, ActiveX DLL全局函数命名空间)
    ComGlobalNs,    // P24-04: VB_GlobalNameSpace promoted函数 (如 VBMAN.Version()中的VBMAN)
};

// 判断是否为Property类型
inline bool isPropertyKind(SymbolKind k) {
    return k == SymbolKind::PropertyGet ||
           k == SymbolKind::PropertyLet ||
           k == SymbolKind::PropertySet;
}

// ============================================================
// 符号
// ============================================================

struct ParameterInfo {
    std::string name;
    Vb6Type type = Vb6Type::Variant;
    // <vbeclipse>: `As <类型名>` 里的那个**名字**（工程类/接口/CoClass 才有值，内建类型为空）。
    // resolveTypeOrDefault 把工程类形参折成 Vb6Type::Variant —— 类名就在这一步丢了, 于是任何
    // 只拿 Vb6Type 说话的发射器（dll 入口的 extern 原型就是）只能把它写成 vb6_VARIANT,
    // 与类模块自己发的 vb6_cls_X* 定义 ABI 不符：调用点照着 vb6_VARIANT 那份原型去装箱,
    // 直接 C2440（实测 cc_demo/itf_via/cls_inh/modulemethod 一系全在这一条上）。
    // 与 MemberInfo::typeRefName 同思路：类型名要跟着符号走, 不要在下游客串里再造一遍。
    std::string typeRefName;
    bool isByVal = false;
    bool isOptional = false;
    bool isParamArray = false;
    // P14.1.4: Optional参数默认值
    bool hasDefaultValue = false;       // 有显式默认值(Optional = expr)
    std::string defaultValueExpr;       // C表达式字符串(如"10", "vb6_BSTR_FromStr(L\"hello\")", "-1"等)
};

struct Symbol {
    std::string name;           // 原始名称 (保留大小写)
    std::string lowerName;      // 小写名称 (用于查找)
    SymbolKind kind;
    Vb6Type type = Vb6Type::Empty;  // 符号类型

    // 位置
    SourceLocation location;

    // 访问级别
    AccessLevel access = AccessLevel::Public;

    // 过程相关
    std::vector<ParameterInfo> params;
    bool isStatic = false;      // Static Sub/Function
    bool isArray = false;       // 数组变量
    int32_t dimCount = 0;       // 数组维度数 (0=非数组, 1=一维, 2+=多维)

    // 常量值 (仅Constant)
    bool hasConstValue = false;
    int64_t constIntValue = 0;
    double constFloatValue = 0.0;
    std::string constStringValue;
    bool constBoolValue = false;
    Vb6Type constType = Vb6Type::Empty;  // 常量值的实际类型

    // 是否已被引用 (用于未使用变量警告)
    bool isReferenced = false;

    // 被 `AddressOf` 取过址的过程 = Win32/COM 回调 (账 #184)。
    // VB6 的 AddressOf 给出的是 __stdcall 调用桩, 不是本体地址: x86 下本体是
    // __cdecl, 直接把它交给 OS 就等于每次回调少弹 N*4 字节 (实测 VBFlexGrid 的
    // SUBCLASSPROC 起窗即堆损坏 0xC0000374)。cgen 见到这个标记就另发一枚
    // __stdcall 转发桩 (见 cgen_delegate.cpp 的 AddressOf 段), 本体保持 cdecl,
    // 只有桩被取址。x64 下 MSVC 忽略该约定 ⇒ 两架构共用一份生成码。
    // 由 Driver::markAddressOfCallbacks() 在 stage 3.5a 置位 (定义模块与所有
    // 引用它的外部副本一起置, 消费方只读自己表里的那份)。
    bool addressOfCallback = false;

    // 是否为内置符号 (由编译器预注册, 非用户代码)
    bool isBuiltin = false;

    // --- 跨模块符号解析 ---
    // isExternal=true 表示该符号定义在其他模块中（Public符号被当前模块引用）
    bool isExternal = false;
    // sourceModule 记录符号定义所在的模块基名（如 "MathUtils"）
    // 仅当 isExternal=true 时有效
    std::string sourceModule;
    // Fix 010r-11: 变量的声明类型名 (如 "cToolsStr", "ADODB.Connection")
    // 仅对 SymbolKind::Variable 有效 — 当变量声明为 As ClassName 时存储类名
    // 用于跨模块解析时传递类类型信息到 consuming 模块的 cgen
    std::string variableTypeName;
    // tB B08c: 局部变量/参数 `As <类型>` 的**原文类型名** (模块级字段同样记录, 与上一条独立)。
    // 为什么不复用 variableTypeName: 那条已被后端十余处按"非空即类实例"消费 (inferClassTypeOfExpr
    // 的 Fix 084g 分支、cgen_with、comwrite…)，给局部变量补上它就是改**发码**，破本批"发码零改动"
    // 的验收口径。本字段目前只有 Protected 越权判定一个读者。
    std::string srcTypeName;
    // Fix 183: 该变量声明为 As New (模块级 Public As New ClassName).
    // 跨模块注入时必须随 Symbol 复制, 否则消费模块不知道需要惰性实例化,
    // 导致 As New 全局对象运行期恒为 NULL → 解引用 0xC0000005.
    bool isNewVar = false;

    // --- 类相关 (仅SymbolKind::Class) ---
    VBInstancing instancing = VBInstancing::Private;  // Instancing属性
    std::vector<std::string> memberNames;              // 类成员名称列表(方法+属性+事件)
    // Fix 015: 成员返回类型表 — 仅记录 Function/PropertyGet 返回类型名为命名类型
    // (即 returnType 是 SimpleTypeRef 的成员) 的方法. key=成员名小写, value=返回类型名(源码原样).
    // 用于 method chaining: 当跨模块 storageKey 冲突导致消费模块符号表中找不到
    // 目标类的 Function 符号时 (如 cCryptoHMAC.DataString 与 cCryptoHash.DataString 冲突,
    // 只保留首个外部符号), getClassMethodReturnType 可从 Class 符号本身读取返回类型,
    // 从而让 HMAC.Secret(...).DataString(...).ReturnHex(...) 三级链式调用能继续.
    // 注: 与 memberNames 不同, 此表不包含 Let/Set/Variable/Event — 它们无返回值.
    //     同名 Property Get/Let/Set 共用同一个 lowerName 键; 只在 Get 时写入,
    //     Let/Set 会覆盖. 因 Get 先于 Let/Set 处理, 同名共存时记录的是 Get 的返回类型
    //     (后者无返回类型, 不写入, 但会覆盖 — 故需在写入前判断 node.propKind==Get).
    std::unordered_map<std::string, std::string> memberReturnTypes;
    // Fix 016: 成员过程类型表 — 解决 prop_get_ 前缀启发式误判 (Fix 014b 遗留).
    // 跨模块 storageKey 冲突场景下 resolveClassMemberCall 的 fallback 需要判断
    // 目标成员在目标类中到底是 Function/Sub (无前缀) 还是 Property Get/Let/Set
    // (对应 prop_get_/prop_let_/prop_set_ 前缀), 之前的启发式用 "作用域中存在
    // <lower>$pg 外部符号" 判断, 在跨类同名混合 (如 cCryptoHMAC.Mode=Function,
    // cDelay.Mode=PropertyGet 共 6 个类的 Mode 互相冲突) 时会误判.
    // 覆盖规则 (语义分析填表时执行, 见 semantic_analyzer.cpp analyze()):
    // - Function    → 写入 Function    (覆盖)
    // - Sub         → 写入 Sub         (覆盖)
    // - PropertyGet → 写入 PropertyGet (覆盖)  // 读上下文优先级最高
    // - PropertyLet → 仅在当前键不存在或为 Let/Set 时写入 (不覆盖 Get/Function)
    // - PropertySet → 仅在当前键不存在或为 Let 时写入    (不覆盖 Get/Function/Let)
    // 读上下文 (resolveClassMemberCall) 优先级: Get > Function > Sub > Let > Set.
    std::unordered_map<std::string, ProcKind> memberProcKinds;
    // Fix 033: 成员参数表 — 解决 IndexOrCallExpr calleeParams 跨模块 storageKey 冲突.
    // 当 cTlsSocket.Create / cAsyncSocket.Create / cPassword.Create 共享 storageKey="create"
    // 时, driver.cpp 的 globalPublicSyms 按 storageKey 去重, 仅首个注册者的 sourceModule
    // 进入消费模块作用域. 此时 findClassMemberCallParams 的 Phase A (按 sourceModule 匹配
    // className) 失败, Phase B 通过 Class 符号自身的 memberParams[lower] 取回参数表 —
    // 与 memberProcKinds 同一优先级 (Get > Function > Sub > Let > Set), 仅在 procKind 写入
    // 胜出时同步更新 memberParams, 确保表中存储的是该类该成员读上下文优先级最高的参数表.
    // 用于 IndexOrCallExpr 中 Optional 参数 _has_ 标志计数匹配, 避免 C2197.
    std::unordered_map<std::string, std::vector<ParameterInfo>> memberParams;
    // Fix 092m: 成员数据字段类型表 — 类模块顶层变量 (仅 As <简单类型> 字段).
    // key=字段名小写, value=源码类型名 (原样, 如 "String"/"Long"/"Date"/"cTlsReMaster").
    // 用途: With 块类字段写 `temp->field = <COM 属性读>` 需按目标字段类型选择 COM
    // 解包函数 (String→ComGetStringProp / Long→ComGetIntProp / Date|Double→
    // ComGetDoubleProp / Object→ComGetObjectProp). 此前无类型信息一律按默认 BSTR
    // 解包 → int32_t 字段收到 BSTR 指针 (cHttpServer 373 `.Port = m_oServer.RemotePort`
    // 编译通过但语义错); 且 COM 标记不消费会泄漏到下一条语句 (374 C2440).
    std::unordered_map<std::string, std::string> memberFieldTypes;
    // Fix 092p: 成员数据字段的**声明原名**表 — key=字段名小写, value=声明时原样名
    // (如 "Socket"). C 结构体成员名按声明生成, 而 VB6 大小写不敏感: 源码里写
    // `.socket` (cHttpServerResponse 487 `Client.socket.SendData`) 也会指向同一字段,
    // cgen 直接 cIdent(源码名) → `me->Client->socket` C2039. 字段访问生成点用
    // canonicalClassFieldName() 规范化回声明名.
    std::unordered_map<std::string, std::string> memberFieldNames;
    // Fix 099: 类模块**显式 Public** 数据字段清单 (ActiveX DLL 的 COM 暴露用).
    // 真 VB6 把类的 Public 字段暴露成 Property Get/Let 对; 此前 C3 只把
    // Sub/Function/Property 收进 memberNames, Public 字段完全不进 IDispatch 表
    // → 客户端 GetIDsOfNames("Router") 失败 → vb6_ComGetProp 返回 NULL → 调用方
    // 拿 NULL 当对象用即崩. 此处独立成表 (不并入 memberNames) 是因为
    // resolveClassMemberCall 用 memberNames 判定"成员访问是否为属性调用",
    // 并入会把字段访问改写成 prop_get_ 调用, 破坏 `Foo(obj.Field)` 的 ByRef 语义.
    // 只收 access==AccessLevel::Public 的显式 Public 关键字: 类模块里的 `Dim x`
    // 在本编译器被解析成 AccessLevel::Default (=Public), 而真 VB6 中 Dim 等价
    // Private, 故必须排除 Default — 否则 cHttpServer 的 `Dim m_Sessions As ...`
    // 这类内部字段会被误暴露. 数组/WithEvents/As New 不暴露 (VB6 亦然).
    // 存声明原名, 保持声明顺序 (生成的表项顺序稳定).
    std::vector<std::string> publicFieldNames;
    // Fix 099: Public 字段的 TypeLib DISPID (driver 在 TypeLib 阶段回写).
    // key=字段名小写, 0/缺省=尚未分配. dll_entry.c 的字段表读它, 保证与
    // TypeLib 中同名字段的 dispid 一致 (否则早绑定客户端按 TypeLib 的 dispid
    // 调 Invoke 会命中错项).
    std::unordered_map<std::string, int32_t> memberFieldDispids;
    // Fix 091a: 成员写方向参数表 — memberParams 按读上下文优先级 (Get > Function >
    // Sub > Let > Set) 只存胜出者, 对「Get 有参 + Let 末参才是 value」的属性
    // (cJson.Item(key)/Let Item(key, Dat As Variant)) 会存成 Get 的 [key], 使
    // Pattern C/D2 属性赋值改写 (prop_get_ → prop_let_) 无法判定 Let 末参是否
    // Variant → 值实参不打包 → C2440. 故独立记录 Let/Set 胜出者参数表 (同类内
    // Let/Set 各自唯一, 无条件覆盖写入), 供 findClassMemberWriteParams 使用.
    // tB 类继承 (B08a): 成员小写键 -> 声明时的访问级别。既有 11 张成员表都不带访问级别,
    // 而 `Protected` 的语义 (家族内可见) 与 `Inherits` 的可见面裁决都要读它; 跨模块拷贝
    // (driver_crossmod) 与继承合并 (driver_classchain) 两处都得带上, 漏一张就是静默错判。
    std::unordered_map<std::string, AccessLevel> memberAccessLevels;
    std::unordered_map<std::string, std::vector<ParameterInfo>> memberLetParams;
    std::unordered_map<std::string, std::vector<ParameterInfo>> memberSetParams;
    bool isInterface = false;                          // 是否为接口类(纯抽象,无实现)
    std::vector<std::string> implementsNames;          // Implements列表: 该类实现的接口名
    // 接口方法(仅isInterface=true时有意义): 必须被实现类覆盖的方法签名
    std::vector<ParameterInfo> interfaceMethodParams;  // 备用: 接口方法参数信息
    std::vector<std::string> interfaceMethodNames;     // 接口方法名列表(小写)

    // --- P6.5 事件相关 ---
    bool isWithEvents = false;                         // 变量是否声明为WithEvents
    std::vector<std::string> eventNames;               // 类声明的事件名列表(仅Class)
    std::string withEventsSourceClass;                 // WithEvents变量的源类名

    // --- COM前期绑定相关 (P6.3, SymbolKind::ComClass/ComInterface) ---
    std::string comIidStr;            // 接口IID字符串 (如 "{2A0A3E20-...}")
    std::string comClsidStr;          // coclass CLSID字符串
    std::string comProgId;            // ProgID (如 "Scripting.FileSystemObject")
    std::string comDefaultIfaceName;  // 默认接口名 (ComClass用)
    std::string comDefaultIfaceIid;   // 默认接口IID (由TypeLib builder回写, 用于早绑定QI)
    int32_t comVtblBase = 7;         // vtable起始偏移 (IDispatch=7, IUnknown=3)
    int32_t comDispid = 0;           // M29: 方法的TypeLib DISPID (由driver在TypeLib阶段回写, 供cgen生成dll_entry.c方法表用, 确保两边dispid一致; 0=未分配)
    bool comIsDual = false;           // 双重接口 (dispinterface + vtable)
    // P13.23: COM event source interface (for WithEvents on external COM objects)
    std::string comSourceIfaceName;    // 默认事件源接口名
    std::string comSourceIfaceIid;     // 事件源接口IID
    bool comHasSourceIface = false;    // 是否有事件源接口
    bool comSourceIfaceIsDispOnly = false;  // source interface 是 dispinterface (true) 还是 vtable 接口 (false)
    std::unordered_map<std::string, int32_t> comEventDispids;  // 事件源方法名(lower)→DISPID

    // COM方法签名 (ComInterface用, 方法名小写→签名)
    struct ComMethodSig {
        std::string realName;         // 原始名称(保留大小写)
        int32_t memid = 0;           // DISPID
        int32_t vtableIndex = -1;    // vtable偏移
        Vb6Type returnType = Vb6Type::Void;  // 返回类型
        bool isPropertyGet = false;
        bool isPropertyPut = false;
        bool isPropertyPutRef = false;
        std::vector<ParameterInfo> params;  // 参数列表(含方向/类型)
    };
    std::unordered_map<std::string, ComMethodSig> comMethods;  // key=小写方法名
    std::unordered_map<std::string, ComMethodSig> comSourceMethods;  // source interface 方法签名 (key=小写方法名)

    // P24-10: COM默认成员名 (DISPID_VALUE=0, 如 Dictionary.Item, Collection._Item)
    // VB6语义: obj(args) 等价于 obj.DefaultMember(args)
    std::string comDefaultMemberName;       // 小写, 用于查找
    std::string comDefaultMemberRealName;   // 原始大小写, 用于代码生成

    // Fix 177b: 被本工程同名类模块遮蔽的类型库 coclass.
    // VB6 语义: 工程内定义优先于引用库, 故 `Dim x As Dictionary` / `New Dictionary`
    // 应当得到工程内那个类. 但引用库的同名 coclass 由类型解析先注入 (builtin),
    // 代码生成期整条晚绑定通路都建立在它上面 (comDefaultMemberName/comMethods 等).
    // 这里只记下"运行期该 coclass 其实由本工程类实现", 供创建点 (comNewExprFor)
    // 改走工程类工厂 + vb6_ComPackVB6InstanceRaw 包装, 其余路径保持不变, 避免
    // 工程类走早绑定通路时大面积替换生成代码.
    // 值为工程类名 (VB_Name, 即 C 侧 vb6_cls_<name> 的 <name>); 空 = 未被遮蔽.
    std::string comProjectImplClass;

    // --- P24-04: COM Module全局函数 (SymbolKind::ComModule) ---
    std::string comModuleDllPath;    // 源DLL路径 (用于LoadLibrary)
    std::unordered_map<std::string, ComMethodSig> comModuleFunctions;  // key=小写函数名

    // --- P24-04: VB_GlobalNameSpace promoted函数 (SymbolKind::ComGlobalNs) ---
    // coclass的VB_GlobalNameSpace=True时, 默认接口的Public方法提升为全局符号
    // 如VBMAN库的sGlobal._sGlobal.VBMAN() → 全局VBMAN函数
    // comClsidStr: GlobalNameSpace coclass的CLSID (如sGlobal的CLSID)
    // comDefaultIfaceIid: 默认接口的IID (如_sGlobal的IID)
    // comMethods[name]: 包含提升的方法签名 (key=小写方法名)
    std::string comGlobalNsMethodName;  // P24-04: 提升的方法名 (原始大小写, 如"VBMAN")

    // --- Delegate 签名 (仅 SymbolKind::Delegate) ---
    // 委托值本身按 Vb6Type::LongPtr 表示 (位兼容), 签名细节存这里,
    // 供赋值/传参/调用点的签名检查与后端桩生成使用.
    ProcKind delegateProcKind = ProcKind::Sub;
    CallConv delegateCallConv = CallConv::StdCall;
    Vb6Type delegateReturnType = Vb6Type::Void;
    // 语义层在 AddressOf 绑定成功时登记的静态目标 (v1 限同模块过程).
    // cgen 按对子生成调用桩: 桩用委托约定的签名, 体内转调 cdecl 过程本体.
    struct DelegateTarget {
        std::string procName;     // VB 过程名 (原样)
        std::string procModule;   // "" = 当前模块 (v1 仅此形态)
        // 绑定的具体重载变体指纹 (O2): "" = 该名字无重载组 (旧行为);
        // 非空 = 组内 fp 精确匹配的那条 (含 head), cgen 据此找回声明并区分桩名.
        std::string procFp;
    };
    std::vector<DelegateTarget> delegateTargets;

    // --- P20-21: UDT成员信息 (仅SymbolKind::UserDefinedType) ---
    struct UdtMemberInfo {
        std::string name;           // 成员名 (保留大小写)
        Vb6Type type = Vb6Type::Empty;  // 成员类型
        std::string typeRefName;    // 若类型为UDT, 保存UDT名称
        int32_t arraySize = 0;      // 固定大小数组: 0=非数组, >0=上界+1
        // Fix 037: 动态数组标记. `memberName() As Type` 语法, arraySize=0 且 isArrayDynamic=true.
        // 用于 UDT C 结构体 emit `vb6_SafeArray1D* Member;` 并支持 obj.member(idx) → VB6_SA_AT.
        bool isArrayDynamic = false;
    };
    std::vector<UdtMemberInfo> udtMembers;

    // --- 重载组 (tB 式 Overloading, O1) ---
    // overloadFp: 由语义层在 Pass1 注册前填好的"可重载过程"签名指纹
    //   (空 = 不参与分组: 类模块成员 / ParamArray 过程 / 非 Sub|Function)。
    // 同 kind 同名的第二个不同指纹变体存进 "<name>$ov$<fp>" 键 (isOverloadVariant),
    // 裸键恒为声明序首个 (head) —— 276 处按名 lookup 的旧语义 = first-in-chain。
    std::string overloadFp;
    bool isOverloadVariant = false;
    int ovlCount = 0;  // 仅 head 有意义: 已挂入的变体数

    Symbol() = default;
    Symbol(SymbolKind k, const std::string& n, Vb6Type t,
           SourceLocation loc, AccessLevel acc = AccessLevel::Public)
        : name(n), lowerName(toLower(n)), kind(k), type(t)
        , location(loc), access(acc) {}

    static std::string toLower(const std::string& s) {
        std::string result = s;
        std::transform(result.begin(), result.end(), result.begin(),
                       [](unsigned char c) { return std::tolower(c); });
        return result;
    }

    // 存储键: Property Get/Let/Set加后缀区分同名共存, Event也加后缀, 其他用lowerName
    std::string storageKey() const {
        if (isPropertyKind(kind)) {
            switch (kind) {
                case SymbolKind::PropertyGet:  return lowerName + "$pg";
                case SymbolKind::PropertyLet:  return lowerName + "$pl";
                case SymbolKind::PropertySet:  return lowerName + "$ps";
                default: break;
            }
        }
        // Event 使用独立存储键, 允许同名 Sub/Function 共存
        // VB6合法: Public Event CloseSck() + Public Sub CloseSck()
        if (kind == SymbolKind::Event) {
            return lowerName + "$ev";
        }
        // 重载变体 (O1/O3): 独立组内键。存储/跨模块注入/getPublicSymbols 全部
        // 经此键自然区分, head 恒占裸键。
        if (isOverloadVariant && (kind == SymbolKind::Sub || kind == SymbolKind::Function)
            && !overloadFp.empty()) {
            return lowerName + "$ov$" + overloadFp;
        }
        return lowerName;
    }

    const char* kindName() const {
        switch (kind) {
            case SymbolKind::Variable:        return "Variable";
            case SymbolKind::Constant:        return "Constant";
            case SymbolKind::Sub:             return "Sub";
            case SymbolKind::Function:        return "Function";
            case SymbolKind::PropertyGet:     return "Property Get";
            case SymbolKind::PropertyLet:     return "Property Let";
            case SymbolKind::PropertySet:     return "Property Set";
            case SymbolKind::Parameter:       return "Parameter";
            case SymbolKind::UserDefinedType: return "Type";
            case SymbolKind::EnumType:        return "Enum";
            case SymbolKind::EnumMember:      return "Enum Member";
            case SymbolKind::DeclareSub:      return "Declare Sub";
            case SymbolKind::DeclareFunc:     return "Declare Function";
            case SymbolKind::Event:           return "Event";
            case SymbolKind::Delegate:        return "Delegate";
            case SymbolKind::Class:           return "Class";
            case SymbolKind::Label:           return "Label";
            case SymbolKind::ComClass:        return "ComClass";
            case SymbolKind::ComInterface:    return "ComInterface";
            case SymbolKind::ComModule:       return "ComModule";
            case SymbolKind::ComGlobalNs:    return "ComGlobalNs";
        }
        return "Unknown";
    }
};

// ============================================================
// 作用域
// ============================================================

enum class ScopeKind : uint8_t {
    Module,     // 模块级 (全局)
    Procedure,  // 过程级 (Sub/Function/Property)
    Block,      // 块级 (If/For/Do/While/With/Select 内)
};

class Scope {
    friend class SymbolTable;  // SymbolTable需要直接访问symbols_
public:
    Scope(ScopeKind kind, Scope* parent = nullptr)
        : kind_(kind), parent_(parent) {}

    // 定义符号, 返回false如果已存在同名符号(Property Get/Let/Set允许同名共存)
    // keyOverride: 显式存储键. 仅用于 Fix 103 的 Type/过程同名共存场景
    //              (类型符号改用 <name>$ty), 留空时按 Symbol::storageKey() 推导.
    bool define(std::unique_ptr<Symbol> sym, const std::string& keyOverride = std::string());

    // 按名称查找 (本作用域, 不递归) - 返回第一个匹配
    Symbol* lookupLocal(const std::string& name) const;

    // 按名称+类别查找 (本作用域, 不递归) - 精确匹配
    Symbol* lookupLocalByKind(const std::string& name, SymbolKind kind) const;

    // 按名称查找 (递归向上搜索所有祖先作用域)
    Symbol* lookup(const std::string& name) const;

    // 获取所有符号
    const std::unordered_map<std::string, std::unique_ptr<Symbol>>& symbols() const {
        return symbols_;
    }

    ScopeKind kind() const { return kind_; }
    Scope* parent() const { return parent_; }

private:
    ScopeKind kind_;
    Scope* parent_;
    std::unordered_map<std::string, std::unique_ptr<Symbol>> symbols_;  // key=storageKey()
};

// ============================================================
// 符号表
// ============================================================

class SymbolTable {
public:
    SymbolTable(Diagnostics& diag);

    // 作用域管理
    void pushScope(ScopeKind kind);
    void popScope();

    // 当前作用域
    Scope* currentScope() const { return current_; }

    // 模块级作用域
    Scope* moduleScope() const { return moduleScope_; }

    // 定义符号 (当前作用域)
    bool define(std::unique_ptr<Symbol> sym);

    // 查找符号 (从当前作用域向上递归)
    Symbol* lookup(const std::string& name) const;

    // 查找符号 (仅当前作用域)
    Symbol* lookupLocal(const std::string& name) const;

    // 查找符号 (按名称+类别, 仅当前作用域)
    Symbol* lookupLocalByKind(const std::string& name, SymbolKind kind) const;

    // 查找模块级符号
    Symbol* lookupModule(const std::string& name) const;

    // 查找模块级符号 (按名称+类别, 用于Property Get/Let精确查找)
    Symbol* lookupModuleByKind(const std::string& name, SymbolKind kind) const;

    // 重载组内按**声明位置**精确定位某一条 Sub/Function 变体 (Pass2/cgen 找回自身)。
    // 无位置命中时回落裸键 head (零重载模块行为完全等价旧 lookupModule)。
    Symbol* lookupModuleOverloadByLoc(const std::string& name,
                                      const SourceLocation& loc) const;

    // 重载组完整候选集: 裸键 head + 全部 "<name>$ov$" 变体 (按 map 迭代序, head 在前)。
    // 组不存在或仅 head 时返回 {head} / {}。
    std::vector<Symbol*> lookupModuleOverloads(const std::string& name) const;

    // 按语义层选定的变体后缀取符号: suffix 空 = 裸键 head; 否则 "<name><suffix>"。
    Symbol* lookupModuleOverloadBySuffix(const std::string& name,
                                         const std::string& suffix) const;

    // 当前作用域深度 (0=模块级)
    int scopeDepth() const;

    // 当前作用域类型
    ScopeKind currentScopeKind() const;

    // 是否在过程中
    bool inProcedure() const;

    // 获取诊断系统
    Diagnostics& diagnostics() { return diag_; }

    // --- 跨模块符号操作 ---

    // 注入一个跨模块外部符号（由Driver在跨模块解析pass中调用）
    // 在模块级作用域定义一个isExternal=true的符号
    // replaceBuiltinCom: 仅当同名占用者是**类型库内建** coclass/接口、且待注入的是
    // 工程 Class 时替换之 (VB6: 工程类遮蔽引用库同名 coclass; Fix <vbeclipse>)
    void defineExternal(std::unique_ptr<Symbol> sym, bool replaceBuiltinCom = false);

    // 获取所有模块级Public符号（供其他模块链接用）
    // 返回 name → Symbol* 的映射（仅Sub/Function/Variable/Constant, Public访问级别）
    std::vector<const Symbol*> getPublicSymbols() const;

    // 获取当前模块引用的所有外部模块名集合
    // 遍历模块级符号, 返回所有 isExternal=true 的 sourceModule
    std::unordered_set<std::string> getExternalModuleNames() const;

    // Fix <vbeclipse>: VB6 隐式变量声明 (工程未写 Option Explicit 时, 首次使用的
    // 裸标识符自动成为 Variant 局部变量)。语义层登记 "<module>$<proc>" → 名字集,
    // 发码层在过程序言按此预声明 C 局部。
    void addImplicitVar(const std::string& moduleLower, const std::string& procLower,
                        const std::string& varName) {
        implicitVars_["<mod>" + moduleLower + "$" + procLower].insert(varName);
    }
    const std::unordered_set<std::string>* implicitVarsFor(const std::string& moduleLower,
                                                           const std::string& procLower) const {
        auto it = implicitVars_.find("<mod>" + moduleLower + "$" + procLower);
        return it == implicitVars_.end() ? nullptr : &it->second;
    }

private:
    Diagnostics& diag_;
    Scope* moduleScope_;
    Scope* current_;
    std::vector<std::unique_ptr<Scope>> scopes_;
    std::unordered_map<std::string, std::unordered_set<std::string>> implicitVars_;
};

} // namespace vb6c3
