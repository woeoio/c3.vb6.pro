#pragma once
// vb6comserver.h - VB6 COM服务端运行时 (P6.6 ActiveX DLL)
// 提供ActiveX DLL所需的基础设施:
//   - 全局引用计数 (g_vb6_cRef + g_vb6_cServerLock)
//   - IClassFactory 实现
//   - DllGetClassObject / DllCanUnloadNow 骨架
//   - 注册表辅助 (DllRegisterServer / DllUnregisterServer)
//   - 每个coclass的自注册信息表
//
// 编译器cgen为每个ActiveX DLL工程生成:
//   1. g_vb6_coclasses[] 表 (coclass描述)
//   2. DllGetClassObject() 实现
//   3. DllRegisterServer() / DllUnregisterServer() 实现
//   4. 每个Public类的IDispatch vtable + 实现

#include <windows.h>
#include <oleauto.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 全局引用计数
// ============================================================

// 全局对象引用计数 (每个AddRef+1, Release-1)
extern LONG g_vb6_cRef;

// 服务器锁计数 (Class Factory LockServer TRUE+1, FALSE-1)
extern LONG g_vb6_cServerLock;

// 增加全局引用计数
static inline void vb6_ServerAddRef(void) {
    InterlockedIncrement(&g_vb6_cRef);
}

// 减少全局引用计数
static inline void vb6_ServerRelease(void) {
    InterlockedDecrement(&g_vb6_cRef);
}

// ============================================================
// coclass描述表 (由cgen生成)
// ============================================================

// P6.6: 事件描述 (source dispinterface中的每个事件)
typedef struct vb6_EventDesc {
    const wchar_t* name;    // 事件名称
    int32_t dispid;         // DISPID
} vb6_EventDesc;

// ai/022 B16: 薄指针 (vb6_ivref_<I>*) 指向的槽表首三槽 = IUnknown。生成码的
// vb6_ivtbl_<I> 与它逐字对齐 (B04 起槽号固定 0/1/2)。
// 本类型与 vb6rtl_class_com.h 里那份是**同一份声明**: 两棵编译树 (vb6rtl / vb6comserver)
// 各自要能用, 而生成码两个头都 include —— 用同一个 #ifndef 守卫, 先到的那个生效。
// `__stdcall` 是 B16 的口径: x86 下不加就是 __cdecl, 与类型库里如实写出的 CC_STDCALL 对不上。
#ifndef VB6_IVTBL_PREFIX_DEFINED
#define VB6_IVTBL_PREFIX_DEFINED
typedef struct vb6_ivtbl_prefix {
    long (__stdcall *QueryInterface)(void* self, const void* riid, void** ppv);
    unsigned long (__stdcall *AddRef)(void* self);
    unsigned long (__stdcall *Release)(void* self);
} vb6_ivtbl_prefix;
#endif

// coclass描述结构
// 每个Public类(instancing >= PublicNotCreatable)生成一条
typedef struct vb6_CoClassDesc {
    const char* progId;           // ProgID (如 "MyLib.MyClass")
    const char* clsidStr;         // CLSID字符串 (如 "{xxxxxxxx-...}")
    const char* classVariable;    // VB6类变量名 (模块名，用于查找New函数)
    // 工厂函数: 创建VB6类实例 (返回void*指针)
    void* (*factoryFunc)(void);   // vb6_cls_<Name>_New()
    // 销毁函数: 释放VB6类实例
    void (*destroyFunc)(void*);   // vb6_cls_<Name>_Destroy()
    // IDispatch vtable (由cgen生成, 可为NULL表示不支持自动化)
    const void* dispatchVtable;
    // 类方法数 (用于IDispatch::GetIDsOfNames查找)
    int methodCount;
    // 方法描述表 (name/dispid/invkind)
    const struct vb6_DispMethodDesc* methods;
    // P12.1: Implements接口IID表 (QI时遍历匹配)
    int ifaceCount;                       // 实现的接口数量 (Implements语句)
    const IID* const* ifaceIids;          // 接口IID指针数组 (每个元素指向一个静态IID)
    // P6.6: 事件源接口 (IConnectionPointContainer)
    const IID* defaultIfaceIid;            // 默认dispinterface IID (指向静态常量, 早绑定QI用)
    const char* sourceIfaceIid;           // source dispinterface IID字符串字符串 (NULL=无事件)
    int eventCount;                       // 事件数量
    const vb6_EventDesc* events;          // 事件描述表 (DISPID + 名称)
    // ai/022 B16: 新式接口 (`Interface ... End Interface`) 的对外面。两个指针同生同灭:
    // cgen 只给"实现了新式接口"的类生成它们, 其余类 (legacy `Implements` 的类模块、
    // 无接口的存量工程) 都是 NULL —— QI 与销毁照旧走胖指针那条, 存量产物零影响。
    // 为什么在**结构体末尾**追加: 除本表和空桩 `{0}` 外没有别的位置初始化者, 但末尾
    // 追加让任何"少写一个初始化器"的旧写法自动得 NULL, 不会错位到别的字段上。
    void* (*ifaceThinPtr)(void* instance, const void* riid);  // 接口 IID -> 薄指针 (vb6_ivref_<I>*)
    void  (*instanceClaimRelease)(void* instance);            // 包装器放手: 清 __comObj + 退底座引用
} vb6_CoClassDesc;

// IDispatch方法描述
typedef struct vb6_DispMethodDesc {
    const wchar_t* name;    // 方法/属性名
    int32_t dispid;         // DISPID
    int32_t invkind;        // INVOKE_KIND (1=Method, 2=PropertyGet, 4=PropertyPut, 8=PropertyPutRef)
    // 调用函数指针 (void* this, VARIANT* args, int argc, VARIANT* result)
    // this=VB6类实例, args=参数数组, argc=参数数, result=返回值(可NULL)
    void (*invokeFunc)(void*, void**, int32_t, void*);
} vb6_DispMethodDesc;

// cgen生成的全局coclass描述表
extern const vb6_CoClassDesc g_vb6_coclasses[];
extern const int g_vb6_coclassCount;

// ============================================================
// VB6 COM对象 (IDispatch包装层)
// ============================================================

// VB6 COM对象结构
// 包装VB6类实例为COM IDispatch对象
typedef struct vb6_ComObject {
    const struct vb6_IDispatchVtable* vtable;
    LONG refCount;
    const vb6_CoClassDesc* desc;
    void* vb6Instance;  // VB6类实例
    int ownsInstance;  // Fix 188: 1=包装器拥有实例(释放时销毁), 0=借用(宿主拥有)
    struct vb6_ConnectionPointContainer* cpc;  // P6.6: 事件连接点容器 (lazy init)
    struct vb6_ProvideClassInfo2* pci;  // P6.6: IProvideClassInfo2 (lazy init)
} vb6_ComObject;

// IDispatch vtable (C风格)
typedef struct vb6_IDispatchVtable {
    // IUnknown
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(vb6_ComObject*, REFIID, void**);
    ULONG (STDMETHODCALLTYPE *AddRef)(vb6_ComObject*);
    ULONG (STDMETHODCALLTYPE *Release)(vb6_ComObject*);
    // IDispatch
    HRESULT (STDMETHODCALLTYPE *GetTypeInfoCount)(vb6_ComObject*, UINT*);
    HRESULT (STDMETHODCALLTYPE *GetTypeInfo)(vb6_ComObject*, UINT, LCID, ITypeInfo**);
    HRESULT (STDMETHODCALLTYPE *GetIDsOfNames)(vb6_ComObject*, REFIID, LPOLESTR*, UINT, LCID, DISPID*);
    HRESULT (STDMETHODCALLTYPE *Invoke)(vb6_ComObject*, DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*);
} vb6_IDispatchVtable;

// 创建VB6 COM对象 (包装VB6实例为IDispatch)
vb6_ComObject* vb6_ComObject_Create(const vb6_CoClassDesc* desc);

// Fix 099: 包装一个**已存在**的 VB6 类实例 (不调 factoryFunc).
// 用于 Public 对象字段 getter —— 字段实例在类内部创建, 是裸结构体指针.
// 实例 __comObj 已置则 AddRef 复用, 否则新建包装并回填.
// 前置条件: 实例所在类的结构体首字段是 __comObj (ExeComBridge 01 起所有类模块
// 都有该字段, DLL 与 EXE 工程一致).
vb6_ComObject* vb6_ComObject_FromInstance(const vb6_CoClassDesc* desc, void* instance);
// Fix 188: 包装**宿主已拥有**的实例 (Public 对象字段 getter / 方法返回工程类实例).
// 释放语义与 FromInstance 不同: 客户端释放到 0 只回收包装器, 不销毁 VB6 实例.
vb6_ComObject* vb6_ComObject_FromBorrowedInstance(const vb6_CoClassDesc* desc, void* instance);

// Fix 099: 按类变量名 (VB6 模块名) 在 g_vb6_coclasses[] 中查描述, 未命中返回 NULL.
const vb6_CoClassDesc* vb6_FindCoClassDesc(const char* classVariable);

// Fix <vbeclipse> rev30: **裸工程类实例 → coclass 描述** 的反查。
// cgen 在每个类工厂 (`vb6_cls_X_New`) 里登记一次, 使晚绑定调用点传进来的
// `vb6_cls_X*` 能被认出来 —— 详见 vb6comserver_obj.c 里那张表的完整论证
// (症状: `vb6_ComCall(裸实例, …)` 恒返回 VT_EMPTY, 因为首字段 __comObj 是 NULL,
//  `vb6_ComIsDispatchable` 判否 ⇒ GetIDsOfNames 失败)。未登记 ⇒ 返回 NULL。
void vb6_RegisterProjectClassInstance(void* inst, const vb6_CoClassDesc* desc);
const vb6_CoClassDesc* vb6_FindProjectClassDesc(const void* inst);

// Fix 099: 从 IDispatch 取回 VB6 实例裸指针; 非本 RTL 产出的对象返回 NULL.
void* vb6_ComObject_GetInstance(void* pdisp);

// ExeComBridge 03: 把工程类实例包成"COM 调用实参格式"的 VARIANT* (VT_DISPATCH).
// cgen 为每个类模块生成 vb6_ComPack_<类名>(void* instance) 包一层本函数; 触发场景
// 是 `New <本工程类>` 出现在 COM 调用实参位置
// (VBMAN_DEMO: `.Router.Reg "Demo", New bHello`). 原先裸 vb6_cls_X* 被直接塞进
// VT_DISPATCH, 对端取值/释放 AddRef 时把结构体首字段当 vtable → 0xC0000005.
// 这里先经 FromInstance 包装成真 IDispatch 再转 VARIANT.
// 引用计数: FromInstance 给出的 1 个引用直接转移给 VARIANT, 不再额外 AddRef
// (额外 AddRef 会让包装器与实例永不释放). 实例所在类未进 coclass 表
// (Private / 非 MultiUse|SingleUse) 时返回 NULL, 等效 VB6 的 Nothing.
void* vb6_ComPackVB6InstanceRaw(const char* classVariable, void* instance);

// ============================================================
// Class Factory (IClassFactory实现)
// ============================================================

typedef struct vb6_ClassFactory {
    const struct vb6_IClassFactoryVtable* vtable;
    LONG refCount;
    const vb6_CoClassDesc* desc;  // 关联的coclass描述
} vb6_ClassFactory;

// IClassFactory vtable (C风格)
typedef struct vb6_IClassFactoryVtable {
    // IUnknown
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(vb6_ClassFactory*, REFIID, void**);
    ULONG (STDMETHODCALLTYPE *AddRef)(vb6_ClassFactory*);
    ULONG (STDMETHODCALLTYPE *Release)(vb6_ClassFactory*);
    // IClassFactory
    HRESULT (STDMETHODCALLTYPE *CreateInstance)(vb6_ClassFactory*, IUnknown*, REFIID, void**);
    HRESULT (STDMETHODCALLTYPE *LockServer)(vb6_ClassFactory*, BOOL);
} vb6_IClassFactoryVtable;

// 获取指定CLSID的Class Factory
// 由DllGetClassObject调用
HRESULT vb6_GetClassFactory(REFCLSID rclsid, REFIID riid, void** ppv, 
                            const vb6_CoClassDesc* coclasses, int count);

// ============================================================
// DLL导出函数 (由cgen生成的包装函数调用)
// ============================================================

// DllCanUnloadNow - 判断DLL是否可卸载
HRESULT vb6_DllCanUnloadNow(void);

// 注册表辅助
// 注册一个coclass (CLSID + ProgID + InprocServer32)
HRESULT vb6_RegisterCoClass(const vb6_CoClassDesc* desc, const wchar_t* dllPath);

// 反注册一个coclass
HRESULT vb6_UnregisterCoClass(const vb6_CoClassDesc* desc);

// 获取当前DLL路径 (用于注册InprocServer32)
HRESULT vb6_GetDllPath(wchar_t* path, DWORD size);

// P6.13: TypeLib注册辅助
// 注册嵌入DLL资源的TypeLib到注册表
HRESULT vb6_RegisterTypeLib(const wchar_t* dllPath);
// 反注册TypeLib
HRESULT vb6_UnregisterTypeLib(const wchar_t* dllPath);


// ============================================================
// P6.6: 服务端 IConnectionPointContainer + IConnectionPoint
// 允许ActiveX DLL的COM对象向客户端触发事件
// ============================================================

typedef struct vb6_ConnectionPoint vb6_ConnectionPoint;
typedef struct vb6_ConnectionPointContainer vb6_ConnectionPointContainer;

// IConnectionPoint vtable
typedef struct vb6_IConnectionPointVtable {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(vb6_ConnectionPoint*, REFIID, void**);
    ULONG (STDMETHODCALLTYPE *AddRef)(vb6_ConnectionPoint*);
    ULONG (STDMETHODCALLTYPE *Release)(vb6_ConnectionPoint*);
    HRESULT (STDMETHODCALLTYPE *GetConnectionInterface)(vb6_ConnectionPoint*, IID*);
    HRESULT (STDMETHODCALLTYPE *GetConnectionPointContainer)(vb6_ConnectionPoint*, void**);
    HRESULT (STDMETHODCALLTYPE *Advise)(vb6_ConnectionPoint*, IUnknown*, DWORD*);
    HRESULT (STDMETHODCALLTYPE *Unadvise)(vb6_ConnectionPoint*, DWORD);
    HRESULT (STDMETHODCALLTYPE *EnumConnections)(vb6_ConnectionPoint*, void**);
} vb6_IConnectionPointVtable;

// IConnectionPoint 实现
struct vb6_ConnectionPoint {
    const vb6_IConnectionPointVtable* vtable;
    LONG refCount;
    IID sourceIid;
    vb6_ConnectionPointContainer* container;
    DWORD* cookies;
    IUnknown** sinks;
    int connCount;
    int connCapacity;
    DWORD nextCookie;
};

// IConnectionPointContainer vtable
typedef struct vb6_IConnectionPointContainerVtable {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(vb6_ConnectionPointContainer*, REFIID, void**);
    ULONG (STDMETHODCALLTYPE *AddRef)(vb6_ConnectionPointContainer*);
    ULONG (STDMETHODCALLTYPE *Release)(vb6_ConnectionPointContainer*);
    HRESULT (STDMETHODCALLTYPE *EnumConnectionPoints)(vb6_ConnectionPointContainer*, void**);
    HRESULT (STDMETHODCALLTYPE *FindConnectionPoint)(vb6_ConnectionPointContainer*, REFIID, void**);
} vb6_IConnectionPointContainerVtable;

// IConnectionPointContainer 实现
struct vb6_ConnectionPointContainer {
    const vb6_IConnectionPointContainerVtable* vtable;
    LONG refCount;
    vb6_ComObject* comObj;
    vb6_ConnectionPoint* connPoint;
};

// 为COM对象创建ConnectionPointContainer
vb6_ConnectionPointContainer* vb6_CPC_Create(vb6_ComObject* comObj);

// 通过COM对象触发事件 (由RaiseEvent调用)
void vb6_FireEvent(vb6_ComObject* comObj, int32_t dispid, VARIANT* args, int argc);
// ============================================================
// P6.6: IProvideClassInfo2 实现
// VBScript通过此接口发现事件源dispinterface
// ============================================================

typedef struct vb6_ProvideClassInfo2 vb6_ProvideClassInfo2;

// IProvideClassInfo2 vtable
typedef struct vb6_IProvideClassInfo2Vtable {
    // IUnknown
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(vb6_ProvideClassInfo2*, REFIID, void**);
    ULONG (STDMETHODCALLTYPE *AddRef)(vb6_ProvideClassInfo2*);
    ULONG (STDMETHODCALLTYPE *Release)(vb6_ProvideClassInfo2*);
    // IProvideClassInfo
    HRESULT (STDMETHODCALLTYPE *GetClassInfo)(vb6_ProvideClassInfo2*, ITypeInfo**);
    // IProvideClassInfo2
    HRESULT (STDMETHODCALLTYPE *GetGUID)(vb6_ProvideClassInfo2*, DWORD, GUID*);
} vb6_IProvideClassInfo2Vtable;

// IProvideClassInfo2 实现
struct vb6_ProvideClassInfo2 {
    const vb6_IProvideClassInfo2Vtable* vtable;
    LONG refCount;
    vb6_ComObject* comObj;
};

// 为COM对象创建IProvideClassInfo2
vb6_ProvideClassInfo2* vb6_PCI_Create(vb6_ComObject* comObj);

#ifdef __cplusplus
}
#endif
