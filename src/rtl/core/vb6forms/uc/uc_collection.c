// uc_collection.c - vb6forms_uc 拆分片：RTL 内建 Collection + For Each 枚举器
//
// 内容 = 拆分前 vb6forms_uc.c 第 964~988 / 989~1185 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_uc_internal.h

#include "vb6forms_uc_internal.h"

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// Fix 112c: RTL 内建 Collection (VB6 内建类)
// ============================================================
// 根因 (Charts 2020): `Set cAxisItem = New Collection` 发射为
// vb6_NewObject(L"Collection") → vb6_CreateObject → CLSIDFromProgID 必失败
// (Collection 是 VB6 语言内建, 无注册 ProgID) → vb6_RaiseError(429) →
// GUI 程序弹模态错误框并 exit(429) → 表现为"白板窗体+挂起".
// 这里提供纯 RTL 的 Collection, 挂进 Fix 112 宿主分派管道
// (vb6_ComCall/GetProp/SetProp/ForEach 已在这些合成对象上工作), 完全绕开 COM.
// 内部存 RTL vb6_VARIANT; 入参出参都是 Windows VARIANT (由 To/FromWinVariant 转换).

typedef struct vb6_CollRec {
    int32_t tag;          // 0xC01C01C0
    int32_t refs;         // 载体引用计数 (第一百二十九刀): New 时为 1, Set/槽进出各一问
    void*   items;        // vb6_VARIANT[] (RTL 简化布局)
    wchar_t** keys;       // BSTR[] 与 items 平行; 可为 NULL (无键集合)
    int32_t count;
    int32_t cap;
} vb6_CollRec;

int32_t vb6_uc_isColl(const void* p) {
    return vb6_uc_ptrReadable(p, sizeof(int32_t)) &&
           ((const vb6_CollRec*)p)->tag == VB6_UC_COLL_TAG;
}

#include "uc_collection_api.inc"

#ifdef __cplusplus
} // extern "C"
#endif
