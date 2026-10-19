// vb6com_pack.c - vb6com 模块拆分: VARIANT 参数打包/解包
// 由 vb6com.c 按 COM 调用层次拆分而来 (纯搬移, 零行为改动)

#include "vb6com.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "vb6com_internal.h"


// ============================================================
// VARIANT 封装 / 解封 (P6.2 cgen使用)
// ============================================================

// 将BSTR封装为VARIANT (返回堆分配的VARIANT*, 需vb6_ComVarClear释放)
void* vb6_ComPackBSTR(const wchar_t* bstr) {
    VARIANT* pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(pv);
    if (bstr) {
        pv->vt = VT_BSTR;
        pv->bstrVal = SysAllocString(bstr);
    } else {
        pv->vt = VT_BSTR;
        pv->bstrVal = NULL;
    }
    return (void*)pv;
}

// 将int32_t封装为VARIANT
void* vb6_ComPackInt(int32_t val) {
    VARIANT* pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(pv);
    pv->vt = VT_I4;
    pv->lVal = val;
    return (void*)pv;
}

// 将VB6 Boolean (int32_t: -1=True, 0=False) 封装为VARIANT VT_BOOL
// VBScript/COM 期望 VT_BOOL 而非 VT_I4
void* vb6_ComPackBool(int32_t val) {
    VARIANT* pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(pv);
    pv->vt = VT_BOOL;
    pv->boolVal = val ? VARIANT_TRUE : VARIANT_FALSE;
    return (void*)pv;
}

// 将double封装为VARIANT
void* vb6_ComPackDouble(double val) {
    VARIANT* pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(pv);
    pv->vt = VT_R8;
    pv->dblVal = val;
    return (void*)pv;
}

// 将void*(IDispatch*)封装为VARIANT (用于对象参数)
void* vb6_ComPackObject(void* obj) {
    VARIANT* pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(pv);
    pv->vt = VT_DISPATCH;
    pv->pdispVal = (IDispatch*)obj;
    if (obj) {
        // czUI fix: RTL 宿主对象 (HWND/UC 实例/集合/字体) 不是真实 COM,
        // 直接 AddRef/Release 会把结构体首字段当 vtable → AV
        // (Charts2020 ClsResizer For-Each Controls 实测)。
        // 宿主对象包一层真实 IDispatch (引用计数 + 晚绑定转发);
        // 真实 COM/ActiveX 对象保持原 AddRef/Release 路径。
        extern int32_t vb6_Host_IsHostObject(void* obj);
        extern void* vb6_UC_WrapHostObject(void* obj);
        if (vb6_Host_IsHostObject(obj)) {
            pv->pdispVal = (IDispatch*)vb6_UC_WrapHostObject(obj);
            pv->vt = VT_DISPATCH;
        } else {
            ((IDispatch*)obj)->lpVtbl->AddRef((IDispatch*)obj);
        }
    }
    return (void*)pv;
}

// Fix 160-com-byref: ByRef 对象出参封装.
// ppObj 是调用方 void* 变量的地址; 生成的 VARIANT 为 VT_DISPATCH|VT_BYREF,
// 其 ppdispVal 指向调用方变量。IDispatch::Invoke 把新对象指针写回 *ppObj,
// 调用方变量即持有新方法返回的对象 (Chilkat ObjectOf2/ArrayOf2 等出参语义)。
// 注意: 不做 AddRef —— 变量本身已是有效引用, ByRef 仅传递地址。
void* vb6_ComPackObjectRef(void** ppObj) {
    VARIANT* pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(pv);
    pv->vt = VT_DISPATCH | VT_BYREF;
    pv->ppdispVal = (IDispatch**)ppObj;
    return (void*)pv;
}

// Fix 104: 将"省略的实参"封装为 VARIANT (VT_ERROR + DISP_E_PARAMNOTFOUND).
// VB6 对 IDispatch 调用省略实参时 (obj.Method a, , c) 传的正是该值, 接收方据此
// 把"未提供"与"显式传 0/空串"区分开. 直接丢弃省略实参会让后续实参前移 (参数错位).
void* vb6_ComPackMissing(void) {
    VARIANT* pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    if (!pv) return NULL;
    VariantInit(pv);
    pv->vt = VT_ERROR;
    pv->scode = DISP_E_PARAMNOTFOUND;   /* 0x80020004 */
    return (void*)pv;
}

// 从VARIANT*解封BSTR (返回BSTR, 调用方需vb6_BSTR_Free释放)
wchar_t* vb6_ComUnpackBSTR(void* variant) {
    if (!variant) return NULL;
    VARIANT* pv = (VARIANT*)variant;
    if (pv->vt == VT_BSTR) {
        if (pv->bstrVal) {
            // 复制BSTR内容, 调用方用vb6_BSTR_Free释放
            return SysAllocString(pv->bstrVal);
        } else {
            // VT_BSTR + bstrVal=NULL = 空BSTR, 返回NULL (与VB6 vb6_BSTR_Empty()一致)
            return NULL;
        }
    }
    // 尝试变体转换
    if (pv->vt != VT_EMPTY && pv->vt != VT_NULL) {
        VARIANT vBstr;
        VariantInit(&vBstr);
        if (SUCCEEDED(VariantChangeType(&vBstr, pv, 0, VT_BSTR))) {
            BSTR tmp = SysAllocString(vBstr.bstrVal);
            VariantClear(&vBstr);
            return tmp;
        }
    }
    return NULL;
}

// 从VARIANT*解封int32_t
int32_t vb6_ComUnpackInt(void* variant) {
    if (!variant) return 0;
    VARIANT* pv = (VARIANT*)variant;
    if (pv->vt == VT_I4) return pv->lVal;
    if (pv->vt == VT_I2) return (int32_t)pv->iVal;
    if (pv->vt == VT_BOOL) return (pv->boolVal == VARIANT_TRUE) ? -1 : 0;
    if (pv->vt == VT_UI1) return (int32_t)pv->bVal;
    /* ai/032 §6.6: 无符号/64 位两档 —— C 截断语义 (与 RTL 取数表同口径)。
     * 不给这几档时落下面的 VariantChangeType(VT_I4): UI4 大值触发 DISP_E_OVERFLOW
     * 被吞成 0 (实测路径: 无符号属性读进 Long)。 */
    /* ai/032 §6.9: 本文件的 pv 是**真 Windows VARIANT***, 所以走它自己的成员名
       (cVal = CHAR 那一个字节 / uiVal = USHORT 那一短字), 不是本仓 vb6_VARIANT 的
       sbVal/usVal —— 那两个只存在于 vb6rtl_variant.h 里, 在这儿写会编不过。 */
    if (pv->vt == VT_I1)  return (int32_t)pv->cVal;    /* ai/032 §6.9 */
    if (pv->vt == VT_UI2) return (int32_t)pv->uiVal;   /* ai/032 §6.9 */
    if (pv->vt == VT_UI4) return (int32_t)pv->lVal;
    if (pv->vt == VT_I8)  return (int32_t)pv->llVal;
    if (pv->vt == VT_UI8) return (int32_t)pv->llVal;
    // 尝试变体转换
    if (pv->vt != VT_EMPTY && pv->vt != VT_NULL) {
        VARIANT vInt;
        VariantInit(&vInt);
        if (SUCCEEDED(VariantChangeType(&vInt, pv, 0, VT_I4))) {
            return vInt.lVal;
        }
    }
    return 0;
}

// 从VARIANT*解封double
double vb6_ComUnpackDouble(void* variant) {
    if (!variant) return 0.0;
    VARIANT* pv = (VARIANT*)variant;
    if (pv->vt == VT_R8) return pv->dblVal;
    if (pv->vt == VT_R4) return (double)pv->fltVal;
    if (pv->vt == VT_I4) return (double)pv->lVal;
    /* ai/032 §6.6: 无符号/64 位档同格 (VariantChangeType 对 UI8 大值也会丢)。 */
    /* 同上: 真 VARIANT 的成员名 (cVal / uiVal), 不是本仓 vb6_VARIANT 的 sbVal/usVal。 */
    if (pv->vt == VT_I1)  return (double)pv->cVal;    /* ai/032 §6.9 */
    if (pv->vt == VT_UI2) return (double)pv->uiVal;   /* ai/032 §6.9 */
    if (pv->vt == VT_UI4) return (double)pv->ulVal;
    if (pv->vt == VT_I8)  return (double)pv->llVal;
    if (pv->vt == VT_UI8) return (double)pv->ullVal;
    // 尝试变体转换
    if (pv->vt != VT_EMPTY && pv->vt != VT_NULL) {
        VARIANT vDbl;
        VariantInit(&vDbl);
        if (SUCCEEDED(VariantChangeType(&vDbl, pv, 0, VT_R8))) {
            return vDbl.dblVal;
        }
    }
    return 0.0;
}

// 从VARIANT*解封对象(void*/IDispatch*)
void* vb6_ComUnpackObject(void* variant) {
    if (!variant) return NULL;
    VARIANT* pv = (VARIANT*)variant;
    if (pv->vt == VT_DISPATCH) {
        // czUI fix: Pack 侧对 RTL 宿主对象 (字体/集合/UC 实例) 包了一层
        // IDispatch 包装器; 解包时必须还原原始对象 — 生成的 ctl 代码对字体
        // 等伪对象按结构体字段直接访问 (vb6_ComIface_Font*), 拿到包装器会
        // 全部失效 (Charts2020 图表标题/百分比文字缺失的根因)
        extern void* vb6_UC_UnwrapHost(void* obj);
        return vb6_UC_UnwrapHost((void*)pv->pdispVal);
    }
    if (pv->vt == VT_UNKNOWN) return (void*)pv->punkVal;
    return NULL;
}

// 释放ComCall/ComGetProp返回的VARIANT* (值类型安全: 释清BSTR/数字等)
void vb6_ComVarClear(void* variant) {
    if (!variant) return;
    VARIANT* pv = (VARIANT*)variant;
    VariantClear(pv);
    free(pv);
}

// 仅释放VARIANT结构体, 不清除内容 (对象所有权转移: pdispVal已被取走)
void vb6_ComVarFree(void* variant) {
    if (!variant) return;
    free(variant);  // 不调用VariantClear, 对象引用已转移给调用方
}
