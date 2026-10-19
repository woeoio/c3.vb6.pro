// uc_propbag.c - vb6forms_uc 拆分片：轻量 PropertyBag (IDispatch)
//
// 内容 = vb6forms_uc.c 在 main 侧新增段（Fix 120-141 / czUI fix）的 PropertyBag 部分，
// 移植自 ferock/0.10.7 合并 origin/main 时的文件尾新增段（原为单文件实现）。
// 运行期读取设计期持久化属性：生成的 UserControl_ReadProperties 通过
// vb6_ComCall(bag, "ReadProperty", ...) late-bound 调用本实现；
// 只有 .ctl 内部的读取后同步逻辑能借此执行
// (czUI.ctl:478 `If m_Checked Then m_AnimPos = 1!` → 开关初始 ON)。
// 跨族共享符号见 vb6forms_uc_internal.h

#include "vb6forms_uc_internal.h"

#ifdef _WIN32
#include <olectl.h>   /* PICTDESC / IPictureDisp / OleCreatePictureIndirect */
#endif

/* Fix P-BMP-3D 那张 Picture 注册表 (vb6rtl_com.c) —— 登进去之后
   `Set 控件.Picture = <图片>` 才能认出这是 COM 对象而不是 GDI 句柄。 */
extern void vb6_PictureRegister(void* picture);

#ifdef __cplusplus
extern "C" {
#endif

typedef struct Vb6BagItem {
    BSTR    name;
    VARIANT val;
} Vb6BagItem;

typedef struct Vb6PropBag {
    IDispatch   disp;          // lpVtbl 在首字段
    ULONG       refs;
    Vb6BagItem* items;
    int32_t     count;
    int32_t     cap;
} Vb6PropBag;

#define VB6_BAG_NAME_MAX 256

static HRESULT STDMETHODCALLTYPE Vb6Bag_QI(IDispatch* self, REFIID riid, void** out) {
    if (!out) return E_POINTER;
    *out = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *out = self;
        self->lpVtbl->AddRef(self);
        return S_OK;
    }
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE Vb6Bag_AddRef(IDispatch* self) { return 1; }
static ULONG STDMETHODCALLTYPE Vb6Bag_Release(IDispatch* self) { return 1; }
static HRESULT STDMETHODCALLTYPE Vb6Bag_GetTypeInfoCount(IDispatch* self, UINT* n) {
    if (n) *n = 0; return S_OK;
}
static HRESULT STDMETHODCALLTYPE Vb6Bag_GetTypeInfo(IDispatch* self, UINT i, LCID l, ITypeInfo** t) {
    (void)t; return E_NOTIMPL;
}
static HRESULT STDMETHODCALLTYPE Vb6Bag_GetIDsOfNames(IDispatch* self, REFIID riid,
        LPOLESTR* names, UINT cNames, LCID lcid, DISPID* out) {
    (void)self; (void)riid; (void)lcid;
    if (!out || !cNames || !names || !names[0]) return E_POINTER;
    wchar_t buf[64];
    int i = 0;
    for (; names[0][i] && i < 63; i++) buf[i] = towlower(names[0][i]);
    buf[i] = 0;
    if (wcscmp(buf, L"readproperty") == 0)  { *out = 1; return S_OK; }
    if (wcscmp(buf, L"writeproperty") == 0) { *out = 2; return S_OK; }
    if (wcscmp(buf, L"contents") == 0)      { *out = 3; return S_OK; }
    return DISP_E_UNKNOWNNAME;
}
static Vb6BagItem* Vb6Bag_Find(Vb6PropBag* bag, const wchar_t* name) {
    for (int32_t i = 0; i < bag->count; i++) {
        if (_wcsicmp(bag->items[i].name, name) == 0) return &bag->items[i];
    }
    return NULL;
}
static void vb6_BagPut(Vb6PropBag* bag, const wchar_t* name, const VARIANT* v);  /* 定义在下方 */
static void vb6_bag_try_deserialize_contents(Vb6PropBag* bag, const VARIANT* v); /* 定义在下方 */
static HRESULT STDMETHODCALLTYPE Vb6Bag_Invoke(IDispatch* self, DISPID id, REFIID riid, LCID lcid,
        WORD flags, DISPPARAMS* pd, VARIANT* result, EXCEPINFO* ei, UINT* argErr) {
    (void)riid; (void)lcid; (void)flags; (void)ei; (void)argErr;
    Vb6PropBag* bag = (Vb6PropBag*)self;
    if (!pd) return E_POINTER;
    if (id == 1) { /* ReadProperty(name[, default]) — rgvarg 逆序 */
        int c = (int)pd->cArgs;
        if (c < 1) return DISP_E_BADPARAMCOUNT;
        VARIANT* vDefault = (c >= 2) ? &pd->rgvarg[0] : NULL;
        VARIANT* vName   = &pd->rgvarg[c - 1];
        if (vName->vt != VT_BSTR || !vName->bstrVal) return DISP_E_TYPEMISMATCH;
        Vb6BagItem* it = Vb6Bag_Find(bag, vName->bstrVal);
        if (!result) return S_OK;
        VariantInit(result);
        if (it) return VariantCopy(result, &it->val);
        if (vDefault) return VariantCopy(result, vDefault);
        return S_OK;
    }
    if (id == 2) { /* WriteProperty(name, value) */
        int c = (int)pd->cArgs;
        if (c < 2) return DISP_E_BADPARAMCOUNT;
        VARIANT* vValue = &pd->rgvarg[0];
        VARIANT* vName  = &pd->rgvarg[1];
        if (vName->vt != VT_BSTR || !vName->bstrVal) return DISP_E_TYPEMISMATCH;
        Vb6BagItem* it = Vb6Bag_Find(bag, vName->bstrVal);
        if (!it) {
            if (bag->count >= bag->cap) {
                int32_t ncap = bag->cap ? bag->cap * 2 : 8;
                Vb6BagItem* ni = (Vb6BagItem*)realloc(bag->items, sizeof(Vb6BagItem) * ncap);
                if (!ni) return E_OUTOFMEMORY;
                bag->items = ni; bag->cap = ncap;
            }
            it = &bag->items[bag->count++];
            memset(it, 0, sizeof(*it));
            it->name = SysAllocString(vName->bstrVal);
        } else {
            VariantClear(&it->val);
        }
        VariantInit(&it->val);
        VariantCopy(&it->val, (VARIANT*)vValue);
        return S_OK;
    }
    if (id == 3) { /* Contents —— 序列化 blob 的读写槽。
                      VB6 语义: Contents 是整块序列化字节, ReadProperty 从中解字段。
                      前半段(czUI 期)只当普通项存取 ⇒ 于是**没写过时就查不到任何字段**,
                      ReadProperty 一律回调用方给的默认值。
                      Fix: 现按真实 VB6 PropertyBag 序列化格式解析 (见 vb6_bag_deserialize),
                      把 blob 里的每个 (name,value) 拍平进本项表 —— 这样
                      `.ctl` 里 `PropBag.Contents = FormBag.ReadProperty("InitX", 0)`
                      之后的 `PropBag.ReadProperty("InitXCount", 0)` 才能拿到真值
                      (此前恒为默认 0 ⇒ ComCtlsDemo 的 StatusBar1.Panels 设计期面板全丢,
                       Command2_Click 里 `Panels(1)` 取回 NULL 直接 AV)。 */
        int c = (int)pd->cArgs;
        if (result && c == 0) {   /* Property Get */
            Vb6BagItem* it = Vb6Bag_Find(bag, L"Contents");
            VariantInit(result);
            if (it) return VariantCopy(result, &it->val);
            return S_OK;
        }
        if (c >= 1) {             /* Property Let / Set (rgvarg[0] 是值, 可能带命名参数) */
            vb6_BagPut(bag, L"Contents", (const VARIANT*)&pd->rgvarg[0]);
            vb6_bag_try_deserialize_contents(bag, (const VARIANT*)&pd->rgvarg[0]);
            return S_OK;
        }
        return DISP_E_BADPARAMCOUNT;
    }
    return DISP_E_MEMBERNOTFOUND;
}

static IDispatchVtbl g_vb6BagVtbl = {
    Vb6Bag_QI, Vb6Bag_AddRef, Vb6Bag_Release,
    Vb6Bag_GetTypeInfoCount, Vb6Bag_GetTypeInfo, Vb6Bag_GetIDsOfNames, Vb6Bag_Invoke
};

void* vb6_UC_PropBagCreate(void) {
    Vb6PropBag* bag = (Vb6PropBag*)calloc(1, sizeof(Vb6PropBag));
    if (!bag) return NULL;
    bag->disp.lpVtbl = &g_vb6BagVtbl;
    bag->refs = 1;
    return bag;
}

void vb6_UC_PropBagFree(void* bagPtr) {
    Vb6PropBag* bag = (Vb6PropBag*)bagPtr;
    if (!bag) return;
    for (int32_t i = 0; i < bag->count; i++) {
        if (bag->items[i].name) SysFreeString(bag->items[i].name);
        VariantClear(&bag->items[i].val);
    }
    free(bag->items);
    free(bag);
}

static void vb6_BagPut(Vb6PropBag* bag, const wchar_t* name, const VARIANT* v) {
    Vb6BagItem* it = Vb6Bag_Find(bag, name);
    if (!it) {
        if (bag->count >= bag->cap) {
            int32_t ncap = bag->cap ? bag->cap * 2 : 8;
            Vb6BagItem* ni = (Vb6BagItem*)realloc(bag->items, sizeof(Vb6BagItem) * ncap);
            if (!ni) return;
            bag->items = ni; bag->cap = ncap;
        }
        it = &bag->items[bag->count++];
        memset(it, 0, sizeof(*it));
        it->name = SysAllocString(name);
    } else {
        VariantClear(&it->val);
    }
    VariantInit(&it->val);
    VariantCopy(&it->val, (VARIANT*)v);
}

/* ============================================================
 * VB6 PropertyBag 序列化格式解析 (Contents 字节数组 / .frx 属性包块)
 * ------------------------------------------------------------
 * 布局 (The trick 2020 逆向 + 本项目 .frx 实测核对):
 *   Header (8B):  WORD wSignature = 0xB293 ; WORD wCatalogOffset ;
 *                 DWORD cbContent   (== 块总长 - 头偏移)
 *   Item:         WORD  iVarType
 *                 WORD  iStrSize      (项名字符数, 低 10 位; 名名字节数 = 字符数*2)
 *                 DWORD lHash
 *                 DWORD lNextItemOffset  (相对**本项起点**到下一项)
 *                 WCHAR name[iStrSize]
 *                 value
 *   取值:  VT_I2(2)=2B  VT_I4(3)=4B  VT_R4(4)=4B  VT_R8(5)=8B  VT_BOOL(11)=2B
 *          VT_ARRAY|VT_UI1(0x2011) = WORD nDims ; DWORD cbData ; DWORD lLBound ; BYTE[cbData]
 *          VT_DISPATCH(9) 等 = 跳过 (图片走 IPersistStream 流, 暂不支持)
 *   项间可能补 0 到 4 字节边界 —— 一律用 lNextItemOffset 前进, 不猜填充。
 * 本实现把解出的 (name, value) **拍平**进本 PropertyBag 的项表 ——
 * 于是 .ctl 里 `PropBag.Contents = FormBag.ReadProperty("InitX", 0)` 之后
 * 的 `PropBag.ReadProperty("InitXCount", 0)` 能拿到真值 (此前恒回默认值)。
 * ============================================================ */
#define VB6_BAG_SIGNATURE 0xB293u

static uint16_t vb6_bag_rd16(const uint8_t* p) { return (uint16_t)(p[0] | (p[1] << 8)); }
static uint32_t vb6_bag_rd32(const uint8_t* p) {
    return (uint32_t)p[0] | ((uint32_t)p[1] << 8) | ((uint32_t)p[2] << 16) | ((uint32_t)p[3] << 24);
}

/* 以 UTF-16 字节串造 BSTR (可含内嵌 NUL); 非文本时返回 NULL */
static BSTR vb6_bag_text_from_utf16le(const uint8_t* p, uint32_t n) {
    if ((n & 1) != 0) return NULL;
    uint32_t cch = n / 2;
    if (cch == 0) return SysAllocString(L"");
    for (uint32_t i = 0; i < cch; i++) {
        uint32_t c = vb6_bag_rd16(p + i * 2);
        if (c == 0x09 || c == 0x0A || c == 0x0D) continue;
        if (c < 0x20 || c == 0xFFFE || c == 0xFFFF) return NULL;   /* 控制字符 => 视为二进制 */
    }
#ifdef _WIN32
    return SysAllocStringLen((const OLECHAR*)p, cch);
#else
    return NULL;
#endif
}

/* VT_DISPATCH 项 = OLE 图片流 (StdPicture 的序列化):
 *   [CLSID StdPicture (16B)][DWORD key='lt' (0x746C)][…头部…][DIB / ICON 数据]
 * 头部各字段的语义未逐个逆向 —— 直接**扫 BITMAPINFOHEADER**
 * (biSize=40, planes=1, bitCount ∈ {1,4,8,16,24,32}, w>0, h≠0) 起算, 用
 * CreateDIBSection 造 HBITMAP 再包成 IPictureDisp (fOwn=TRUE, 图归图管)。
 * 实测 Builds/ImageList + Builds/StatusBar 的 .frx 图片块即 4bpp 带调色板 DIB。
 * 认不出来就返回 NULL ⇒ 该项不进表 (.ctl 侧拿到默认 Nothing)。 */
static void* vb6_bag_make_picture(const uint8_t* p, uint32_t n) {
    if (n < 32) return NULL;
    if (vb6_bag_rd32(p + 16) != 0x746Cu) return NULL;      /* 'lt' 键 */
    const uint8_t* dib = NULL;
    for (uint32_t i = 20; i + 40 <= n; i++) {
        if (vb6_bag_rd32(p + i) != 40) continue;
        int32_t w = (int32_t)vb6_bag_rd32(p + i + 4);
        int32_t h = (int32_t)vb6_bag_rd32(p + i + 8);
        uint16_t planes = vb6_bag_rd16(p + i + 12);
        uint16_t bits = vb6_bag_rd16(p + i + 14);
        if (w <= 0 || w > 8192 || h == 0 || planes != 1) continue;
        if (bits != 1 && bits != 4 && bits != 8 && bits != 16 && bits != 24 && bits != 32) continue;
        dib = p + i;
        break;
    }
    if (!dib) return NULL;

    int32_t w = (int32_t)vb6_bag_rd32(dib + 4);
    int32_t h = (int32_t)vb6_bag_rd32(dib + 8);
    uint16_t bits = vb6_bag_rd16(dib + 14);
    uint32_t clrUsed = vb6_bag_rd32(dib + 32);
    uint32_t absH = (uint32_t)(h < 0 ? -(int64_t)h : (int64_t)h);
    uint32_t palEntry = (bits <= 8) ? (clrUsed ? clrUsed : (1u << bits)) : 0u;
    uint32_t hdrTotal = 40u + palEntry * 4u;
    uint32_t stride = (((uint32_t)w * (uint32_t)bits) + 31u) / 32u * 4u;
    uint32_t maskStride = (((uint32_t)w + 31u) / 32u) * 4u;   /* 1bpp 行按 4 字节对齐 */
    uint32_t imgSize = vb6_bag_rd32(dib + 20);
    uint32_t off = (uint32_t)(dib - p);
    if (off + hdrTotal >= n) return NULL;
    uint32_t avail = n - off - hdrTotal;
    if (imgSize == 0 || imgSize > avail) imgSize = avail;

    /* ⚠⚠ StdPicture 的位图块常见的是 **DDB (颜色位 [+ 1bpp AND 掩码])**, 不是 DIB:
     * `biHeight` 是 DDB 高度 = 图高的两倍, 于是 biSizeImage 不等于 stride*biHeight。
     * 旧实现直接拿 biHeight 当图高建 DIB ⇒ 图被拉成两倍高、掩码字节被当颜色解出来,
     * 症状是图标挤在格子底部、其余全黑 (ComCtlsDemo 的 ImageList1/2/3 全线:
     * Toolbar / TabStrip / TreeView / ListView / ImageCombo 的图标)。
     * 三种布局都用 biSizeImage 自证 (字节数互不相同), 不猜:
     *   ① stride*absH                    —— 纯 DIB, 图高 = biHeight
     *   ② stride*(absH/2) + maskStride*(absH/2) —— DDB 带 AND 掩码 (ImageList1/2)
     *   ③ stride*(absH/2)                —— 只存色行、头高度仍翻倍 (ImageList3 的 32x32)
     * 行序一律是 BMP 那套自下而上, 所以**保持 biHeight 为正**建 DIB 就不必翻转。 */
    uint32_t plainBytes = stride * absH;
    uint32_t half = (absH % 2u) ? 0u : absH / 2u;
    uint32_t colourHalf = half ? stride * half : 0u;
    uint32_t ddbBytes = half ? (colourHalf + maskStride * half) : 0u;
    uint32_t imgH = absH;
    int halfHeight = 0, hasMask = 0;
    if (bits <= 8 && imgSize != plainBytes && half) {
        /* 只有调色板图才可能是 DDB —— 下面的位操作按每字节 8/bits 个像素算, 16bpp 以上不成立 */
        if (ddbBytes && imgSize == ddbBytes) { imgH = half; halfHeight = 1; hasMask = 1; }
        else if (colourHalf && imgSize == colourHalf) { imgH = half; halfHeight = 1; }
    }

    void* bitsPtr = NULL;
    HBITMAP hbm;
    if (!halfHeight) {
        hbm = CreateDIBSection(NULL, (BITMAPINFO*)dib, DIB_RGB_COLORS, &bitsPtr, NULL, 0);
        if (!hbm || !bitsPtr) { if (hbm) DeleteObject(hbm); return NULL; }
        uint32_t copy = imgSize < plainBytes ? imgSize : plainBytes;
        if (copy) memcpy(bitsPtr, dib + hdrTotal, copy);
    } else {
        /* 掩码位 1 = 透明, 且**掩码段与色行同序** (都自下而上) —— 两条独立核出来的:
         * MainForm.frx 的第 1 枚 16x16/4bpp 图标, 图行 0 与掩码行 15 相配 (掩码行 15
         * = ff 01, 只有 x=8..15 不透明, 正好对上 VB6 截图里那一行露出的 7 个灰点),
         * 图行 10 与掩码行 5 (c0 07) 相配 (x=0/1/14/15 透明, VB6 截图那两端确为空)。
         * 透明格涂 **&HC0C0C0 银灰**: ComCtls ImageList 默认 `PropMaskColor = &HC0C0C0`,
         * 填充走 `ImageList_AddMasked(himl, .Handle, WinColor(PropMaskColor))`, 于是
         * 洞被键成透明, 与 VB6 一致。落白则整圈白方块 (工具条底色是 240 灰, 一眼可见)。
         * 已知残留: 图标**自带的银灰**与 MaskColor 同值, 那一档像素会被一起键掉
         * (实测这枚 16x16 图标 256 格里差 29 格)。改调色板第 0 项能救回, 但那样
         * 未被掩码的 index 0 (真黑) 也变成银灰一起被键掉, 实测反而差 137 格 —— 因为
         * VB6 那些像素是真黑, 靠 AddMasked 的取色键根本达不到, 要一致得让 .Handle
         * 交出带 AND 掩码的 DDB (OleCreatePictureIndirect 没有带掩码的 PICTDESC)。
         * 现状取两者里好的那个, 记在这里备查。
         * 扩成 24bpp: 原调色板里没有"透明"这一档, 且 Common.bas 的 RenderPicture
         * 见到 32bpp 就改走 GdiAlphaBlend 那一支 —— 那支在 C3 里从没跑过
         * (VB6 递的是 24bpp 位图, 见 MainForm.frx 的 Picture1), 踩了直接报错误 9 退出。
         * DIB 用正的 biHeight = 自下而上, 于是 r 同时是色行号与掩码行号。
         * ⚠ 调色板要**从 DIB 头**取 (dib+40), 不是从像素区 src —— src 已经过了头+调色板。 */
        BITMAPINFO bi24;
        memset(&bi24, 0, sizeof(bi24));
        bi24.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
        bi24.bmiHeader.biWidth = w;
        bi24.bmiHeader.biHeight = (LONG)imgH;        /* 正 = 自下而上, 与色行同序 */
        bi24.bmiHeader.biPlanes = 1;
        bi24.bmiHeader.biBitCount = 24;
        bi24.bmiHeader.biCompression = BI_RGB;
        hbm = CreateDIBSection(NULL, &bi24, DIB_RGB_COLORS, &bitsPtr, NULL, 0);
        if (!hbm || !bitsPtr) { if (hbm) DeleteObject(hbm); return NULL; }
        const uint8_t* pal = dib + 40;                 /* 调色板跟在头后面 */
        const uint8_t* src = dib + hdrTotal;
        const uint8_t* msk = hasMask ? (src + colourHalf) : NULL;
        uint8_t* dst = (uint8_t*)bitsPtr;
        uint32_t dstStride = (((uint32_t)w * 24u) + 31u) / 32u * 4u;
        for (uint32_t r = 0; r < imgH; r++) {         /* r = DIB 行, 0 = 图的底行 */
            const uint8_t* line = src + (size_t)r * stride;
            const uint8_t* mskRow = msk ? (msk + (size_t)r * maskStride) : NULL;
            uint8_t* out = dst + (size_t)r * dstStride;
            for (int32_t x = 0; x < w; x++) {
                uint8_t r8 = 192, g8 = 192, b8 = 192;   /* 透明格 = MaskColor 银灰 */
                if (!mskRow || !((mskRow[x / 8] >> (7 - (x % 8))) & 1)) {
                    if (bits <= 8) {
                        int pxPerByte = 8 / (int)bits;
                        int shift = 8 - (int)bits - (x % pxPerByte) * (int)bits;
                        uint32_t idx = (uint32_t)((line[x / pxPerByte] >> shift)
                                                  & ((1u << bits) - 1u));
                        if (idx < palEntry) {
                            b8 = pal[idx * 4 + 0];
                            g8 = pal[idx * 4 + 1];
                            r8 = pal[idx * 4 + 2];
                        }
                    } else {
                        size_t bo = (size_t)x * (bits / 8u);
                        b8 = line[bo + 0];
                        g8 = line[bo + 1];
                        r8 = line[bo + 2];
                    }
                }
                out[x * 3 + 0] = b8;
                out[x * 3 + 1] = g8;
                out[x * 3 + 2] = r8;
            }
        }
    }

    PICTDESC pd;
    memset(&pd, 0, sizeof(pd));
    pd.cbSizeofstruct = sizeof(pd);
    pd.picType = PICTYPE_BITMAP;
    pd.bmp.hbitmap = hbm;
    pd.bmp.hpal = NULL;

    static int oleInitedBag = 0;
    if (!oleInitedBag) {
        if (SUCCEEDED(OleInitialize(NULL))) oleInitedBag = 1;
        else { CoInitialize(NULL); oleInitedBag = 2; }
    }
    /* IID_IPictureDisp = {7BF80981-BF32-101A-8BBB-00AA00300CAB}
       (与 vb6forms_picture_prop.c 同一枚本地定义; 必须用 Disp 那枚 ——
        IPicture 的 vtable 槽位与 IDispatch 不同, 下游 GetIDsOfNames 会落到
        get_Type 上 ⇒ 第二十三层踩过的类型混淆) */
    static const GUID local_IID_IPictureDisp =
        {0x7BF80981, 0xBF32, 0x101A, {0x8B, 0xBB, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB}};
    IPictureDisp* pNew = NULL;
    HRESULT hr = OleCreatePictureIndirect(&pd, &local_IID_IPictureDisp, TRUE, (void**)&pNew);
    if (FAILED(hr) || !pNew) { DeleteObject(hbm); return NULL; }
    vb6_PictureRegister((void*)pNew);
    return (void*)pNew;      /* 自持一枚引用, 调用方负责 */
}

static void vb6_bag_put_u8array(Vb6PropBag* bag, const wchar_t* name,
                                const uint8_t* p, uint32_t n) {
    VARIANT v; VariantInit(&v);
    SAFEARRAYBOUND bound; bound.cElements = n; bound.lLbound = 0;
    SAFEARRAY* psa = SafeArrayCreate(VT_UI1, 1, &bound);
    if (!psa) return;
    void* dst = NULL;
    if (SUCCEEDED(SafeArrayAccessData(psa, &dst)) && dst) {
        if (n) memcpy(dst, p, n);
        SafeArrayUnaccessData(psa);
        v.vt = VT_ARRAY | VT_UI1;
        v.parray = psa;
        vb6_BagPut(bag, name, &v);   /* VariantCopy 深拷一份进表 */
        VariantClear(&v);            /* 交还本函数这一枚 (表里那份独立) */
    } else {
        SafeArrayDestroy(psa);
    }
}

/* 返回是否解析成功 (失败时不动 bag) */
static int vb6_bag_deserialize(Vb6PropBag* bag, const uint8_t* data, uint32_t n) {
    if (!bag || !data || n < 8 + 12) return 0;
    /* 找 header: 签名 0xB293 且 cbContent 正好把块补满 (实测 h 恒 = 28) */
    const uint32_t kScanMax = 64;
    const uint8_t* h = NULL;
    uint32_t csize = 0;
    uint32_t scanEnd = (n < kScanMax) ? n : kScanMax;
    for (uint32_t i = 0; i + 8 <= n && i <= scanEnd; i++) {
        if (vb6_bag_rd16(data + i) != VB6_BAG_SIGNATURE) continue;
        uint32_t cs = vb6_bag_rd32(data + i + 4);
        if (cs >= 8 && cs <= n && i + cs == n) { h = data + i; csize = cs; break; }
    }
    if (!h) return 0;

    const uint8_t* end = h + csize;
    const uint8_t* p = h + 8;          /* wCatalogOffset 实测恒 0 */
    int parsed = 0;
    int guard = 0;
    while (p + 12 <= end && guard++ < 4096) {
        uint16_t vt     = vb6_bag_rd16(p);
        uint16_t ssize  = vb6_bag_rd16(p + 2);
        uint32_t lnext  = vb6_bag_rd32(p + 8);
        uint32_t nchars = ssize & 0x3FFu;
        uint32_t nbytes = nchars * 2u;
        const uint8_t* vp = p + 12 + nbytes;
        if (p + 12 + nbytes > end || nchars == 0) break;

        /* 项名: UTF-16LE -> wchar_t* (栈上临时) */
        wchar_t name[VB6_BAG_NAME_MAX];
        uint32_t k = 0;
        for (; k < nchars && k < VB6_BAG_NAME_MAX - 1; k++)
            name[k] = (wchar_t)vb6_bag_rd16(p + 12 + k * 2);
        name[k] = 0;

        switch (vt) {
        case 2: {   /* VT_I2 */
            if (vp + 2 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_I4;
                v.lVal = (int16_t)vb6_bag_rd16(vp); vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case 3: {   /* VT_I4 */
            if (vp + 4 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_I4;
                v.lVal = (int32_t)vb6_bag_rd32(vp); vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case 4: {   /* VT_R4 */
            if (vp + 4 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_R8;
                float f; uint32_t bits = vb6_bag_rd32(vp);
                memcpy(&f, &bits, 4); v.dblVal = (double)f;
                vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case 5: {   /* VT_R8 */
            if (vp + 8 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_R8;
                double d; uint64_t bits = (uint64_t)vb6_bag_rd32(vp)
                                        | ((uint64_t)vb6_bag_rd32(vp + 4) << 32);
                memcpy(&d, &bits, 8); v.dblVal = d;
                vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case 11: {  /* VT_BOOL */
            if (vp + 2 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_BOOL;
                v.boolVal = vb6_bag_rd16(vp) ? VARIANT_TRUE : VARIANT_FALSE;
                vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case 6: {   /* VT_CY —— Currency (VB6 的 Currency 属性走这一档),
                       序列化形状 = int64 定点 (万分之一), 交回 R8 */
            if (vp + 8 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_R8;
                int64_t cy; uint64_t bits = (uint64_t)vb6_bag_rd32(vp)
                                          | ((uint64_t)vb6_bag_rd32(vp + 4) << 32);
                memcpy(&cy, &bits, 8); v.dblVal = (double)cy / 10000.0;
                vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case 7: {   /* VT_DATE */
            if (vp + 8 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_DATE;
                double d; uint64_t bits = (uint64_t)vb6_bag_rd32(vp)
                                        | ((uint64_t)vb6_bag_rd32(vp + 4) << 32);
                memcpy(&d, &bits, 8); v.date = d;
                vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case VT_DECIMAL: {  /* 14 —— 16B DECIMAL 原样落进 VARIANT, 转换交给下游读口
                       (vb6_VarToDouble_internal / TypeName / IsNumeric 都答这一档) */
            if (vp + 16 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_DECIMAL;
                memcpy(&v.decVal, vp, 16);
                vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case VT_I8: {  /* 20 —— VB6 自己写不出这一档 (无 LongLong), 但宿主/外部
                       属性包会; 原样保留, 别回落 default 抹成 0 */
            if (vp + 8 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_I8;
                int64_t q; uint64_t bits = (uint64_t)vb6_bag_rd32(vp)
                                         | ((uint64_t)vb6_bag_rd32(vp + 4) << 32);
                memcpy(&q, &bits, 8); v.llVal = q;
                vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case VT_UI4: {  /* 19 —— ai/032 §6.6: 无符号两档同理, 原样保留 */
            if (vp + 4 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_UI4;
                v.ulVal = vb6_bag_rd32(vp);
                vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case VT_UI8: {  /* 21 */
            if (vp + 8 <= end) { VARIANT v; VariantInit(&v); v.vt = VT_UI8;
                uint64_t bits2 = (uint64_t)vb6_bag_rd32(vp)
                               | ((uint64_t)vb6_bag_rd32(vp + 4) << 32);
                memcpy(&v.ullVal, &bits2, 8);
                vb6_BagPut(bag, name, &v); parsed++; }
            break; }
        case (VT_ARRAY | VT_UI1): {   /* 0x2011 */
            if (vp + 10 <= end) {
                uint32_t cbData = vb6_bag_rd32(vp + 2);
                const uint8_t* ap = vp + 10;
                if (cbData && ap + cbData <= end) {
                    BSTR s = vb6_bag_text_from_utf16le(ap, cbData);
                    if (s) {
                        VARIANT v; VariantInit(&v); v.vt = VT_BSTR; v.bstrVal = s;
                        vb6_BagPut(bag, name, &v); VariantClear(&v);
                    } else {
                        vb6_bag_put_u8array(bag, name, ap, cbData);
                    }
                    parsed++;
                }
            }
            break; }
        case 9: {   /* VT_DISPATCH —— 图片流 (见 vb6_bag_make_picture) */
            uint32_t vAvail = (lnext > 12u + nbytes) ? (lnext - 12u - nbytes)
                                                     : (uint32_t)(end - vp);
            void* pic = vb6_bag_make_picture(vp, vAvail);
            if (pic) {
                VARIANT v; VariantInit(&v);
                v.vt = VT_DISPATCH; v.pdispVal = (IDispatch*)pic;
                vb6_BagPut(bag, name, &v);       /* VariantCopy 再 AddRef 一枚给表 */
                VariantClear(&v);                /* 交还本函数这一枚 */
                parsed++;
            }
            break; }
        default:
            break;   /* 其它档暂不支持 */
        }

        if (lnext == 0) break;
        if (p + lnext <= p || p + lnext > end) break;
        p += lnext;
    }
    return parsed;
}

/* Contents 赋值时: 若是字节数组 (或 BSTR), 试按 PropertyBag 序列化格式拍平进项表 */
static void vb6_bag_try_deserialize_contents(Vb6PropBag* bag, const VARIANT* v) {
    if (!bag || !v) return;
    if (v->vt == (VT_ARRAY | VT_UI1) && v->parray) {
        void* pd = NULL;
        if (SUCCEEDED(SafeArrayAccessData(v->parray, &pd)) && pd) {
            uint32_t cb = (uint32_t)v->parray->rgsabound[0].cElements;
            vb6_bag_deserialize(bag, (const uint8_t*)pd, cb);
            SafeArrayUnaccessData(v->parray);
        }
        return;
    }
    if (v->vt == VT_BSTR && v->bstrVal) {
        vb6_bag_deserialize(bag, (const uint8_t*)v->bstrVal,
                            (uint32_t)SysStringByteLen(v->bstrVal));
    }
}

void vb6_UC_BagPutBin(void* bag, const wchar_t* name, const char* bytes, int32_t len) {
    if (!bag || !name) return;
    vb6_bag_put_u8array((Vb6PropBag*)bag, name, (const uint8_t*)(bytes ? bytes : ""),
                        (uint32_t)(len > 0 ? len : 0));
}


void vb6_UC_BagPutStr(void* bag, const wchar_t* name, const wchar_t* value) {
    if (!bag || !name) return;
    VARIANT v; VariantInit(&v);
    v.vt = VT_BSTR; v.bstrVal = SysAllocString(value ? value : L"");
    vb6_BagPut((Vb6PropBag*)bag, name, &v);
    VariantClear(&v);
}
void vb6_UC_BagPutInt(void* bag, const wchar_t* name, int32_t value) {
    if (!bag || !name) return;
    VARIANT v; VariantInit(&v);
    v.vt = VT_I4; v.lVal = value;
    vb6_BagPut((Vb6PropBag*)bag, name, &v);
    VariantClear(&v);
}
void vb6_UC_BagPutDbl(void* bag, const wchar_t* name, double value) {
    if (!bag || !name) return;
    VARIANT v; VariantInit(&v);
    v.vt = VT_R8; v.dblVal = value;
    vb6_BagPut((Vb6PropBag*)bag, name, &v);
    VariantClear(&v);
}
void vb6_UC_BagPutBool(void* bag, const wchar_t* name, int32_t value) {
    if (!bag || !name) return;
    VARIANT v; VariantInit(&v);
    v.vt = VT_BOOL; v.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    vb6_BagPut((Vb6PropBag*)bag, name, &v);
    VariantClear(&v);
}

#ifdef __cplusplus
} // extern "C"
#endif
