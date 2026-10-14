// vb6forms_uc 拆分片段 uc_collection_api (函数体)
// 内容 = 拆分前 vb6forms_uc.c 第 989~1185 行，逐行未改
// 在函数体内 #include；路径一律用 basename

void* vb6_Collection_New(void) {
    vb6_CollRec* c = (vb6_CollRec*)calloc(1, sizeof(vb6_CollRec));
    if (c) c->tag = VB6_UC_COLL_TAG;
    return c;
}

int32_t vb6_Collection_IsCollection(void* p) { return vb6_uc_isColl(p); }

static void vb6_coll_push(void* coll, const vb6_VARIANT* v, const wchar_t* key) {
    vb6_CollRec* c = (vb6_CollRec*)coll;
    if (!vb6_uc_isColl(c) || !v) return;
    if (c->count >= c->cap) {
        int32_t nc = c->cap ? c->cap * 2 : 8;
        vb6_VARIANT* ni = (vb6_VARIANT*)realloc(c->items, sizeof(vb6_VARIANT) * (size_t)nc);
        if (!ni) return;
        c->items = ni;
        wchar_t** nk = (wchar_t**)realloc(c->keys, sizeof(wchar_t*) * (size_t)nc);
        if (!nk) return;
        c->keys = nk;
        c->cap = nc;
    }
    if (!c->keys) {
        c->keys = (wchar_t**)calloc((size_t)(c->cap ? c->cap : 8), sizeof(wchar_t*));
    }
    vb6_VARIANT* dst = &((vb6_VARIANT*)c->items)[c->count];
    memset(dst, 0, sizeof(*dst));
    /* 账 #305 §B165 (第五十四刀): 改前只有串那一档另分配，其余裸结构赋值 ⇒ x86 的 Decimal
       存进来就是**别人那 16 字节的指针**，取回时早悬空(实测读回 2066035336255469780997 那一类
       数)。存值口径问那唯一一枚出口，与 vb6_Host_ClearVariant 的放那一对配平。 */
    vb6_VariantValueCopy(dst, v);
    if (c->keys) c->keys[c->count] = key ? SysAllocString(key) : NULL;
    c->count++;
}

// 实参是 Windows VARIANT* (vb6_ComPackXxx 打包), 转成 RTL 形式后存入
void vb6_Collection_Add(void* coll, const void* winVar) {
    vb6_VARIANT hv;
    vb6_Host_FromWinVariant(winVar, &hv);
    vb6_coll_push(coll, &hv, NULL);
    vb6_Host_ClearVariant(&hv);
}

// 带键 Add (VB6: 集合.Add Item, Key); 键重复时报 457
void vb6_Collection_AddKeyed(void* coll, const void* winVar, const wchar_t* key) {
    vb6_CollRec* c = (vb6_CollRec*)coll;
    if (vb6_uc_isColl(c) && key && c->keys) {
        for (int32_t i = 0; i < c->count; i++) {
            if (c->keys[i] && _wcsicmp(c->keys[i], key) == 0) {
                vb6_ErrRaiseNumber(457);
                return;
            }
        }
    }
    vb6_VARIANT hv;
    vb6_Host_FromWinVariant(winVar, &hv);
    vb6_coll_push(coll, &hv, key);
    vb6_Host_ClearVariant(&hv);
}

int32_t vb6_Collection_Count(void* coll) {
    return vb6_uc_isColl(coll) ? ((vb6_CollRec*)coll)->count : 0;
}

// Fix 134: VB6 Collection.Add 的 Before/After 位置插入。
// Charts 2020 ucTreeMaps.AddLineSeries 用 `cValues.Remove j : cValues.Add v, , i`
// 做降序排序 —— 此前 Add 忽略 Before, 排序不生效, Squarified 布局顺序与 VB6 不符。
// before1/after1 为 1 基序号; before1 使新元素占据第 before1 位, after1 插到其后。
void vb6_Collection_AddAt(void* coll, const void* winVar, int32_t pos1) {
    vb6_CollRec* c = (vb6_CollRec*)coll;
    if (!vb6_uc_isColl(c) || !winVar) return;
    vb6_Collection_Add(coll, winVar);           // 先追加到尾部
    int32_t n = c->count;
    if (n <= 1) return;
    if (pos1 < 1) pos1 = 1;
    if (pos1 > n) pos1 = n;
    vb6_VARIANT* items = (vb6_VARIANT*)c->items;
    vb6_VARIANT tmp = items[n - 1];             // 刚追加的元素
    memmove(&items[pos1 - 1 + 1], &items[pos1 - 1],
            sizeof(vb6_VARIANT) * (size_t)(n - pos1));
    items[pos1 - 1] = tmp;
    if (c->keys) {
        wchar_t* kt = c->keys[n - 1];
        memmove(&c->keys[pos1 - 1 + 1], &c->keys[pos1 - 1],
                sizeof(wchar_t*) * (size_t)(n - pos1));
        c->keys[pos1 - 1] = kt;
    }
}

void vb6_Collection_AddKeyedAt(void* coll, const void* winVar, const wchar_t* key, int32_t pos1) {
    vb6_CollRec* c = (vb6_CollRec*)coll;
    if (!vb6_uc_isColl(c) || !winVar) return;
    vb6_Collection_AddKeyed(coll, winVar, key);  // 先追加 (含键重复检查)
    int32_t n = c->count;
    if (n <= 1) return;
    if (pos1 < 1) pos1 = 1;
    if (pos1 > n) pos1 = n;
    vb6_VARIANT* items = (vb6_VARIANT*)c->items;
    vb6_VARIANT tmp = items[n - 1];
    memmove(&items[pos1 - 1 + 1], &items[pos1 - 1],
            sizeof(vb6_VARIANT) * (size_t)(n - pos1));
    items[pos1 - 1] = tmp;
    if (c->keys) {
        wchar_t* kt = c->keys[n - 1];
        memmove(&c->keys[pos1 - 1 + 1], &c->keys[pos1 - 1],
                sizeof(wchar_t*) * (size_t)(n - pos1));
        c->keys[pos1 - 1] = kt;
    }
}

void vb6_Collection_Remove(void* coll, int32_t idx1) {
    vb6_CollRec* c = (vb6_CollRec*)coll;
    if (!vb6_uc_isColl(c) || idx1 < 1 || idx1 > c->count) return;
    vb6_VARIANT* items = (vb6_VARIANT*)c->items;
    vb6_VARIANT* v = &items[idx1 - 1];
    vb6_VariantValueFree(v);   /* 账 #305 §B165: 释放口径与 vb6_coll_push 同源(对象档照旧不动) */
    if (c->keys) {
        if (c->keys[idx1 - 1]) SysFreeString(c->keys[idx1 - 1]);
        memmove(&c->keys[idx1 - 1], &c->keys[idx1], sizeof(wchar_t*) * (size_t)(c->count - idx1));
    }
    memmove(&items[idx1 - 1], &items[idx1], sizeof(vb6_VARIANT) * (size_t)(c->count - idx1));
    c->count--;
}

// 按键删除 (VB6: 集合.Remove Key)
void vb6_Collection_RemoveByKey(void* coll, const wchar_t* key) {
    vb6_CollRec* c = (vb6_CollRec*)coll;
    if (!vb6_uc_isColl(c) || !key || !c->keys) return;
    for (int32_t i = 0; i < c->count; i++) {
        if (c->keys[i] && _wcsicmp(c->keys[i], key) == 0) {
            vb6_Collection_Remove(coll, i + 1);
            return;
        }
    }
}

// Item(i): 1-based; 取回的那一格自己持有一份 (宿主管道随后按同一口径放掉)
static void vb6_coll_itemCopy(const vb6_VARIANT* src, vb6_VARIANT* out) {
    memset(out, 0, sizeof(*out));
    vb6_VariantValueCopy(out, src);   /* 账 #305 §B165: 与 vb6_coll_push 同一枚出口 */
}

void vb6_Collection_Item(void* coll, int32_t idx1, void* outV) {
    vb6_VARIANT* out = (vb6_VARIANT*)outV;
    vb6_CollRec* c = (vb6_CollRec*)coll;
    if (!vb6_uc_isColl(c) || idx1 < 1 || idx1 > c->count) {
        vb6_ho_setVariantEmpty(out);
        return;
    }
    vb6_coll_itemCopy(&((const vb6_VARIANT*)c->items)[idx1 - 1], out);
}

// 按键查找 (VB6: 集合(Key)); 未找到时报错 5 (供 On Error Resume Next + Err 判定)
int32_t vb6_Collection_ItemByKey(void* coll, const wchar_t* key, void* outV) {
    vb6_VARIANT* out = (vb6_VARIANT*)outV;
    vb6_CollRec* c = (vb6_CollRec*)coll;
    if (vb6_uc_isColl(c) && key && c->keys) {
        for (int32_t i = 0; i < c->count; i++) {
            if (c->keys[i] && _wcsicmp(c->keys[i], key) == 0) {
                vb6_coll_itemCopy(&((const vb6_VARIANT*)c->items)[i], out);
                return 1;
            }
        }
    }
    vb6_ho_setVariantEmpty(out);
    vb6_ErrRaiseNumber(5);   // 无效的过程调用或参数 (键不存在)
    return 0;
}

// cgen 对 `As New Collection` 生成 vb6_cls_<Name>_New(); 内建 Collection 无项目类,
// 以此别名桥接 (见 cgen_decl_var.cpp / cgen_localdecl.cpp 的 Collection 注册).
void* vb6_cls_Collection_New(void) { return vb6_Collection_New(); }

// For Each 枚举器: {coll, nextIdx} 两元素句柄, 语义同 Controls 枚举
// 注意: 句柄首元素存的是指针, 必须用 intptr_t 存储 —— x64 下按 int32_t 存会
// 截断高 32 位, vb6_ForEach_Next 里再当指针解引用即 0xC0000005.
void* vb6_Collection_EnumInit(void* coll) {
    intptr_t* e = (intptr_t*)malloc(sizeof(intptr_t) * 2);
    if (e) { e[0] = (intptr_t)coll; e[1] = 0; }
    return e;
}

int32_t vb6_Collection_EnumNext(void* enumPtr, void* outV) {
    vb6_VARIANT* out = (vb6_VARIANT*)outV;
    intptr_t* e = (intptr_t*)enumPtr;
    if (!e) return 0;
    vb6_CollRec* c = (vb6_CollRec*)e[0];
    if (!vb6_uc_isColl(c) || e[1] >= c->count) return 0;
    vb6_coll_itemCopy(&((const vb6_VARIANT*)c->items)[e[1]++], out);
    return 1;
}
