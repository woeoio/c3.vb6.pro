// owned_temp_keeps.hpp — 账 #454 §B164 形① 那张 KEEPS 表的**发码侧可问形态**（单一权威）
//
// 这一张表回答一个问题：**一条语句里现造的带权临时量，传进去之后谁还留着它**。
// 三条轴，各答各的：
//   keeps_pointer  指针/槽位/引用有没有**带出调用**（0 = 没带出 / 1 = 带出了 / U = 证不出）
//                 ⇒ **回收只看这一轴**（放掉它之前必须问这一句）
//   persists_data  这条形参的**数据**有没有落到比调用活得久的地方（注册表写 / 文件写 / 剪贴板）
//                 ⇒ 这一轴**不参与回收**（数据持久化 ≠ 指针留存）
//   gate           closed 的行**不进提名面**（注册表那一族：运行期破坏未归因 ⇒ 停机）
// 另有一张**借用轴**（返回侧）：这一枚出口**交出去的指针是借来的**，不归调用方所有 ⇒ 调用方不许放它。
//
// **为什么要两份形态**（一张表在 scripts/ 里、一份在 src/common/ 里）：
//   * `scripts/owned_temp_keeps_164.txt` = **推导出来**的整张表（出口面出自 RTL 的 24 个权威头，
//     由 `scripts/check_owned_temp_keeps.ps1` 现算）⇒ 它是**证据**。
//   * 本头 = **给发码侧问的那一份**（只有例外行：KEEPS=1 / U / closed / 借用族）⇒ 它是**出口**。
//   两者由那道哨兵**双向**对账（少一行/多一行/分类不同都红）⇒ 漂不了。
//   这与仓里既有的做法同形：`kBorrowedBstrExports167` 是手写的**出口**，
//   `check_bstr_borrow_slot.ps1` 的 Q6 从 RTL **推导**后双向对账。
//
// **本题的由来**：撤回那一版（9546dbfe / 6bf5989 / be5d6abf）**只问了 mint 那一侧**
// （「这一格是不是现造的带权临时量」），**没问 borrow 那一侧**（「这一格是借来的吗」），
// 也**从没问消费者那一侧**（「这个消费者会不会把指针留下来」）⇒ 它把
// `vb6_BSTR_Concat` 换成 `vb6_BSTR_ConcatFree`，放掉了**借来的** `vb6_ErrDescription()`
// （RTL 全局 vb6_err 的指针）⇒ 全局悬空 ⇒ 读数变堆垃圾（`test_bstr_borrow_slot`）、
// `test_host_decimal`(x86) 直接 0xC0000374。归因见台账 §B227。
//
// 规矩：**消费者那一侧答不出（U）⇒ 不许提名**；**借用族 ⇒ 一律不许放**；
// **closed ⇒ 一律不许提名**（未归因 ⇒ 停机）。
#pragma once
#include <cstddef>
#include <cstring>

namespace vb6c3 {

// 一行 = 一枚例外（出口名 + 形参序号 + 分类 + 逃逸通道）
struct OwnedTempException {
    const char* exitName;   // 出口名（vb6_*）
    int         slot;       // 形参序号（0 起）
    char        keeps;      // '1' = 指针带出调用 / 'U' = 证不出
    const char* escape;     // 逃逸通道：store / ret / write / handoff-*（为什么是这一枚）
};

// 例外行：**只有**这些不许提名（其余的出口默认 keeps_pointer=0 ⇒ 提名面内）。
// 26 枚 KEEPS=1 + 1 枚 U（vb6_DecSetTag 第 0 格）—— 与台账 §B227 的表逐行对齐。
inline const OwnedTempException* ownedTempExceptions() {
    static const OwnedTempException kRows[] = {
        { "vb6_ArraySetVariant",  2, '1', "write" },
        { "vb6_BSTR_AssignMove",  1, '1', "write" },
        { "vb6_CDec",             0, '1', "ret" },
        { "vb6_DecimalAsVariant", 0, '1', "write" },
        { "vb6_DecimalToBSTR",    0, '1', "store" },
        { "vb6_DecimalToR8",      0, '1', "store" },
        { "vb6_DecSetTag",        0, 'U', "handoff-unknown:vb6_DEC_VT" },
        { "vb6_ErrRaise",         1, '1', "store" },
        { "vb6_ErrRaise",         2, '1', "store" },
        { "vb6_ho_setVariantBstr",1, '1', "store" },
        { "vb6_IIfVariant",       1, '1', "ret+write" },
        { "vb6_IIfVariant",       2, '1', "ret+write" },
        { "vb6_LoadResData",      0, '1', "write" },
        { "vb6_LoadResData",      1, '1', "write" },
        { "vb6_LoadResPicture",   0, '1', "write" },
        { "vb6_LoadResPicture",   1, '1', "write" },
        { "vb6_LoadResString",    0, '1', "write" },
        { "vb6_RaiseError",       1, '1', "store" },
        { "vb6_SendKeys",         0, '1', "write" },
        { "vb6_VariantArrayGetVal",0, '1', "write" },
        { "vb6_VariantArraySet",  2, '1', "write" },
        { "vb6_VariantAssign",    1, '1', "write" },
        { "vb6_VariantIdentity",  0, '1', "ret" },
        { "vb6_VariantString",    0, '1', "store" },
        { "vb6_VariantToBool",    0, '1', "store" },
        { "vb6_VariantToLong",    0, '1', "store" },
        { "vb6_VariantToLongPtr", 0, '1', "write" },
        { nullptr, 0, '0', nullptr },   // 空哨兵：遍历靠它收尾
    };
    return kRows;
}

// closed 的那一族（注册表）：运行期破坏未归因 ⇒ 不进提名面。
// 名单与 scripts/owned_temp_keeps_164.txt 的 gate=closed 那一节逐行对齐（哨兵双向对账）。
inline const char* const* ownedTempClosedExits() {
    static const char* const kClosed[] = {
        "vb6_SaveSetting", "vb6_GetSetting", "vb6_DeleteSetting", "vb6_GetAllSettings",
        0,                                // 空哨兵（遍历靠它收尾；0 = 空指针常量）
    };
    return kClosed;
}

// 借用族（返回侧）：这一枚出口**交出去的指针是借来的** ⇒ 调用方不许放它。
// 与编译器侧 kBorrowedBstrExports167 同名单（哨兵双向对账；那份名单本身由
// check_bstr_borrow_slot.ps1 的 Q6 从 RTL 双向推导钉住）。
inline const char* const* ownedTempBorrowedExits() {
    static const char* const kBorrowed[] = {
        "vb6_ErrDescription", "vb6_ErrSource", "vb6_ho_variantToBstr",
        0,                                // 空哨兵（遍历靠它收尾；0 = 空指针常量）
    };
    return kBorrowed;
}

// ---------------------------------------------------------------------------
// **消费轴**（账 #463 §B164 研究新增）：消费者**自己放掉**自己的形参的那几格。
//
// 为什么要有这一轴：keeps_pointer 答的是「指针有没有**带出**调用」（留下那一面），
// 它**没有**答「消费者会不会**自己放掉**这一格」。那两件事不同：
//     keeps_pointer = '1' ⇒ 消费者把指针留在外面 ⇒ 调用方不能放（否则双放/悬空）
//     consumes      = 1   ⇒ 消费者自己放掉这一格 ⇒ 调用方**也不能**再放（否则**双放** ⇒ 堆损坏）
// 例：vb6_BSTR_ConcatFree 放掉它的**左**形参（vb6rtl_string.c:51 `vb6_BSTR_Free(a)`）——
//     照旧表（那一格 keeps=0）提名，就是「ConcatFree 放一次 + 语句末再放一次」⇒ 双放。
// 名单由 scripts/check_owned_temp_keeps.ps1 从 RTL **实读**（函数体里对自己的形参调
// vb6_BSTR_Free / SysFreeString / vb6_VariantClear）⇒ 与 scripts/owned_temp_keeps_164.txt
// 的 [consumes] 节**双向对账** ⇒ 漂不了。名单是**数据**不是逻辑：加一格 = 加一行。
//
// 一格 = (出口名, 形参序号)；「条件性」那一列只作记录：条件性的那几格（IIf 那一族）
//     放不放**看运行期分支** ⇒ 调用方无从知道 ⇒ 两格都**不许**提名。
struct OwnedTempConsumes {
    const char* exitName;   // 出口名（vb6_*）
    int         slot;       // 形参序号（0 起）
    char        conditional; // '0' = 无条件放掉；'1' = 看运行期分支（调用方无从知道）
};
inline const OwnedTempConsumes* ownedTempConsumes() {
    static const OwnedTempConsumes kConsumes[] = {
        { "vb6_BSTR_ConcatFree", 0, '0' },
        { "vb6_BSTR_Free",       0, '0' },
        { "vb6_IIfBSTR",         1, '1' },
        { "vb6_IIfBSTR",         2, '1' },
        { "vb6_IIfVariant",      1, '1' },
        { "vb6_IIfVariant",      2, '1' },
        { "vb6_LSetFree",        0, '0' },
        { "vb6_VariantAssign",   0, '0' },
        { nullptr, 0, '0' },    // 空哨兵：遍历靠它收尾
    };
    return kConsumes;
}

// 这一格会不会被**消费者自己放掉**？会 ⇒ 调用方不许再放（否则双放 ⇒ 堆损坏）。
inline bool ownedTempConsumerConsumes(const char* exitName, int slot) {
    if (!exitName) return true;
    for (const OwnedTempConsumes* c = ownedTempConsumes(); c->exitName; ++c) {
        if (c->slot == slot && std::strcmp(c->exitName, exitName) == 0) return true;
    }
    return false;
}

// 消费者那一侧的答案：这一枚出口的**这一格**会不会把指针带出调用？
//   '0' = 不会（可提名）  '1' = 会（禁止提名）  'U' = 证不出（禁止提名）
inline char ownedTempKeepsPointerKind(const char* exitName, int slot) {
    if (!exitName) return 'U';
    for (const OwnedTempException* r = ownedTempExceptions(); r->exitName; ++r) {
        if (r->slot == slot && std::strcmp(r->exitName, exitName) == 0) return r->keeps;
    }
    return '0';
}

// 这一族进不进提名面？closed ⇒ 不进（未归因 ⇒ 停机）。
inline bool ownedTempConsumerClosed(const char* exitName) {
    if (!exitName) return true;
    for (const char* const* p = ownedTempClosedExits(); *p; ++p) {
        if (std::strcmp(*p, exitName) == 0) return true;
    }
    return false;
}

// 这一枚出口交出去的是**借来的**指针吗？借来的 ⇒ 调用方不许放它。
inline bool ownedTempExitIsBorrowed(const char* exitName) {
    if (!exitName) return true;
    for (const char* const* p = ownedTempBorrowedExits(); *p; ++p) {
        if (std::strcmp(*p, exitName) == 0) return true;
    }
    return false;
}

// **返回侧的所有权**（账 #456 §B227 第五格）：这一枚出口交出去的是不是**新造且归调用方所有**？
//   owned = 新造（调用方**可以**放它）· alias = 别名（指着调用方的内存或一个全局 ⇒ 不许放）
//            U = 答不出（安全默认：当作不是新造的 ⇒ 不许回收）
// 回收刀的白名单就是这一份 owned 名单：只有**从这些出口出来的**值才谈得上「现造」，
// 其余一律不许碰 —— 撤回那一版把「顶层是个调用」当成「现造」⇒ 放掉了借来的那一枚 ⇒ 堆损坏。
// 这份名单由 `scripts/check_owned_temp_keeps.ps1` 从 RTL **推导**（mint ⇒ owned；返回参数/全局
// 的成员 ⇒ alias；答不出 ⇒ U），推导结果落在 scripts/owned_temp_keeps_164.txt 的 [own] 节，
// 哨兵 **K13** 与本名单**双向**对账 ⇒ 漂不了。名单是**数据**不是逻辑：加一档 = 加一行。
inline const char* const* ownedTempOwnedReturnExits() {
    static const char* const kOwned[] = {
        // >>> GENERATED-OWNED-EXITS <<<（推导结果见 scripts/owned_temp_keeps_164.txt 的 [own] 节；哨兵 K13 双向对账）
        "vb6_App_EXEName",
        "vb6_App_HelpFile",
        "vb6_App_Path",
        "vb6_BSTR_Concat",
        "vb6_BSTR_ConcatFree",
        "vb6_BSTR_Empty",
        "vb6_BSTR_FixedSTR",
        "vb6_BSTR_FromANSIBuf",
        "vb6_BSTR_FromBSTR",
        "vb6_BSTR_FromStr",
        "vb6_BSTR_FromStrN",
        "vb6_ByteArrayToString",
        "vb6_CallByName",
        "vb6_Chr",
        "vb6_ChrB",
        "vb6_ChrW",
        "vb6_Clipboard_GetText",
        "vb6_Command",
        "vb6_ConReadKey",
        "vb6_ConReadLine",
        "vb6_Console_ReadKey",
        "vb6_Console_ReadLine",
        "vb6_CStr",
        "vb6_CStrBool",
        "vb6_CStrByte",
        "vb6_CStrDate",
        "vb6_CStrDateFromBSTR",
        "vb6_CStrDbl",
        "vb6_CStrDblFromBSTR",
        "vb6_CStrDblFromVariant",
        "vb6_CStrLong",
        "vb6_CStrLongFromVariant",
        "vb6_CStrLongLong",
        "vb6_CStrSingle",
        "vb6_CStrULong",
        "vb6_CStrULongFromVariant",
        "vb6_CStrULongLong",
        "vb6_CStrULongLongFromVariant",
        "vb6_CurDir",
        "vb6_CVErr",
        "vb6_DecimalToBSTR",
        "vb6_Dir",
        "vb6_Environ",
        "vb6_FormatCurrency",
        "vb6_FormatDateTime",
        "vb6_FormatNumber",
        "vb6_FormatPercent",
        "vb6_GetSetting",
        "vb6_Hex",
        "vb6_InputB",
        "vb6_InputBox",
        "vb6_InputString",
        "vb6_Join",
        "vb6_LCase",
        "vb6_LCase_str",
        "vb6_Left",
        "vb6_LeftB",
        "vb6_LineInput",
        "vb6_LoadResString",
        "vb6_LSet",
        "vb6_LSetFree",
        "vb6_LTrim",
        "vb6_LTrimVar",
        "vb6_Mid",
        "vb6_MidB",
        "vb6_MonthName",
        "vb6_NumToBSTRDefault",
        "vb6_Oct",
        "vb6_PA_GetBSTR",
        "vb6_Partition",
        "vb6_Replace",
        "vb6_Right",
        "vb6_RightB",
        "vb6_RSet",
        "vb6_RTrim",
        "vb6_Space",
        "vb6_Spc",
        "vb6_Str",
        "vb6_StrConv",
        "vb6_StrConvFromByteArray",
        "vb6_String",
        "vb6_StrReverse",
        "vb6_Tab",
        "vb6_Trim",
        "vb6_TypeName",
        "vb6_UCase",
        "vb6_UCase_str",
        "vb6_VarAddV",
        "vb6_VarDivV",
        "vb6_VariantBool",
        "vb6_VariantByte",
        "vb6_VariantDate",
        "vb6_VariantDouble",
        "vb6_VariantEmpty",
        "vb6_VariantFromComResult",
        "vb6_VariantInt",
        "vb6_VariantLong",
        "vb6_VariantLongLong",
        "vb6_VariantNull",
        "vb6_VariantSByte",
        "vb6_VariantSingle",
        "vb6_VariantToString",
        "vb6_VariantUInteger",
        "vb6_VariantULong",
        "vb6_VariantULongLong",
        "vb6_VarMulV",
        "vb6_VarSubV",
        "vb6_WeekdayName",
        // >>> END-GENERATED-OWNED-EXITS <<<
        0,                                // 空哨兵（遍历靠它收尾；0 = 空指针常量）
    };
    return kOwned;
}

// 这一枚出口交出去的是**新造且归调用方所有**吗？不在 owned 名单里（答不出 / 是别名）⇒ 不许回收。
inline bool ownedTempReturnIsOwned(const char* exitName) {
    if (!exitName) return false;
    for (const char* const* p = ownedTempOwnedReturnExits(); *p; ++p) {
        if (std::strcmp(*p, exitName) == 0) return true;
    }
    return false;
}

// **owned 名单里那些交出 VARIANT 的一档**（其余 owned 交出 BSTR）。
// 档号**只由出口的返回档定** —— 源表达式的 VB 档不作数：`Len(v)` 的源档是 Variant，
// 交出去的却是 BSTR；照源档声明就是 `vb6_VARIANT t = <BSTR>` ⇒ cl C2440（账 #454 的教训）。
// 这张名单是 owned 那张的**子集**（哨兵 K13 双向对账 ⇒ 漂不了）。
inline const char* const* ownedTempOwnedVariantExits() {
    static const char* const kOwnedVariant[] = {
        "vb6_CallByName",
        "vb6_CVErr",
        "vb6_LoadResString",
        "vb6_VarAddV",
        "vb6_VarDivV",
        "vb6_VarMulV",
        "vb6_VarSubV",
        "vb6_VariantBool",
        "vb6_VariantByte",
        "vb6_VariantDate",
        "vb6_VariantDouble",
        "vb6_VariantEmpty",
        "vb6_VariantFromComResult",
        "vb6_VariantInt",
        "vb6_VariantLong",
        "vb6_VariantLongLong",
        "vb6_VariantNull",
        "vb6_VariantSByte",
        "vb6_VariantSingle",
        "vb6_VariantUInteger",
        "vb6_VariantULong",
        "vb6_VariantULongLong",
        0,                                // 空哨兵（遍历靠它收尾；0 = 空指针常量）
    };
    return kOwnedVariant;
}

// 这一枚 mint 交出来的是哪一档？'B' = BSTR / 'V' = VARIANT / 0 = 不是新造的（不许回收）。
inline char ownedTempMintKind(const char* exitName) {
    if (!exitName) return 0;
    for (const char* const* p = ownedTempOwnedVariantExits(); *p; ++p) {
        if (std::strcmp(*p, exitName) == 0) return 'V';
    }
    for (const char* const* p = ownedTempOwnedReturnExits(); *p; ++p) {
        if (std::strcmp(*p, exitName) == 0) return 'B';
    }
    return 0;
}

// **提名这一枚实参吗？** —— 回收点只许问这一位主人，不再自己判。
//   返回 'B' / 'V' = 可提名（档号）  0 = 不许提名
// 四道闸，各答各的（账 #458 §B228 把前三道收成一位主人；账 #463 §B164 研究补上第四道）：
//   ① 档号：这一枚是不是**现造**的（出口交出来的是不是新造且归我们所有）—— 答不出 ⇒ 不提名
//      （这挡掉「借来的那一枚」：Err.Description / 数组元素读 / 任何别名；
//        账 #463 补上「返回值是形参本身」那几形：vb6_IIfBSTR / vb6_IIfVariant /
//        vb6_CDec / vb6_VariantIdentity —— 它们交回的是**别人手里**的那一枚）
//   ② 消费者：消费者**那一格**会不会把指针带出调用 —— '1' 或 'U' ⇒ 不提名
//   ③ 那一族：closed（注册表那一族，运行期破坏未归因）⇒ 不提名 ⇒ **预测集合与 closed 必不相交**
//   ④ **消费轴**（账 #463 新增）：消费者会不会**自己放掉**这一格 —— 会 ⇒ 不提名
//      （ConcatFree 放掉左操作数 / LSetFree 放掉被对齐的那一枚 / IIf 那一族**条件性**地
//        放掉没选中的那一支 ⇒ 调用方无从知道是哪一支 ⇒ 两格都不许提名）
// 另有一条由 §B227 的归因加进来的硬规则，落在**站点**那一侧：
//   **放掉一次的责任只许一处** —— 拼接站（`vb6_BSTR_ConcatFree` 放掉右操作数）与语句末的按档回收
//   **不许对同一枚临时量各放一次**（撤回那一版就是这么双放的 ⇒ 堆损坏 ⇒ 注册表那一族的红是它的下游受害者）。
//   ⇒ 实参站提名的那一枚，**只许**由语句末回收；拼接站自己放掉的那一枚，**不进**名单。
inline char ownedTempNominationVerdict(const char* consumer, int slot, const char* mintCall) {
    if (!consumer || !mintCall) return 0;
    const char kind = ownedTempMintKind(mintCall);
    if (!kind) return 0;                                  // ① 不是现造的（含借来的/别名）⇒ 不许放
    const char keeps = ownedTempKeepsPointerKind(consumer, slot);
    if (keeps != '0') return 0;                           // ② 消费者留不留 ⇒ '1'/'U' 不许放
    if (ownedTempConsumerClosed(consumer)) return 0;       // ③ closed ⇒ 不许提名
    if (ownedTempConsumerConsumes(consumer, slot)) return 0; // ④ 消费者自己放掉 ⇒ 不许再放（双放）
    return kind;                                          // ⇒ 可提名（档号由出口的返回档定）
}
}  // namespace vb6c3
