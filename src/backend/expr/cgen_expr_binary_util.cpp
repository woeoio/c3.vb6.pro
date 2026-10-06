#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>
#include <cstdio>

namespace vb6c3 {

// --- cgen_expr_binary_util.cpp: BSTR 包装 + 二元运算符映射 ---
// 由 src/backend/expr/cgen_expr_binary.cpp 拆出（2026-09-17），纯搬移、零行为改动。

// Fix 100a: 返回 BSTR 的内建"顶层"清单 — 表达式以此开头即为已是 BSTR, 无需包装.
// 这是 092a 修法的同族应用: 用**顶层前缀**替代旧的**未锚定子串**判定.
// 数组有序按最长前缀优先匹配无影响 (全部互不为前缀).
static const char* kBstrReturningCalls[] = {
    "vb6_BSTR_",            // vb6_BSTR_FromStr / vb6_BSTR_Empty / vb6_BSTR_Concat ... 全族
    "vb6_CStr",             // vb6_CStr / vb6_CStrLong / vb6_CStrDbl / vb6_CStrBool ...
    "vb6_ChrW(", "vb6_Chr(",
    "vb6_Mid(", "vb6_Left(", "vb6_Right(", "vb6_LTrim(", "vb6_RTrim(", "vb6_Trim(",
    "vb6_UCase(", "vb6_LCase(",
    "vb6_Format(", "vb6_StrConv(", "vb6_Str(", "vb6_String(", "vb6_Space(",
    "vb6_Replace(",
    "vb6_Dir(", "vb6_Command(", "vb6_Environ(", "vb6_CurDir(", "vb6_InputBox(",
    "vb6_Console_ReadLine(", "vb6_Console_ReadKey(",  // Fix 161: Console 输入返回 BSTR
    "vb6_ErrDescription(", "vb6_ErrSource(",
    "vb6_App_Path(", "vb6_App_EXEName(", "vb6_App_HelpFile(",
    "vb6_GetControlText(", "vb6_GetControlCaption(",
    // Fix 194: 下列 RTL 取值函数的 C 声明是 `void*`(RTL 统一风格), **语义上是 BSTR**。
    // 漏登记会被判成 Variant → `vb6_CStr(vb6_VariantFromValue(void*))` → _Generic 无
    // void* 分支, 落到 default `vb6_VariantObject` → CStr 得空串。
    // 实测: Form1.frm 的 `B = List2.List(J)` 恒不相等 ("aaa" vs ""), 对比数据把所有项
    // 都当成"未找到"塞进 List3 (应为 List4)。
    "vb6_GetListItem(",
    "vb6_GetSelText(", "vb6_GetToolTipText(", "vb6_GetMenuCaption(",
    // P20-39: ListImage.Key (原生复刻的 ImageList) —— 同上, C 签名是 void*、语义是 BSTR。
    "vb6_GetImageListKeyAt(",
    "vb6_GetImageListKeyByKey(",
    // P20-40: StatusBar 的 Panels 取值函数同样是 `void*` 声明 / BSTR 语义
    // (vb6_StatusBar_GetPanelText / Key / ToolTip / SimpleText ...) —— 漏登记
    // 的表现很隐蔽: 数字类属性 (Count/Style/Width/MinWidth) 全对, 只有所有
    // 字符串恒为空串, 一眼看去还像"面板没建起来"。
    "vb6_StatusBar_GetPanelText",
    "vb6_StatusBar_GetPanelTextByKey",
    "vb6_StatusBar_GetPanelKey",
    "vb6_StatusBar_GetPanelKeyByKey",
    "vb6_StatusBar_GetPanelToolTip",
    "vb6_StatusBar_GetSimpleText",
    // P20-42: SSTab 的 TabCaption(i) —— C 签名 `void*`、语义 BSTR, 同一类。
    // 漏登记时 Debug.Print 会打出空串, 看着像"页标题没设进去"。
    "vb6_SSTab_GetTabCaption",
    // Task #44: SSTab 的 TabToolTipText(i) —— 同类 (void* 声明 / BSTR 语义)。
    // 漏登记的表现: Set 后立刻 Get 打出空串 (Variant 包装吞掉了真 BSTR)。
    "vb6_SSTab_GetTabToolTipText",
    // P20-44: OLE 拖放 DataObject 取值 —— 同类 (void* 声明 / BSTR 语义)
    "vb6_oleDD_GetText",
};

bool CCodeGen::isBstrReturningCall(const std::string& expr) {
    for (const char* pfx : kBstrReturningCalls) {
        if (expr.compare(0, strlen(pfx), pfx) == 0) return true;
    }
    return false;
}

// Fix 159-B (声明见 cgen.hpp / cgen_helpers.inc)
bool CCodeGen::isScalarImplicitStr159(Vb6Type t) {
    switch (t) {
        case Vb6Type::Date:
        case Vb6Type::Double:
        case Vb6Type::Single:
        case Vb6Type::Long:
        case Vb6Type::Integer:
        case Vb6Type::Byte:
        case Vb6Type::Boolean:
        case Vb6Type::Currency:  return true;
        default:                 return false;
    }
}

std::string CCodeGen::wrapToBSTR(const std::string& expr, Expr& node) {
    // P24-01: 后期绑定COM调用返回VARIANT*, 需解包为BSTR(必须在vb6_BSTR检查之前)
    // P24-02: Variant数组索引返回vb6_VARIANT, 需转BSTR
    if (expr.find("vb6_VariantArrayGet") == 0) {
        return "vb6_VariantToString(" + expr + ")";
    }
    if (expr.find("vb6_ComCall(") == 0) {
        return "vb6_VariantToString(vb6_VariantFromComResult(" + expr + "))";
    }
    // Fix 092w: 链式默认属性访问结果 (vb6_VariantFromComResult(vb6_ComCall(...))) 是
    // vb6_VARIANT — BSTR 上下文需提取为字符串, 否则 C2440 (Demo_Database 506)。
    // 必须放在 vb6_BSTR 子串检查之前 (实参中可能含 vb6_BSTR_FromStr)。
    if (expr.find("vb6_VariantFromComResult(") == 0) {
        return "vb6_VariantToString(" + expr + ")";
    }
    // P24-01: COM属性返回int/double, 需转BSTR
    // ⚠ 必须是**顶层前缀**判定 (与 Fix 100a / 100d 同一族修法)。原来这里是未锚定的
    // find(): 实参里带一个 COM 整数读的法向调用 —— 实测 `TF(tv1.Nodes.Count = 3)`,
    // 其中 TF 是 `Function TF(...) As String` —— 会整条命中, 于是返回 String 的项目
    // 函数被当数值处理, 发成 `vb6_CStrLong(vb6_TF(...))`: 把 BSTR 指针当数字打出来
    // (读数全是一串 -99xxxxxxx)。
    auto topHead = [&](const char* h) {
        return expr.compare(0, strlen(h), h) == 0;
    };
    if (topHead("vb6_ComGetIntProp(") || topHead("vb6_ComVtableGetInt(")) {
        return "vb6_CStrLong(" + expr + ")";
    }
    if (topHead("vb6_ComGetDoubleProp(") || topHead("vb6_ComVtableGetDouble(")) {
        return "vb6_CStrDbl(" + expr + ")";
    }
    // Fix 100a (原 092a 同族): 此处原为 expr.find("vb6_BSTR") != npos 的**未锚定子串**
    // 检查, 实参中含 vb6_BSTR_FromStr 就会整体命中 → vb6_Len("abc") / vb6_InStr(...) /
    // vb6_Val(...) / vb6_Asc(...) 等"带字符串实参的标量函数"被判为"已是 BSTR"直通,
    // 绕过后面的数值清单与类型推断 (整数型静默误编→0xC0000005, 浮点型 C2440)。
    // 判定改为**顶层前缀**: 只有表达式本身就是 BSTR 返回函数调用才直通 (092a 的修法)。
    if (isBstrReturningCall(expr)) return expr;
    // Fix 049a: VB6_SA_AT(BSTR, arr, idx) returns BSTR — no wrapping needed
    if (expr.compare(0, 16, "VB6_SA_AT(BSTR,") == 0) return expr;
    // vb6_Now() returns double (Date), NOT BSTR — removed early return
    // vb6_Now will fall through to inferExprType → Vb6Type::Date → vb6_CStrDate()
    // BSTR变量: 已知BSTR变量或者vb6_Module1_xxx 格式的BSTR
    // 简化: 如果以vb6_开头且非数值函数, 假定是BSTR
    if (expr.compare(0, 4, "vb6_") == 0) {
        // Fix 100c (B3): 取顶层函数名精确匹配, 替换原未锚定 find() ——
        // 原 `expr.find("vb6_Time")` 会把 vb6_Timer 吞进 CStrDate 分支;
        // 数值清单同样未锚定且一律归 vb6_CStrLong, 导致 Sqr/Abs/CDbl/CSng/
        // Timer/Rnd/Fix/Int/Val 的 double/float 返回值被截断 (MSVC C4244,
        // "x=" & Sqr(2) 输出 "1" 而非 "1.4142135623731"). 按 src/rtl 真实
        // 返回类型拆成 CStrDbl / CStrLong 两组.
        const size_t pfn = expr.find('(');
        const std::string fn = (pfn != std::string::npos) ? expr.substr(0, pfn) : expr;
        // Date类: vb6_Now/vb6_Date/vb6_Time 返回double(Date)
        if (fn == "vb6_Now" || fn == "vb6_Date" || fn == "vb6_Time") {
            return "vb6_CStrDate(" + expr + ")";
        }
        // 数值函数: RTL 返回 double/float → vb6_CStrDbl
        if (fn == "vb6_Abs" || fn == "vb6_CDbl" || fn == "vb6_CSng" ||
            fn == "vb6_Timer" || fn == "vb6_Rnd" || fn == "vb6_Sqr" ||
            fn == "vb6_Fix" || fn == "vb6_Int" || fn == "vb6_Val") {
            return "vb6_CStrDbl(" + expr + ")";
        }
        // 数值函数: RTL 返回整数 → vb6_CStrLong
        if (fn == "vb6_CLng" || fn == "vb6_CInt" || fn == "vb6_CBool" ||
            fn == "vb6_CByte" || fn == "vb6_Len" || fn == "vb6_LenB" ||
            fn == "vb6_InStr" || fn == "vb6_InStrRev" || fn == "vb6_Sgn" ||
            fn == "vb6_VarType") {
            return "vb6_CStrLong(" + expr + ")";
        }
        // Fix 091i: 项目内返回 Variant 的函数 (driver 预扫描 variantReturnFuncs_,
        // 含 vb6_<cls>_prop_get_<name>) 不能按"其他 vb6_ 函数假定为 BSTR"直通 —
        // 需 vb6_VariantToString 提取. 此前字符串拼接 / BSTR 形参处生成
        // vb6_BSTR_Concat(L"...", vb6_cDataBase_LastInsertId(...)) C2440
        // (Demo_Database.c 310/506/678).
        if (variantReturnFuncs_) {
            size_t p091i = expr.find('(');
            if (p091i != std::string::npos
                && variantReturnFuncs_->count(expr.substr(0, p091i))) {
                return "vb6_VariantToString(" + expr + ")";
            }
        }
        // Fix 100b: 原为 `return expr; // 其他vb6_函数假定为BSTR` —— 清单外且无字符串
        // 实参的标量返回函数 (vb6_ErrNumber / vb6_UBound / vb6_Screen_Width 等) 被
        // 直接当 BSTR 直通: 整数型只报 C4047 警告→运行期把整数当指针 0xC0000005,
        // 浮点型 C2440 编译失败. 兜底不再假定 BSTR, 落到下方 inferExprType(node):
        //   识别出具体类型 → 对应 vb6_CStrXxx (Err.Number 已由 cgen_util_type.cpp:109
        //                    判为 Vb6Type::Long);
        //   识别不出 → Variant → vb6_CStr(vb6_VariantFromValue(expr)), 由 _Generic
        //              按实参 C 类型分派, 对 BSTR 返回同样安全.
        // 即不再存在"静默传标量", 风险从"猜错就 UB"变为"多套一层转换".
    }
    // string literal L"..." —— 必须**顶层前缀**判定 (Fix 100d, B1-A1 同族收尾):
    // 原为未锚定 find(), 实参含字符串字面量的标量函数 (vb6_Asc(vb6_BSTR_FromStr(L"A")))
    // 会整体命中而直通, 绕过数值清单与 inferExprType → 整数当指针传, 运行期 0xC0000005
    // (与 Fix 100a 同一处缺陷类, A-1 只改了前者, 此处的未锚定检查当时被漏掉).
    if (expr.compare(0, 17, "vb6_BSTR_FromStr(") == 0) return expr;
    // 推断类型
    Vb6Type t = inferExprType(node);
    switch (t) {
        case Vb6Type::String: return expr;
        case Vb6Type::Integer:
        case Vb6Type::Long:   return "vb6_CStrLong(" + expr + ")";
        // Fix 084m: LongLong 走 64 位格式化 —— 落到 default 的 vb6_CStrLong 会先截成
        // int32_t (实测 4000000000 → -294967296)。
        case Vb6Type::LongLong: return "vb6_CStrLongLong(" + expr + ")";
        // LongPtr 复用 LongLong 的 64 位格式化: x64 下 intptr_t 就是 int64_t; x86 下
        // intptr_t 为 32 位, 但该分支只在 LongPtr 变量作字符串拼接时命中, 传 intptr_t
        // 给 int64_t 形参是合法提升, 无截断风险 (x86 值本来就只有 32 位)。
        case Vb6Type::LongPtr:  return "vb6_CStrLongLong((int64_t)" + expr + ")";
        case Vb6Type::Single:
        case Vb6Type::Double: return "vb6_CStrDbl(" + expr + ")";
        case Vb6Type::Boolean: return "vb6_CStrBool(" + expr + ")";
        case Vb6Type::Byte:   return "vb6_CStrByte(" + expr + ")";
        case Vb6Type::Date:   return "vb6_CStrDate(" + expr + ")";
        case Vb6Type::Variant: {
            // Fix 049a: inferExprType falls back to Variant for unknown symbols.
            // If the AST node is a known BSTR variable, return as-is to avoid
            // generating vb6_CStr(BSTR_expr) which causes C2440 (BSTR→VARIANT).
            // For all other cases (including actual Variant variables), keep
            // vb6_CStr(expr) which correctly converts VARIANT→BSTR.
            if (node.kind == ASTNodeKind::IdentifierExpr) {
                auto& id = static_cast<IdentifierExpr&>(node);
                std::string lower = id.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                if (knownBstrVars_.count(lower)) return expr;
            }
            // Fix 084: inferExprType 回退 Variant 的表达式可能是标量(int32_t等)、
            // BSTR 或已是 vb6_VARIANT. 直接 vb6_CStr(expr) 会因实参类型不匹配触发
            // C2440 (如 vb6_CStr((int32_t)GetCurrentThreadId())). 用 vb6_VariantFromValue
            // 按 C 实参类型 _Generic 自动包装: 标量→VariantLong, BSTR→VariantString,
            // 已是 vb6_VARIANT→identity 直通, 再交给 vb6_CStr 统一转 BSTR.
            return "vb6_CStr(vb6_VariantFromValue(" + expr + "))";
        }
        default: return "vb6_CStrLong(" + expr + ")";  // fallback
    }
}

// Fix 194: 字符串型控件属性 (Caption/Text/ToolTipText/MenuCaption) 的写入值必须转 BSTR。
//
// 现象 (用户报的 Form1.frm): `Label1(0).Caption = List1.ListCount` 点一下就
// 0xC000041D (用户回调未处理异常, 进程直接消失); `Label1(1).Caption = List4.ListCount`
// 当 List4 为空时"没反应"(不崩也不改)。
//
// 根因: 生成 `vb6_SetControlText(vb6_CtrlArr_GetAt(&vb6_arr_Label1, 0), vb6_GetListCount(...))`
// —— 第二参形参是 `void* bstr`, 实参却是 int32 计数。于是
//   * 计数 ≠ 0 → SetWindowTextW(hwnd, (BSTR)3) 解引用地址 3 → 访问违例;
//   * 计数 == 0 → 传 NULL → 函数内 `if (bstr)` 为假 → **静默不生效**
//     (这正是"点了没什么反应"的来源, 也是最难查的一种)。
//
// 原来只有"默认属性写入"和"With 块"两条路径手动调了 wrapToBSTR, 控件数组元素具名
// 属性 (`Label1(0).Caption = v`)、WithEvents 变量、显式 Let/Set 这几条都漏了 ——
// 同一缺陷类散在 6 处, 修一处补一处迟早再漏。收敛到本函数: 凡写字符串型控件属性,
// 值一律过这里; 非字符串型属性 (Picture/Value/ListIndex …) 原样返回。
std::string CCodeGen::controlPropValueExpr(const std::string& writeFn,
                                           const std::string& valExpr, Expr& valueNode) {
    if (writeFn.find("SetControlText") != std::string::npos
        || writeFn.find("SetMenuCaption") != std::string::npos) {
        return wrapToBSTR(valExpr, valueNode);
    }
    // Fix <vbeclipse> 2026-10-06: Picture 属性写遇到 Variant 载体实参 (如
    // `Set imgA.Picture = LoadResPicture(...)` — LoadRes* 实装后返回 vb6_VARIANT)
    // 时按对象解包, 与赋值路 (Fix 038b-6 名单) 同口径; 否则 C2172 "实参不是指针"。
    if (writeFn == "vb6_SetControlPicture" || writeFn == "vb6_SetControlPictureFromCom") {
        if (cExprIsVariant(valExpr)) {
            return "vb6_VariantToObjectVal(" + valExpr + ")";
        }
    }
    return valExpr;

}


std::string CCodeGen::mapBinaryOp(BinaryOp op) const {
    switch (op) {
        case BinaryOp::Or:     return "|";    // VB6 Or = 位或
        case BinaryOp::Xor:    return "^";    // VB6 Xor = 位异或
        case BinaryOp::And:    return "&";    // VB6 And = 位与
        case BinaryOp::Eq:     return "==";
        case BinaryOp::Neq:    return "!=";
        case BinaryOp::Lt:     return "<";
        case BinaryOp::Gt:     return ">";
        case BinaryOp::Le:     return "<=";
        case BinaryOp::Ge:     return ">=";
        case BinaryOp::Add:    return "+";
        case BinaryOp::Sub:    return "-";
        case BinaryOp::Mod:    return "%";
        case BinaryOp::Mul:    return "*";
        case BinaryOp::Div:    return "/";
        case BinaryOp::Is:     return "==";   // 对象引用比较
        // 以下运算符在visit(BinaryExpr&)中已特殊处理, 此处不应到达
        case BinaryOp::Concat: return "/* CONCAT */";
        case BinaryOp::IntDiv: return "/* INTDIV */";
        case BinaryOp::Pow:    return "/* POW */";
        case BinaryOp::Eqv:    return "/* EQV */";
        case BinaryOp::Imp:    return "/* IMP */";
        case BinaryOp::Like:   return "/* LIKE */";
        default:               return "/* unhandled BinaryOp */";
    }
}

} // namespace vb6c3
