#include "backend/cgen.hpp"
#include "driver/detail/phase_timing.hpp"   // ai/030 续: inferExprType 累计统计
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// Fix <VBFlexGridDemo/GA 36766856082>: 内置函数与"控件自己声明的同名 Property"撞名时的
// 类型裁决。
//
// 背景: Task #40 (40960eb) 给 parseIdentifierOrCall 加了"使用点剥类型后缀",
// `Left$(Value, 1)` 到 cgen 时名字已是 `Left`; VBFlexGrid.ctl 自己声明了
// `Public Property Get Left() As Single`, 于是 inferExprType 的符号表支路命中这个
// **属性**符号 ⇒ 答 Single ⇒ `float _vb6_select_81 = vb6_Left(Temp, 1)` 把 BSTR
// 强转 float → C2440 (27 处)。剥后缀前名字是 `Left$`, 查不到该属性符号, 天然不撞。
//
// 口径 (刻意收窄): 只有"符号 kind 是 Property* **且** 该名字在内置函数表里"才启用,
// 其余一切照旧 —— 用户 Function 同名、普通变量、真·属性访问都不受影响。
// 返回类型按 VB6 内置函数语义给: 字符串类答 String, 整数类答 Long, 其余答 Variant
// (Variant 是既有兜底, 不会引入新的硬错)。
#include "backend/detail/expr/cgen_expr_ident_builtin_table.inc"

static Vb6Type builtinFuncReturnTypeForShadowedProp(const std::string& nameLower) {
    static const std::unordered_set<std::string> kStrFuncs = {
        "left", "mid", "right", "trim", "ltrim", "rtrim", "lcase", "ucase",
        "space", "string", "str", "chr", "hex", "oct", "format", "strconv",
        "input", "command", "curdir", "environ", "dir", "date", "time", "monthname",
        "weekdayname", "error"
    };
    if (kStrFuncs.count(nameLower)) return Vb6Type::String;
    static const std::unordered_set<std::string> kLongFuncs = {
        "len", "lenb", "instr", "instrb", "asc", "cint", "clng", "cbyte",
        "freefile", "erl", "hour", "minute", "second", "weekday", "day",
        "month", "year", "abs", "sgn", "int", "fix"
    };
    if (kLongFuncs.count(nameLower)) return Vb6Type::Long;
    return Vb6Type::Variant;
}

// 判据: 符号是 Property* **且** 该名字在内置函数表里 ⇒ 这是"属性遮蔽了内置函数",
// 类型要按内置函数答 (见上注释)。剥 $ 后缀后比, 与 ident_builtin 的 lookupName 同口径。
static bool isPropShadowingBuiltinFunc(const Symbol* sym, const std::string& idName) {
    if (!sym) return false;
    if (sym->kind != SymbolKind::PropertyGet && sym->kind != SymbolKind::PropertyLet
        && sym->kind != SymbolKind::PropertySet) return false;
    std::string n = Symbol::toLower(idName);
    if (!n.empty() && n.back() == '$') n.pop_back();
    return kBuiltinFuncNames.count(n) != 0;
}

static Vb6Type cgenTypeForShadowedBuiltin(const std::string& idName) {
    std::string n = Symbol::toLower(idName);
    if (!n.empty() && n.back() == '$') n.pop_back();
    return builtinFuncReturnTypeForShadowedProp(n);
}

// --- cgen_util_type.cpp: 表达式类型推断 + Variant 判定 + 运行时参数 C 类型 ---


// ============================================================
void CCodeGen::visit(SimpleTypeRef& node) {}
void CCodeGen::visit(ArrayTypeRef& node) {}
void CCodeGen::visit(FixedStringTypeRef& node) {}

// ============================================================
// 模块 visit (generate()已按类别分派, 此处为空)
// ============================================================

void CCodeGen::visit(Module& node) {}

// ============================================================
// ai/032: 四档无符号/窄整型 (SByte/UInteger/ULong/ULongLong) 的映射唯一入口。
// 位宽口径: 阶梯对齐 8/16/32/64 —— U<x> 恒等于「<x> 的无符号版」,
// 与既有 Vb6Type::ULong=19 → uint32_t 完全吻合 (LongLong=64 位 ⇒ ULongLong=64 位无符号)。
// C 型这一侧与 mapType (cgen_base_type.cpp) 是同一份答案, 两处必须同步。
// ============================================================
bool CCodeGen::isNarrowIntVbType(Vb6Type t) {
    return t == Vb6Type::SByte || t == Vb6Type::UInteger
        || t == Vb6Type::ULong || t == Vb6Type::ULongLong;
}

Vb6Type CCodeGen::narrowIntTypeOfCType(const std::string& cType) {
    if (cType == "int8_t")   return Vb6Type::SByte;
    if (cType == "uint16_t") return Vb6Type::UInteger;
    if (cType == "uint32_t") return Vb6Type::ULong;
    if (cType == "uint64_t") return Vb6Type::ULongLong;
    return Vb6Type::Unknown;
}

// ai/032: 把操作数按**无符号**口径交给 helper (形参是 uint32_t / uint64_t)。
// Variant 操作数先 vb6_VariantToLongPtr 解包 —— 不用 vb6_VariantToLong: 后者只认
// VT_I4 家族, 对 VT_I8/VT_UI8 会答 0 (vb6_Shl 那一族同理, 见 cgen_expr_binary.cpp)。
std::string CCodeGen::asUnsignedOperand(const std::string& cExpr, const Expr& astExpr,
                                        const char* ctype) const {
    bool isVar = isDefinitelyVariantExpr(const_cast<Expr&>(astExpr));
    if (!isVar) isVar = cExprIsVariant(cExpr);
    if (isVar) {
        return "(" + std::string(ctype) + ")vb6_VariantToLongPtr(" + cExpr + ")";
    }
    return "(" + std::string(ctype) + ")(" + cExpr + ")";
}

Vb6Type CCodeGen::intDivResultType(Vb6Type lt, Vb6Type rt) const {
    // MSDN "Data Types of Operator Results" 的 \ 表: 两个窄整型 (Boolean/SByte/Byte/
    // Short/UShort/Integer/UInteger) 的结果一律是 Long(64) —— 映射到 C3 的阶梯就是
    // "窄档一律落 Long(32)"。再宽的档按 promote 走, 与 `+ - *` 同一份有符号/无符号
    // 提升规则, 免得同一个表达式里"商"与"和"两个口径。
    Vb6Type p = TypeSystem::promote(lt, rt);
    switch (p) {
        case Vb6Type::Long:      return Vb6Type::Long;       // 窄档统一落这里
        case Vb6Type::LongLong:  return Vb6Type::LongLong;
        case Vb6Type::ULong:     return Vb6Type::ULong;
        case Vb6Type::ULongLong: return Vb6Type::ULongLong;  // 含 64 位混符号的兜底
        // SByte/Byte/Integer/UInteger/Boolean → Long; 浮点/Currency/Decimal/Variant →
        // Long (VB: `\` 的操作数先转整型, 结果类型是 Long)。
        default:                 return Vb6Type::Long;
    }
}

// ============================================================
// 表达式类型推断 (简化版, 用于Select Case等场景)
// ============================================================

// ai/030 续 (§性能, 2026-10-14): inferExprType 的按节点缓存。
// 背景 (VBFlexGridDemo 读数): inferExprType 263,641 次 / 18.4s —— 70µs/次。根因是
// 同一枚节点在一条语句里被反复下潜重推 (对象子树 × 若干辅助各推一遍, 每推一遍还带
// 全符号表线性扫描与 toLower 分配)。缓存口径:
//   · epoch 单调不减且**全局** —— AST 节点可能被两个 CCodeGen 实例先后访问 (EXE 的
//     逐模块实例与 DLL 的 dll_entry 实例), 各实例自行计数会撞车;
//   · 推 epoch 的两个边界: 每个过程入口 (visit(Sub/Function/Property)) 与
//     emitStmtList 的每条语句 —— 会改推断答案的表 (known*Vars_) 在体内中途的注册
//     (Dim/For Each/With 临时) 全落在语句边界上, 语句内表格不动, 正是缓存可复用的窗口;
//   · epoch==0 (首次推之前) 只算不存 —— 该窗口的表格状态不保证稳定。
static uint32_t g_inferCacheEpoch = 0;

void CCodeGen::bumpInferEpoch() { ++g_inferCacheEpoch; }

Vb6Type CCodeGen::inferExprType(Expr& expr) const {
    C3ScopeTimer _c3t(1);   // ai/030 续: 累计统计计在包装层 (缓存命中也算一次问)
    const uint32_t ep = g_inferCacheEpoch;
    if (ep != 0 && expr.inferCacheEpoch == ep) return static_cast<Vb6Type>(expr.inferCacheType);
    Vb6Type t = inferExprTypeImpl(expr);
    expr.inferCacheEpoch = ep;
    expr.inferCacheType = static_cast<uint16_t>(t);
    return t;
}

Vb6Type CCodeGen::inferExprTypeImpl(Expr& expr) const {
    switch (expr.kind) {
        case ASTNodeKind::IdentifierExpr: {
            auto& id = static_cast<IdentifierExpr&>(expr);
            // 优先检查已知的变量类型集合 (局部变量在符号表中作用域可能不可达)
            std::string lower = id.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            // Fix 161f: 无 As Type 但整数可折叠的 Const (符号表里 type=Variant)。
            // C 侧已 emit 成 `#define NAME (4127)`, 标识符也内联为数值; 若这里
            // 仍答 Variant, Declare 调用的实参会被包成
            // vb6_VariantToLong(LVM_GETHEADER) → C2440 "无法从 int 转换为
            // vb6_VARIANT" (extlist MListView SendMessage 实测)。按 Long 答。
            if (moduleIntConstValues_.count(lower)) return Vb6Type::Long;
            if (knownBstrVars_.count(lower)) return Vb6Type::String;
            if (knownSingleVars_.count(lower)) return Vb6Type::Single;
            // Fix 175: Date 必须先于 Double 判 (Date 变量同时登记在 knownDoubleVars_
            // 里以复用既有 double 取值路径, 口径同 Fix 117c 的 Single)。
            if (knownDateVars_.count(lower)) return Vb6Type::Date;
            if (knownDoubleVars_.count(lower)) return Vb6Type::Double;
            // ai/022 W1: 必须先于 knownLongVars_ 判 (口径同 Fix 175 的 Date) ——
            // As Boolean 的 C 型与 Integer 同串, 只按 C 类型登记就永远看不见布尔。
            if (knownBoolVars_.count(lower)) return Vb6Type::Boolean;
            // 账 #123: 口径同 Fix 175 的 Date / W1 的 Boolean —— 登记过就必须由这张表答 Byte,
            // 让局部/形参 Byte 与模块级 Byte (走符号表那条支路) 给出同一份答案。
            if (knownByteVars_.count(lower)) return Vb6Type::Byte;
            // ai/032: SByte/UInteger/ULong/ULongLong 的局部/形参/返回槽。C 型
            // (int8_t/uint16_t/uint32_t/uint64_t) 不在上面任何一支, 且局部变量在
            // symTab_ 里不可达 ⇒ 以前落到 default 答 Variant (CStr/Debug.Print 就
            // 因此按有符号 32 位打, &HFFFFFFFF 变 -1)。必须由专属表先答。
            if (auto itNI = knownNarrowIntVars_.find(lower); itNI != knownNarrowIntVars_.end())
                return itNI->second;
            if (knownLongVars_.count(lower)) return Vb6Type::Long;
            if (knownLongPtrVars_.count(lower)) return Vb6Type::LongPtr;
            // Fix <vbeclipse> rev20 → 账 #159: 设计器 (.ctl/.pag) 模块内裸写的宿主
            // 伪成员 (ScaleWidth / hDC / hWnd / ScaleMode / Enabled …) 在 C 侧是 RTL
            // 全局量 (extern int32_t 等, vb6rtl_userctl.h), **不是** vb6_VARIANT。
            // 此前它们落到下面的符号表回退 —— 符号表里没有 (不是 VB6 声明的变量),
            // 于是答 Variant, 比较就发成 `vb6_VarCmpLongLt(&vb6_UserControl_ScaleWidth,
            // …)` (ucTabStrip.ctl:146 实测): 拿 int32_t* 当 vb6_VARIANT* 传 (RTL 签名
            // 是 vb6_VARIANT*, 16 字节), 读 4 字节对象的头 16 字节当 vt + lVal ⇒ 比较
            // 结果是垃圾。此前能编过只是因为 MSVC 把它当 C4133 指针类型不兼容警告放行。
            // 成员名与类型现在由 kHostPseudoRows 一处回答; requireBare=true 与本文件
            // 发射侧的 HPF_BARE 门同口径 (两条路必须答同一个数)。判据要求无同名局部
            // (Dim ScaleWidth As Long 要保住自己的类型), 与发射侧的 knownLocalVars_ 门一致。
            if (isDesignerModule_ && !knownLocalVars_.count(lower)) {
                Vb6Type hpT = Vb6Type::Unknown;
                if (hostPseudoValueType(isPropertyPageDesigner_ ? "PropertyPage" : "UserControl",
                                        lower, hpT, true)) {
                    return hpT;
                }
            }
            if (knownVariantVars_.count(lower)) return Vb6Type::Variant;
            // 检查符号表
            auto* sym = symTab_.lookup(id.name);
            if (!sym) sym = symTab_.lookupModule(id.name);
            // Fix <VBFlexGridDemo>: 属性遮蔽内置函数时按内置函数的类型答 (见文件头注释)。
            // 但**函数体内 `Name = expr` 引用返回槽**这一支例外 —— VB6 里 `Left = Extender.Left`
            // 在 `Public Property Get Left() As Single` 内部指的是 vb6_ret_Left (float),
            // 不是内置 Left$()。判据同 cgen_assign_value_sem.inc:5-12 —— currentProc_ 是
            // Function/PropertyGet 且名字等于本标识符时, 走返回槽 (sym->type)。
            // 修前: inferExprType 答 String ⇒ vb6_BSTR_AssignMove(&vb6_ret_Left, vb6_CStrDbl(...))
            //       ⇒ BSTR_Free 把 float 位当指针 ⇒ MainForm Form_Resize 一调 VBFlexGrid1.Left
            //       即 0xC0000005, 启动看不到界面。
            // 只改 IdentifierExpr 分支; IndexOrCallExpr 那一支 (本文件 178 行附近) 不动 ——
            // 那里 Left(Temp, 1) 就是真·内置函数调用。
            if (sym && isPropShadowingBuiltinFunc(sym, id.name)) {
                const bool isRetSlotSelfRef =
                    currentProc_
                    && (currentProc_->kind == SymbolKind::Function
                        || currentProc_->kind == SymbolKind::PropertyGet)
                    && Symbol::toLower(currentProc_->name) == Symbol::toLower(id.name);
                if (!isRetSlotSelfRef) return cgenTypeForShadowedBuiltin(id.name);
            }
            if (sym) return sym->type;
            break;
        }
        case ASTNodeKind::LiteralExpr: {
            auto& lit = static_cast<LiteralExpr&>(expr);
            if (lit.literalKind == LiteralKind::String) return Vb6Type::String;
            if (lit.literalKind == LiteralKind::Double || lit.literalKind == LiteralKind::Single
                || lit.literalKind == LiteralKind::Currency || lit.literalKind == LiteralKind::Decimal)
                return Vb6Type::Double;
            if (lit.literalKind == LiteralKind::Boolean) return Vb6Type::Boolean;
            if (lit.literalKind == LiteralKind::Date) return Vb6Type::Date;
            // 装不进 32 位的 Long 字面量, cgen 实际发的是 64 位 (LL 后缀), 这里
            // 必须跟着答 LongLong。原先一律答 Long, 于是收窄检查把它当成"装得下"
            // 跳过 —— `l As Long: l = -9223372036854775807` 就从 Error 6 变成
            // 静默截成 1。类型口径要和实际发出的位宽一致, 否则检查等于没做。
            if (lit.literalKind == LiteralKind::Long
                && (lit.longValue < INT32_MIN || lit.longValue > INT32_MAX))
                return Vb6Type::LongLong;
            return Vb6Type::Long;
        }
        case ASTNodeKind::BinaryExpr: {
            auto& bin = static_cast<BinaryExpr&>(expr);
            // 字符串连接运算符 → String
            if (bin.op == BinaryOp::Concat) return Vb6Type::String;
            // 比较运算符 → Boolean
            if (bin.op == BinaryOp::Eq || bin.op == BinaryOp::Neq ||
                bin.op == BinaryOp::Lt || bin.op == BinaryOp::Gt ||
                bin.op == BinaryOp::Le || bin.op == BinaryOp::Ge ||
                bin.op == BinaryOp::Like || bin.op == BinaryOp::Is)
                return Vb6Type::Boolean;
            // VB6 的 And/Or/Xor/Eqv/Imp 是**位运算**，结果类型只在
            // TypeSystem::bitwiseResult 一处写 (账 #216：以前这份 oracle 无条件答
            // Boolean，于是 `CStr(a Or b)` 打成 True、`vb6_ComPack*(hDC Or 0)` 装箱成
            // VT_BOOL —— 值一直是位的数，只是类型被问错)。
            if (bin.op == BinaryOp::And || bin.op == BinaryOp::Or ||
                bin.op == BinaryOp::Xor || bin.op == BinaryOp::Eqv ||
                bin.op == BinaryOp::Imp) {
                return TypeSystem::bitwiseResult(inferExprType(*bin.left),
                                                  inferExprType(*bin.right));
            }
            // 浮点除法 → Double。**但有一侧是 Variant 时必须答 Variant**（账 #305 §B145）：
            // 那一族的发码自第三十六刀起交运行期出口 vb6_VarDivV(...)，C 表达式的类型是整只
            // vb6_VARIANT；这里再答 Double 就是「oracle 与 C 表达式的型不一致」，实物 = 门 #517
            // 唯一那条真红：Charts 2020 ClsResizer.cls 的 COM 实参打包按 inferExprType 选型 ⇒
            // `vb6_ComPackDouble(一枚槽)` ⇒ C2440，两架构同形。
            // ⚠ 第三十七刀的订正（§B147）：这一问**必须与发码改道共用同一条判据的 AST 那一半**
            // (variantOperandAST)，不能拿 inferExprType 的递归答案当「操作数是槽」。改前就是两
            // 条答案：oracle 看 inferExprType、改道看 C 文本前缀表，于是「AST 说是 Variant 而
            // 文本表读不出」那一族（实测：Variant 数组的元素访问、UC 宿主伪成员 TextHeight、
            // Screen_TwipsPerPixelY）压根没改道，却被这条 oracle 答成 Variant ⇒ 消费点不提取
            // 也不取整 ⇒ 浮点进窄整型目标静默截断。改吃 AST 那一半之后方向性是构造性的：
            // 改道 = AST ∪ 文本 ⊇ oracle = AST。
            if (bin.op == BinaryOp::Div) {
                if (variantOperandAST(bin.left.get()) || variantOperandAST(bin.right.get()))
                    return Vb6Type::Variant;
                return Vb6Type::Double;
            }

            // 移位 Shl/Shr: 结果类型 = **左操作数**类型, 窄整型抬到 32 位。
            // MSDN 的移位表就是"把移位当一元运算作用在左操作数上" —— 右操作数只是
            // 计数, 不参与结果类型; 改前这里走 promote(左, 右) 会把 `u >> 4`
            // (u As ULong) 按"混符号"提升成 LongLong, 与发码层投回的 uint32_t 两个口径。
            // 口径与 cgen_expr_binary.cpp 的 castBack 同一份 (那边也只看左操作数)。
            if (bin.op == BinaryOp::Shl || bin.op == BinaryOp::Shr) {
                Vb6Type lS = inferExprType(*bin.left);
                switch (lS) {
                    case Vb6Type::SByte: case Vb6Type::Byte:
                    case Vb6Type::UInteger: case Vb6Type::Integer:
                    case Vb6Type::Boolean: case Vb6Type::Long:
                        return Vb6Type::Long;          // 窄档抬到 32 位
                    case Vb6Type::ULong: case Vb6Type::LongLong:
                    case Vb6Type::ULongLong: case Vb6Type::LongPtr:
                        return lS;
                    default: return Vb6Type::LongLong; // 浮点/Variant/Unknown: C 侧是 int64 中间量
                }
            }

            // 算术运算符: 提升左右类型
            {
                Vb6Type lt = inferExprType(*bin.left);
                Vb6Type rt = inferExprType(*bin.right);
                // Fix 175: VB6 日期算术 —— `Date ± 数值` 仍是 Date, `Date - Date` 是
                // 天数差 (Double), `*` `\` `Mod` `^` 无日期语义按数值处理。
                // TypeSystem::isNumeric(Date)==false, 交给 promote 会掉出数值阶梯。
                if (lt == Vb6Type::Date || rt == Vb6Type::Date) {
                    if (bin.op == BinaryOp::Add) return Vb6Type::Date;
                    if (bin.op == BinaryOp::Sub)
                        return (lt == Vb6Type::Date && rt == Vb6Type::Date)
                                   ? Vb6Type::Double : Vb6Type::Date;
                    return Vb6Type::Double;
                }
                // ai/032: `\` / `Mod` 的结果类型与"发哪一条 helper"必须是**同一份判据**
                // (intDivResultType, 定义在 cgen_util_type.cpp 的 asUnsignedOperand 旁)。
                // 改前这里自带一套"任一侧 ULong/ULongLong 就答无符号", 与 promote 的
                // 有符号/无符号提升规则各写一遍 —— `2 \ u` 这类换边写法就会漂移。
                if (bin.op == BinaryOp::IntDiv || bin.op == BinaryOp::Mod)
                    return intDivResultType(lt, rt);
                return TypeSystem::promote(lt, rt);
            }
        }
        case ASTNodeKind::UnaryExpr: {
            auto& un = static_cast<UnaryExpr&>(expr);
            // 同上：`Not` 对数值是按位取反，口径只在 TypeSystem::logicalNotResult 一处。
            if (un.op == UnaryOp::Not) {
                return TypeSystem::logicalNotResult(inferExprType(*un.operand));
            }
            return inferExprType(*un.operand);
        }
        case ASTNodeKind::IndexOrCallExpr: {
            // 函数调用: 返回函数返回类型
            auto& call = static_cast<IndexOrCallExpr&>(expr);

            // C3 扩展 (ai/032): `If(cond, a, b)` 三元的结果类型必须**自己算**, 不能
            // 沿用 addBuiltinFunc 登记的 Variant 兜底 —— 那个登记只是让语义层知道
            // "这个名字存在"。发码层发的是 C 的 `?:`, 真实 C 类型由两个分支按通常
            // 算术转换决定; 判成 Variant 会让下游套 vb6_VariantToString(<BSTR>) →
            // C2440 (实测: `P If(0 > 0, "p", "n")` 传给 String 形参时炸)。
            // 口径与 C 对齐: 同型取该型; 都数值取提升; 任一支 String 取 String;
            // 其余 (对象/UDT/混不入) 回 Variant, 由调用点自己想办法。
            if (call.callee && call.callee->kind == ASTNodeKind::IdentifierExpr
                && call.positional.size() == 3) {
                auto& idIf = static_cast<IdentifierExpr&>(*call.callee);
                if (Symbol::toLower(idIf.name) == "if") {
                    Vb6Type tTrue = inferExprType(*call.positional[1]);
                    Vb6Type tFalse = inferExprType(*call.positional[2]);
                    if (tTrue == tFalse) return tTrue;
                    if (TypeSystem::isNumeric(tTrue) && TypeSystem::isNumeric(tFalse))
                        return TypeSystem::promote(tTrue, tFalse);
                    if (tTrue == Vb6Type::String || tFalse == Vb6Type::String)
                        return Vb6Type::String;
                    return Vb6Type::Variant;
                }
            }

            if (call.callee && call.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& id = static_cast<IdentifierExpr&>(*call.callee);
                auto* sym = symTab_.lookup(id.name);
                if (!sym) sym = symTab_.lookupModule(id.name);
                // Fix <VBFlexGridDemo>: 同上 —— 属性遮蔽内置函数时按内置函数的类型答,
                // 否则 `float _vb6_select_81 = vb6_Left(Temp, 1)` 把 BSTR 强转 float (C2440)。
                if (sym && isPropShadowingBuiltinFunc(sym, id.name)) {
                    return cgenTypeForShadowedBuiltin(id.name);
                }
                if (sym) return sym->type;
            }
            // P26: 类实例方法调用 a.Method(args) → callee是MemberAccessExpr
            // 需要查询成员函数的返回类型, 而非直接fallback到Variant
            if (call.callee && call.callee->kind == ASTNodeKind::MemberAccessExpr) {
                return inferExprType(*call.callee);
            }
            break;
        }
        case ASTNodeKind::MemberAccessExpr: {
            auto& ma = static_cast<MemberAccessExpr&>(expr);
            // C29-1a + P20-42 + 账 #229 + 账 #231: 对象是窗体上的已知控件 (单枚**或**数组
            // 元素) 时, 属性类型**只问 controlPropType 这一张表**。
            // 为什么必须问在这条 case 的**最前面**: 兜底那条 lookupModule(memberName) 按成员
            // **裸名**查模块级符号, 凡是与模块级/内置符号同名的控件属性都会被顶掉 ——
            // 实测 `SSTab1.Tab` 撞上返回 BSTR 的内置函数 `Tab` → 判成 String → Debug.Print
            // 拼接不套 vb6_CStr(vb6_VariantFromValue(...)) → 而 RTL 的 vb6_SSTab_GetTab 返回
            // int32_t → 整数当 BSTR 指针解引用 → 0xC0000005。控件数组那一形 (`uArr(0).Left`)
            // 以前根本不进这条规则, 掉进同一个坑 (账 #229)。
            // 只认内建控件类型: 工程内 .ctl 宿主同名属性由它自己的符号表说话 (表内那道闸)。
            {
                FrmControlType ctlType = FrmControlType::Unknown;
                if (ctrlTypeOfMemberObject(ma.object.get(), ctlType)) {
                    Vb6Type pt = controlPropType(ctlType, ma.memberName);
                    if (pt != Vb6Type::Unknown) return pt;
                }
            }
            // P24-12: Err对象特殊处理
            if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objId = static_cast<IdentifierExpr&>(*ma.object);
                std::string objLower = objId.name;
                std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
                if (objLower == "err") {
                    std::string memLower = ma.memberName;
                    std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                    if (memLower == "number") return Vb6Type::Long;
                    if (memLower == "description" || memLower == "source") return Vb6Type::String;
                    if (memLower == "helppath" || memLower == "helpfile" || memLower == "helpcontext") return Vb6Type::String;
                    if (memLower == "lastdllerror") return Vb6Type::Long;
                }
            }
            // 账 #159: 限定形态的宿主伪成员 (<UserControl|PropertyPage|Extender|
            // Ambient>.<成员>) 与裸名那一路**同源** —— 都问 kHostPseudoRows。
            // 此前这里是第五份成员名清单 (7 枚, 一律答 Long, 且不含 hWnd/hDC/
            // ScaleMode/Extender/Ambient) ⇒ 同一个值写 `ScaleWidth` 答 Long、写
            // `UserControl.hWnd` 两边都不答 ⇒ 装箱比较恒假。MouseUp 里
            // `X < ScaleWidth` 恒假 → RaiseEvent Click 永不触发 (czUI 实测) 就是它。
            if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                const IdentifierExpr& objIdHp = static_cast<const IdentifierExpr&>(*ma.object);
                const std::string objLowerHp = Symbol::toLower(objIdHp.name);
                if (hostPseudoIsObject(objLowerHp)) {  // 账 #278 §B120: 这一问由那张表答
                    Vb6Type hpT = Vb6Type::Unknown;
                    if (hostPseudoValueType(objLowerHp, ma.memberName, hpT, false)) {
                        return hpT;
                    }
                }
            }
            // Fix 220d: `<宿主>.<宿主对象>` 形 (UserControl.Ambient / UserControl.Extender /
            // PropertyPage.Ambient …) —— 这个成员访问的结果**仍是宿主伪对象本身**, 它的成员
            // (如 .Font / .Left) 在生成码里是结构体字段/独立全局 (vb6_UserControl_Ambient.Font),
            // 不是后期绑定读。
            // 必须在这里返回 Unknown 拦住下面的 Fix 220: 否则 Fix 220 会顺着
            // `inferExprType(UserControl) == Variant` (兜底) 把 `.Ambient` 答成 Variant,
            // 外层 `.Font` 再答 Variant ⇒ comPackExpr 选 vb6_ComPackValue 而非
            // vb6_ComPackObject ⇒ 生成的 exe 起来即崩。
            // 实测 Charts 2020 生成码 diff 的残留就是这个形状
            //   `- vb6_ComPackObject(vb6_UserControl_Ambient.Font)`
            //   `+ vb6_ComPackValue(vb6_UserControl_Ambient.Font)`
            // (LabelPlus.c / ucTreeMaps.c)。
            if (hostPseudoIsObject(Symbol::toLower(ma.memberName))) return Vb6Type::Unknown;
            // Fix 081i: UDT字段访问 — 先查找UDT成员类型，避免lookupModule
            // 匹配到内置函数(如Left→String)导致类型推断错误
            {
                std::string udtCType = inferUdtTypeOfExpr(*ma.object);
                if (!udtCType.empty()) {
                    const std::string prefix = "vb6_type_";
                    if (udtCType.size() > prefix.size()
                        && udtCType.compare(0, prefix.size(), prefix) == 0) {
                        std::string udtName = udtCType.substr(prefix.size());
                        Symbol* udtSym = symTab_.lookupModule(udtName);
                        if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                            std::string memLower = ma.memberName;
                            std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                            for (auto& mi : udtSym->udtMembers) {
                                std::string miLower = mi.name;
                                std::transform(miLower.begin(), miLower.end(), miLower.begin(), ::tolower);
                                if (miLower == memLower) {
                                    return mi.type;  // 找到UDT字段，返回其Vb6Type
                                }
                            }
                            // Fix 084o-2: UDT 已确认但字段未找到 → 字段类型未知, 返回 Unknown.
                            // 不要回退 lookupModule(memberName): 成员名是字段名而非模块符号,
                            // 可能误匹配模块级同名符号 (如 SourceFile 变量) 导致类型误判为 String.
                            // Fix 090ab: 也不应返回 Variant — 否则 UDT 字段 (如 COMSTAT.fBitFields,
                            // 跨模块 Type 符号缺失) 被 isDefinitelyVariantExpr 误判为 Variant,
                            // 实参生成时套 vb6_VariantToLong(int32 字段) → C2440.
                            return Vb6Type::Unknown;
                        }
                    }
                    // Fix 090ab: UDT C 类型已知 (knownUdtVars_) 但符号不可解析 (跨模块
                    // Public Type 在符号表注册表不可达) → 字段类型未知, 保守返回 Unknown,
                    // 禁止回退到模块级同名符号或默认 Variant (否则 COMSTAT.fBitFields 等
                    // 会被误判为 Variant 而错误套用 vb6_VariantToLong → C2440).
                    return Vb6Type::Unknown;
                }
            }
            // 类实例成员函数返回类型 (tB 泛型类特化高发的跨类同名碰撞):
            // 按 object 推断出的宿主类查 memberReturnTypes, 内置标量名直接映射.
            // 下面的裸名 lookupModule(memberName) 在多个类同名成员时只命中
            // storageKey 去重胜出的第一个模块版本 (如 box_g1_long 的 GetVal→Long),
            // String 版被误判 → Debug.Print 把 BSTR 指针截断成垃圾数 (Fix 176 同源).
            if (ma.object && symTab_.moduleScope()) {
                std::string clsName = inferClassTypeOfExpr(*ma.object);
                if (!clsName.empty()) {
                    const std::string clsLower = Symbol::toLower(clsName);
                    const std::string memLowerCls = Symbol::toLower(ma.memberName);
                    for (const auto& [k, csym] : symTab_.moduleScope()->symbols()) {
                        if (csym->kind != SymbolKind::Class) continue;
                        bool hit = csym->isExternal
                            ? (Symbol::toLower(csym->sourceModule) == clsLower)
                            : (Symbol::toLower(csym->name) == clsLower);
                        if (!hit) continue;
                        auto it = csym->memberReturnTypes.find(memLowerCls);
                        if (it != csym->memberReturnTypes.end()) {
                            std::string rn = Symbol::toLower(it->second);
                            if (rn == "string")  return Vb6Type::String;
                            if (rn == "long")    return Vb6Type::Long;
                            if (rn == "integer") return Vb6Type::Integer;
                            if (rn == "boolean") return Vb6Type::Boolean;
                            if (rn == "single")  return Vb6Type::Single;
                            if (rn == "double")  return Vb6Type::Double;
                            if (rn == "date")    return Vb6Type::Date;
                            if (rn == "byte")    return Vb6Type::Byte;
                            if (rn == "currency")return Vb6Type::Currency;
                            if (rn == "variant") return Vb6Type::Variant;
                            if (rn == "object")  return Vb6Type::Object;
                        }
                        break;
                    }
                }
            }
            // 账 #305 §B170 (第六十一刀): 「这枚成员读是哪一档」先问**宿主类**那张字段表
            // (classFieldKind = §B168 那趟预扫描建的表, §B169 收成一处的那一问), 再落到底下
            // 那条按**裸成员名**查模块级符号的兜底。兜底那一问不认识宿主: 字段名撞上内置函数名
            // 时它交回的是内置函数的返回档 —— 实测 `Public PV As Variant` 撞上 PV() (金融函数,
            // builtin_funcs.inc:317 答 Double) ⇒ `CStr(h.PV)` 发 vb6_CStrDbl(那枚槽), RTL 的
            // _Generic 按数值解它 (VarR8) ⇒ 串值读回 0|1 (探针 K61-2, 两架构同形); 同一形里
            // `Public PV As String` 也答 Double (K61-4) —— 声明档整个被吞掉。350 行控件那一档
            // 早就按宿主问表 (账 #229/#231: SSTab1.Tab 撞内置 Tab), 本刀把同一条口径补到类字段。
            // 只在表答得出 (非 Unknown) 时接管: 数组/定长串/对象字段不入表 (§B168 那条刻意口径),
            // 同名方法/属性也不许被字段答案抢走 ⇒ 与 §B169 同一枚 resolveClassMemberCall 否决。
            if (ma.object) {
                const std::string cls61 = inferClassTypeOfExpr(*ma.object);
                if (!cls61.empty() && resolveClassMemberCall(cls61, ma.memberName).empty()) {
                    const Vb6Type fld61 = classFieldKind(cls61, ma.memberName);
                    if (fld61 != Vb6Type::Unknown) return fld61;
                }
            }
            // Fix 220: 对象是 **Variant 载体** (集合 .Item(i) 返回 Variant / Variant 局部变量
            // 持有对象 / ComCall 结果 …) 时, 成员访问是后期绑定属性读, 结果**本身也是
            // Variant** —— 必须与 codegen 侧 Fix 219 (cgen_expr_member_class_fallback.inc /
            // cgen_expr_member_voidptr_com.inc) 发出的
            //   `vb6_VariantFromComResult(vb6_ComGetProp(vb6_VariantToObjectVal(obj), L"member"))`
            // (返回 vb6_VARIANT 值) 严格一致。
            // 若这里落到下方 `lookupModule(memberName)` 用成员的**声明类型** (如 Long) 答, 两条路
            // 就错位: comPackExpr 据此选 vb6_ComPackInt, 生成
            //   `vb6_ComPackInt(vb6_VariantFromComResult(...))`
            // → C2440 "vb6_VARIANT 无法转换为 int32_t" (CoolBar.c:1230/1235 等, 共约 170 处
            // VARIANT→int32_t/float/double)。本修复与早期绑定互不干扰: 早期绑定只在 codegen 侧
            // 已知对象类时触发, 那时 inferExprType(object) 会答出该类而非 Variant, 不会进此分支。
            // Fix 220b/c/f: **必须先把三类"其实不是 Variant 载体"的对象摘出去**,
            // 否则下面那条规则会把它们的成员访问也答成 Variant, 下游 (comPackExpr /
            // 赋值转换 / 实参包装) 据此生成错代码, 生成的 exe 起来就崩。
            //   220b —— 裸标识符若符号表里查不到, 它更可能是**命名空间/库限定符**
            //          (VBA / Constants …): `VBA.DateAdd(...)`。若答 Variant,
            //          cgen_assign_com_prop.inc 的 Fix 159-B (isScalarImplicitStr159)
            //          不命中 ⇒ 裸拼 vb6_DateAdd(...) 给 BSTR 形参 ⇒ C2440
            //          (VBFlexGridDemo UserEditingForm.c 2 处)。
            //   220c —— **宿主伪对象本身** (Ambient / Extender / UserControl /
            //          PropertyPage): `Ambient.Font` / `Extender.Left` 在生成码里是
            //          结构体字段/独立全局 (vb6_Ambient_Font / vb6_Extender_Left),
            //          不是后期绑定读。若答 Variant, comPackExpr 选 vb6_ComPackValue
            //          而非 vb6_ComPackObject ⇒ Charts 2020 起来即崩 (0xC000041D)。
            //   220f —— 对象能被推成**具体类** ⇒ 早期绑定的直接调用, 同样不是后期
            //          绑定读 (Charts 2020 PropPagLP.pag:401
            //          `m_oUCImage.PictureFromStream LabelPlus1.PictureGetStream`)。
            // `Me` 例外保留 (它同样可能不在符号表里, 但确是实例载体)。
            {
                bool objIsRuntimeEntity = true;
                if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                    const std::string& objNm =
                        static_cast<IdentifierExpr&>(*ma.object).name;
                    const std::string objLower220 = Symbol::toLower(objNm);
                    if (objLower220 == "me") {
                        // `Me` 例外保留
                    } else if (hostPseudoIsObject(objLower220)) {
                        objIsRuntimeEntity = false;                 // 220c
                    } else if (!inferClassTypeOfExpr(*ma.object).empty()) {
                        objIsRuntimeEntity = false;                 // 220f
                    } else {
                        const Symbol* objSym = symTab_.lookup(objNm);
                        if (!objSym) objSym = symTab_.lookupModule(objNm);
                        if (!objSym) objIsRuntimeEntity = false;    // 220b
                    }
                }
                if (objIsRuntimeEntity
                    && inferExprType(*ma.object) == Vb6Type::Variant) {
                    return Vb6Type::Variant;
                }
            }
            // 查找成员函数/属性的返回类型
            auto* memSym = symTab_.lookupModule(ma.memberName);
            if (memSym) return memSym->type;
            break;
        }
        // Fix 084o-2: With 块内 UDT 字段访问 (.SourceFile 等) 的类型推断.
        // WithMemberExpr 由 With 语句展开 (_vb6_with_N->field), 必须按 UDT 字段查
        // udtMembers, 且不能回退 lookupModule(memberName) — 与 MemberAccessExpr 同理.
        case ASTNodeKind::WithMemberExpr: {
            auto& wm = static_cast<WithMemberExpr&>(expr);
            if (withObjectInfoStack_.empty() || withObjectVars_.empty()) return Vb6Type::Variant;
            const auto& winfo = withObjectInfoStack_.back();
            if (winfo.kind != WithObjKind::Unknown) {
                // 非 UDT With (类实例/COM 对象): memberName 是属性/方法 → 查成员符号
                // 账 #305 §B170 (第六十一刀): With 体内那一读与点号那一路同一个病、同一个问 ——
                // 下面按裸成员名查, 于是 `With h` 里的 `CStr(.PV)` 也发成数值档 (探针 K61-6)。
                // 宿主类名在 With 元信息栈里就有 (winfo.className, Fix 011r-1; 只有类实例那一档
                // 非空, 控件/COM/内置那一档交空 ⇒ 既有答案一字不动)。
                if (!winfo.className.empty()
                    && resolveClassMemberCall(winfo.className, wm.memberName).empty()) {
                    const Vb6Type fld61w = classFieldKind(winfo.className, wm.memberName);
                    if (fld61w != Vb6Type::Unknown) return fld61w;
                }
                auto* memSym2 = symTab_.lookupModule(wm.memberName);
                if (memSym2) return memSym2->type;
                return Vb6Type::Variant;
            }
            std::string tempLower = Symbol::toLower(withObjectVars_.back());
            auto it = knownUdtVars_.find(tempLower);
            if (it == knownUdtVars_.end()) return Vb6Type::Variant;
            const std::string prefix = "vb6_type_";
            const std::string& udtCType = it->second;
            if (udtCType.size() <= prefix.size()
                || udtCType.compare(0, prefix.size(), prefix) != 0) return Vb6Type::Variant;
            std::string udtName = udtCType.substr(prefix.size());
            Symbol* udtSym = symTab_.lookupModule(udtName);
            if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                std::string memLower = Symbol::toLower(wm.memberName);
                for (auto& mi : udtSym->udtMembers) {
                    if (Symbol::toLower(mi.name) == memLower) return mi.type;
                }
            }
            return Vb6Type::Variant;
        }
        default:
            break;
    }
    return Vb6Type::Variant;
}


// Fix 029: 严格 Variant 推断. 仅当表达式明确为 Variant 时返回 true.
// 与 inferExprType 的差异: 内置函数 (sym==null) 与 UDT 字段访问 (lookupModule 失败) 等
// 通过 fallback 返回 Variant 的情形, 此处视为非 Variant, 避免对 int/BSTR 等实参误包装.
bool CCodeGen::isDefinitelyVariantExpr(Expr& expr, bool* isArrOut) const {
    C3ScopeTimer _c3t(3);   // ai/030 续: isDefinitelyVariantExpr 累计统计 (slot 3)
    if (isArrOut) *isArrOut = false;
    uint16_t variantArrRaw = static_cast<uint16_t>(Vb6Type::Variant)
                           | static_cast<uint16_t>(Vb6Type::Array);

    // 多态内置函数 denylist: symTab 注册为 Variant, 但 codegen 按上下文
    // 发出类型化版本 (vb6_IIfBSTR/Long/Double, Choose 嵌套三元, Switch 嵌套三元),
    // 实际 C 返回类型不是 vb6_VARIANT. 视为非 Variant 以避免错误包装.
    //
    // "if"(VB.NET 三元 If(c,t,f), 账 #256 新增): 与 iif 同族 —— 同样在
    // builtin_funcs.inc 注册为 Variant (仅为"名字存在"以便表达式解析放行), 实际
    // 由 cgen 发**裸 C 三目** ((cond) ? a : b), 真 C 类型由两支共同决定
    // (两 String 支 → BSTR, 两数值支 → 提升后的数值). 若不在此处否认, 会落到
    // symTab_.lookup("If") 命中 Variant → 对 BSTR 结果套 vb6_VariantToString
    // → C2440 "无法从 BSTR 转换为 vb6_VARIANT" (FeatSem.bas if_nested 实测).
    // 注: 真结果类型由 inferExprType 的 IndexOrCallExpr 分支同一份 denylist 语义
    // 配套算出, 两处口径必须一致.
    auto isPolymorphicBuiltin = [](const std::string& name) -> bool {
        std::string lower = name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        return lower == "iif" || lower == "choose" || lower == "switch"
            || lower == "if";
    };

    auto checkSym = [&](Symbol* sym) -> bool {
        if (!sym) return false;
        if (sym->type == Vb6Type::Variant) return true;
        if (static_cast<uint16_t>(sym->type) == variantArrRaw) {
            if (isArrOut) *isArrOut = true;
            return true;
        }
        return false;
    };

    switch (expr.kind) {
        case ASTNodeKind::IdentifierExpr: {
            auto& id = static_cast<IdentifierExpr&>(expr);
            std::string lower = id.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            // 已知数组变量: C 类型已是 vb6_SafeArray1D* (元素为具体类型或 VARIANT),
            // 不是 vb6_VARIANT. 防止 symTab_ lookupModule 回退命中其他模块同名
            // Variant 符号, 导致 UBound(arr) 被错误包装成 vb6_VariantToSafeArray1D(arr)
            // (C2440: 无法从 vb6_SafeArray1D* 转换为 vb6_VARIANT, 如 cDataBase WhereIn).
            if (knownArrays_.count(lower)) {
                return false;
            }
            // 已知 Variant 局部变量集合
            if (knownVariantVars_.count(lower)) {
                // 无法区分 Variant 与 Variant(), 视作普通 Variant
                return true;
            }
            // Fix 161f: 无 As Type 但整数可折叠的 Const — 符号表里 type=Variant,
            // 走下面的回退会被当成 Variant 表达式。C 侧是数值 (见 inferExprType
            // 同 Fix), 不是 Variant (extlist MListView SendMessage 实参 C2440)。
            if (moduleIntConstValues_.count(lower)) return false;
            // Fix 049b: 如果已知为非 Variant 具体类型 (BSTR/Long/Double),
            // 不应回退到符号表查找 (可能命中其他模块的同名 Variant 符号)
            // 账 #123: 补 Byte 那一档。缺它时局部 `Dim bt As Byte` 掉到下面的符号表回退 ⇒
            // 被判成 Variant ⇒ 比较发成 vb6_VarCmpLongEq(&bt, …) (拿 1 字节对象的地址当
            // vb6_VARIANT* 传) ⇒ 实测 `(bt = 65)` 返回 False。
            if (knownBstrVars_.count(lower) || knownLongVars_.count(lower)
                || knownDoubleVars_.count(lower) || knownSingleVars_.count(lower)
                || knownByteVars_.count(lower)
                // ai/032: SByte/UInteger/ULong/ULongLong 的局部/形参/返回槽同样是
                // 具体标量, 不是 Variant —— 不显式挡一道的话会落到下面的符号表回退,
                // 命中不住时虽然仍答 false, 但命中同名 Variant 符号时就会误包装。
                || knownNarrowIntVars_.count(lower)) {
                return false;
            }
            // 符号表查询
            auto* sym = symTab_.lookup(id.name);
            if (!sym) sym = symTab_.lookupModule(id.name);
            // Fix <ComCtlsDemo>: 属性遮蔽内置函数时按内置函数类型答 (非 Variant) ——
            // 与 inferExprType 的 IdentifierExpr 护栏同口径. 例: DTPicker 自带
            // `Property Get Month() As Variant` 与内置 Month 撞名, 裸引用 Me 之外的
            // `Month` 若落到该 Variant 属性符号会被误判 Variant.
            if (sym && isPropShadowingBuiltinFunc(sym, id.name)) return false;
            return checkSym(sym);
        }
        case ASTNodeKind::IndexOrCallExpr: {
            auto& call = static_cast<IndexOrCallExpr&>(expr);
            if (call.callee && call.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& cid = static_cast<IdentifierExpr&>(*call.callee);
                // 多态内置函数: 实际返回类型与 symTab 注册不同, 不视为 Variant
                if (isPolymorphicBuiltin(cid.name)) return false;
                auto* sym = symTab_.lookup(cid.name);
                if (!sym) sym = symTab_.lookupModule(cid.name);
                // Fix <ComCtlsDemo>: 名字在内置函数表 (kBuiltinFuncNames) 的**调用**一律
                // 按内置 (typed) 处理, 不当 Variant —— 发码侧就是发这个内置 RTL
                // (vb6_Month 返回 int32_t 等), 与"是否撞名类模块的 Variant 属性"无关.
                // 旧护栏 isPropShadowingBuiltinFunc 只认 PropertyGet/Let/Set 符号, 而
                // DTPicker 自带的 Property Get Month()/Year()/Day()/... As Variant 其
                // 符号 kind 未必是 PropertyGet (可能为 Function/Variable 装载 Variant),
                // 于是漏判 → 误当 Variant → vb6_VariantToLong(vb6_Month(...)) C2440
                // "无法从 int32_t 转换为 vb6_VARIANT". 这里直接按"名字是内置"判,
                // 覆盖面更全, 且与发码侧 (cgen_expr_ident_builtin_table) 同一张表、同一口径.
                // 注: 多态内置 (IIf/Choose/Switch/If) 不在此表, 仍由上方 isPolymorphicBuiltin
                // 处理; 此表只含具类型内置, 返回非 Variant.
                std::string bname = Symbol::toLower(cid.name);
                if (!bname.empty() && bname.back() == '$') bname.pop_back();
                if (kBuiltinFuncNames.count(bname)) return false;
                // 关键差异: sym==null (内置函数) 视为非 Variant
                return checkSym(sym);
            }
            // Fix <ComCtlsDemo>: callee 是成员访问调用 (形如 VBA.Month(x) / obj.Month(x))
            // 且成员名是内置函数 (Month/Year/Day/Hour/Minute/Second 等) 时, 按内置
            // (typed) 处理, 不当 Variant —— 发码侧就是发这个内置 RTL (vb6_Month 返回
            // int32_t 等), 与"是否撞名类模块的 Variant 属性"无关.
            // 例: DTPicker.ctl:1546 `DateSerial(Value, VBA.Month(PropValue), ...)`.
            // 此前 callee 是 MemberAccessExpr, 落到下方递归分支, lookupModule("Month")
            // 命中 DTPicker 自带 `Property Get Month() As Variant` → 误判 Variant →
            // vb6_VariantToLong(vb6_Month(...)) C2440 "无法从 int32_t 转换为 vb6_VARIANT".
            // 这里按"成员名是内置"判 (kBuiltinFuncNames), 与上方 IdentifierExpr 分支
            // 同一张表、同一口径; 上方 IdentifierExpr 分支只覆盖裸名 Month(...), 此处
            // 补上带限定符 VBA.Month(...) 这一形.
            if (call.callee && call.callee->kind == ASTNodeKind::MemberAccessExpr) {
                auto& ma2 = static_cast<MemberAccessExpr&>(*call.callee);
                std::string mname = Symbol::toLower(ma2.memberName);
                if (!mname.empty() && mname.back() == '$') mname.pop_back();
                if (kBuiltinFuncNames.count(mname)) return false;
            }
            // 类方法调用 a.Method(): 递归推断 callee 类型
            if (call.callee && call.callee->kind == ASTNodeKind::MemberAccessExpr) {
                bool arr = false;
                bool v = isDefinitelyVariantExpr(*call.callee, &arr);
                if (v && isArrOut) *isArrOut = arr;
                return v;
            }
            return false;
        }
        case ASTNodeKind::MemberAccessExpr: {
            auto& ma = static_cast<MemberAccessExpr&>(expr);
            // Err 对象特殊处理 (与 inferExprType 一致)
            if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objId = static_cast<IdentifierExpr&>(*ma.object);
                std::string objLower = objId.name;
                std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
                if (objLower == "err") {
                    // Err.* 的具体类型见 inferExprType, 均为具体类型 (Long/String), 非 Variant
                    return false;
                }
            }
            // Fix 090l: UDT 字段访问 — 字段声明 As Variant (如 cZipArchive
            // ZipVfsType.BufferArray / SourceFileInfo) 是明确 Variant 表达式.
            // 此前 memSym==null (UDT 字段) 被一律视为非 Variant, 导致
            // UBound(vb6_ret_X.BufferArray) 等不包装 vb6_VariantToSafeArray1D →
            // C2440 (无法从 vb6_VARIANT 转 vb6_SafeArray1D*).
            if (!inferUdtTypeOfExpr(*ma.object).empty()) {
                return inferExprType(expr) == Vb6Type::Variant;
            }
            // 账 #305 §B171 (第六十二刀): 宿主是工程类时, 这一问先问 §B168 那张字段表 ——
            //   与 §B170 给 inferExprType 补的那一格同一条口径、同一枚否决。下面那条按**裸成员名**
            //   查模块符号的兜底不认识宿主: 成员名撞上内置函数名时它交回内置函数的档位
            //   (实测 `Public PV As Variant` 撞 PV() ⇒ 这里答「不是槽」⇒ `Len(h.PV)` 发裸的
            //   vb6_VARIANT 进 vb6_Len(BSTR) = C2440; 不撞名的 PZ 一直答对)。
            //   表答得出 (非 Unknown) 才接管: 数组 / 定长串 / 对象字段不入表, 仍归兜底。
            if (ma.object) {
                const std::string cls62 = inferClassTypeOfExpr(*ma.object);
                if (!cls62.empty() && resolveClassMemberCall(cls62, ma.memberName).empty()) {
                    const Vb6Type fld62 = classFieldKind(cls62, ma.memberName);
                    if (fld62 != Vb6Type::Unknown) return fld62 == Vb6Type::Variant;
                }
            }
            // 符号表查找成员 (Property/Function)
            auto* memSym = symTab_.lookupModule(ma.memberName);
            // 关键差异: memSym==null (UDT 字段访问或外部类成员) 视为非 Variant
            return checkSym(memSym);
        }
        default:
            // 其他表达式 (BinaryExpr/UnaryExpr/LiteralExpr 等) 不会明确返回 Variant,
            // 除非其子表达式明确为 Variant. 此处不递归, 保持严格性.
            return false;
    }
}

// 账 #238: 比较发码那一族"能不能把这个操作数按 vb6_VARIANT* 交出去"的唯一判据。
// 只回答**裸名字**那一形 (其余形状今天的取址判据不动, 交给调用方自己的形状测试):
// 名字是否 vb6_VARIANT 那份存储, 问 isDefinitelyVariantExpr —— 它读声明那几张表
// (knownVariantVars_ / knownBstrVars_ / knownLongVars_ / knownDoubleVars_ /
// knownSingleVars_ / knownByteVars_ / knownArrays_ / moduleIntConstValues_) 再加符号表,
// 正是这几张表把 `Dim d As Double` 钉成 double 的。以前 vb6_VarCmp*(&A, &B) 的四个取址点
// 各自按"看着像左值"就 &，于是标量局部的地址被当 VARIANT* 递进 RTL: 按 VARIANT 的布局读一个
// 8 字节标量 ⇒ 相等的两个数答 False (实测), 且读过头 (越界读)。
bool CCodeGen::cmpOperandMayTakeAddr(const std::string& c, Expr* ast) const {
    if (c.empty()) return true;
    if (!(std::isalpha(static_cast<unsigned char>(c[0])) || c[0] == '_')) return true;
    for (char ch : c) {
        if (!std::isalnum(static_cast<unsigned char>(ch)) && ch != '_') return true;
    }
    return ast && isDefinitelyVariantExpr(*ast);
}


// ============================================================
// Fix 038b-1: 基于 C 表达式字符串的 Variant 检测
// ============================================================
// 补充 isDefinitelyVariantExpr 的 AST 级检测. 当 codegen 生成的 C 表达式
// 包含已知返回 vb6_VARIANT 的函数调用时, 判定为 Variant.
// 仅检查顶层表达式 (去除前导括号/空白后), 避免对子表达式误判.

// 账 #305 §B172 (第六十三刀): 「把一枚类数据字段的读发出去」的唯一出口 —— 尾注由这一处拼,
// 三站只递参数。为什么必须收在这里: §B171 把「认槽」那一问教成「看尾注那个 variant 词」,
// 而词当时只有**两条**发码路会盖 (点号走 knownClassVars_ 的那一站 + With 体内那一站);
// 第三条 (Fix 088c: 接收者是"文本没进 knownClassVars_ 的类实例" —— 类形参 `UseAll(h As K63Holder)`
// 的标识符发成 `(*h)`) 连尾注都不写, 于是同一枚字段换一种接收者声明方式就一律答「不是槽」,
// InStr(h.PV, …) 直接把整只 vb6_VARIANT 递给 BSTR 形参 = C2440 (两台实测, 都没 exe)。
// 档位与 voidptr 都仍由 §B169 那张表回答: 本函数只是把那张表的答案**投影**成 C 文本。
std::string CCodeGen::classFieldReadExpr(const std::string& objC, const std::string& clsName,
                                         const std::string& memberName, bool isVoidPtr,
                                         const std::string& markerLabel) const {
    // Fix 092p 的口径留在这一处: VB6 大小写不敏感, 源码拼写要规范化回声明名, 否则
    // `...->socket` (声明 Socket) 是 C2039。
    const std::string base = objC + "->" + cIdent(canonicalClassFieldName(clsName, memberName));
    if (isVoidPtr) {
        return base + "  /* " + markerLabel + " ." + memberName + " field voidptr */";
    }
    return base + "  /* " + markerLabel + " ." + memberName + " field"
         + (classFieldIsVariantSlot(clsName, memberName) ? " variant" : "") + " */";
}

// 账 #305 §B174 (第六十五刀): 「数值实参交进 VB 的字符串形参 ⇒ VB6 会隐式 CStr」这一条规则,
// 发码侧此前有**两位主人**，本刀收成一位：
//   主人一 = Fix 175 (cgen_expr_call_arg_emit.inc) —— 按**形参表**那一档答的 (ByVal String)，
//             形状走 wrapToBSTR (全仓唯一的"数值→BSTR"形状出口)。
//   主人二 = Fix 108c (cgen_expr_call_builtin_argtype.inc) —— 按一张 13 枚名单答的, 形状自己
//             写了一遍 if/else, 而且做在 argList **已拼成一整串之后** (靠"从上次命中往后找下
//             一次出现"对齐实参下标)。
// 本站点只留主人一，且它以前答不到的那一格补上了：InStr 这类**故意不声明形参表**的内在函数
// (builtin_funcs.inc:74-76 写明原因 = 可选 start 会让位号错位) 没有形参可问 ⇒ 改问下面那张表。
// 档位集合 = 两位主人原来两份清单的**并集** (13 枚名单的 SByte/UInteger/ULong/ULongLong/Decimal
// + Fix 175 的 Byte/LongPtr) —— 每一档都有一条现成的清单作底, 没量过的 (如 LongLong) 不加。
// 返回 false = 这一档不折 (串 / 对象 / 数组 / Variant / 未知, 以及清单外没量过的档)。
bool CCodeGen::typeTakesImplicitCStr(Vb6Type t) const {
    switch (t) {
        case Vb6Type::Date:
        case Vb6Type::Double:
        case Vb6Type::Single:
        case Vb6Type::Currency:
        case Vb6Type::Decimal:
        case Vb6Type::Long:
        case Vb6Type::Integer:
        case Vb6Type::Byte:
        case Vb6Type::Boolean:
        case Vb6Type::LongPtr:
        case Vb6Type::SByte:
        case Vb6Type::UInteger:
        case Vb6Type::ULong:
        case Vb6Type::ULongLong:
            return true;
        default:
            return false;
    }
}

// 账 #305 §B174 (第六十五刀): 「**没有形参表**的内在函数 callee 的第 pos 位 (VB 实参下标) 是不是
// VB 语义上的字符串形参」。只管 builtin_funcs.inc 里用 addBuiltinFunc (只登记返回类型) 那一族,
// 带类型化形参表的内在函数与工程过程都由 Fix 175 直接问形参表, 不吃这张表。
// 这一问**不许**从运行时 C 签名表 (getRuntimeParamCType) 反推, 两条都在本刀里实测过:
//  ① `vb6_LenB` 在表里那格 "BSTR" 只是 _Generic 宏的第一个分支 (vb6rtl_builtin.h:19) —— VB6 的
//     Len/LenB 对数值交**字节宽度**、不做 CStr, 发码侧另有按类型的 sizeof 分派 (Fix 065/084h/125/160w)。
//     草稿按那张表裸答时, A/B 在 VBFlexGridDemo 上当场打出 `LenB(Version)` 的
//     `sizeof(vb6_type_DLLVERSIONINFO)` 变成 `sizeof(int32_t)` (折出来的串让那条按名字认 UDT 的分派失手)。
//  ② 那张表也答不对**位号**: InStr 的 C 签名是 (start, haystack, needle), 而 VB 写两个实参时 cgen
//     要补 start=1、写三个而首参是串时又换一条路 (见 builtin_fixup 选 InStr/InStrC 那段)。
// ⇒ 表按 VB 手册一枚枚写, 每行都是"这一位收了数值实参 VB6 会隐式 CStr"那条事实。
// Len / LenB / Asc / StrPtr 刻意**不进表** —— ①就是理由, 不是来不及。
// InStrB 不进表: RTL 那两形是 void* (Byte() 或 BSTR), 数值进不去也就没人折; 没量过 = 不加。
bool CCodeGen::builtinNoParamTableStringSlot(const std::string& callee, size_t argc, size_t pos,
                                             Vb6Type firstArgType) const {
    // §B174 的正主: InStr 的字符串位随"有没有写 start"整体移位, 依据与 builtin_fixup 选
    // vb6_InStr / vb6_InStrC 时那一问同源 (首参是不是串 —— 那里问 AST, 这里也问 AST)。
    //   InStr(s1, s2)                   → 串在 0,1
    //   InStr(s1, s2, compare)          → 串在 0,1   (末位是 compare, 不许折)
    //   InStr(start, s1, s2)            → 串在 1,2
    //   InStr(start, s1, s2, compare)   → 串在 1,2   (末位同上)
    if (callee != "vb6_InStr") return false;
    if (argc == 2 || (argc == 3 && firstArgType == Vb6Type::String)) {
        return pos == 0 || pos == 1;
    }
    if (argc == 3 || argc == 4) return pos == 1 || pos == 2;
    return false;
}

// 账 #305 §B171 (第六十二刀): 「按 C 文本认槽」这一问漏认的那一族 —— 类字段的 Variant 槽读。
// 两条发码路 (§B168/§B169 那张表的两处成员读) 现在在尾注里带上 ` variant` 这个词, 而**只有**
// Variant 那一档带 —— 所以这一问的答案仍然只有那张表一个来源, 这里只是把它读回来。
// 判据钉在**串尾**: 成员读的 C 就是 `obj->fld  /* ... */`, 尾注后头什么都不许有 ——
// 拿 "包含" 判会把 `vb6_CStr(h->PZ /*...*/)` 那种**已经消费过槽**的表达式也答成槽 (反向 C2440)。
bool CCodeGen::cExprIsVariantSlotFieldRead(const std::string& cExpr) const {
    static const std::string tail = " field variant */";
    return cExpr.size() >= tail.size()
        && cExpr.compare(cExpr.size() - tail.size(), tail.size(), tail) == 0;
}

bool CCodeGen::cExprIsVariant(const std::string& cExpr) const {
    // 去除前导空白和括号
    size_t start = 0;
    while (start < cExpr.size()) {
        char c = cExpr[start];
        if (c == '(' || c == ' ' || c == '\t' || c == '\n' || c == '\r') {
            start++;
        } else {
            break;
        }
    }
    if (start >= cExpr.size()) return false;

    // 账 #305 §B171 (第六十二刀): 类字段的 Variant 槽读 —— 此前这一问只认得「函数交槽」那一族
    //   (下面的前缀表) 与「数组元素是槽」那一族, 成员读一律答否, 于是实参装箱的两站把整只
    //   vb6_VARIANT 塞进标量槽 (`.bstrVal=h->PZ` ⇒ cl C2440, 两台都没 exe)。
    if (cExprIsVariantSlotFieldRead(cExpr)) return true;

    // 已知返回 vb6_VARIANT 的函数前缀。
    //
    // ⚠ 这张表**不是权威** —— 权威是 RTL 本身: 「一个内置出口交出来的 C 是不是整只槽」问的是
    //   RTL 里那句 `vb6_VARIANT vb6_X(...)` 声明。手写这张表就会与 RTL 走散, 而两个方向的走散都贵:
    //   少一名 = 下游按裸标量发码 (账 #305 §B146 的实物是 `CDec(...) + 1` 直接 C2088「+ 对于
    //   struct 非法」、`d = CDec("2.5")` C2440, 两台真编译都编不过); 多一名 = 死名躺着骗人
    //   (本表清掉了 11 条: RTL 里没有、400 份语料发码里也一次都没出现过 —— 逐名核过。
    //    第 12 条 `vb6_DispCallByVtbl(` 被这张 census 报成死名, 实测**不是**, 见下面那条的注释)。
    //   对账 = 第 54 道哨兵 scripts/check_builtin_slot_returns.ps1: 它从 RTL 派生「返回 vb6_VARIANT
    //   的函数」与「后端会 emit」两个集合, 两个方向逐名比这张表。
    //
    // _Generic 宏只有**每一档变目都交槽**时才登记 (vb6_VariantFromValue 是: 含
    //   `vb6_VARIANT: vb6_VariantIdentity` 档); vb6_CDbl / vb6_CInt / vb6_LenB / vb6_InStr 那一族
    //   宏交的是标量, 按名字登记会一律答错, 所以它们不进表。
    //   指针返回也不进表: vb6_VariantArrayElemPtr 交回 vb6_VARIANT*(槽的地址), 不是槽。
    static const std::vector<std::string> variantPrefixes = {
        "vb6_VariantArrayGet(",
        // 账 #305 §B150: 秩 ≥ 2 的晚绑定槽元素读走这一枚新出口 (variantSlotElemRead 按实参
        //   个数二选一)。不登记 = 表只认识单下标那一形, 新出口发出去下游按裸标量读它。
        "vb6_VariantArrayGetND(",
        "vb6_VariantFromComResult(",
        "vb6_VariantFromStackVARIANT(",
        "vb6_VariantFromValue(",
        "vb6_VariantEmpty(",
        "vb6_VariantLong(",
        "vb6_VariantDouble(",
        "vb6_VariantString(",
        "vb6_VariantBool(",
        "vb6_VariantArray(",
        "vb6_VariantObject(",
        "vb6_VariantNull(",
        "vb6_LoadResData(",       // Fix 090bz: VBA LoadResData → vb6_VARIANT;
                                 //   Dim D() As Byte: D = LoadResData(...) 赋值
                                 //   需 VariantToSafeArray1D 提取 (cLang LoadData/LoadInfo C2440).
        // vbeclipse: LoadResPicture 与 LoadResData 同族, RTL 签名都是
        //   vb6_VARIANT vb6_LoadResPicture(vb6_VARIANT, vb6_VARIANT) (vb6rtl_runtime.h).
        //   不登记会让 `Function getResourceIcon(...) As IPictureDisp` 的
        //   `Set getResourceIcon = LoadResPicture(...)` 直接 `vb6_ret_X =
        //   vb6_LoadResPicture(...)` → C2440 (modResources.c 22/24/33). 登记后走
        //   Set 的 Fix 038b-6 分支包 vb6_VariantToObjectVal 提取对象指针.
        // Fix <vbeclipse> (2026-10-06): LoadRes 实参/返回全面 Variant 化 (2026-10-05
        //   实测 LoadResData("BIN1","CUSTOM") 的 BSTR 实参被截进 int32 形参), 三函数
        //   返回都是 vb6_VARIANT — LoadResString 也一样, 不登记则 `s = LoadResString(1)`
        //   发 vb6_BSTR_Assign 直收 Variant → C2440.
        "vb6_LoadResPicture(",
        "vb6_LoadResString(",
        // ⚠ 这一名**不是** RTL 出口, 是全表唯一的例外, 所以「从 RTL 派生」的那把 census
        //   看不见它 —— 它来自语料自己的工程: mdQRCodegen.bas:1407 写着
        //   `Private Function DispCallByVtbl(...) As Variant`, 而单文件 (非多模块) 发码时工程
        //   过程的名是裸的 `vb6_<proc>` (cProcName: 只有多模块才拼模块名), 于是
        //   variantReturnFuncs_ 也够不着 (那张集合里只有 `vb6_<模块>_<proc>` 和 `vb6_pv_<proc>`)。
        //   实测撤掉这一行: mdQRCodegen 五处 `hResult = DispCallByVtbl(...)` 少掉
        //   vb6_VariantToLong 提取, 整只槽进 vb6_ChkLong(int64_t) —— 这是 400 份 A/B 里唯一
        //   一条归因不到的移动。同形缺口 (工程的 Variant 函数直接赋给非 Variant 目标, 单模块)
        //   两台 BASE/NEW 都还发 `vb6_ChkLong(vb6_MakeDec())`, 另立 §B148。
        "vb6_DispCallByVtbl(",  // Fix 068: DispCallByVtbl returns Variant
        // Fix 110w: VB6 CallByName 返回 vb6_VARIANT (见 vb6rtl_class_com.h) —
        // 参与算术/关系运算或需 BSTR 时必须按 Variant 处理, 否则 C2088
        // ("*" 对于 struct 非法; Charts 2020 ClsResizer.cls:142/148
        //  CallByName(oCtrl, ..., VbGet) * 100).
        "vb6_CallByName(",
        // Fix <vbeclipse> rev37: ParamArray 元素按**声明类型**解包 (rev37 前一律
        //   GetLong ⇒ String 实参静默读成 0)。Variant 元素的解包函数返回整只
        //   VARIANT, 必须登记成 Variant 表达式, 否则下游 BSTR/Variant 目标
        //   不会走 vb6_VariantToString / wrapVariantValue, 直接 C2440
        //   (ClsResizer.AddControlFont 的 `.PropFont = PropFont(i)`)。
        "vb6_PA_GetVariant(",
        // 账 #305 §B145: Variant 参与 + - * / 的四枚运行期出口（vb6rtl_com.c），返回整只
        //   vb6_VARIANT。不登记在这里 = 下游的解包/装箱/比较全按「C 原生标量」发
        //   (int32_t)(...) ⇒ C2440/C2088；登记之后它们一律走 Variant 那一条路。
        "vb6_VarAddV(",
        "vb6_VarSubV(",
        "vb6_VarMulV(",
        "vb6_VarDivV(",
        // 账 #305 §B146: 下面这十名都是「RTL 声明返回 vb6_VARIANT、后端确实会 emit」而表里漏着的
        // (派生自 src/rtl/** 的声明逐名核过; 前两版 census 各错过一次: 一把漏认 `static inline`
        //  前缀、一把把 `vb6_VARIANT *f(...)` 的指针返回当成槽 —— 所以 vb6_VariantArrayElemPtr
        //  刻意**不进**这张表, 它交回的是槽的地址而不是槽)。
        "vb6_CDec(",                // CDec: 唯一能造出精确 Decimal 的内置出口 (vb6rtl_misc.c:271)
        "vb6_CVErr(",              // CVErr -> vb6_VARIANT (vt=10, vb6rtl_conv.c:389)
        "vb6_GetAllSettings(",     // 注册表一组值 -> vb6_VARIANT (vb6rtl_registry.c:113)
        "vb6_IIfVariant(",         // IIf 的两个分支都是槽时选中的那一族
        "vb6_VariantArrayGetVal(", // 晚绑定数组元素读值 (vb6rtl_com.c:698)
        "vb6_VariantByte(",
        "vb6_VariantInt(",
        "vb6_VariantSingle(",
        "vb6_VariantDate(",
        "vb6_VariantIdentity(",
        //   ↑ 上面五枚是构造器 (vb6rtl_variant.h 的 static inline)。表里本来就有 Empty/Long/
        //     Double/String/Bool/Null/Object/Array 那一批, 缺的正是这五档 ——
        //     cgen_util_ctrl.cpp:2419 早就写着「vb6_VariantByte( 不在 cExprIsVariant 的清单里」
        //     并为此另立了第二份名单 (cExprIsVariantCarrierStr); 今天把名补进权威这张表。
    };
    for (const auto& prefix : variantPrefixes) {
        if (cExpr.compare(start, prefix.size(), prefix) == 0) return true;
    }

    // Fix 040b: VB6_SA_AT(vb6_VARIANT, arr, idx) expands to an array element
    // of type vb6_VARIANT — also a Variant expression.
    if (cExpr.compare(start, 21, "VB6_SA_AT(vb6_VARIANT") == 0) return true;

    // Fix 045: 检查项目函数是否返回 Variant — 通过 driver.cpp 预扫描构建的
    // C 函数名集合. 提取 C 表达式中的函数名 (从 start 到第一个 '(') 并查集合.
    if (variantReturnFuncs_) {
        size_t parenPos = cExpr.find('(', start);
        if (parenPos != std::string::npos) {
            std::string funcName = cExpr.substr(start, parenPos - start);
            if (variantReturnFuncs_->count(funcName)) return true;
        }
    }

    // 检查 (&(vb6_VARIANT){...}) 复合字面量 — 也是 VARIANT 类型
    // 但这种形式通常作为 ByRef 参数传递, 不需要再转换, 故不检测.

    return false;
}

// 账 #305 §B147: 「这枚表达式交出来的 C 是不是整只槽」的 AST 那一半 (另一半天然是
// 上面那张文本前缀表)。口径就一条: **发码改道吃 AST ∪ 文本, 类型 oracle 只吃 AST**,
// 于是「oracle 答 Variant」必然蕴含「改道真的发火、真的交出 vb6_VARIANT」。
// 反过来 (oracle 说槽、发码给裸标量) 就是第三十六刀造出的那一族静默回归: 消费点按
// Variant 那一档既不提取也不取整, 直接把裸 double 递进 vb6_ChkByte/ChkInt/ChkLong =
// C 在调用边界截断, 错误 6 也不报了 (§B147 的读数)。
// isArrOut 为真 = C 类型是 vb6_SafeArray1D* (Variant **数组本体**), 那不是槽, 不算。
bool CCodeGen::variantOperandAST(const Expr* e) const {
    if (!e) return false;
    bool isArr = false;
    return isDefinitelyVariantExpr(const_cast<Expr&>(*e), &isArr) && !isArr;
}

// ============================================================
// <vbeclipse>: C 表达式是否**已经是 SafeArray1D\* 载体**
// ============================================================
// Split/Filter/Array 这些内置函数在 VB6 侧的类型是 Variant, 但 codegen 发的是
// 直接返回 vb6_SafeArray1D\* 的 RTL 调用。数组槽实参 (Join 首参 / UBound/LBound
// 首参) 的按 Variant 提取 (vb6_VariantToSafeArray1D) 若套在它们外面就是
// C2440 (vb6_SafeArray1D\* → vb6_VARIANT) —— 实测这三条形全中:
//   Join(Split(s, ","), "|") / UBound(Split(s, ",")) / Join(Filter(a, "x"), "|")
// 与 Fix 092g 的 _arr_N 特例同源、同解法, 区别只是这里包的是内置函数调用。
bool CCodeGen::cExprIsSafeArrayCarrier(const std::string& cExpr) const {
    size_t start = 0;
    while (start < cExpr.size()) {
        char c = cExpr[start];
        if (c == '(' || c == ' ' || c == '\t' || c == '\n' || c == '\r' || c == '*') {
            start++;
        } else {
            break;
        }
    }
    // 匿名数组临时量 (`Array(...)` → _arr_N, Fix 092g) 本身就是载体。
    if (cExpr.compare(start, 5, "_arr_") == 0) return true;

    static const std::vector<std::string> carrierPrefixes = {
        "vb6_Split(",                  // vb6_SafeArray1D* vb6_Split(...)
        "vb6_Filter(",
        "vb6_ArrayCreate(",
        "vb6_ArrayAssign1D(",          // 整体数组赋值的深拷贝 (Fix 170)
        "vb6_SafeArrayCreate1D(",
        "vb6_SafeArrayReDim1D(",
        "vb6_SafeArrayReDimPreserve1D(",
        // Fix <vbeclipse> rev3: 带 elemType 的新入口。注意**不能**指望上一行前缀
        // 命中 —— `vb6_SafeArrayReDimPreserve1D(` 与 `...1D_T(` 在第 25 个字符处分叉。
        "vb6_SafeArrayReDimPreserve1D_T(",
        "vb6_VariantToSafeArray1D(",   // 已提取过, 再包一层就是双重解引用
        "vb6_VariantToByteArray(",
        "vb6_StringToByteArray(",
        "vb6_StrConvToByteArray(",
        "vb6_ComCallByteArray(",
    };
    for (const auto& prefix : carrierPrefixes) {
        if (cExpr.compare(start, prefix.size(), prefix) == 0) return true;
    }
    return false;
}


// ============================================================
// Fix 038b-5: 运行时函数参数 C 类型查找表
// ============================================================
// 当 calleeParams 为空 (运行时/内置函数) 时, 通过函数名和参数索引查找
// 期望的 C 类型. 返回空字符串表示未知.

std::string CCodeGen::getRuntimeParamCType(const std::string& funcName, size_t paramIdx) {
    static const std::unordered_map<std::string, std::vector<std::string>> table = {
        // Array 创建/设置
        {"vb6_ArraySetLong",    {"vb6_SafeArray1D*", "int32_t", "int32_t"}},
        {"vb6_ArraySetBSTR",    {"vb6_SafeArray1D*", "int32_t", "BSTR"}},
        {"vb6_ArraySetDouble",  {"vb6_SafeArray1D*", "int32_t", "double"}},
        {"vb6_ArraySetVariant", {"vb6_SafeArray1D*", "int32_t", "vb6_VARIANT"}},
        {"vb6_ArrayGetLong",    {"vb6_SafeArray1D*", "int32_t"}},
        {"vb6_ArrayGetBSTR",    {"vb6_SafeArray1D*", "int32_t"}},
        {"vb6_ArrayGetDouble",  {"vb6_SafeArray1D*", "int32_t"}},
        // Fix <vbeclipse>: Unload / LoadPicture — 形参是对象指针 (HWND / IDispatch);
        // 实参是 COM 后期绑定读数 (VARIANT) 时按此表解包 (vb6_UnloadForm(l_View.View))。
        {"vb6_UnloadForm",      {"void*"}},
        {"vb6_LoadPictureEx",   {"BSTR"}},
        {"vb6_ArrayGetVariant", {"vb6_SafeArray1D*", "int32_t"}},
        // BSTR 操作
        {"vb6_BSTR_Assign",     {"BSTR*", "BSTR"}},
        {"vb6_BSTR_Concat",     {"BSTR", "BSTR"}},
        {"vb6_BSTR_ConcatFree", {"BSTR", "BSTR"}},
        {"vb6_BSTR_FromStr",    {"const wchar_t*"}},
        {"vb6_BSTR_Empty",      {}},
        {"vb6_BSTR_ToANSI",     {"BSTR"}},
        {"vb6_BSTR_Free",       {"BSTR*"}},
        {"vb6_BSTR_Clone",      {"BSTR"}},
        // 字符串比较/操作
        {"vb6_StrCmp",          {"BSTR", "BSTR"}},
        {"vb6_StrComp",         {"BSTR", "BSTR", "int32_t"}},
        {"vb6_Val",             {"BSTR"}},
        {"vb6_Trim",            {"BSTR"}},
        {"vb6_LTrim",           {"BSTR"}},
        {"vb6_RTrim",           {"BSTR"}},
        {"vb6_Left",            {"BSTR", "int32_t"}},
        {"vb6_Right",           {"BSTR", "int32_t"}},
        {"vb6_Mid",             {"BSTR", "int32_t", "int32_t"}},
        {"vb6_Len",             {"BSTR"}},
        {"vb6_LenB",            {"BSTR"}},
        {"vb6_InStr",           {"BSTR", "BSTR"}},
        {"vb6_Replace",         {"BSTR", "BSTR", "BSTR"}},
        {"vb6_Split",           {"BSTR", "BSTR"}},
        {"vb6_Join",            {"vb6_SafeArray1D*", "BSTR"}},
        {"vb6_UCase",           {"BSTR"}},
        {"vb6_LCase",           {"BSTR"}},
        {"vb6_Space",           {"int32_t"}},
        // Fix 091e: StrConv(BSTR, int32_t, int32_t) — 实参为 Variant 时需
        // vb6_VariantToString (cAesCBC.c 25 StrConv(LoadResData(...), 64, 0) C2440)
        {"vb6_StrConv",         {"BSTR", "int32_t", "int32_t"}},
        {"vb6_String",          {"int32_t", "int32_t"}},
        {"vb6_Chr",             {"int32_t"}},
        {"vb6_Asc",             {"BSTR"}},
        {"vb6_Hex",             {"int32_t"}},
        {"vb6_Oct",             {"int32_t"}},
        // 类型转换
        {"vb6_CStr",            {"vb6_VARIANT"}},
        {"vb6_CLng",            {"double"}},
        {"vb6_CInt",            {"double"}},
        {"vb6_CDbl",            {"double"}},
        {"vb6_CSng",            {"double"}},
        {"vb6_CBool",           {"vb6_VARIANT"}},
        {"vb6_CByte",           {"vb6_VARIANT"}},
        {"vb6_CDate",           {"vb6_VARIANT"}},
        {"vb6_CCur",            {"vb6_VARIANT"}},
        // ai/032: 四档无符号/窄整型转换函数 —— 与 vb6_CLng 同口径 (形参是 double,
        // Variant 实参由 cgen_expr_call_conv_numeric.inc 改写为 *V 版本)。
        {"vb6_CSByte",          {"double"}},
        {"vb6_CUInt",           {"double"}},
        {"vb6_CULng",           {"double"}},
        {"vb6_CULngLng",        {"double"}},
        // twinBASIC 形 CLngPtr: 转 LongPtr, 同族 double 入口 (Variant 实参由
        // cgen_expr_call_conv_numeric.inc 改写为 *V 版本)。
        {"vb6_CLngPtr",         {"double"}},
        // 数组操作
        {"vb6_UBound",          {"vb6_SafeArray1D*", "int32_t"}},
        {"vb6_LBound",          {"vb6_SafeArray1D*", "int32_t"}},
        {"vb6_ArrayCreate",     {"int32_t"}},
        // 消息框
        {"vb6_MsgBox",          {"BSTR"}},
        {"vb6_MsgBox1",         {"BSTR"}},
        // 错误处理
        {"vb6_ErrRaise",        {"int32_t", "BSTR", "BSTR"}},
        {"vb6_ErrNumber",       {}},
        {"vb6_ErrDescription",  {}},
        {"vb6_ErrSource",       {}},
        {"vb6_ErrClear",        {}},
        // IsMissing
        // Fix 091k: RTL 实际签名 int32_t vb6_IsMissing(SAFEARRAY* psa) (ParamArray 专用,
        // 判断是否未传实参). 原表项 vb6_VARIANT* 与 RTL 不符, 导致 IsMissing(<Variant>)
        // 时表项无效, 裸传 Variant 值 → C2440.
        {"vb6_IsMissing",       {"SAFEARRAY*"}},
        // Variant 提取
        {"vb6_VariantToLong",      {"vb6_VARIANT"}},
        {"vb6_VariantToDouble",    {"vb6_VARIANT"}},
        {"vb6_VariantToString",    {"vb6_VARIANT"}},
        {"vb6_VariantToBool",      {"vb6_VARIANT"}},
        {"vb6_VariantToSafeArray1D", {"vb6_VARIANT"}},
        {"vb6_VariantToObjectVal", {"vb6_VARIANT"}},
        // Fix 046: IIf family + Variant-aware conversion functions
        {"vb6_IIfBSTR",         {"int32_t", "BSTR", "BSTR"}},
        {"vb6_IIfLong",         {"int32_t", "int32_t", "int32_t"}},
        {"vb6_IIfDouble",       {"int32_t", "double", "double"}},
        {"vb6_IIfVariant",      {"int32_t", "vb6_VARIANT", "vb6_VARIANT"}},
        {"vb6_CLngV",           {"vb6_VARIANT"}},
        {"vb6_CIntV",           {"vb6_VARIANT"}},
        {"vb6_IntDiv",          {"int32_t", "int32_t"}},
        // Debug
        {"vb6_DebugPrint",      {"BSTR"}},
        {"vb6_DebugWriteLong",  {"int32_t"}},
        // 对象操作
        {"vb6_StrPtr",          {"BSTR"}},   // Fix 084o-7: StrPtr(Variant) → vb6_VariantToString 先行
        // Fix 155: ObjPtr 实参签名是 `vb6_ObjPtr(void* obj)` (vb6rtl_builtin.h:288),
        // 此前登记 "uintptr_t" (那是**返回**类型, 非形参). 运行时提取分支按形参类型
        // 匹配 ("void*" → vb6_VariantToObjectVal), "uintptr_t" 无对应分支 → ObjPtr
        // 收到 vb6_VARIANT (如 ObjPtr(ParentControls.Item(0)) 的 COM 结果) 时不做
        // 提取, 把结构体裸传给 void* 形参 → C2172 "实参不是指针". 改为真实形参类型.
        {"vb6_ObjPtr",          {"void*"}},
        {"vb6_ReleaseObject",   {"void**"}},
        {"vb6_NewObject",       {"const wchar_t*"}},
        {"vb6_CallByName",      {"void*", "BSTR", "int32_t"}},
        // Fix 113: UserControl 宿主内建方法 (vb6rtl_userctl.h) — 裸名书写,
        // cgen 映射为 vb6_UserControl_<Member> (cgen_expr_ident_builtin.inc).
        // 此前不在运行时参数表, 实参为 vb6_ComCall(...) 等 COM Variant 时缺 BSTR
        // 提取 → 传入 GetTextExtentPoint32W/字符串 API 崩溃. 注册 BSTR 形参.
        {"vb6_UserControl_TextWidth",      {"BSTR"}},
        {"vb6_UserControl_TextHeight",     {"BSTR"}},
        // 账 #196: 控件那一对（Form / PictureBox 的 .TextWidth/.TextHeight）同一条规矩 ——
        // 第一个实参是句柄(void*)，第二个必须是 BSTR；实参是 Variant 时缺这条就
        // 把 vb6_VARIANT 结构体裸传给 GetTextExtentPoint32W ⇒ 崩（Fix 113 记的那一味）。
        {"vb6_ControlTextWidth",           {"void*", "BSTR"}},
        {"vb6_ControlTextHeight",          {"void*", "BSTR"}},
        {"vb6_UserControl_AsyncRead",      {"BSTR", "int32_t", "BSTR", "int32_t"}},
        {"vb6_UserControl_PropertyChanged", {"BSTR"}},
        {"vb6_UserControl_CancelAsyncRead", {"BSTR"}},
    };
    auto it = table.find(funcName);
    if (it != table.end() && paramIdx < it->second.size()) {
        return it->second[paramIdx];
    }
    return "";
}

// ============================================================
// Fix 092w: Byte 数组赋值右侧改写
// ============================================================
// VB6/twinbasic 允许 Dim ba() As Byte = "..." 或 = StrConv(s, 64/128), 语义为
// 生成含字节内容的动态数组. RTL 的 vb6_StrConv 返回 BSTR, 不能直接赋给
// vb6_SafeArray1D*. 此处在编译期改用返回字节数组的 RTL helper.
std::string CCodeGen::rewriteByteArrayValue(const std::string& value) const {
    if (value.empty() || value == "NULL") return value;

    // StrConv(...) 用于构造字节数组 → 改用 vb6_StrConvToByteArray(...)
    if (value.compare(0, 12, "vb6_StrConv(") == 0) {
        return "vb6_StrConvToByteArray" + value.substr(11);
    }
    // 已经是字节数组/数组表达式 → 原样返回
    if (value.find("vb6_StrConvToByteArray(") != std::string::npos ||
        value.find("vb6_StringToByteArray(") != std::string::npos ||
        value.compare(0, 14, "vb6_SafeArray") == 0 ||
        value.compare(0, 22, "vb6_VariantToByteArray") == 0 ||
        value.compare(0, 25, "vb6_VariantToSafeArray1D") == 0) {
        return value;
    }
    // 字符串/BSTR 表达式 → 复制为字节数组 (原始 UTF-16LE 字节)
    // Fix 140: 补充 Ambient.DisplayName (BSTR) — LabelPlus.ctl UserControl_InitProperties
    // 里 `m_Caption = Ambient.DisplayName` 生成 `me->m_Caption = vb6_Ambient_DisplayName`,
    // 是 BSTR 赋给 Byte() 字段, 需同样改写成字节数组.
    bool isBstrExpr = value.compare(0, 9, "vb6_BSTR_") == 0 ||
                      value.compare(0, 20, "vb6_VariantToString(") == 0 ||
                      value == "vb6_Ambient_DisplayName";
    if (isBstrExpr) {
        return "vb6_StringToByteArray(" + value + ")";
    }
    // Fix 140: ByRef 解引用形态 `(*Param)` — 形参 C 类型为 BSTR* (ByRef String),
    // `(*Param)` 即真实 BSTR. 若内层名字是已知 BSTR 变量则按"字符串→字节数组"
    // 改写. 场景: LabelPlus.ctl `Property Let Caption(ByRef New_Caption As String)`
    // 内 `m_Caption = New_Caption` 生成 `me->m_Caption = (*New_Caption)`, 若不改写
    // 会把 BSTR 当 SafeArray1D* 赋给 Byte() 字段 → caption 读不到, 卡片空白.
    if (value.size() > 4 && value[0] == '(' && value[1] == '*'
        && value[value.size() - 1] == ')') {
        std::string inner = value.substr(2, value.size() - 3);
        if (inner.compare(0, 4, "me->") == 0) inner = inner.substr(4);
        if (inner.find_first_of("( .->") == std::string::npos) {
            std::string lower = inner;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (knownBstrVars_.count(lower)) {
                return "vb6_StringToByteArray(" + value + ")";
            }
        }
        return value;
    }
    // 裸变量名: 若为已知 BSTR 变量则包装
    std::string name = value;
    if (name.compare(0, 4, "me->") == 0) name = name.substr(4);
    if (name.find_first_of("( .") == std::string::npos) {
        std::string lower = name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownBstrVars_.count(lower)) {
            return "vb6_StringToByteArray(" + value + ")";
        }
    }
    return value;
}

// Fix 170: VB6 **整体数组引用** `A()` —— IndexOrCallExpr 无实参, callee 是已知数组变量或
// 模块级数组。判定口径与 cgen_expr_call_prelude.inc 的 isArrayAccess 一致 (先 knownArrays_,
// 再查模块符号表 sym->isArray), 因此 `GetTickCount()` / `Command()` 这类零参调用不会误判。
bool CCodeGen::isWholeArrayRef(const Expr* e) const {
    if (!e || e->kind != ASTNodeKind::IndexOrCallExpr) return false;
    auto& n = static_cast<const IndexOrCallExpr&>(*e);
    if (!n.positional.empty() || !n.named.empty()) return false;
    if (!n.callee) return false;
    if (n.callee->kind == ASTNodeKind::IdentifierExpr) {
        auto& id = static_cast<const IdentifierExpr&>(*n.callee);
        if (knownArrays_.count(Symbol::toLower(id.name))) return true;
        Symbol* sym = symTab_.lookupModule(id.name);
        return sym && sym->kind == SymbolKind::Variable && sym->isArray;
    }
    // Fix 178: `.Cols() = VBFlexGridDefaultCols.Cols()` —— UDT/With 的动态数组成员
    // 用空括号整体赋值。原先只认裸数组变量, 这类目标退化成载体指针直赋 → 别名。
    return isDynamicArrayMemberCallee(n.callee.get(), nullptr);
}

// 账 #305 §B161/§B159 (第五十刀): 数组载体实参这一问的唯一主人 (声明处记着四个排除项)。
// 刻意**不看** inferExprType: 它对模块级数组答的是**元素**档 (Long/Byte/String), 那两个消费者要的
// 恰恰是「别按元素档处置」—— 第四十九刀的实读数 (3/"Long"、17/"Byte"、8/"String") 就是照着它抄的结果。
bool CCodeGen::isArrayCarrierArgName(const std::string& lowerName,
                                     const std::string& cExpr) const {
    if (knownArrays_.count(lowerName) == 0) return false;
    if (knownObjectVars_.count(lowerName)) return false;
    if (knownVariantVars_.count(lowerName)) return false;
    if (arrayClassElemTypes_.count(lowerName)) return false;
    if (cExprIsVariant(cExpr)) return false;
    return true;
}

// Fix 178: callee 是「UDT 的动态数组成员」(`.Cols` / `Default.Cols`) 时返回 true,
// 并按需带出元素 UDT 的 C 类型 (元素是标量类型时保持原值)。
bool CCodeGen::isDynamicArrayMemberCallee(const Expr* callee, std::string* elemUdt) const {
    if (!callee) return false;
    std::string parentUdt, memName;
    if (callee->kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<const MemberAccessExpr&>(*callee);
        if (!ma.object) return false;
        parentUdt = inferUdtTypeOfExpr(*ma.object);
        memName = ma.memberName;
    } else if (callee->kind == ASTNodeKind::WithMemberExpr) {
        if (withObjectInfoStack_.empty() || withObjectVars_.empty()) return false;
        if (withObjectInfoStack_.back().kind != WithObjKind::Unknown) return false;
        auto it = knownUdtVars_.find(Symbol::toLower(withObjectVars_.back()));
        if (it == knownUdtVars_.end()) return false;
        parentUdt = it->second;
        memName = static_cast<const WithMemberExpr&>(*callee).memberName;
    } else {
        return false;
    }
    const std::string prefix = "vb6_type_";
    if (parentUdt.size() <= prefix.size()
        || parentUdt.compare(0, prefix.size(), prefix) != 0) return false;
    Symbol* sym = symTab_.lookupModule(parentUdt.substr(prefix.size()));
    if (!sym || sym->kind != SymbolKind::UserDefinedType) return false;
    std::string memLower = Symbol::toLower(memName);
    for (const auto& mi : sym->udtMembers) {
        if (Symbol::toLower(mi.name) != memLower) continue;
        if (!mi.isArrayDynamic) return false;
        if (elemUdt && mi.type == Vb6Type::UserDefinedType && !mi.typeRefName.empty())
            *elemUdt = "vb6_type_" + cIdent(mi.typeRefName);
        return true;
    }
    return false;
}

// Fix 170: 整体数组赋值 `A() = expr` 的右侧收口。
// 只放行**必然新建载体**的 helper (vb6_StringToByteArray / StrConvToByteArray /
// Array(...) 物化 / Split / Filter / COM ByteArray 解封)；其余形态一律走
// vb6_ArrayAssign1D 深拷贝：
//   · 裸数组变量 `A() = B()` —— 直接赋是两个名字别名同一载体;
//   · vb6_VariantToByteArray / vb6_VariantToSafeArray1D —— Variant 持数组时**返回
//     v.parray 本身** (见 vb6rtl_compat.c:187)，不拷贝就是与那个 Variant 共用载体。
std::string CCodeGen::wrapWholeArrayAssign(const std::string& target,
                                           const std::string& rhs,
                                           const Expr* targetNode) const {
    static const std::vector<std::string> freshCarrier = {
        "vb6_StringToByteArray(", "vb6_StrConvToByteArray(",
        "vb6_ArrayCreate(", "vb6_ArrayAssign1D(", "vb6_ComCallByteArray(",
        "vb6_VariantArray(", "vb6_Split(", "vb6_Filter(",
    };
    // 跳过空白与左括号/解引用 (`(*Arr)` 形态的 ByRef 参数取的是载体本身)
    size_t s = rhs.find_first_not_of(" \t\r\n(*");
    if (s != std::string::npos) {
        for (const auto& p : freshCarrier)
            if (rhs.compare(s, p.size(), p) == 0) return rhs;
    }
    // Fix 178: 元素是含所有权成员的 UDT 时, 载体克隆还要逐元素深拷贝, 否则两侧元素
    // 共享同一 BSTR / 子数组 (VBFlexGridCells.Rows(i).Cols 的别名就是这么来的)。
    std::string elemUdt;
    if (targetNode && targetNode->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& n = static_cast<const IndexOrCallExpr&>(*targetNode);
        if (n.callee) {
            if (n.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& id = static_cast<const IdentifierExpr&>(*n.callee);
                auto it = arrayUdtElemTypes_.find(Symbol::toLower(id.name));
                if (it != arrayUdtElemTypes_.end()) elemUdt = it->second;
            } else {
                isDynamicArrayMemberCallee(n.callee.get(), &elemUdt);
            }
        }
    }
    if (!elemUdt.empty() && elemUdt.compare(0, 9, "vb6_type_") == 0
        && udtHasOwnedMembers(elemUdt)) {
        requestUdtCopy(elemUdt, true);
        return "vb6_ArrayAssign1D_Cb(" + target + ", " + rhs + ", vb6_udtcpy_"
             + elemUdt.substr(9) + "_v)";
    }
    return "vb6_ArrayAssign1D(" + target + ", " + rhs + ")";
}

// ============================================================
// ai/009 §5.10 (P3) 溢出检查
// ============================================================
// 判"右值类型是否可能越出目标范围"。用**比特宽**比, 而不是逐类型列白名单:
// 目标宽度以外的整型、以及一切浮点 (Single/Double/Currency/Date), 都能越界;
// 目标宽度以内的 (含 Boolean —— 值域只有 -1/0), 永远不可能, 套检查纯属噪声。
// Unknown/Variant 一律按"可能越界"算 —— 宁可多查, 不可漏查, 因为漏查就是静默错编。
// 账 #230 §B138b: ptrBits = 目标架构的指针位宽 (x86 → 32, x64 → 64)。
//   **LongPtr 与 VBA 里一样是跟平台走的**: x86 下就是 32 位, x64 下才是 64 位。
//   改前这里把它和 Long/Single 一起写死在 32 位档 —— 与紧邻那句注释
//   ("LongLong/LongPtr 是 64 位整数") 自相矛盾; 后果是 x64 上把 LongPtr 存进
//   Byte/Integer/Long 时, srcBits(32) <= tgtBits(32) ⇒ "装得下" ⇒ **不套检查**,
//   静默截掉高 32 位 (正是 VarPtr/StrPtr/ObjPtr 那三枚交地址的典型下场)。
//   **LongLong 则始终 64 位**, 与平台无关 (C 侧固定 int64_t), 故不进这一档。
static int cgenIntBits(Vb6Type t, int ptrBits) {
    switch (t) {
    case Vb6Type::Byte: case Vb6Type::Boolean: return 8;
    case Vb6Type::Integer: return 16;
    case Vb6Type::Long: case Vb6Type::ULong:
    case Vb6Type::Single: return 32;
    case Vb6Type::LongPtr: return ptrBits;
    // Currency/Date 在 C 侧是 double, LongLong 恒 64 位整数, 都可能越界
    case Vb6Type::Double: case Vb6Type::Currency: case Vb6Type::Date:
    case Vb6Type::LongLong: return 64;
    // C3 扩展 (ai/032): 新整型按实际位宽入档。注意 SByte 与 Byte 同为 8 位 ——
    // 这里问的是"值可能越出目标范围吗", 只看宽度, 不看符号。
    case Vb6Type::SByte: return 8;
    case Vb6Type::UInteger: return 16;
    case Vb6Type::ULongLong: return 64;
    default: return 0;   // Unknown / Variant / String / Object ... 由调用点决定
    }
}

// 账 #261 (§B77②): 「这枚左值是哪档窄整型」只在这里回答一次。
// 三条口径:
//   1) 答案只许是 Byte / Integer / Long 三档之一。认不出 (Array 旗标 / UDT 整体 /
//      String / Object / Variant / 推不出) 一律答 Unknown, 调用方原样发 —— 与改前同形。
//      猜错会把合法赋值判成越界 (那比不检查更糟), 所以宁可答 Unknown。
//   2) 数组元素按**声明的元素类型**答 (arrayElemTypes_ 那张表由六条声明路填),
//      不是按索引表达式。
//   3) UDT 字段问 inferUdtFieldVb6Type —— 那是字段声明类型的既有唯一出口; 它带着
//      Array 旗标 (整型数组成员) 时在这里退回 Unknown, 绝不套标量检查 (rev36 那枚
//      run-time error 6 就是这么来的: 指针被当 Byte 值送进 vb6_ChkByte)。
Vb6Type CCodeGen::narrowTargetTypeOf(Expr* target) const {
    if (!target) return Vb6Type::Unknown;

    if (target->kind == ASTNodeKind::MemberAccessExpr ||
        target->kind == ASTNodeKind::WithMemberExpr) {
        // ⚠ 数组成员不是标量槽。`Data() As Byte` 这类字段的 mi.type 存的是**元素**档
        //   (数组性在 mi.isArrayDynamic / mi.arraySize 上), 所以光看档位认不出它 ——
        //   实测把 `With x : .Data = baData` 包成 `vb6_ChkByte(数组描述符指针)`，
        //   VbQRCodegen 的 Project1 从此启动期 Unhandled VB6 Error #6 (BASE 同一份源不报)。
        //   与 rev36 那枚雷同一形状, 只是换了一条入口。
        bool fldIsArray = false;
        Vb6Type ft = inferUdtFieldVb6Type(target, &fldIsArray);
        if (fldIsArray) return Vb6Type::Unknown;
        if (ft == Vb6Type::Byte || ft == Vb6Type::Integer || ft == Vb6Type::Long) return ft;
        return Vb6Type::Unknown;
    }

    if (target->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& ioc = static_cast<IndexOrCallExpr&>(*target);
        if (!ioc.callee) return Vb6Type::Unknown;
        // ⚠ **空下标 = 整体数组赋值** (Fix 170 那一形: `dst() = src()` / `bb() = s`)。
        //   它发成的 C 是 `dst = vb6_ArrayAssign1D(dst, src)` —— 右边是**数组描述符指针**，
        //   不是元素值。实测把它按元素档套上检查 ⇒ `vb6_ChkByte(vb6_StringToByteArray(s))`
        //   ⇒ 指针送进标量闸 ⇒ run-time error 6 (test_array.bas 改后 EXIT=0x00000006,
        //   wa-clone/wa-ub/wa-str/wa-rt 四行整片不打印)。与 rev36 那条同一个雷, 只是
        //   换了一条入口, 所以这一档先问"有没有下标"再问元素类型。
        if (ioc.positional.empty()) return Vb6Type::Unknown;
        if (ioc.callee->kind == ASTNodeKind::MemberAccessExpr ||
            ioc.callee->kind == ASTNodeKind::WithMemberExpr) {
            // 账 #262 之后才敢接这一形: 成员数组的**元素**槽 (`p.Pixels(0)` / `m.M(0, 1)`)。
            // 以前它的左值本身就是错的 —— 多维被折成一维, 三个下标写进同一格, 再套一层
            // 取整只会把两格缺陷搅成一格读数。现在那几格是确定的了。
            // 必须问出**数组性**: 认不出是数组就答 Unknown (不是数组的带括号左值各有各的
            // 解析链, 猜一档会把合法赋值判成越界)。
            bool fldIsArray = false;
            Vb6Type et = inferUdtFieldVb6Type(ioc.callee.get(), &fldIsArray);
            if (!fldIsArray) return Vb6Type::Unknown;
            if (et == Vb6Type::Byte || et == Vb6Type::Integer || et == Vb6Type::Long) return et;
            return Vb6Type::Unknown;
        }
        // 裸标识符 = 本模块数组的元素。
        if (ioc.callee->kind != ASTNodeKind::IdentifierExpr) return Vb6Type::Unknown;
        auto& cid = static_cast<IdentifierExpr&>(*ioc.callee);
        std::string cLower = cid.name;
        std::transform(cLower.begin(), cLower.end(), cLower.begin(), ::tolower);
        auto it = arrayElemTypes_.find(cLower);
        if (it == arrayElemTypes_.end()) return Vb6Type::Unknown;
        Vb6Type et = it->second;
        if (et == Vb6Type::Byte || et == Vb6Type::Integer || et == Vb6Type::Long) return et;
        return Vb6Type::Unknown;
    }

    if (target->kind != ASTNodeKind::IdentifierExpr) return Vb6Type::Unknown;

    // 这里**故意不复用** inferExprType 来定目标类型: 它把 Integer 答成 Long、
    // 把 Byte 答成 Variant/Unknown, 那是几十个消费点共同依赖的既有口径, 动它
    // 等于给 Debug.Print / Variant 装箱等一整条链换答案。这里只为溢出检查
    // 单独查一遍精确的窄整型, 顺带靠这个局部性把影响面关在收窄赋值里。
    auto& id = static_cast<IdentifierExpr&>(*target);
    std::string lower = id.name;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);

    // Fix <vbeclipse> rev36: **Byte 数组成员**不是标量, 绝不套窄整型溢出检查。
    //
    // 症状 (Charts 2020 实测, LabelPlus.ctl `Dim m_Caption() As Byte` +
    // `Property Let Caption`): 生成
    //   me->m_Caption = vb6_ChkByte(vb6_StringToByteArray((*New_Caption)));
    // vb6_StringToByteArray 返回 `vb6_SafeArray1D*` (头文件 vb6rtl_builtin.h:362),
    // x86 下这个指针被当 Byte 值送进 vb6_ChkByte ⇒ 必然落在 0~255 之外 ⇒
    // run-time error 6 "Overflow", 启动即弹框 (实测 exit=0x00000006)。
    // ⚠ 右侧**转换本身是对的** (Fix 140 已把 BSTR → vb6_StringToByteArray 改写对),
    //   错的只是外面多套的那层标量检查 —— 对照 getter 是 `vb6_ByteArrayToString(me->m_Caption)`,
    //   没有检查。所以这里必须**只**去掉检查, 不要动右值改写。
    //
    // 为什么在这里拦: 下面所有分支都只认**标量**类型表 (knownByteVars_ 等装的
    // 是 `As Byte` 标量), 数组成员一个都不在 ⇒ 落到 `inferExprType` 兜底,
    // 而它按元素类型答 Byte ⇒ 被当标量。`knownByteArrayVars_` /
    // `classByteArrayMembers_` 正是为"这是字节数组"准备的登记表(Fix 140 起),
    // 全代码库十几处消费点都在查它, **唯独这里漏了** —— 那才是本条的真根因。
    // `classByteArrayMembers_` 要一并查: 类字段在过程入口从 knownByteArrayVars_
    // copy 回, 但那是 copy 不是同一容器, 且保守起见两张都查。
    //
    // 判据用 `lower` 全名(含 m_ 前缀)。属性形参(如 New_Caption)不在这两张表里,
    // 不受影响 —— 只有**被赋值的左值**进这条路径。
    //
    // ⚠ `id.name` 的形态要归一: 成员赋值的目标在这里既可能带 `me->` 前缀也可能不带,
    //   而 classByteArrayMembers_ 登记的是 `m_X` / `X` 两种**裸名**(见
    //   cgen_base_generate_state_scan.inc:53-54 的 mLower/oLower)。三处都试一遍,
    //   否则带前缀的那条(实测就是它)永远命不中 —— 漏这一步会让本修复看起来"没生效"。
    if (knownByteArrayVars_.count(lower) || classByteArrayMembers_.count(lower))
        return Vb6Type::Unknown;
    {
        std::string bare = lower;
        if (bare.compare(0, 4, "me->") == 0) bare = bare.substr(4);
        else if (bare.size() > 4 && bare[0] == '(' && bare[1] == '*'
                 && bare.back() == ')') bare = bare.substr(2, bare.size() - 3);
        if (bare != lower
            && (knownByteArrayVars_.count(bare) || classByteArrayMembers_.count(bare)))
            return Vb6Type::Unknown;
    }

    if (knownByteVars_.count(lower))        return Vb6Type::Byte;
    if (knownIntVars_.count(lower))         return Vb6Type::Integer;
    if (knownBoolVars_.count(lower))        return Vb6Type::Unknown;  // 值域只有 -1/0
    if (knownLongVars_.count(lower))        return Vb6Type::Long;
    return inferExprType(*target);
}

std::string CCodeGen::narrowCheckAssign(Expr* target, Expr* value,
                                        const std::string& cValue) const {
    if (!target || cValue.empty()) return cValue;
    // 账 #261: 目标的窄整型档只从上面那一处出口问来 (改前这里自带一道
    // `kind != IdentifierExpr` 的闸门, 成员/数组元素两形因此整片不经检查)。
    Vb6Type tt = narrowTargetTypeOf(target);

    const char* fn = nullptr;
    int tgtBits = 0;
    switch (tt) {
    case Vb6Type::Byte:    fn = "vb6_ChkByte"; tgtBits = 8;  break;
    case Vb6Type::Integer: fn = "vb6_ChkInt";  tgtBits = 16; break;
    case Vb6Type::Long:    fn = "vb6_ChkLong"; tgtBits = 32; break;
    default: return cValue;
    }

    Vb6Type vt = value ? inferExprType(*value) : Vb6Type::Unknown;
    // 账 #248: 浮点源**必须在这一套检查之前先过那一份取整出口**。
    // 直接把 double 递进 vb6_ChkLong(int64_t) 会让 C 在调用边界上截断 —— 实测
    // `l = 6.73` 交 6、`l = 7 / 2` 交 3，而**同一句**写成 `l = CLng(7 / 2)` 交 4
    // (vb6_CLng 内部走 round())。同一件事两个答案就是"同一个决定抄了两遍"那一族。
    // 也**不能**靠下面那句"装得下, 不套"放过：cgenIntBits(Single) 答 32，Long 目标 32，
    // 于是 `l = 某Single` 以前连检查都不进 —— 截得更彻底。
    if (vt == Vb6Type::Single || vt == Vb6Type::Double || vt == Vb6Type::Currency)
        return std::string(fn) + "(vb6_FltToLng(" + cValue + "))";

    // §B138b: LongPtr 的宽度跟目标架构走 (x86=32 / x64=64), 见 cgenIntBits 处注释。
    const int ptrBits138 = (targetArch_ == "x86") ? 32 : 64;
    int srcBits = value ? cgenIntBits(vt, ptrBits138) : 0;
    if (srcBits != 0 && srcBits <= tgtBits) return cValue;   // 装得下, 不套

    return std::string(fn) + "(" + cValue + ")";
}

// ============================================================
// 账 #305 §B150 (第四十刀): 晚绑定槽 (Variant 变量 / 形参 / 字段) 的元素读与写
// ============================================================
// 秩数由**实参个数**决定: 一枚走 vb6_VariantArrayGet/Set (既有语义, 存量产物逐字节不变),
// 两枚及以上走 GetND/SetND —— 那两枚出口把下标一次交全, 上下界与错误 9 由 RTL 里唯一一枚
// vb6_SaNdElemPtr 管 (账 #209 那一族的口径)。
std::string CCodeGen::variantSlotElemRead(const std::string& slotArg,
                                          const std::vector<ExprPtr>& positional) {
    std::vector<std::string> idxs;
    for (const auto& e : positional) {
        if (!e) continue;
        emitExpr(*e);
        // Fix 084o: 索引为 Variant 时转 Long (这些出口的形参都是 int32_t)
        idxs.push_back(toLongIfVariant(std::move(lastExpr_), e.get()));
    }
    if (idxs.size() <= 1) {
        return "vb6_VariantArrayGet(" + slotArg + ", " + (idxs.empty()
                          ? std::string("") : idxs[0]) + ")";
    }
    std::string call = "vb6_VariantArrayGetND(" + slotArg + ", " +
                       std::to_string(idxs.size());
    for (const auto& s : idxs) call += ", " + s;
    return call + ")";
}

// 账 #305 §B150 (下半): 一张表管「槽元素读」的三种名 —— 读、取址、写。参数形态三者
// 完全同形 (槽 [, 秩数] , 秩数个下标 [, 末尾一枚盒]) —— 秩数那格就是为「只换函数名」
// 这一件事设计的 (GetND/SetND 的形参表见 vb6rtl_com.c)。发码侧凡是
// 拿**已发出的 C 文本**问「这是不是一枚槽元素读」的地方一律进这张表: 以前六处各抄一遍,
// 而两种抄法 (带不带尾括号) 对本刀的新名答案相反 ⇒ 同一形在 VarPtr 路是 C2440, 在 Set
// 路是整条赋值被静默丢掉。取址那条越界交错误 9 (vb6_SaNdElemPtr 的口径, 账 #209 那一族),
// 与一维那枚返 NULL 的既有语义刻意不同 —— 一维那条不在本格。
namespace {
struct SlotElemForms {
    const char* get;
    const char* addr;
    const char* set;
};
const SlotElemForms kSlotElemForms[] = {
    { "vb6_VariantArrayGet(",   "vb6_VariantArrayElemPtr(",   "vb6_VariantArraySet("   },
    { "vb6_VariantArrayGetND(", "vb6_VariantArrayElemPtrND(", "vb6_VariantArraySetND(" },
};
} // namespace

bool CCodeGen::isVariantSlotElemRead(const std::string& text) const {
    for (const auto& f : kSlotElemForms)
        if (text.compare(0, std::char_traits<char>::length(f.get), f.get) == 0) return true;
    return false;
}

std::string CCodeGen::variantSlotElemAddrRead(const std::string& text) const {
    for (const auto& f : kSlotElemForms) {
        const size_t n = std::char_traits<char>::length(f.get);
        if (text.compare(0, n, f.get) != 0) continue;
        if (text.size() <= n || text.back() != ')') return std::string();
        return std::string(f.addr) + text.substr(n, text.size() - n - 1) + ")";
    }
    return std::string();
}

bool CCodeGen::variantSlotAssignLine(const std::string& target, const std::string& boxedValue,
                                     std::string& outLine) const {
    for (const auto& f : kSlotElemForms) {
        const size_t n = std::char_traits<char>::length(f.get);
        if (target.compare(0, n, f.get) != 0) continue;
        if (target.size() <= n || target.back() != ')') return false;
        outLine = std::string(f.set) + target.substr(n, target.size() - n - 1)
                  + ", " + boxedValue + ")";
        return true;
    }
    return false;
}
} // namespace vb6c3
