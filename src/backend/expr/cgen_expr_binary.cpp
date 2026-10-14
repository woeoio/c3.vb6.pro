#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>
#include <cstdio>

namespace vb6c3 {

// --- cgen_expr_binary.cpp: 二元表达式求值 + BSTR 包装 + 二元运算符映射 ---

// ai/032 rev2: 有符号/无符号混算的显式加宽 —— 只在"C 会算错"的那一档返回非空。
//
// C 的通常算术转换对"同宽异符号"取**无符号** (uint32_t + int32_t → uint32_t,
// uint64_t + int64_t → uint64_t), 而 VB.NET 的规则是"结果取有符号型, 范围至少不窄于
// 两个操作数" —— 正好相反。这一个方向差不是"类型被问错", 是**值会算错**:
//   · `u As ULong : CStr(u + 1000000000)` (u = 3000000000) 在 C 里会经 uint32_t
//     回绕 (实测得 705032704), 而 VB.NET 提升到 Long(64) 后是 4000000000;
//   · `u > -1` 在 C 里按无符号比 (4294967295 > 4294967295 → False), VB.NET 是 True。
// 所以当 promote 给出的结果**严格宽于两个操作数**时, 把两侧显式转成它。
//
// 刻意不动的三类 (返回 nullptr):
//   · 同符号 —— C 自己取更宽的那个, 方向一致;
//   · 提升结果是"两侧之一本来就有的宽度"(包含关系) —— 例如 `LongLong + ULong`
//     (LongLong 完整容纳 ULong), C 会把 uint32_t 提到 int64_t, 已经对了;
//   · 64 位异符号 —— 没有更宽的有符号型, promote 兜底给 ULongLong, 而 C 给的也是
//     uint64_t, 两侧一致。
// 后两条同时保证了**存量语料零影响**: 既有类型里唯一的异符号组合是 Byte vs
// Integer/Long/LongLong (都是有符号侧更宽 = 包含关系), 一个都不落进这一档。
static const char* mixedSignWidenCType(Vb6Type lt, Vb6Type rt) {
    int lb = 0, rb = 0;
    bool ls = true, rs = true;
    if (!TypeSystem::intShape(lt, &lb, &ls)) return nullptr;
    if (!TypeSystem::intShape(rt, &rb, &rs)) return nullptr;
    if (ls == rs) return nullptr;                       // 同符号: C 已经对
    int pb = 0;
    bool ps = true;
    const Vb6Type p = TypeSystem::promote(lt, rt);
    if (!TypeSystem::intShape(p, &pb, &ps)) return nullptr;   // 浮点/Decimal: 不走这里
    if (pb <= (lb > rb ? lb : rb)) return nullptr;      // 不宽于两者 ⇒ 包含关系
    switch (pb) {
        case 16: return "int16_t";                      //  8 位异符号
        case 32: return "int32_t";                      // 16 位异符号
        case 64: return "int64_t";                      // 32 位异符号
        default: return nullptr;
    }
}

// M22: 将非BSTR表达式包装为BSTR (用于字符串连接 & 运算符)
void CCodeGen::visit(BinaryExpr& node) {
    // 账 #88: 非 & 的那一路以前一律按 "Long" 解封 COM 读 —— 字符串成员因此在**比较**里
    // 被发成 vb6_ComGetIntProp + vb6_CStrLong (拿 BSTR 指针当数字去和字面量比) ⇒ 恒 False。
    // 实测 `tv1.Nodes(2).Text = "子乙"` 读回 N，而把同一个值先存进局部 String 变量再比是 Y。
    // 口径只看**对侧**: 等值/大小比较且对侧是字符串 (字面量或推得出 String) ⇒ 这一侧按 BSTR
    // 解封; 算术 (+ - * / 等) 与其余情形维持 "Long" 不动 —— 那些形状今天都是好的。
    auto isCmp88 = [](BinaryOp op) {
        return op == BinaryOp::Eq || op == BinaryOp::Neq || op == BinaryOp::Lt
            || op == BinaryOp::Gt || op == BinaryOp::Le || op == BinaryOp::Ge;
    };
    auto hint88 = [&](const ExprPtr& sib) -> const char* {
        if (!sib || !isCmp88(node.op)) return "Long";
        if (sib->kind == ASTNodeKind::LiteralExpr
            && static_cast<const LiteralExpr&>(*sib).literalKind == LiteralKind::String) {
            return "BSTR";
        }
        if (inferExprType(*sib) == Vb6Type::String) return "BSTR";
        return "Long";
    };
    emitExpr(*node.left);
    // COM标记解析: 如果左操作数是COM属性, 解析为值
    // P24-02: 算术运算默认Long解包
    if (isComMarker_) {
        if (node.op == BinaryOp::Concat) resolveComValue();
        else resolveComValue(hint88(node.right));
    }
    std::string left = std::move(lastExpr_);
    emitExpr(*node.right);
    // COM标记解析: 如果右操作数是COM属性, 解析为值
    if (isComMarker_) {
        if (node.op == BinaryOp::Concat) resolveComValue();
        else resolveComValue(hint88(node.left));
    }
    std::string right = std::move(lastExpr_);

    // 字符串连接运算: VB6 & → vb6_BSTR_Concat / vb6_BSTR_ConcatFree
    // 嵌套Concat时用ConcatFree释放中间临时BSTR，避免内存泄漏
    // M22: 非BSTR操作数自动包装为BSTR (int→vb6_CStr(vb6_CLng(x)), double→vb6_CStr(vb6_CDbl(x)))
    if (node.op == BinaryOp::Concat) {
        left = wrapToBSTR(left, *node.left);
        right = wrapToBSTR(right, *node.right);
        bool leftIsConcat = (left.find("vb6_BSTR_Concat") != std::string::npos);
        if (leftIsConcat) {
            lastExpr_ = "vb6_BSTR_ConcatFree(" + left + ", " + right + ")";
        } else {
            lastExpr_ = "vb6_BSTR_Concat(" + left + ", " + right + ")";
        }
        return;
    }

    // Fix 113g: VB6 `+` 语义 — 一侧是 String 而另一侧是**数值**时, VB6 把字符串
    // 按数值解释后做**数值加法** (经典陷阱: "2000" + i 得 2005 而非 "2000i"),
    // 结果再在目标上下文按需转字符串.
    // 原实现只覆盖 String+String 与 String+<BSTR返回调用>(Fix 092i), 数值侧直接
    // 落到下面的算术 '+', 生成 `vb6_BSTR_FromStr(L"2000") + i` = **指针算术** →
    // 垃圾 BSTR → SysStringLen/GetTextExtentPoint32W 读越界崩溃.
    //   Charts 2020 Form2.frm:506
    //     ucPieChart1.AddItem "2000" + i, Random(10, 30), CLng(cPalette(i + 1))
    //   i=0 侥幸正确, i>=1 → 0xC0000005 @ _vb6_UserControl_TextHeight (标签测量).
    // 生成 vb6_CStrDbl(vb6_Val(<str>) + (double)(<num>)) — VB6 数值语义且直接得 BSTR.
    // (vb6_Val 对非数字串返回 0; VB6 本会报错 13, 此处宽松处理, 不视为致命.)
    // 注: 本形态若被用于**数值**上下文 (如 "1"+x > 5) 会得到 BSTR → 编译期 C2440,
    // 属"响亮的失败"而非静默错误; 该写法在真实 VB6 代码中几乎只出现在 String 上下文.
    if (node.op == BinaryOp::Add) {
        auto isBstrCExpr113g = [](const std::string& e) -> bool {
            return e.rfind("vb6_BSTR_", 0) == 0 || e.rfind("vb6_CStr", 0) == 0
                || e.rfind("vb6_Chr(", 0) == 0 || e.rfind("vb6_Trim", 0) == 0
                || (e.size() >= 2 && e[0] == 'L' && e[1] == '"');
        };
        auto isNumericType113g = [](Vb6Type t) -> bool {
            return t == Vb6Type::Long || t == Vb6Type::Integer
                || t == Vb6Type::Byte || t == Vb6Type::Boolean
                || t == Vb6Type::Double || t == Vb6Type::Single
                || t == Vb6Type::Currency || t == Vb6Type::LongPtr
                || t == Vb6Type::ULong
                // C3 扩展 (ai/032): 4 个新整型同样是"数值侧", 漏掉它们会让
                // `"2000" + u` (u 是 ULong) 掉进指针算术那条坑。
                || t == Vb6Type::SByte || t == Vb6Type::UInteger
                || t == Vb6Type::ULongLong || t == Vb6Type::LongLong;
        };
        bool leftStr113g = (inferExprType(*node.left) == Vb6Type::String
                            || isBstrCExpr113g(left));
        bool rightStr113g = (inferExprType(*node.right) == Vb6Type::String
                             || isBstrCExpr113g(right));
        bool leftNum113g = isNumericType113g(inferExprType(*node.left));
        bool rightNum113g = isNumericType113g(inferExprType(*node.right));
        bool mixStrNum113g =
            (!leftStr113g && rightStr113g && leftNum113g)
            || (leftStr113g && !rightStr113g && rightNum113g);
        if (mixStrNum113g) {
            std::string strSide = leftStr113g ? left : right;
            std::string numSide = leftStr113g ? right : left;
            lastExpr_ = "vb6_CStrDbl(vb6_Val(" + strSide + ") + (double)("
                      + numSide + "))";
            return;
        }
    }

    // P14.1.1: VB6 + 运算符 — 两端String时等同&拼接
    // VB6允许 "a" + "b" 作为字符串连接，语义与 & 相同
    // Fix 092i: 右操作数是返回 BSTR 的内建调用 (Chr$/Mid/Trim/...) 时,
    // inferExprType 无法把整个 Add 推断为 String → 退化为算术 + :
    //   cToolsHttp 286 (源码 GB_UrlDecode + Chr$(d), GB_UrlDecode As String)
    //   → 生成 (GB_UrlDecode + vb6_Chr(d)) 触发 C2110 指针相加 +
    //     vb6_BSTR_Assign 参数太少; 而同处 288 的 GB_UrlDecode + c (c As String)
    //     正常走 Concat. 左操作数为 String 且右侧是 BSTR 返回调用时按拼接处理.
    bool addIsConcat092i = (inferExprType(node) == Vb6Type::String);
    if (!addIsConcat092i && inferExprType(*node.left) == Vb6Type::String) {
        static const char* bstrReturnCalls092i[] = {
            "vb6_Chr(", "vb6_ChrW(", "vb6_Mid(", "vb6_Left(", "vb6_Right(",
            "vb6_Trim(", "vb6_LTrim(", "vb6_RTrim(", "vb6_UCase(", "vb6_LCase(",
            "vb6_Replace(", "vb6_String(", "vb6_Space(", "vb6_StrConv(",
            "vb6_Format(", "vb6_VariantToString(", "vb6_BSTR_"};
        for (auto* bcp092i : bstrReturnCalls092i) {
            if (right.compare(0, strlen(bcp092i), bcp092i) == 0) {
                addIsConcat092i = true;
                break;
            }
        }
    }
    if (node.op == BinaryOp::Add && addIsConcat092i) {
        left = wrapToBSTR(left, *node.left);
        right = wrapToBSTR(right, *node.right);
        bool leftIsConcat = (left.find("vb6_BSTR_Concat") != std::string::npos);
        if (leftIsConcat) {
            lastExpr_ = "vb6_BSTR_ConcatFree(" + left + ", " + right + ")";
        } else {
            lastExpr_ = "vb6_BSTR_Concat(" + left + ", " + right + ")";
        }
        return;
    }

    // 幂运算: VB6 ^ → pow()
    if (node.op == BinaryOp::Pow) {
        lastExpr_ = "vb6_Pow(" + left + ", " + right + ")";
        return;
    }

    // 整除: VB6 \ → vb6_IntDiv (确保整数截断)
    // Fix 084o: 操作数为 Variant 时需显式转 Long (vb6_IntDiv 形参是 int32_t),
    // 否则 cToolsHttp 等文件中 "nAsc \ 2^6" (nAsc As Variant) 产生 C2440.
    // ai/032 rev2: 发哪一条 helper 与"结果类型"由同一份判据 intDivResultType 给出
    // (inferExprType 也调它), 四档: int32 / int64 / uint32 / uint64。为什么不能只留
    // int32: vb6_IntDiv 的形参是 int32_t, `&H80000000 As ULong \ 2` 会先被按有符号
    // 解释 —— (-2147483648)/2 回存 uint32_t 得 3221225472 (实测)。
    if (node.op == BinaryOp::IntDiv) {
        Vb6Type ltD = inferExprType(*node.left);
        Vb6Type rtD = inferExprType(*node.right);
        switch (intDivResultType(ltD, rtD)) {
            case Vb6Type::ULongLong:
                lastExpr_ = "vb6_IntDivULongLong("
                          + asUnsignedOperand(left,  *node.left,  "uint64_t") + ", "
                          + asUnsignedOperand(right, *node.right, "uint64_t") + ")";
                return;
            case Vb6Type::LongLong:
                lastExpr_ = "vb6_IntDivLongLong("
                          + asUnsignedOperand(left,  *node.left,  "int64_t") + ", "
                          + asUnsignedOperand(right, *node.right, "int64_t") + ")";
                return;
            case Vb6Type::ULong:
                lastExpr_ = "vb6_IntDivULong("
                          + asUnsignedOperand(left,  *node.left,  "uint32_t") + ", "
                          + asUnsignedOperand(right, *node.right, "uint32_t") + ")";
                return;
            default:
                lastExpr_ = "vb6_IntDiv(" + toLongIfVariant(left, node.left.get())
                          + ", " + toLongIfVariant(right, node.right.get()) + ")";
                return;
        }
    }

    // 账 #305 §B145: 「这一侧的 C 表达式是不是一枚 Variant」—— Fix 086 的数值提取与下面
    // 四则运算的改道**共用这一条判据**（三条形状: 前缀表 / As Variant 名单 / (*name) 形参），
    // 以前这个形状内联在 extractVar86 里，现在不再抄第二遍。
    auto isVariantOperand86 = [&](const std::string& cExpr, const Expr* astExpr) -> bool {
        if (cExprIsVariant(cExpr)) return true;
        if (astExpr && astExpr->kind == ASTNodeKind::IdentifierExpr) {
            auto& id86 = static_cast<IdentifierExpr&>(const_cast<Expr&>(*astExpr));
            if (knownVariantVars_.count(Symbol::toLower(id86.name))) return true;
        }
        if (cExpr.size() > 4 && cExpr[0] == '(' && cExpr[1] == '*'
            && cExpr.back() == ')' && cExpr.find('(') == std::string::npos) {
            // (*name) 解引用形式 — ByRef Variant 参数
            std::string inner86 = Symbol::toLower(cExpr.substr(2, cExpr.size() - 3));
            if (knownVariantVars_.count(inner86)) return true;
            if (currentProc_) {
                for (const auto& p : currentProc_->params) {
                    if (Symbol::toLower(p.name) == inner86
                        && (p.type == Vb6Type::Variant || p.type == Vb6Type::Empty)) {
                        return true;
                    }
                }
            }
        }
        return false;
    };
    // 装箱: 已经是 Variant 的直接交出去，其余按 vb6_VariantFromValue 的 _Generic 造一枚槽
    // （它自带 vb6_VARIANT 那档的恒等臂，所以两头都不用知道对方是什么）。
    auto boxVariant86 = [&](const std::string& cExpr, const Expr* astExpr) -> std::string {
        return isVariantOperand86(cExpr, astExpr)
                   ? cExpr : "vb6_VariantFromValue(" + cExpr + ")";
    };
    // 账 #305 §B147: **改道那一问**吃「AST ∪ 文本」，而上面的 isVariantOperand86 刻意维持
    // 「文本 ∪ 已知变量」不变 —— 它还供着 Fix 086 那条提取 (vb6_VariantToLong(...))，那一档
    // 只能对**真的是槽**的 C 文本出手，把 AST 的宽答案灌进去就是 vb6_VariantToLong(裸 double)
    // = C2440。方向性是这一格的要点：Div 的类型 oracle 只吃 variantOperandAST，这里吃它的
    // 超集 ⇒「oracle 答 Variant」必然蕴含「改道接走了、交出的是整只槽」，§B147 那一族
    // (oracle 说是槽、发码给裸标量 ⇒ 消费点既不提取也不取整) 就此按构造消失。
    auto routeOperand86 = [&](const std::string& cExpr, const Expr* astExpr) -> bool {
        return isVariantOperand86(cExpr, astExpr) || variantOperandAST(astExpr);
    };

    // 账 #305 §B145: Variant 参与 + - * / 时**不再替运行期猜档**。
    // 改前这四条运算符把 Variant 操作数发成 vb6_VariantToLong(...)（加减乘）或
    // vb6_VariantToDouble(...)（除），实测两架构一致: 盒里是 Double 时 1.5+1 答 3、1.5*2 答 4；
    // 是 Decimal 时 1.25+1 答 2、0.1+0.2 答 0 —— 小数与 Decimal 的精确值整个丢掉
    // （探针 .build/b1095_arith/probe5.bas，读数在台账 §B145）。
    // 现在两侧装箱后交 RTL 那一族运行期出口，数值档由运行期的 vt 决定（= oleaut32 的
    // VarAdd 一族，VB6 本人的 Variant 算术引擎），于是 Decimal 的 96 位精确、Double 的小数、
    // Null 传播、Date 加天数一起继承过来。
    // 刻意**不接管**任一侧静态类型是 String 的那一族: VB6 的 + 只要有一侧是串就走拼接，
    // 而 VarAdd 是按数值强制（实测 "5"+1 交 6）—— 那一条路留在上面 113g 那张表上。
    if (node.op == BinaryOp::Add || node.op == BinaryOp::Sub
        || node.op == BinaryOp::Mul || node.op == BinaryOp::Div) {
        const bool strL86 = (inferExprType(*node.left) == Vb6Type::String);
        const bool strR86 = (inferExprType(*node.right) == Vb6Type::String);
        const bool anyVar86 = routeOperand86(left, node.left.get())
                           || routeOperand86(right, node.right.get());
        if (anyVar86 && !strL86 && !strR86) {
            const char* fn86 = node.op == BinaryOp::Add  ? "vb6_VarAddV"
                             : node.op == BinaryOp::Sub  ? "vb6_VarSubV"
                             : node.op == BinaryOp::Mul  ? "vb6_VarMulV"
                             :                             "vb6_VarDivV";
            lastExpr_ = std::string("(") + fn86 + "("
                      + boxVariant86(left, node.left.get()) + ", "
                      + boxVariant86(right, node.right.get()) + "))";
            return;
        }
    }

    // 浮点除法: VB6 / → (double)left / (double)right
    // Fix 156: 操作数为 Variant 时 C 强转非法 (C2440 vb6_VARIANT→double),
    // 改走 vb6_VariantToDouble 提取 (与上方 IntDiv 的 Fix 084o 同构).
    if (node.op == BinaryOp::Div) {
        // Task #44 → 变量除数也补上: VB6 语义是运行期错误 11 而非 IEEE inf。
        // 此前只有字面量 0 除数走 vb6_Num_Div, 变量除数仍落裸 C 除法得 inf
        // (vb6_Num_Div 的注释里记为"运行期语义缺口")。既然统一走 helper,
        // 常量折叠 C2124 也顺带消失, 无条件走即可。
        lastExpr_ = "vb6_Num_Div((double)(" + toDoubleIfVariant(left, node.left.get())
                  + "), (double)(" + toDoubleIfVariant(right, node.right.get()) + "))";
        return;
    }

    // Fix 039: VB6 Eqv → ~(a^b), cast to int32_t for non-integer operands
    // (double from vb6_Pow, pointer from BSTR/void*/SafeArray*)
    // Fix 039b: For Variant operands, use vb6_VariantToLong() instead of (int32_t)() cast.
    if (node.op == BinaryOp::Eqv) {
        // ai/032: 无符号档按位宽收口 (同下面 And/Or/Xor 那条; Eqv/Imp 是 VB6 遗留
        // 位运算, 但 ULongLong 被 (int32_t) 拦腰截断同样是错的)。
        const Vb6Type ltE = inferExprType(*node.left);
        const Vb6Type rtE = inferExprType(*node.right);
        const bool wideUE = (ltE == Vb6Type::ULongLong || rtE == Vb6Type::ULongLong);
        const bool wideUE32 = !wideUE && (ltE == Vb6Type::ULong || rtE == Vb6Type::ULong);
        auto castBitwise = [&](const std::string& cExpr, const Expr* astExpr) -> std::string {
            if (wideUE || wideUE32) {
                const char* ctE = wideUE ? "uint64_t" : "uint32_t";
                return "((" + std::string(ctE) + ")(" + cExpr + "))";
            }
            if (cExprIsVariant(cExpr)) return "vb6_VariantToLong(" + cExpr + ")";
            if (astExpr && astExpr->kind == ASTNodeKind::IdentifierExpr) {
                auto& ident = static_cast<IdentifierExpr&>(const_cast<Expr&>(*astExpr));
                std::string lower = ident.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                if (knownVariantVars_.count(lower)) return "vb6_VariantToLong(" + cExpr + ")";
            }
            return "(int32_t)(" + cExpr + ")";
        };
        lastExpr_ = "(~(" + castBitwise(left, node.left.get()) + " ^ " + castBitwise(right, node.right.get()) + "))";
        return;
    }

    // Fix 039: VB6 Imp → (~a | b), cast to int32_t for non-integer operands
    if (node.op == BinaryOp::Imp) {
        // ai/032: 无符号档按位宽收口 (同上面 Eqv / 下面 And/Or/Xor)
        const Vb6Type ltI = inferExprType(*node.left);
        const Vb6Type rtI = inferExprType(*node.right);
        const bool wideUI = (ltI == Vb6Type::ULongLong || rtI == Vb6Type::ULongLong);
        const bool wideUI32 = !wideUI && (ltI == Vb6Type::ULong || rtI == Vb6Type::ULong);
        auto castBitwise = [&](const std::string& cExpr, const Expr* astExpr) -> std::string {
            if (wideUI || wideUI32) {
                const char* ctI = wideUI ? "uint64_t" : "uint32_t";
                return "((" + std::string(ctI) + ")(" + cExpr + "))";
            }
            if (cExprIsVariant(cExpr)) return "vb6_VariantToLong(" + cExpr + ")";
            if (astExpr && astExpr->kind == ASTNodeKind::IdentifierExpr) {
                auto& ident = static_cast<IdentifierExpr&>(const_cast<Expr&>(*astExpr));
                std::string lower = ident.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                if (knownVariantVars_.count(lower)) return "vb6_VariantToLong(" + cExpr + ")";
            }
            return "(int32_t)(" + cExpr + ")";
        };
        lastExpr_ = "((~" + castBitwise(left, node.left.get()) + ") | " + castBitwise(right, node.right.get()) + ")";
        return;
    }

    // Like运算: VB6 Like → vb6_Like
    if (node.op == BinaryOp::Like) {
        // <vbeclipse>: Option Compare Text 模块里 Like 的大小写与字符区间都按文本形
        lastExpr_ = std::string("vb6_Like") + (optionCompareText_ ? "C" : "")
                    + "(" + left + ", " + right + (optionCompareText_ ? ", 1" : "") + ")";
        return;
    }

    // P24-07: 字符串比较运算 — BSTR不能用==做指针比较, 需用vb6_StrCmp
    if (node.op == BinaryOp::Eq || node.op == BinaryOp::Neq ||
        node.op == BinaryOp::Lt || node.op == BinaryOp::Gt ||
        node.op == BinaryOp::Le || node.op == BinaryOp::Ge) {
        Vb6Type lt = inferExprType(*node.left);
        Vb6Type rt = inferExprType(*node.right);
        if (lt == Vb6Type::String || rt == Vb6Type::String) {
            // 两侧都需要是BSTR: 非BSTR端用wrapToBSTR转换
            if (lt != Vb6Type::String) left = wrapToBSTR(left, *node.left);
            if (rt != Vb6Type::String) right = wrapToBSTR(right, *node.right);
            std::string cmpOp;
            switch (node.op) {
                case BinaryOp::Eq:       cmpOp = "== 0"; break;
                case BinaryOp::Neq:    cmpOp = "!= 0"; break;
                case BinaryOp::Lt:        cmpOp = "< 0";  break;
                case BinaryOp::Gt:     cmpOp = "> 0";  break;
                case BinaryOp::Le:   cmpOp = "<= 0"; break;
                case BinaryOp::Ge:cmpOp = ">= 0"; break;
                default: cmpOp = "== 0"; break;
            }
            // <vbeclipse>: 本模块 Option Compare Text ⇒ 走恒文本入口 vb6_StrCmpT
            // (VB6 的比较模式是按模块的编译期属性, 不能靠进程唯一的全局)
            lastExpr_ = "(-(vb6_StrCmp" + std::string(optionCompareText_ ? "T" : "")
                        + "(" + left + ", " + right + ") " + cmpOp + "))";
            return;
        }
    }

    // Fix 158n: BinaryOp::Is 的 Variant 操作数 — `Is` 映射为 C `==`, 直接比较
    // vb6_VARIANT 结构体 → C2088. 三种形态均需改写:
    //   1. VBFlexGrid.ctl FindTag `If Buffer Is Value` (两操作数均 As Variant,
    //      VT=vbObject 分支) → vb6_VarCmpEq(&Buffer, &Value)
    //   2. VBFlexGrid.ctl `If Value Is Nothing` / `If Not Value Is Nothing`
    //      (Value As Variant, Nothing→NULL) → vb6_IsNothing(vb6_VariantToObject(&Value))
    //   3. 变体 vs 其它标量 → vb6_VarCmpLongEq(取值契形参)
    // 普通对象/指针 Is (void* <=> NULL) 不满足 varLike, 保持原样, 不受影响.
    if (node.op == BinaryOp::Is) {
        auto varLike158n = [&](const std::string& c, Expr* ast) -> bool {
            // Fix <vbeclipse> rev6: `Property Get Foo() As <工程类>` 的返回变量
            // (**vb6_cls_X\* 类型**) 不是 Variant, 尽管 cExprIsVariant /
            // isDefinitelyVariantExpr 都会说它是 —— 因为 resolveTypeOrDefault 把
            // `As <类名>` 折成了 Variant (类名只留在 typeRefName, 见 MEMORY)。
            // 后果: `If ActivePerspective Is Nothing` 生成
            //   vb6_IsNothing(vb6_VariantToObject(&vb6_ret_ActivePerspective))
            // 把裸类指针当 vb6_VARIANT* 读 → vt 字段是指针低 2 字节 (非 0) →
            // 返回 NULL → IsNothing 判**真** → 抛 "No active perspective!"
            // (实测 play78.exe: ucPerspective.ctl 的 Property Get ActivePerspective)。
            // currentReturnCType_ 是权威 (发定义那侧用它)。
            std::string ct158n = c;
            while (ct158n.size() >= 2 && ct158n.front() == '(' && ct158n.back() == ')')
                ct158n = ct158n.substr(1, ct158n.size() - 2);
            if (!currentReturnVar_.empty() && ct158n == currentReturnVar_
                && currentReturnCType_.compare(0, 8, "vb6_cls_") == 0) {
                return false;
            }
            if (cExprIsVariant(c)) return true;
            if (!ast) return false;
            if (ast->kind == ASTNodeKind::IdentifierExpr) {
                // 局部/形参标识符: knownVariantVars_ 只在 `As Variant` 声明时登记
                // (cgen_localdecl 172 / cgen_decl_func 112), 类/接口成员不会误入 →
                // 直接信 isDefinitelyVariantExpr, 不套类门 (否则 As Variant 形参
                // 若另有类符号登记, 会误回对象路径 → C2088).
                return isDefinitelyVariantExpr(*ast);
            }
            if (isDefinitelyVariantExpr(*ast)) {
                // Fix 158p: 类/COM 接口成员 (VBFlexGrid.ctl `Private VBFlexGridFlexDataSource
                // As IVBFlexDataSource`) 被符号表误注册为 Variant 时, inferClassTypeOfExpr
                // 非空 → 实为对象指针, `Is Nothing` 应走 (X == NULL). 否则 158n 会把
                // 对象取址塞进 vb6_VARIANT 临时变量 → C2440.
                if (!inferClassTypeOfExpr(*ast).empty()) return false;
                return true;
            }
            return false;
        };
        auto isNothingSentinel158n = [](const std::string& s) -> bool {
            std::string t = s;
            while (t.size() >= 2 && t.front() == '(' && t.back() == ')')
                t = t.substr(1, t.size() - 2);
            return t == "NULL" || t == "0";
        };
        auto simpIdent158n = [](const std::string& s) -> bool {
            if (s.empty()) return false;
            if (!(std::isalpha(static_cast<unsigned char>(s[0])) || s[0] == '_')) return false;
            for (char c : s) {
                if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_' && c != '.') return false;
            }
            return true;
        };
        // Fix 196: 裸 `vb6_VARIANT tmp = X;` 对标量 (vb6_ComGetIntProp→int32_t,
        // cHeartbeat `If Not oClient.Heartbeat Is Nothing`) 与类指针 (cSSE
        // `(*Client)`) 触发 C2440. 统一 vb6_VariantFromValue 包装: _Generic 下
        // 已是 VARIANT 的表达式恒等直传 (Fix 158u 场景不受影响), 标量/指针自动
        // 包装; 裸 vb6_ComCall( 结果是 VARIANT*, 须走 VariantFromComResult
        // 解引用 (与 Fix 132 同规则).
        auto wrapVariantRvalue196 = [&](const std::string& e) -> std::string {
            if ((e.find("vb6_ComCall(") != std::string::npos
                 || e.find("vb6_ComCallByDispid(") != std::string::npos)
                && e.find("vb6_VariantFromComResult(") == std::string::npos)
                return "vb6_VariantFromComResult(" + e + ")";
            if (e.find("vb6_VariantFromComResult(") != std::string::npos
                || e.find("vb6_VariantFromValue(") != std::string::npos)
                return e;
            return "vb6_VariantFromValue(" + e + ")";
        };
        // Variant 表达式的取址: 左值标识符/成员/VB6_SA_AT(...) 直接 &,
        // 其余 rvalue (vb6_VariantFromComResult 等) 用临时变量存上再取址.
        // 账 #238: 裸名字那一形还要问"它是不是 vb6_VARIANT 那份存储" (判据在一处)。
        auto variantAddr158n = [&](const std::string& s, Expr* ast) -> std::string {
            std::string t = s;
            while (t.size() >= 2 && t.front() == '(' && t.back() == ')')
                t = t.substr(1, t.size() - 2);
            // Fix <vbeclipse>: 枚举常量 (vb6_enum_eMsgWhen_MSG_BEFORE) 被当作
            // Variant 比较左值取址 → C2101 "常量上的&" (MagneticWnd.ctl
            // `eMsgWhen.MSG_BEFORE = When`, 枚举成员无 constIntValue 时以
            // vb6_enum_<Mod>_<Member> 裸名发出)。isConstIdent 只认 #define 宏,
            // 认不出 vb6_enum_*. 枚举成员不可取址, 也不能当 vb6_VARIANT* 传
            // (vb6_VarCmpLongEq(vb6_VARIANT*, int32_t)), 构造可寻址的
            // vb6_VARIANT 临时变量 (同下方 rvalue 通道的 _vcmp_N 模式).
            if (t.rfind("vb6_enum_", 0) == 0) {
                std::string tmp = "_vcmp_" + std::to_string(vcmpCounter_++);
                c_.emitLine("vb6_VARIANT " + tmp + " = vb6_VariantLong(" + s + ");");
                return "&" + tmp;
            }
            if (simpIdent158n(t) && !isConstIdent(t) && cmpOperandMayTakeAddr(t, ast)) return "&" + s;
            if (t.rfind("VB6_SA_AT(", 0) == 0) return "&" + s;
            // Fix 158u: Variant 比较左值语义落在 COM 对象指针成员 (me->VBFlexGrid
            // FlexDataSource 等 union 成员, 声类型 ComIface*/void*) 时, 裸 `vb6_VARIANT
            // tmp = me->X;` 触发 C2440 (无法从 ComIface* 转换到 vb6_VARIANT).
            // 包 vb6_VariantFromValue → 对指针走 vb6_VariantObject(VT_DISPATCH),
            // 对已是 Variant 的 rvalue (vb6_VariantFromComResult 等, 无 "->") 保持直拷.
            if (t.find("->") != std::string::npos) {
                std::string tmp = "_vcmp_" + std::to_string(vcmpCounter_++);
                c_.emitLine("vb6_VARIANT " + tmp + " = vb6_VariantFromValue(" + s + ");");
                return "&" + tmp;
            }
            std::string tmp = "_vcmp_" + std::to_string(vcmpCounter_++);
            c_.emitLine("vb6_VARIANT " + tmp + " = " + wrapVariantRvalue196(s) + ";");
            return "&" + tmp;
        };
        bool lv158n = varLike158n(left, node.left.get());
        bool rv158n = varLike158n(right, node.right.get());
        if (lv158n || rv158n) {
            if (lv158n && rv158n) {
                // Fix: 两侧都是 Variant 时也必须按**对象同一性**比 —— VB6 的 `Is`
                // 问的是"两个对象引用是否同一只"。旧发码走 vb6_VarCmpEq(值比较),
                // 而它对非同串档会退化成
                //   vb6_VarToDouble_internal(a) == vb6_VarToDouble_internal(b)
                // 空串 → 0.0、对象 → 0.0 ⇒ **恒等**。
                // 实测代价: ComCtlsDemo/Builds/CoolBar/CbrBand.cls
                //   Public Property Set Tag(ByVal Value As Variant)
                //     If Value Is ShadowCoolBar Then Err.Raise 35700 ...
                // 右侧工程类调用被折成 Variant ⇒ 走的就是这一支 ⇒ 空串与 CoolBar
                // "相等" ⇒ Err.Raise 35700 ⇒ 设计期 band 重放整只进程退出。
                // 取 `两边都是对象 且 对象指针相等` (VB6 里 `Is` 的非对象操作数本是
                // 编译期错误; 这里退化为 false, 安全侧)。
                std::string la = variantAddr158n(left, node.left.get());
                std::string ra = variantAddr158n(right, node.right.get());
                lastExpr_ = "(vb6_IsObject(*(" + la + ")) && vb6_IsObject(*(" + ra
                          + ")) && (void*)vb6_VariantToObject(" + la
                          + ") == (void*)vb6_VariantToObject(" + ra + "))";
            } else if (lv158n && isNothingSentinel158n(right)) {
                lastExpr_ = "(vb6_IsNothing(vb6_VariantToObject(" + variantAddr158n(left, node.left.get()) + ")))";
            } else if (rv158n && isNothingSentinel158n(left)) {
                lastExpr_ = "(vb6_IsNothing(vb6_VariantToObject(" + variantAddr158n(right, node.right.get()) + ")))";
            } else if (lv158n) {
                // Fix: `<Variant> Is <对象/指针>` —— VB6 的 `Is` 问的是**对象同一性**
                // (指针相等), 不是值相等。旧发码
                //   vb6_VarCmpLongEq(&v, (int32_t)(other))
                // 有两处硬伤: (a) 把对象指针截成 int32 再按值比较 (x64/gc 下必错);
                // (b) vb6_VarCmp 会把对象转成它的默认值 — 于是**空串 Variant 与
                // 对象指针"相等"**。实测代价: ComCtlsDemo/Builds/CoolBar/CbrBand.cls:96
                //   Public Property Set Tag(ByVal Value As Variant)
                //     If Value Is ShadowCoolBar Then Err.Raise 35700 ...
                // 设计期 band 重放走 `.Tag = InitBands(i).Tag`(空串) 时误判"是
                // ShadowCoolBar" ⇒ Err.Raise 35700 ⇒ 未处理错误, 整只进程退出。
                // 现按 `IsObject(v) && 两边对象指针相等` 发 (VB6 里 `Is` 非对象操作数
                // 是编译期错误, 这里退化为 false —— 安全侧)。
                std::string la = variantAddr158n(left, node.left.get());
                lastExpr_ = "(vb6_IsObject(*(" + la + ")) && (void*)vb6_VariantToObject(" + la
                          + ") == (void*)(" + right + "))";
            } else {
                std::string ra = variantAddr158n(right, node.right.get());
                lastExpr_ = "(vb6_IsObject(*(" + ra + ")) && (void*)vb6_VariantToObject(" + ra
                          + ") == (void*)(" + left + "))";
            }
            return;
        }
    }

    // P24-Bug2: Variant比较运算 — vb6_VARIANT不能用C内置比较运算符
    if (node.op == BinaryOp::Eq || node.op == BinaryOp::Neq ||
        node.op == BinaryOp::Lt || node.op == BinaryOp::Gt ||
        node.op == BinaryOp::Le || node.op == BinaryOp::Ge) {
        Vb6Type lt = inferExprType(*node.left);
        Vb6Type rt = inferExprType(*node.right);
        // P25: inferExprType对后期绑定COM属性返回Variant, 但实际代码生成的是类型化getter
        // 修正类型以避免对rvalue取地址或选择错误的比较函数
        if (lt == Vb6Type::Variant) {
            if (left.find("vb6_ComGetIntProp") == 0 || left.find("vb6_ComVtableGetInt") == 0) lt = Vb6Type::Long;
            else if (left.find("vb6_ComGetDoubleProp") == 0 || left.find("vb6_ComVtableGetDouble") == 0) lt = Vb6Type::Double;
            else if (left.find("vb6_ComGetStringProp") == 0) lt = Vb6Type::String;
            else if (left.find("vb6_ComGetObjectProp") == 0) lt = Vb6Type::Object;
        }
        if (rt == Vb6Type::Variant) {
            if (right.find("vb6_ComGetIntProp") == 0 || right.find("vb6_ComVtableGetInt") == 0) rt = Vb6Type::Long;
            else if (right.find("vb6_ComGetDoubleProp") == 0 || right.find("vb6_ComVtableGetDouble") == 0) rt = Vb6Type::Double;
            else if (right.find("vb6_ComGetStringProp") == 0) rt = Vb6Type::String;
            else if (right.find("vb6_ComGetObjectProp") == 0) rt = Vb6Type::Object;
        }
        // Fix 110aa: 与上面相反方向的修正 — inferExprType 给出具体类型, 但**实际发射
        // 的 C 表达式**是 Variant 值 (vb6_VariantEmpty() / vb6_VariantFromComResult(...)
        // / vb6_CallByName(...) 等). 此时原生 C 比较作用于结构体 → C2088
        // ("==" 对于 struct 非法). 按 C 表达式判定, 把该侧提升为 Variant 走
        // vb6_VarCmp* 通道. Charts 2020 ucProgressCircular.ctl:949
        //   If hBrush = 0 Or Count = 0            (Count = 隐式 Variant 局部)
        //   → (vb6_VariantEmpty() == 0) C2088.
        // 类型化 getter (vb6_ComGetIntProp 等) 不在 cExprIsVariant 前缀表内, 不受影响.
        // Fix 158m: inferExprType 对某些 Variant 表达式误报具体类型 (Variant 形参
        // VBFlexGrid.ctl `If Value Is Nothing`、Variant() 数组元素 VTableHandle.bas
        // `If VTableIPAO(0) = 0`), 直接 C 比较 vb6_VARIANT 结构体 → C2088. 用
        // isDefinitelyVariantExpr (AST 语义表: knownVariantVars_/符号表 Variant 返回)
        // 补充升级; 对上面的类型化 getter 返回 false, 不会回滚 P25 的降级修正.
        auto isVariantOperand158m = [&](const std::string& c, Expr* ast) -> bool {
            if (cExprIsVariant(c)) return true;
            return ast && isDefinitelyVariantExpr(*ast);
        };
        if (lt != Vb6Type::Variant && isVariantOperand158m(left, node.left.get())) lt = Vb6Type::Variant;
        if (rt != Vb6Type::Variant && isVariantOperand158m(right, node.right.get())) rt = Vb6Type::Variant;
        // Bug #2 fix: also check knownLongVars_/knownLongPtrVars_ for simple variable names
        // because inferExprType may return Variant for optional params or out-of-scope variables
        auto isSimpleIdent = [](const std::string& s) -> bool {
            if (s.empty()) return false;
            if (!(std::isalpha(static_cast<unsigned char>(s[0])) || s[0] == '_')) return false;
            for (char c : s) { if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_') return false; }
            return true;
        };
        if (lt == Vb6Type::Variant && isSimpleIdent(left)) {
            std::string lower = left;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (knownLongPtrVars_.count(lower)) lt = Vb6Type::LongPtr;
            else if (knownLongVars_.count(lower)) lt = Vb6Type::Long;
        }
        if (rt == Vb6Type::Variant && isSimpleIdent(right)) {
            std::string lower = right;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (knownLongPtrVars_.count(lower)) rt = Vb6Type::LongPtr;
            else if (knownLongVars_.count(lower)) rt = Vb6Type::Long;
        }
        // Bug #2 fix: LongPtr (intptr_t) comparisons should use direct C operators
        // instead of VarCmpLong which treats the operand as vb6_VARIANT*
        // Fix 158o: 若某侧经 158m 提升为 Variant (C 表达式确实是 vb6_VARIANT 结构体,
        // 如 LongPtr 数组误发成 VB6_SA_AT(vb6_VARIANT,...)), 另一侧 LongPtr 仍不能裸
        // `==` (C2088 结构体比较). 双侧 LongPtr 才可直接比较; 混合态交给下方 Variant 块,
        // 其中 LongPtr 侧按标量经 scalarArg 传给 VarCmpLong (VTableHandle.bas
        // `If VTableIPAO(0) = NULL_PTR`, NULL_PTR=0).
        if ((lt == Vb6Type::LongPtr || rt == Vb6Type::LongPtr) && lt != Vb6Type::Variant && rt != Vb6Type::Variant) {
            std::string op;
            switch (node.op) {
                case BinaryOp::Eq:  op = "=="; break;
                case BinaryOp::Neq: op = "!="; break;
                case BinaryOp::Lt:  op = "<";  break;
                case BinaryOp::Gt:  op = ">";  break;
                case BinaryOp::Le:  op = "<="; break;
                case BinaryOp::Ge:  op = ">="; break;
                default: op = "=="; break;
            }
            // Fix 092v: LongPtr 关系比较同样产生 C 的 0/1 → 取负为 -1/0 (VB6 Boolean)
            lastExpr_ = "(-(" + left + " " + op + " " + right + "))";
            return;
        }
        if (lt == Vb6Type::Variant || rt == Vb6Type::Variant) {
            // 确定比较函数后缀
            std::string cmpFn;
            switch (node.op) {
                case BinaryOp::Eq:  cmpFn = "Eq";  break;
                case BinaryOp::Neq: cmpFn = "Ne";  break;
                case BinaryOp::Lt:  cmpFn = "Lt";  break;
                case BinaryOp::Gt:  cmpFn = "Gt";  break;
                case BinaryOp::Le:  cmpFn = "Le";  break;
                case BinaryOp::Ge:  cmpFn = "Ge";  break;
                default: cmpFn = "Eq"; break;
            }
            // Variant vs NonVariant: 使用VarCmpLong快捷函数
            // Fix 158m-2: 标量侧可能是 VB Nothing → NULL (void*) (VBFlexGrid.ctl
            // `If Value Is Nothing`), 直接传 VarCmpLongEq(vb6_VARIANT*, int32_t)
            // → C2440 (不能 void* → int). 裸 NULL 包 (int32_t)(...) 再传.
            auto scalarArg158m = [](const std::string& s) -> std::string {
                std::string t = s;
                while (t.size() >= 2 && t.front() == '(' && t.back() == ')')
                    t = t.substr(1, t.size() - 2);
                if (t.find("NULL") != std::string::npos) return "(int32_t)(" + s + ")";
                return s;
            };
            if (lt == Vb6Type::Variant && rt != Vb6Type::Variant) {
                Vb6Type rActual = rt;
                // Fix 158o: LongPtr 侧也按标量处理 (Variant vs LongPtr 常量).
                if (rActual == Vb6Type::Long || rActual == Vb6Type::Integer || rActual == Vb6Type::Boolean || rActual == Vb6Type::LongPtr) {
                    // P25: left可能是VARIANT rvalue(vb6_VariantFromComResult), 需要临时变量
                    // Fix 084aa: 常量宏 (#define) 不可取址 → 视为非左值走临时变量
                    // Fix <vbeclipse>: 枚举常量 (vb6_enum_*) 也是裸标识符但不可取址 —
                    // isConstIdent 只认 #define 宏, 认不出 vb6_enum_MsgWhen_MSG_BEFORE →
                    // `&vb6_enum_...` C2101 (MagneticWnd.ctl `eMsgWhen.MSG_BEFORE = When`)
                    bool leftIsLvalue = !left.empty() && (std::isalpha(static_cast<unsigned char>(left[0])) || left[0] == '_') && !isConstIdent(left) && left.rfind("vb6_enum_", 0) != 0;
                    if (leftIsLvalue) { for (char c : left) { if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_') { leftIsLvalue = false; break; } } }
                    if (leftIsLvalue && cmpOperandMayTakeAddr(left, node.left.get())) {
                        lastExpr_ = "(vb6_VarCmpLong" + cmpFn + "(&" + left + ", " + scalarArg158m(right) + "))";
                    } else {
                        std::string tmp = "_vcmp_" + std::to_string(vcmpCounter_++);
                        // Fix 024: left 被 inferExprType 误判为 Variant, 但实际标量 (LenB/Asc/int 等).
                        // 用 vb6_VariantFromValue 在编译期按实类型选择 variant 构造函数, 消除 C2440.
                        // Fix 132: 裸 COM 结果必须走 VariantFromComResult (见下方说明)。
                        std::string wrapL = ((left.find("vb6_ComCall(") != std::string::npos
                                              || left.find("vb6_ComCallByDispid(") != std::string::npos)
                                             && left.find("vb6_VariantFromComResult(") == std::string::npos)
                                            ? ("vb6_VariantFromComResult(" + left + ")") : ("vb6_VariantFromValue(" + left + ")");
                        c_.emitLine("vb6_VARIANT " + tmp + " = " + wrapL + ";");
                        lastExpr_ = "(vb6_VarCmpLong" + cmpFn + "(&" + tmp + ", " + scalarArg158m(right) + "))";
                    }
                    return;
                }
            }
            if (rt == Vb6Type::Variant && lt != Vb6Type::Variant) {
                Vb6Type lActual = lt;
                if (lActual == Vb6Type::Long || lActual == Vb6Type::Integer || lActual == Vb6Type::Boolean || lActual == Vb6Type::LongPtr) {
                    // 反转比较方向: Long op Variant → Variant reverseOp Long
                    std::string revCmpFn;
                    switch (node.op) {
                        case BinaryOp::Lt: revCmpFn = "Gt"; break;
                        case BinaryOp::Gt: revCmpFn = "Lt"; break;
                        case BinaryOp::Le: revCmpFn = "Ge"; break;
                        case BinaryOp::Ge: revCmpFn = "Le"; break;
                        default: revCmpFn = cmpFn; break;  // Eq/Ne是对称的
                    }
                    // Fix 084aa: 常量宏不可取址 → 视为非左值
                    // Fix <vbeclipse>: 枚举常量 (vb6_enum_*) 同样不可取址 (对称于 493 行)
                    bool rightIsLvalue = !right.empty() && (std::isalpha(static_cast<unsigned char>(right[0])) || right[0] == '_') && !isConstIdent(right) && right.rfind("vb6_enum_", 0) != 0;
                    if (rightIsLvalue) { for (char c : right) { if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_') { rightIsLvalue = false; break; } } }
                    if (rightIsLvalue && cmpOperandMayTakeAddr(right, node.right.get())) {
                        lastExpr_ = "(vb6_VarCmpLong" + revCmpFn + "(&" + right + ", " + scalarArg158m(left) + "))";
                    } else {
                        std::string tmp = "_vcmp_" + std::to_string(vcmpCounter_++);
                        // Fix 024: right 被 inferExprType 误判为 Variant, 但实际标量. 用 FromValue 包装.
                        // Fix 132: 裸 COM 结果必须走 VariantFromComResult。
                        std::string wrapR = ((right.find("vb6_ComCall(") != std::string::npos
                                              || right.find("vb6_ComCallByDispid(") != std::string::npos)
                                             && right.find("vb6_VariantFromComResult(") == std::string::npos)
                                            ? ("vb6_VariantFromComResult(" + right + ")") : ("vb6_VariantFromValue(" + right + ")");
                        c_.emitLine("vb6_VARIANT " + tmp + " = " + wrapR + ";");
                        lastExpr_ = "(vb6_VarCmpLong" + revCmpFn + "(&" + tmp + ", " + scalarArg158m(left) + "))";
                    }
                    return;
                }
            }
            // Variant vs Variant: 使用VarCmp函数
            // P25: 检测rvalue, 非左值需要存临时变量
            // Fix 084aa: 常量宏 (#define) 不可取址 → 视为非左值走临时变量
            auto isLvalue = [this](const std::string& s) -> bool {
                if (s.empty()) return false;
                if (isConstIdent(s)) return false;
                if (!(std::isalpha(static_cast<unsigned char>(s[0])) || s[0] == '_')) return false;
                for (char c : s) { if (!std::isalnum(static_cast<unsigned char>(c)) && c != '_') return false; }
                return true;
            };
            // Fix 132: 裸 COM 调用结果 (vb6_ComCall 返回 VARIANT*) 不能用
            // vb6_VariantFromValue 包装 —— 它按"值指针"解释 void* 得 VT_EMPTY,
            // 使 `If coll.Item(i) <= 0` 恒真 (TreeMaps 所有值被判 <=0 而写成 0.0001,
            // 蓝色梯度消失)。COM 结果必须走 vb6_VariantFromComResult 解引用。
            auto wrapOperand132 = [&](const std::string& e) -> std::string {
                if ((e.find("vb6_ComCall(") != std::string::npos
                     || e.find("vb6_ComCallByDispid(") != std::string::npos)
                    && e.find("vb6_VariantFromComResult(") == std::string::npos)
                    return "vb6_VariantFromComResult(" + e + ")";
                if (e.find("vb6_VariantFromComResult(") != std::string::npos
                    || e.find("vb6_VariantFromValue(") != std::string::npos)
                    return e;
                return "vb6_VariantFromValue(" + e + ")";
            };
            // 账 #238: 取址之前还要问"这个名字是不是 vb6_VARIANT 那份存储" —— 判据在一处
            // (CCodeGen::cmpOperandMayTakeAddr), 裸标识符的形状测试本身不够: `Dim d As Double`
            // 的地址当 vb6_VARIANT* 递进 RTL 会按 VARIANT 布局读一个 8 字节标量。
            std::string leftAddr = (isLvalue(left) && cmpOperandMayTakeAddr(left, node.left.get()))
                ? ("&" + left)
                : ([&]{ std::string tmp = "_vcmp_" + std::to_string(vcmpCounter_++); c_.emitLine("vb6_VARIANT " + tmp + " = " + wrapOperand132(left) + ";"); return "&" + tmp; }());
            std::string rightAddr = (isLvalue(right) && cmpOperandMayTakeAddr(right, node.right.get()))
                ? ("&" + right)
                : ([&]{ std::string tmp = "_vcmp_" + std::to_string(vcmpCounter_++); c_.emitLine("vb6_VARIANT " + tmp + " = " + wrapOperand132(right) + ";"); return "&" + tmp; }());
            lastExpr_ = "(vb6_VarCmp" + cmpFn + "(" + leftAddr + ", " + rightAddr + "))";
            return;
        }
    }

    std::string op = mapBinaryOp(node.op);

    // Fix 039: VB6 And/Or/Xor are bitwise operators that require integer operands.
    // Cast to int32_t to handle double (from vb6_Pow) and pointer (BSTR, void*, SafeArray*)
    // operands. VB6 semantics: And/Or/Xor convert operands to Long before bitwise op.
    // Fix 039b: For Variant operands (knownVariantVars_ or cExprIsVariant), use
    // vb6_VariantToLong() instead of (int32_t)() cast, since VARIANT can't be cast to int.
    if (node.op == BinaryOp::And || node.op == BinaryOp::Or || node.op == BinaryOp::Xor) {
        // Fix 082: VBA7 LongPtr 字面量 (^ 后缀) 参与位运算时必须保持指针宽度.
        // LongPtr 是平台相关宽度 (32 位机 4 字节 / 64 位机 8 字节) -> intptr_t.
        // 一律 (int32_t) 截断会丢弃符号位:
        //   &H8000000000000000^ -> 0 (符号位被截掉),
        //   intptr_t 变量 -> 32 位 -> 再赋值回 intptr_t 时符号扩展 -> 结果错误.
        // (Common.bas: UnsignedAdd / UnsignedSub / Get_Wheel_Delta_wParam)
        // 无 LongPtr 字面量时保持原有 int32_t 行为不变.
        bool wide82 = false;
        for (const Expr* e82 : {node.left.get(), node.right.get()}) {
            if (e82 && e82->kind == ASTNodeKind::LiteralExpr &&
                static_cast<const LiteralExpr*>(e82)->literalKind == LiteralKind::LongPtr) {
                wide82 = true;
                break;
            }
        }
        // ai/032: 无符号档 (ULong/ULongLong) 参与 & | ^ —— 一律 (int32_t) 会截断,
        // ULongLong 的高 32 位整段丢掉 (实测 `u3 = (1 << 40) Or 255` 得 255),
        // ULong 的 bit31 会被当符号位。按档改用 uint32_t / uint64_t。
        // 两档的判据与 TypeSystem::bitwiseResult / inferExprType 一致 (结果类型那边
        // 由 promote 提高位宽, 这里是同一件事的"发码侧"一半)。
        const Vb6Type ltB = inferExprType(*node.left);
        const Vb6Type rtB = inferExprType(*node.right);
        const bool wideU64 = (ltB == Vb6Type::ULongLong || rtB == Vb6Type::ULongLong);
        const bool wideU32 = !wideU64 && (ltB == Vb6Type::ULong || rtB == Vb6Type::ULong);
        auto castBitwise = [&](const std::string& cExpr, const Expr* astExpr) -> std::string {
            if (wideU64 || wideU32) {
                const char* ctU = wideU64 ? "uint64_t" : "uint32_t";
                bool isVarU = cExprIsVariant(cExpr);
                if (!isVarU && astExpr && astExpr->kind == ASTNodeKind::IdentifierExpr) {
                    auto& idU = static_cast<IdentifierExpr&>(const_cast<Expr&>(*astExpr));
                    if (knownVariantVars_.count(Symbol::toLower(idU.name))) isVarU = true;
                }
                // Variant 提取走 LongPtr 版 (64 位档用 ToLong 会先截成 int32_t)
                if (isVarU) return "(" + std::string(ctU) + ")vb6_VariantToLongPtr(" + cExpr + ")";
                return "((" + std::string(ctU) + ")(" + cExpr + "))";
            }
            if (wide82) {
                // Variant 提取用 vb6_VariantToLongPtr (返回 intptr_t) 而非
                // vb6_VariantToLong (int32_t), 否则会先截断再拓宽.
                if (cExprIsVariant(cExpr)) return "vb6_VariantToLongPtr(" + cExpr + ")";
                if (astExpr && astExpr->kind == ASTNodeKind::IdentifierExpr) {
                    auto& ident82 = static_cast<IdentifierExpr&>(const_cast<Expr&>(*astExpr));
                    if (knownVariantVars_.count(Symbol::toLower(ident82.name))) {
                        return "vb6_VariantToLongPtr(" + cExpr + ")";
                    }
                }
                return "((intptr_t)(" + cExpr + "))";
            }
            if (cExprIsVariant(cExpr)) return "vb6_VariantToLong(" + cExpr + ")";
            // Fix 110m: 本函数返回变量 (vb6_ret_X, C 类型 vb6_VARIANT) 参与位运算.
            // VB6 中函数名即返回变量, 且返回值是 Variant —
            //   Function ARGB(...): ARGB = ARGB Or CLng(Red) * &H10000 Or ...
            // 里读 ARGB 得到 Variant. VARIANT 结构体不能直接做 | (C2440 "无法从
            // vb6_VARIANT 转换为 int32_t", 且会使外层 vb6_VariantBool 报 C2198).
            // (Charts 2020 ppProgressCircular.pag 550)
            if (!currentReturnVar_.empty() && cExpr == currentReturnVar_
                && currentReturnCType_ == "vb6_VARIANT") {
                return "vb6_VariantToLong(" + cExpr + ")";
            }
            if (astExpr && astExpr->kind == ASTNodeKind::IdentifierExpr) {
                auto& ident = static_cast<IdentifierExpr&>(const_cast<Expr&>(*astExpr));
                std::string lower = ident.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                if (knownVariantVars_.count(lower)) return "vb6_VariantToLong(" + cExpr + ")";
                if (currentProc_ && Symbol::toLower(currentProc_->name) == lower
                    && currentReturnCType_ == "vb6_VARIANT") {
                    return "vb6_VariantToLong(" + cExpr + ")";
                }
            }
            return "(int32_t)(" + cExpr + ")";
        };
        lastExpr_ = "(" + castBitwise(left, node.left.get()) + " " + op + " " + castBitwise(right, node.right.get()) + ")";
        return;
    }

    // Fix 086: 算术运算 (+,-,*) 的 Variant 操作数提取 — VARIANT 结构体不能直接
    // 参与 C 算术运算 (C2088). 用 vb6_VariantToLong 提取 (覆盖 ByRef Variant
    // 参数解引用 (*maxLen) 与 vb6_VariantArrayGet 等返回 Variant 的表达式).
    // 账 #305 §B145: 走到这里说明上面那条改道**没接管**这一对操作数（两侧都不是 Variant，
    // 或有一侧静态是 String）—— 那两种形状维持 Fix 086 的原样，逐字节未动。
    if (node.op == BinaryOp::Add || node.op == BinaryOp::Sub || node.op == BinaryOp::Mul) {
        auto extractVar86 = [&](std::string& cExpr, const Expr* astExpr) {
            if (isVariantOperand86(cExpr, astExpr)) {
                cExpr = "vb6_VariantToLong(" + cExpr + ")";
            }
        };
        extractVar86(left, node.left.get());
        extractVar86(right, node.right.get());
    }

    // Fix 108: 走到这里说明左右操作数都将按 C 原生运算符发射 (类型特化分支都未命中).
    // 无类型 COM 调用 (vb6_ComCall) 的静态 C 类型是 void* (指向 VARIANT),
    // 直接参与算术/关系运算会报 C2440 (void* 与 float/int 比较) 或 C2111
    // (指针 +/-). 实例: Charts 2020 ucChartBar GetMax
    //   `If M < m_Serie(i).Values(j) Then` → `M < vb6_ComCall(...)`
    // 这里补一次显式数值解包. 仅识别无类型形式, vb6_ComCallInt/Double/BSTR/
    // Object 前缀不同, 不受影响; `Is` (对象引用比较) 排除在外.
    if (node.op != BinaryOp::Is) {
        auto unwrapBareComCall108 = [](const std::string& cExpr) -> std::string {
            if (cExpr.compare(0, 12, "vb6_ComCall(") == 0
                || cExpr.compare(0, 20, "vb6_ComCallByDispid(") == 0) {
                return "vb6_VariantToDouble(vb6_VariantFromComResult(" + cExpr + "))";
            }
            return cExpr;
        };
        left = unwrapBareComCall108(left);
        right = unwrapBareComCall108(right);
    }

    // C3 扩展 (ai/032): 移位 `<<` / `>>`。
    //
    // 为什么不能直接发 C 的 `(a << b)`: 三处语义都与 VB.NET 不同, 且两处是 UB。
    //   1) 移位计数。VB.NET 把计数**屏蔽到左操作数位宽内** (Long 上 `1 << 32` 得 1),
    //      C 则对 计数 >= 位宽 是未定义行为 —— 变量计数时这是真会咬人的 (编译期
    //      优化可以把 UB 变成任意结果)。RTL 的 vb6_Shl/vb6_Shr* 在 64 位中间量上
    //      按 63 取模, 再经下面的"投回左操作数类型"截断, 对 8/16/32 位操作数得到的
    //      低位与"按自身位宽取模"完全一致。
    //   2) 有符号左移溢出。`&H80000000 << 1` 在 C 里是 int32 有符号溢出 UB, 在
    //      RTL 里按 uint64 位模式做完再截断, 结果是确定的。
    //   3) `>>` 的有符号性。VB.NET 对有符号整型的 >> 是**算术**右移 (符号扩展),
    //      对无符号整型 (Byte/UInteger/ULong/ULongLong) 是**逻辑**右移。C 的 >> 对
    //      有符号数是实现定义, 所以按左操作数的有符号性分别发 vb6_Shr / vb6_ShrU。
    // 结果类型 = 两操作数提升后的类型, 但窄整型抬到 32 位 (理由见下面 castBack 处);
    // 这样 inferExprType (走 promote) 与生成的 C 类型自洽, 上层 CStr/装箱/收窄检查
    // 不会看到意料外的宽度。
    if (node.op == BinaryOp::Shl || node.op == BinaryOp::Shr) {
        Vb6Type ltShift = inferExprType(*node.left);
        // Variant 操作数不能直接进 (int64_t) 强转 —— 先取数值面。VB6 字面/变量混进
        // Variant 的写法太常见, 漏这一步就是 C 编译错而不是运行期错。
        auto shiftOperand = [&](std::string cExpr, const Expr* astExpr) -> std::string {
            bool isVar = cExprIsVariant(cExpr);
            if (!isVar && astExpr && astExpr->kind == ASTNodeKind::IdentifierExpr) {
                auto& id = static_cast<IdentifierExpr&>(const_cast<Expr&>(*astExpr));
                if (knownVariantVars_.count(Symbol::toLower(id.name))) isVar = true;
            }
            // 用 LongPtr 版提取器 (它认得 VT_I8/VT_UI8), 别用 ToLong 把 64 位砍成 32
            if (isVar) return "vb6_VariantToLongPtr(" + cExpr + ")";
            return cExpr;
        };
        std::string lhsShift = shiftOperand(left, node.left.get());
        std::string rhsShift = shiftOperand(right, node.right.get());

        bool lhsUnsigned = (ltShift == Vb6Type::Byte || ltShift == Vb6Type::UInteger
                            || ltShift == Vb6Type::ULong || ltShift == Vb6Type::ULongLong);
        std::string core;
        if (node.op == BinaryOp::Shl) {
            core = "vb6_Shl((int64_t)(" + lhsShift + "), (int64_t)(" + rhsShift + "))";
        } else if (lhsUnsigned) {
            core = "vb6_ShrU((uint64_t)(" + lhsShift + "), (int64_t)(" + rhsShift + "))";
        } else {
            core = "vb6_Shr((int64_t)(" + lhsShift + "), (int64_t)(" + rhsShift + "))";
        }
        // 结果类型: **左操作数**类型, 但窄整型一律抬到 32 位。
        // MSDN 的移位表就是把移位当"一元运算作用在左操作数上"(VB.NET 甚至要求右
        // 操作数是 Integer), 所以计数不参与结果类型 —— 改前这里写的是
        // promote(左, 右), 右操作数的类型会污染结果: `u >> 4` (u As ULong, 4 是
        // Long 字面量) 按新的混符号规则会被提升成 LongLong, 与这里投回的 uint32_t
        // 两个口径。类型口径与 inferExprType (cgen_util_type.cpp 的 Shl/Shr 分支)
        // 是同一份, 两处都只看左操作数。
        // 为什么不照抄 VB.NET 的"结果 = 左操作数类型": C3 的 Integer 是 **16 位**
        // (VB6 口径), 而 VB.NET 的是 32 位。照搬的话 `1 << 20` 会被截成 0
        // (1<<20 = 0x100000, 低 16 位全 0), 而任何人的直觉与 VB.NET 实测都是
        // 1048576 —— 因为 VB.NET 里字面量 1 是 32 位。抬到 32 位既符合 C 的整型
        // 提升方向, 又让 C 实际类型与 inferExprType 自洽。
        // 认不出类型 (Unknown/Variant/浮点/字符串) 时不投, 让 C 自己按 int64 走。
        const char* castBack = nullptr;
        switch (ltShift) {
            case Vb6Type::SByte: case Vb6Type::Byte:
            case Vb6Type::UInteger: case Vb6Type::Integer:
            case Vb6Type::Boolean: case Vb6Type::Long:
                castBack = "int32_t";   break;
            case Vb6Type::ULong:     castBack = "uint32_t"; break;
            case Vb6Type::LongLong:  castBack = "int64_t";  break;
            case Vb6Type::ULongLong: castBack = "uint64_t"; break;
            case Vb6Type::LongPtr:   castBack = "intptr_t"; break;
            default:                 break;
        }
        lastExpr_ = castBack ? ("(" + std::string(castBack) + ")(" + core + ")") : core;
        return;
    }

    // Fix 126 (rev2): Currency 现在与 Date 一样按**值语义**映射为 double
    // (见 cgen_base_type.cpp), 因此混合运算无需任何缩放 —— 两侧已经是数值。
    // 早前的"Currency 一侧 /10000.0"补丁必须移除, 否则会二次缩放 (图形坐标/颜色
    // 被除到近 0, 表现为"整个图表只剩文字").
    // 放大整数 (cyVal = 值×10000) 只存在于 VT_CY 的 VARIANT 里, 由读取侧除回来
    // (rtl: vb6_VariantToDouble / vb6_Format 的 vb6_vtCurrency 分支)。

    // Fix 108b: VB6 Mod 的语义是"操作数先转 Long 再取余"; C 的 % 不接受浮点
    // 操作数 (C2296). 这里显式取整, 与 VB6 一致.
    if (node.op == BinaryOp::Mod) {
        // Task #44 → 变量除数也补上: 同 Div, 统一走 vb6_Num_Mod (错误 11 语义)。
        // ai/032 rev2: 与 `\` 完全同一份分档 (intDivResultType) —— vb6_Num_Mod 的形参
        // 也是 int32_t, 无符号/64 位操作数会被先按有符号解释。
        Vb6Type ltM = inferExprType(*node.left);
        Vb6Type rtM = inferExprType(*node.right);
        switch (intDivResultType(ltM, rtM)) {
            case Vb6Type::ULongLong:
                lastExpr_ = "vb6_Num_ModULongLong("
                          + asUnsignedOperand(left,  *node.left,  "uint64_t") + ", "
                          + asUnsignedOperand(right, *node.right, "uint64_t") + ")";
                return;
            case Vb6Type::LongLong:
                lastExpr_ = "vb6_Num_ModLongLong("
                          + asUnsignedOperand(left,  *node.left,  "int64_t") + ", "
                          + asUnsignedOperand(right, *node.right, "int64_t") + ")";
                return;
            case Vb6Type::ULong:
                lastExpr_ = "vb6_Num_ModULong("
                          + asUnsignedOperand(left,  *node.left,  "uint32_t") + ", "
                          + asUnsignedOperand(right, *node.right, "uint32_t") + ")";
                return;
            default:
                lastExpr_ = "vb6_Num_Mod((int32_t)(" + left + "), (int32_t)(" + right + "))";
                return;
        }
    }

    // VB6的And/Or/Not是逻辑运算也是位运算（取决于操作数类型）
    // 简化处理: 直接映射为C位运算, VB6语义兼容
    // Fix 092v: VB6 关系比较 (=,<>,<,>,<=,>=,Is) 结果为 Boolean(-1/0),
    // 而 C 原生比较为 0/1. 取负转 -1/0, 使上层 Not(位反)/And/Xor/算术
    // 与 VB6 一致. 只有关系运算符需要转; 算术(+-*/等)保持原样.
    // ---- ai/032 rev2: 有符号/无符号混算的显式加宽 (判据见 mixedSignWidenCType) ----
    // 走到这里的 = 算术 `+ - *` 与关系比较 —— 两者在 C 里都是裸发运算符, 而 C 对
    // "同宽异符号"取无符号, 方向与 VB.NET 相反 (见上面那个函数的注释)。
    // 插在这一步之前: 后面只剩"取负转 Boolean(-1/0)"这一道收口。
    if (const char* widenC = mixedSignWidenCType(inferExprType(*node.left),
                                                inferExprType(*node.right))) {
        left  = "((" + std::string(widenC) + ")(" + left  + "))";
        right = "((" + std::string(widenC) + ")(" + right + "))";
    }

    static const std::unordered_set<std::string> relOps092v = {
        "==", "!=", "<", ">", "<=", ">="
    };
    if (relOps092v.count(op)) {
        lastExpr_ = "(-(" + left + " " + op + " " + right + "))";
    } else {
        lastExpr_ = "(" + left + " " + op + " " + right + ")";
    }
}

} // namespace vb6c3
