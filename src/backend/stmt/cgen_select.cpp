#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_select.cpp: Select Case 语句生成 ---

void CCodeGen::visit(SelectCaseStmt& node) {
    // 推断测试表达式的类型
    Vb6Type testType = inferExprType(*node.testExpr);
    bool isStringSelect = TypeSystem::isString(testType);
    bool isFloatSelect = TypeSystem::isFloat(testType);

    emitExpr(*node.testExpr);
    if (isComMarker_) {
        // Fix 161d: 解锁类型必须与下面 tempType **同源**。原先这里是无参调用 ⇒ 走
        // resolveComValue 的默认参数 "BSTR" ⇒ vb6_ComGetStringProp 返回 wchar_t*,
        // 而 tempType 对非字符串测试是 int32_t ⇒ `int32_t v = (wchar_t*)ptr` ——
        // **裸指针被当数值**(x86 截断成低 32 位打出 -1916936344 一类值; x64 是
        // C4047 warning)。这就是 ai/029:429 记的"未登记控件属性按数值读漏裸指针",
        // 只有**未登记**属性中招: 已登记属性走 getControlPropReadFn 的专属 getter
        // (返回 int32_t), 根本不进这个 marker 分支。
        // 同族的两个先例: Fix 092n (For 的 start/end/step) 与 Fix 092r (UnaryExpr),
        // 口径一致 —— 按目标类型解锁, 不按"默认最通用"解锁。
        // 注: 未登记属性的 inferExprType 兜底是 Variant ⇒ isString/isFloat 均为假
        // ⇒ 落 Long 分支, 与 tempType=int32_t 对齐; 读不到值时 getter 给 0。
        if (isStringSelect)      resolveComValue("BSTR");
        else if (isFloatSelect)  resolveComValue("Double");
        else                     resolveComValue("Long");
    }
    std::string testVar = lastExpr_;

    // 为test创建临时变量
    std::string tempVar = "_vb6_select_" + std::to_string(tempCounter_++);
    c_.emitLine("{");
    c_.indent();

    // 根据测试表达式类型选择临时变量类型
    std::string tempType;
    if (isStringSelect) {
        tempType = "BSTR";
    } else if (isFloatSelect) {
        // Fix 136: Single 测试必须按 **float** 承载并参与 Case 比较。
        // 之前临时变量用 double: `Select Case iStep: Case Is = 0.2 * nDec`
        // (ucChartBar 轴步长) 左边 0.2f=0.20000000298…, 右边 double 0.2 →
        // 永不相等 → Case 永不命中 → nDec 循环 26 万次不退出 (白屏死循环)。
        // VB6 两边都是 Single, 相等。Case 侧表达式也 cast (float)。
        tempType = "float";
    } else {
        tempType = "int32_t";
    }

    // 声明并初始化临时变量, 保存测试表达式的值
    // Fix 038b-6: 如果测试表达式是 Variant 但临时变量是具体类型, 插入提取函数
    // 仅使用 cExprIsVariant (C 字符串级) 和 knownVariantVars_ 检测.
    {
        std::string initExpr = testVar;
        bool testIsVariant = cExprIsVariant(testVar);
        if (!testIsVariant && node.testExpr && node.testExpr->kind == ASTNodeKind::IdentifierExpr) {
            auto& id = static_cast<IdentifierExpr&>(*node.testExpr);
            std::string idLower = id.name;
            std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
            if (knownVariantVars_.count(idLower)) testIsVariant = true;
        }
        if (testIsVariant && !isStringSelect && !isFloatSelect) {
            // int32_t temp = Variant → vb6_VariantToLong(Variant)
            initExpr = "vb6_VariantToLong(" + testVar + ")";
        } else if (testIsVariant && isStringSelect) {
            // BSTR temp = Variant → vb6_VariantToString(Variant)
            initExpr = "vb6_VariantToString(" + testVar + ")";
        } else if (testIsVariant && isFloatSelect) {
            // double temp = Variant → vb6_VariantToDouble(Variant)
            initExpr = "vb6_VariantToDouble(" + testVar + ")";
        }
        c_.emitLine(tempType + " " + tempVar + " = " + initExpr + ";");
    }

    // 用if-else if链代替switch (VB6 Select Case支持范围比较和字符串)
    bool first = true;
    for (auto& caseClause : node.cases) {
        // 构建条件: 同一Case子句的多个值用||连接 (Case 1, 2, 3 → val==1 || val==2 || val==3)
        std::string combinedCond;

        for (size_t vi = 0; vi < caseClause->values.size(); vi++) {
            auto& cv = caseClause->values[vi];
            std::string cond;

            if (cv.isIsClause) {
                // Case Is > 0 → tempVar > 0
                // cv.value 是 BinaryExpr(IdentifierExpr("Is"), op, rightOperand)
                if (cv.value && cv.value->kind == ASTNodeKind::BinaryExpr) {
                    auto& binExpr = static_cast<BinaryExpr&>(*cv.value);
                    emitExpr(*binExpr.right);
                    std::string rightVal = std::move(lastExpr_);
                    if (isStringSelect) {
                        // 字符串比较: vb6_StrCmp(tempVar, rightVal) op 0
                        // <vbeclipse>: 本模块 Option Compare Text → 恒文本入口
                        cond = std::string(optionCompareText_ ? "vb6_StrCmpT(" : "vb6_StrCmp(") + tempVar + ", " + rightVal + ") " + mapBinaryOp(binExpr.op) + " 0";
                    } else if (isFloatSelect) {
                        // Fix 136: Single 语义 — Case 侧也按 float 求值
                        cond = "(float)" + tempVar + " " + mapBinaryOp(binExpr.op) + " (float)(" + rightVal + ")";
                    } else {
                        cond = tempVar + " " + mapBinaryOp(binExpr.op) + " " + rightVal;
                    }
                } else {
                    // Case Is (无比较符) → 非零/非空
                    if (isStringSelect) {
                        cond = tempVar + " != NULL && " + tempVar + "[0] != 0";
                    } else {
                        cond = tempVar + " != 0";
                    }
                }
            } else if (cv.toValue) {
                // Case 1 To 10 → tempVar >= 1 && tempVar <= 10
                emitExpr(*cv.value);
                std::string lo = std::move(lastExpr_);
                emitExpr(*cv.toValue);
                std::string hi = std::move(lastExpr_);
                if (isStringSelect) {
                    // 字符串范围比较: wcscmp >= lo && wcscmp <= hi
                    cond = std::string(optionCompareText_ ? "vb6_StrCmpT(" : "vb6_StrCmp(") + tempVar + ", " + lo + ") >= 0 && "
                           + (optionCompareText_ ? "vb6_StrCmpT(" : "vb6_StrCmp(") + tempVar + ", " + hi + ") <= 0";
                } else if (isFloatSelect) {
                    // Fix 136: Single 语义范围比较
                    cond = "(float)" + tempVar + " >= (float)(" + lo + ") && (float)" + tempVar + " <= (float)(" + hi + ")";
                } else {
                    cond = tempVar + " >= " + lo + " && " + tempVar + " <= " + hi;
                }
            } else {
                // 精确匹配
                emitExpr(*cv.value);
                if (isStringSelect) {
                    cond = std::string(optionCompareText_ ? "vb6_StrCmpT(" : "vb6_StrCmp(") + tempVar + ", " + lastExpr_ + ") == 0";
                } else if (isFloatSelect) {
                    // Fix 136: Single 语义精确匹配
                    cond = "(float)" + tempVar + " == (float)(" + lastExpr_ + ")";
                } else {
                    cond = tempVar + " == " + lastExpr_;
                }
            }

            if (!combinedCond.empty()) combinedCond += " || ";
            combinedCond += "(" + cond + ")";
        }

        if (first) {
            c_.emitLine("if (" + combinedCond + ") {");
            first = false;
        } else {
            c_.emitLine("} else if (" + combinedCond + ") {");
        }
        c_.indent();
        emitStmtList(caseClause->body);
        c_.dedent();
    }

    if (!node.elseCase.empty()) {
        c_.emitLine("} else {");
        c_.indent();
        emitStmtList(node.elseCase);
        c_.dedent();
    }

    if (!first) {
        c_.emitLine("}");
    }
    c_.dedent();
    c_.emitLine("}");
}

void CCodeGen::visit(CaseClause& node) {
    // 由SelectCaseStmt内部处理
}

} // namespace vb6c3
