#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_select.cpp: Select Case 语句生成 ---

// 账 #305 第311刀: 「Case 标签落在哪一档」的**唯一**回答处（声明在
// detail/util/cgen_helpers.inc）。契约：一个标签都不许漏 —— `Case Is > x` 的右操作数、
// `Case a To b` 的两端都是标签；`Case Else` 不在 node.cases 里，不用管。
// 三类答案：AllString = 每一枚都答得出 String；Mixed = 至少一枚 String 且并非全是；
// NoString = 没有 String 标签（含拿不到标签/类型问不出）—— 后两档以外**宁可退回现行为也不猜**。
CCodeGen::SelectCaseLabelClass CCodeGen::selectCaseLabelClass(SelectCaseStmt& node) const {
    bool sawLabel = false;
    bool sawString = false;
    bool allString = true;
    for (auto& clause : node.cases) {
        for (auto& cv : clause->values) {
            if (!cv.value || !TypeSystem::isString(inferExprType(*cv.value))) allString = false;
            else sawString = true;
            if (cv.toValue && !TypeSystem::isString(inferExprType(*cv.toValue))) allString = false;
            sawLabel = true;
        }
    }
    if (!sawLabel) return SelectCaseLabelClass::NoString;
    if (allString) return SelectCaseLabelClass::AllString;
    return sawString ? SelectCaseLabelClass::Mixed : SelectCaseLabelClass::NoString;
}

void CCodeGen::visit(SelectCaseStmt& node) {
    // 推断测试表达式的类型
    Vb6Type testType = inferExprType(*node.testExpr);
    bool isStringSelect = TypeSystem::isString(testType);
    bool isFloatSelect = TypeSystem::isFloat(testType);

    // 账 #305 第311刀「实施」= 主语静态答 **Variant** 时，承载档改由 Case 标签回答。
    // VB6 对 Variant 主语的 Select Case 是按 Case 项的类型逐项做 VarCmp；我方以前只问
    // inferExprType 一档 ⇒ 答 Variant 就落 int32_t 档 + vb6_VariantToLong 解包，而标签
    // 是串 ⇒ 发成 `int32_t == vb6_BSTR_FromStr(L"…")`（C 模式只警告不报红），**每一档都
    // 假** ⇒ 恒落 Case Else。靶子 = ComCtlsDemo 工具栏 `Select Case Button.Tag`
    // （TbrButton.cls:73 `Property Get Tag() As Variant`）：点 "Show ListBoxW1" 该让
    // SysListView32 隐身、ListBox 上位，实测只弹了一句 "clicked button 5"、零枚窗口翻转。
    const SelectCaseLabelClass labelClass = selectCaseLabelClass(node);
    if (!isStringSelect && !isFloatSelect && testType == Vb6Type::Variant
        && labelClass == SelectCaseLabelClass::AllString) {
        isStringSelect = true;
    }
    // 账 #305 第319刀 = 311 留下的那一形：串数**混合**标签。同一份标签里既有 Case "a"
    // 又有 Case 1 时，311 的两条整档出口（串承载 / Long 解包）都不成立 —— 只有逐项
    // 按该标签自己的档比才对。承载档改成整枚槽，逐项走 VarCmp（与 binary 比较同一份
    // 出口，不在这里另开打包器）；标签的装箱走 wrapVariantValue 唯一那处。
    // 只放开「主语**真是**槽」这一形（下面 testIsVariant 那一道闸），类型化 getter
    // （inferExprType 答 Variant 而 C 侧是 int32_t 的那档）仍走 ToLong。
    const bool wantVariantCarrier = !isStringSelect && !isFloatSelect
        && testType == Vb6Type::Variant && labelClass == SelectCaseLabelClass::Mixed;

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
        // 注: 未登记属性的 inferExprType 兜底是 Variant ⇒ 上方两档均为假 ⇒ 落 Long 分支,
        // 与 tempType=int32_t 对齐; 读不到值时 getter 给 0。**第311刀之后**多一条岔口：
        // 同一条 Select 的 Case 标签若是**全部**串档，承载档已按标签改走 BSTR，这里跟着
        // 一起答 BSTR —— 两条问句读的是同一个 isStringSelect，不许各抄一份。
        // **第319刀**再补一条：混合档要的是**整枚槽**，解包出去就再也比不了串标签。
        if (isStringSelect)            resolveComValue("BSTR");
        else if (isFloatSelect)        resolveComValue("Double");
        else if (wantVariantCarrier)   resolveComValue("Variant");
        else                           resolveComValue("Long");
    }
    /* Fix 296 / 4b: `Select Case Extender.Align` 的判据临时量也走实例入口
       (ToolBar.ctl 发成 `int32_t _vb6_select_10 = vb6_Extender_Align;`, 实测 1 处)。
       换道落在**临时量的初始化**这一处 —— 下游 cond / switch 读的都是 tempVar,
       所以 getter 只调一次, 不会每个 Case 调一遍。 */
    std::string testVar = extenderMemberReadExpr(lastExpr_);

    // 为test创建临时变量
    std::string tempVar = "_vb6_select_" + std::to_string(tempCounter_++);
    c_.emitLine("{");
    c_.indent();

    // 测试表达式是不是**槽**（两档：C 表达式本身 + 按名表）。
    // Fix 038b-6 的既有口径，第319刀把它提到承载档选择之前 —— 混合档只在主语真是槽时成立。
    bool testIsVariant = cExprIsVariant(testVar);
    if (!testIsVariant && node.testExpr && node.testExpr->kind == ASTNodeKind::IdentifierExpr) {
        auto& id = static_cast<IdentifierExpr&>(*node.testExpr);
        std::string idLower = id.name;
        std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
        if (knownVariantVars_.count(idLower)) testIsVariant = true;
    }
    // 混合档的第二道闸：类型化 getter（inferExprType 答 Variant、C 侧却是 int32_t 那档）
    // 不满足 testIsVariant ⇒ 退回 ToLong，与第311刀以前的行为一致。
    const bool isVariantSelect = wantVariantCarrier && testIsVariant;

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
    } else if (isVariantSelect) {
        // 第319刀混合档：整枚槽直拷（VarCmp 只读，浅拷贝不动所有权）。
        tempType = "vb6_VARIANT";
    } else {
        tempType = "int32_t";
    }

    // 声明并初始化临时变量, 保存测试表达式的值
    {
        std::string initExpr = testVar;
        if (isVariantSelect) {
            // 已是槽: 原样拷进临时量, 不解包 (解包就把标签的档丢了)
        } else if (testIsVariant && !isStringSelect && !isFloatSelect) {
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

    // 第319刀混合档：主语是槽、标签串数混合 ⇒ 每一项按标签自己那一档比。
    // 比较出口复用 binary 那一份（vb6_VarCmp*，账 #238 的取址判据在那边已经定死），
    // 标签的装箱复用 wrapVariantValue（唯一那处）—— 不在这里另开打包器。
    // 形状：Case Is op x → VarCmp<Op>(&subj, &x)；Case a To b → VarCmpGe(&subj,&a) &&
    // VarCmpLe(&subj,&b)；精确 → VarCmpEq(&subj,&x)。标签临时量必须提到 if 链**之前**声明，
    // 否则第二枚标签的声明会落在 `}` 与 `else if` 之间（非法 C）。
    if (isVariantSelect) {
        auto cmpSuffix = [](BinaryOp op) -> std::string {
            switch (op) {
                case BinaryOp::Eq:  return "Eq";
                case BinaryOp::Neq: return "Ne";
                case BinaryOp::Lt:  return "Lt";
                case BinaryOp::Gt:  return "Gt";
                case BinaryOp::Le:  return "Le";
                case BinaryOp::Ge:  return "Ge";
                default:            return "Eq";
            }
        };
        std::vector<std::string> clauseConds;
        std::vector<std::string> labelDecls;
        // 账 #604: 「这枚标签箱收下了所有权没有」只问 variantRhsOwnsValue 一处（§B164/§B167
        // 那位主人）—— 它说 owns 的（字面量 / 现造的调用 / 对象档取了引用）在 if 链尾部
        // vb6_VariantClear 放掉；它说 borrowed 的（裸串变量、Variant 槽的浅别名）**一个都不许清**
        // —— wrapVariantValue 对它们是 vb6_VariantString(s) / vb6_VariantFromValue(v) 这类
        // 别名箱，清下去放的就是别人还活着的那只 BSTR / 那个 Dispatch。
        std::vector<std::string> labelClears;
        const std::string subjAddr = "&" + tempVar;
        auto labelVariantAddr = [&](Expr* ast, const std::string& expr) -> std::string {
            std::string lbl = "_vb6_select_lbl_" + std::to_string(tempCounter_++);
            labelDecls.push_back("vb6_VARIANT " + lbl + " = " + wrapVariantValue(ast, expr) + ";");
            if (variantRhsOwnsValue(ast, expr)) labelClears.push_back(lbl);
            return "&" + lbl;
        };

        for (auto& caseClause : node.cases) {
            std::string combinedCond;
            for (auto& cv : caseClause->values) {
                std::string cond;
                if (cv.isIsClause) {
                    if (cv.hasRelOp && cv.value) {
                        emitExpr(*cv.value);
                        std::string rightVal = std::move(lastExpr_);
                        cond = "vb6_VarCmp" + cmpSuffix(cv.relOp) + "(" + subjAddr + ", "
                             + labelVariantAddr(cv.value.get(), rightVal) + ")";
                    } else {
                        // 无比较符的 `Case Is`：与下方非槽档同义 —— 非零/非空
                        cond = "vb6_VariantToBool(" + tempVar + ") != 0";
                    }
                } else if (cv.toValue) {
                    emitExpr(*cv.value);
                    std::string lo = std::move(lastExpr_);
                    emitExpr(*cv.toValue);
                    std::string hi = std::move(lastExpr_);
                    // 两个标签的临时量分开算: `a + f() + b + g()` 的求值顺序在 C++ 里
                    // 未指定, 声明顺序会跟着编译器漂 (发码字节要稳定)
                    std::string loAddr = labelVariantAddr(cv.value.get(), lo);
                    std::string hiAddr = labelVariantAddr(cv.toValue.get(), hi);
                    cond = "vb6_VarCmpGe(" + subjAddr + ", " + loAddr + ") && vb6_VarCmpLe("
                         + subjAddr + ", " + hiAddr + ")";
                } else {
                    emitExpr(*cv.value);
                    cond = "vb6_VarCmpEq(" + subjAddr + ", "
                         + labelVariantAddr(cv.value.get(), lastExpr_) + ")";
                }
                if (!combinedCond.empty()) combinedCond += " || ";
                combinedCond += "(" + cond + ")";
            }
            clauseConds.push_back(combinedCond);
        }

        for (auto& decl : labelDecls) c_.emitLine(decl);
        bool firstVariant = true;
        for (size_t ci = 0; ci < node.cases.size(); ++ci) {
            if (firstVariant) {
                c_.emitLine("if (" + clauseConds[ci] + ") {");
                firstVariant = false;
            } else {
                c_.emitLine("} else if (" + clauseConds[ci] + ") {");
            }
            c_.indent();
            emitStmtList(node.cases[ci]->body);
            c_.dedent();
        }
        if (!node.elseCase.empty()) {
            c_.emitLine("} else {");
            c_.indent();
            emitStmtList(node.elseCase);
            c_.dedent();
        }
        if (!firstVariant) {
            c_.emitLine("}");
        }
        // 账 #604: 收回上面那几枚「自己拿下了所有权」的标签箱（放掉的只是箱里那只 BSTR /
        // 那枚 Dispatch 引用）。放在 if 链之后 = 每一枚标签都求过值、都过了比较；
        // Case 体里 Exit Sub / GoTo 跳出去的那条路会跳过这里（与 §B164 ①②③⑤ 同一族，
        // 宁可不收也不许把收尾发成悬空 —— 本处不欠新账）。借来的那几枚（见上）不在此列。
        for (auto& lblClear : labelClears) {
            c_.emitLine("vb6_VariantClear(&" + lblClear + ");");
        }
        c_.dedent();
        c_.emitLine("}");
        return;
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
                // Case Is > 0 → tempVar > 0 (比较符在 cv.relOp, 右操作数在 cv.value)
                if (cv.hasRelOp && cv.value) {
                    emitExpr(*cv.value);
                    std::string rightVal = std::move(lastExpr_);
                    if (isStringSelect) {
                        // 字符串比较: vb6_StrCmp(tempVar, rightVal) op 0
                        // <vbeclipse>: 本模块 Option Compare Text → 恒文本入口
                        cond = std::string(optionCompareText_ ? "vb6_StrCmpT(" : "vb6_StrCmp(") + tempVar + ", " + rightVal + ") " + mapBinaryOp(cv.relOp) + " 0";
                    } else if (isFloatSelect) {
                        // Fix 136: Single 语义 — Case 侧也按 float 求值
                        cond = "(float)" + tempVar + " " + mapBinaryOp(cv.relOp) + " (float)(" + rightVal + ")";
                    } else {
                        cond = tempVar + " " + mapBinaryOp(cv.relOp) + " " + rightVal;
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
