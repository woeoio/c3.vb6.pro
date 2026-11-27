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
CCodeGen::SelectCaseLabelClass CCodeGen::selectCaseLabelClass(SelectCaseStmt& node,
                                                             bool* outSlotLabel) const {
    bool sawLabel = false;
    bool sawString = false;
    bool allString = true;
    bool sawSlotLabel = false;
    // 槽标签的问法与主语同一条（第319刀那道闸）：只认**按名在册的 Variant 槽**这一形。
    // 调用 / 成员读那一档 C 侧是 int32_t 还是整枚槽问不出（要发了码才知道），宁可不改口。
    auto labelIsSlot = [&](const Expr* e) -> bool {
        if (!e || e->kind != ASTNodeKind::IdentifierExpr) return false;
        std::string n = static_cast<const IdentifierExpr&>(*e).name;
        std::transform(n.begin(), n.end(), n.begin(), ::tolower);
        return knownVariantVars_.count(n) != 0;
    };
    for (auto& clause : node.cases) {
        for (auto& cv : clause->values) {
            if (!cv.value || !TypeSystem::isString(inferExprType(*cv.value))) allString = false;
            else sawString = true;
            if (cv.toValue && !TypeSystem::isString(inferExprType(*cv.toValue))) allString = false;
            sawLabel = true;
            if (labelIsSlot(cv.value.get()) || labelIsSlot(cv.toValue.get())) sawSlotLabel = true;
        }
    }
    if (outSlotLabel) *outSlotLabel = sawSlotLabel;
    if (!sawLabel) return SelectCaseLabelClass::NoString;
    if (allString) return SelectCaseLabelClass::AllString;
    return sawString ? SelectCaseLabelClass::Mixed : SelectCaseLabelClass::NoString;
}

void CCodeGen::visit(SelectCaseStmt& node) {
    // 推断测试表达式的类型
    Vb6Type testType = inferExprType(*node.testExpr);
    bool isStringSelect = TypeSystem::isString(testType);
    bool isFloatSelect = TypeSystem::isFloat(testType);
    // §11.190 (第二百零七刀): Date 主语的 Select Case —— 承载档必须是 **double**。
    // 改前落 int32_t 档 (`double d` 被截成 36xxx), 再把标签当 BSTR 去 `==`:
    // 真机 Y6 (`Select Case d: Case "2021-03-04"`) 真值 1, 改前答 0 (既截断又比错类型)。
    // 标签是串时按同一条口径折成日期序列 (vb6_DateVal, 与 binary 比较那腿同一枚入口)。
    // §11.195 (第二百一十二刀): 主语是 **Date 数组元素 / UDT 成员** 时 inferExprType 答
    // Variant ⇒ 旧判据把它当 Variant 主语, 标签全是串时又被拉去串承载, 发成
    // `BSTR _vb6_select = VB6_SA_AT(double, q, 0);` (真机 P259 实测一条 C2440)。
    // 判据与 binary 比较那腿同形: 先问语义类型, 再补「裸名数组」与「UDT 成员」两形。
    bool isDateSelect = (testType == Vb6Type::Date);
    if (!isDateSelect && node.testExpr) {
        Expr* subj195 = node.testExpr.get();
        if (subj195->kind == ASTNodeKind::IndexOrCallExpr) {
            const auto& ic195 = static_cast<const IndexOrCallExpr&>(*subj195);
            if (ic195.callee && ic195.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto it195 = arrayElemTypes_.find(Symbol::toLower(
                    static_cast<const IdentifierExpr&>(*ic195.callee).name));
                isDateSelect = (it195 != arrayElemTypes_.end() && it195->second == Vb6Type::Date);
            } else if (ic195.callee
                       && (ic195.callee->kind == ASTNodeKind::MemberAccessExpr
                           || ic195.callee->kind == ASTNodeKind::WithMemberExpr)) {
                bool isArr195 = false;
                isDateSelect = inferUdtFieldVb6Type(ic195.callee.get(), &isArr195) == Vb6Type::Date
                               && isArr195;
            }
        } else if (subj195->kind == ASTNodeKind::MemberAccessExpr
                   || subj195->kind == ASTNodeKind::WithMemberExpr) {
            bool isArr195 = false;
            isDateSelect = inferUdtFieldVb6Type(subj195, &isArr195) == Vb6Type::Date && !isArr195;
        }
    }

    // §11.196 (第二百一十三刀): VB6 的 `Now` / `Date` / `Time` / `DateSerial` / `DateValue` /
    // `TimeValue` / `DateAdd` / `DateDiff` / `DatePart` / `TimeSerial` 返回值是 **Variant (Date)**,
    // 不是 Date —— 真机 P262 的两格只有「按 Variant 比」才答得对：
    //   H1 `Select Case DateSerial(2021,3,4): Case "2021-03-04"` ⇒ **不命中** (按文本比),
    //   H4 `Select Case Now: Case CStr(Now)` ⇒ **命中** (文本相等; 带时刻的串解不成日期)。
    // 只在本站点降成 Variant：形状仍走 §11.191 那枚 wrapVariantValue (VT_DATE 装箱),
    // `inferExprType` 与其它站的答案一个字节都不动。
    bool dateBuiltinSubj196 = false;
    if (node.testExpr && node.testExpr->kind == ASTNodeKind::IndexOrCallExpr) {
        const auto& ic196 = static_cast<const IndexOrCallExpr&>(*node.testExpr);
        if (ic196.callee && ic196.callee->kind == ASTNodeKind::IdentifierExpr) {
            const std::string n196 = Symbol::toLower(
                static_cast<const IdentifierExpr&>(*ic196.callee).name);
            static const char* kVariantDate196[] = {
                "now", "date", "time", "dateserial", "datevalue", "timevalue",
                "dateadd", "datediff", "datepart", "timeserial"};
            for (const char* cand196 : kVariantDate196) {
                if (n196 == cand196) { dateBuiltinSubj196 = true; break; }
            }
        }
    }
    if (dateBuiltinSubj196) { testType = Vb6Type::Variant; isDateSelect = false; }

    // VB6 对 Variant 主语的 Select Case 是按 Case 项的类型逐项做 VarCmp；我方以前只问
    // inferExprType 一档 ⇒ 答 Variant 就落 int32_t 档 + vb6_VariantToLong 解包，而标签
    // 是串 ⇒ 发成 `int32_t == vb6_BSTR_FromStr(L"…")`（C 模式只警告不报红），**每一档都
    // 假** ⇒ 恒落 Case Else。靶子 = ComCtlsDemo 工具栏 `Select Case Button.Tag`
    // （TbrButton.cls:73 `Property Get Tag() As Variant`）：点 "Show ListBoxW1" 该让
    // SysListView32 隐身、ListBox 上位，实测只弹了一句 "clicked button 5"、零枚窗口翻转。
    bool slotLabel = false;
    const SelectCaseLabelClass labelClass = selectCaseLabelClass(node, &slotLabel);
    // §11.192 (第二百零九刀): 标签里有 **Date 形**（日期字面量 / Date 变量 / Date 数组元素）时,
    // 主语是槽的话必须走整枚槽那条链 —— 真机 P250 的 A7/A8 读数是「文本日期命中、Null 不命中
    // 且不报错」, 这正好是 §11.191 给 vb6_VarCmp* 定下的口径（VT_DATE 与文本静默解析、
    // VT_NULL 无序）；而旧的 int32_t + vb6_VariantToLong 两格都错（A7 答 0、A8 抛 94 误命中）。
    bool dateLabelPresent = false;
    for (auto& cc : node.cases) {
        for (auto& v190 : cc->values) {
            Expr* e190 = v190.value.get();
            if (e190 && inferExprType(*e190) == Vb6Type::Date) { dateLabelPresent = true; break; }
            Expr* t190 = v190.toValue.get();
            if (t190 && inferExprType(*t190) == Vb6Type::Date) { dateLabelPresent = true; break; }
        }
        if (dateLabelPresent) break;
    }
    if (!isStringSelect && !isFloatSelect && !isDateSelect && !dateBuiltinSubj196
        && testType == Vb6Type::Variant
        && labelClass == SelectCaseLabelClass::AllString) {
        isStringSelect = true;
    }
    // 账 #305 第324刀 = 311/319 留下的最后那一档（账 #606）：标签里有**整枚 Variant 槽**时，
    // 三条整档承载（串承载 / 浮点承载 / Long 解包）没有一条成立 —— 它们都把主语折成一个
    // 标量，再拿标量去 `==` / `vb6_StrCmp` 那枚结构体，发出来的是**编不过的 C**
    // （实测 `error C2088: “==”: 对于 struct 非法`，四种主语档同形：Variant / Long / String / Double）。
    // VB6 那一侧本来就是按 Case 项逐项 VarCmp，所以承载档一律改走整枚槽 + 下面那条 VarCmp 链。
    if (slotLabel) { isStringSelect = false; isFloatSelect = false; }
    // 账 #305 第319刀 = 311 留下的那一形：串数**混合**标签。同一份标签里既有 Case "a"
    // 又有 Case 1 时，311 的两条整档出口（串承载 / Long 解包）都不成立 —— 只有逐项
    // 按该标签自己的档比才对。承载档改成整枚槽，逐项走 VarCmp（与 binary 比较同一份
    // 出口，不在这里另开打包器）；标签的装箱走 wrapVariantValue 唯一那处。
    // 只放开「主语**真是**槽」这一形（下面 testIsVariant 那一道闸），类型化 getter
    // （inferExprType 答 Variant 而 C 侧是 int32_t 的那档）仍走 ToLong。
    const bool wantVariantCarrier = !isStringSelect && !isFloatSelect
        && (slotLabel || dateBuiltinSubj196
            || (testType == Vb6Type::Variant
                && (labelClass == SelectCaseLabelClass::Mixed || dateLabelPresent)));

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
    bool testIsVariant = isVariantSlot(testVar, nullptr);  // P1.3: ast 不可得 (testVar 经 extenderMemberReadExpr 变换)
    if (!testIsVariant && node.testExpr && node.testExpr->kind == ASTNodeKind::IdentifierExpr) {
        auto& id = static_cast<IdentifierExpr&>(*node.testExpr);
        std::string idLower = id.name;
        std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
        if (knownVariantVars_.count(idLower)) testIsVariant = true;
    }
    // 混合档的第二道闸：类型化 getter（inferExprType 答 Variant、C 侧却是 int32_t 那档）
    // 不满足 testIsVariant ⇒ 退回 ToLong，与第311刀以前的行为一致。
    // 第324刀放宽一处：槽标签那一档不要求主语是槽 —— 主语不是槽就**装箱**（下面那处
    // wrapVariantValue 唯一出口），因为要迁就的是标签的档，不是主语的档。
    const bool isVariantSelect = wantVariantCarrier && (testIsVariant || slotLabel || dateBuiltinSubj196);

    // 根据测试表达式类型选择临时变量类型
    std::string tempType;
    if (isDateSelect) {
        tempType = "double";
    } else if (isStringSelect) {
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
    bool subjectBoxOwned = false;
    {
        std::string initExpr = testVar;
        if (isDateSelect) {
            // §11.192 (第二百零九刀): Date 承载档优先于“整枚槽”那一档。标签是 Variant 时
            // `isVariantSelect` 也会为真 (slotLabel), 但主语 Date 的 tempType 是 double ——
            // 再走 wrapVariantValue 就发成 `double x = vb6_VariantDate(d);` (真机 P250 实测 6 条 C2440)。
            initExpr = testIsVariant ? ("vb6_CDate(" + testVar + ")") : testVar;
        } else if (isVariantSelect && !testIsVariant) {
            // 第324刀：主语不是槽而标签是槽 ⇒ 把主语装箱成整枚槽（唯一那处 wrapVariantValue，
            // 与标签同一个打包器）。它说 owns 的（现造的调用 / 对象档取了引用）要在链尾收，
            // 说 borrowed 的（裸变量别名）一个都不许清 —— 同 §B164 那条主人。
            initExpr = wrapVariantValue(node.testExpr.get(), testVar);
            subjectBoxOwned = variantRhsOwnsValue(node.testExpr.get(), testVar);
        } else if (isVariantSelect) {
            // 已是槽: 原样拷进临时量, 不解包 (解包就把标签的档丢了)
        } else if (testIsVariant && !isStringSelect && !isFloatSelect) {
            // int32_t temp = Variant → vb6_VariantToLong(Variant)
            initExpr = "vb6_VariantToLong(" + testVar + ")";
        } else if (testIsVariant && isDateSelect) {
            // §11.190: 槽主语进 Date 承载 —— 走 CDate 那一档 (文本按日期解析、Null 抛 94)
            initExpr = "vb6_CDate(" + testVar + ")";
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
    if (isVariantSelect && !isDateSelect) {
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
        /* 比较出口的选择（唯一一处）：标签是**静态类型**的字面量 ⇒ 主语换成标签那一档再比
           （vb6_VarSelCmp*）；标签本身是 Variant 槽 ⇒ Variant-Variant 那一档（vb6_VarCmp*）。
           依据是真 VB6 /make 实测（2026-11-07）：主语 7 + Case "07" 不命中（串比），
           主语 "07" + Case 7 命中（数比）—— 二进制比较那条「数值恒小于字符串」在这里不成立。 */
        auto cmpPrefix = [&](Expr* lbl) -> std::string {
            // §11.192: 标签是 Date 形 ⇒ 走 Variant-Variant 那一档（vb6_VarCmp*）——
            // 它的那套口径（VT_NULL 无序、VT_DATE 与文本静默解析，§11.191）才是真机
            // 在 P250 A7/A8 上的读数；把主语往 Date 档拉的 VarSelCmp 会把 Null 抛成
            // 94（实测 A8 误命中），与真机“不命中、不报错”相反。
            if (lbl && inferExprType(*lbl) == Vb6Type::Date) return "vb6_VarCmp";
            if (lbl && inferExprType(*lbl) != Vb6Type::Variant) return "vb6_VarSelCmp";
            return "vb6_VarCmp";
        };
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
                        cond = cmpPrefix(cv.value.get()) + cmpSuffix(cv.relOp) + "(" + subjAddr + ", "
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
                    cond = cmpPrefix(cv.value.get()) + "Ge(" + subjAddr + ", " + loAddr + ") && "
                         + cmpPrefix(cv.toValue.get()) + "Le("
                         + subjAddr + ", " + hiAddr + ")";
                } else {
                    emitExpr(*cv.value);
                    cond = cmpPrefix(cv.value.get()) + "Eq(" + subjAddr + ", "
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
        // 第324刀：装箱的主语同样只收「箱自己拿下所有权」的那一枚（声明在标签之前 ⇒ 收在之后）
        if (subjectBoxOwned) {
            c_.emitLine("vb6_VariantClear(&" + tempVar + ");");
        }
        c_.dedent();
        c_.emitLine("}");
        return;
    }

    // 用if-else if链代替switch (VB6 Select Case支持范围比较和字符串)
    // 先**只建条件**、再发 if 链：条件求值时的物化语句（如 ParamArray 的
    // `SAFEARRAY* _pa_N = NULL;`）发在当前输出位置 —— 边建边发的话，第二条
    // Case 的声明会落在上一条 Case 的**体内**，`} else if (...)` 那一行却在
    // 体外观测不到它 ⇒ C2065（实测 Units.c 的 `Select Case g_Language.
    // TranslateMessage("px")` 一族 12 条）。变体档那条链早就按这个次序做
    // （见上面 labelDecls 的注释），这里与它对齐；没有物化语句时发码逐字节不变。
    std::vector<std::string> clauseConds;
    for (auto& caseClause : node.cases) {
        // 构建条件: 同一Case子句的多个值用||连接 (Case 1, 2, 3 → val==1 || val==2 || val==3)
        std::string combinedCond;

        for (size_t vi = 0; vi < caseClause->values.size(); vi++) {
            auto& cv = caseClause->values[vi];
            std::string cond;
            // §11.190/§11.192: Date 主语的标签 —— 串档折成日期序列, Variant 标签走 CDate 那一档
            // (文本按日期解析、数值原样、Null 抛 94、坏文本抛 13; 真机 P250 的 A3/A4 就是这两档),
            // 其余按数值原样。
            // §11.193: 串主语的标签若是 Date 形 —— 先折成 VB6 自己的短日期文本再按串比。
            // 真机 p253 的读数是「按文本比」: T3 (`CStr(CDate(...))` = "2021/3/4") 命中,
            // T1 "9999" > #1/1/2021# 命中, T5 "2021-03-05" 不命中 (`-` < `/`)。
            auto strLabel193 = [&](Expr* lbl, const std::string& txt) -> std::string {
                if (!isStringSelect) return txt;
                if (lbl && inferExprType(*lbl) == Vb6Type::Date)
                    return "vb6_CStrDate(" + txt + ")";
                return txt;
            };
            auto dateLabel190 = [&](Expr* lbl, const std::string& txt) -> std::string {
                if (!isDateSelect) return txt;
                // 「这格是 Date 吗」与比较站同一个问句形状: 先问语义类型, 再补数组元素那一形
                // (inferExprType 对数组元素答 Variant, 与 §11.190 的 Y3 同因)。
                auto isDate190 = [&](Expr* e) -> bool {
                    if (!e) return false;
                    if (inferExprType(*e) == Vb6Type::Date) return true;
                    if (e->kind == ASTNodeKind::IndexOrCallExpr) {
                        const auto& ic = static_cast<const IndexOrCallExpr&>(*e);
                        if (ic.callee && ic.callee->kind == ASTNodeKind::IdentifierExpr) {
                            auto it = arrayElemTypes_.find(Symbol::toLower(
                                static_cast<const IdentifierExpr&>(*ic.callee).name));
                            return it != arrayElemTypes_.end() && it->second == Vb6Type::Date;
                        }
                    }
                    return false;
                };
                Vb6Type lt190 = lbl ? inferExprType(*lbl) : Vb6Type::Double;
                if (isDate190(lbl)) return txt;
                if (lt190 == Vb6Type::String) return "vb6_DateVal(" + txt + ")";
                if (lt190 == Vb6Type::Variant) return "vb6_CDate(" + txt + ")";
                return txt;
            };

            if (cv.isIsClause) {
                // Case Is > 0 → tempVar > 0 (比较符在 cv.relOp, 右操作数在 cv.value)
                if (cv.hasRelOp && cv.value) {
                    emitExpr(*cv.value);
                    std::string rightVal = std::move(lastExpr_);
                    if (isDateSelect) {
                        cond = tempVar + " " + mapBinaryOp(cv.relOp) + " " + dateLabel190(cv.value.get(), rightVal);
                    } else if (isStringSelect) {
                        // 字符串比较: vb6_StrCmp(tempVar, rightVal) op 0
                        // <vbeclipse>: 本模块 Option Compare Text → 恒文本入口
                        // §11.193: 标签是日期时先折成 VB6 自己的短日期文本 (vb6_CStrDate — 与
                        // CStr(Date) 同一个出口), 再按串比 (真机 p253: T3 命中、T1/T2/T5 全对)。
                        cond = std::string(optionCompareText_ ? "vb6_StrCmpT(" : "vb6_StrCmp(") + tempVar + ", " + strLabel193(cv.value.get(), rightVal) + ") " + mapBinaryOp(cv.relOp) + " 0";
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
                if (isDateSelect) {
                    cond = tempVar + " >= " + dateLabel190(cv.value.get(), lo) + " && " + tempVar
                           + " <= " + dateLabel190(cv.toValue.get(), hi);
                } else if (isStringSelect) {
                    // 字符串范围比较: wcscmp >= lo && wcscmp <= hi
                    cond = std::string(optionCompareText_ ? "vb6_StrCmpT(" : "vb6_StrCmp(") + tempVar + ", " + strLabel193(cv.value.get(), lo) + ") >= 0 && "
                           + (optionCompareText_ ? "vb6_StrCmpT(" : "vb6_StrCmp(") + tempVar + ", " + strLabel193(cv.toValue.get(), hi) + ") <= 0";
                } else if (isFloatSelect) {
                    // Fix 136: Single 语义范围比较
                    cond = "(float)" + tempVar + " >= (float)(" + lo + ") && (float)" + tempVar + " <= (float)(" + hi + ")";
                } else {
                    cond = tempVar + " >= " + lo + " && " + tempVar + " <= " + hi;
                }
            } else {
                // 精确匹配
                emitExpr(*cv.value);
                if (isDateSelect) {
                    cond = tempVar + " == " + dateLabel190(cv.value.get(), lastExpr_);
                } else if (isStringSelect) {
                    cond = std::string(optionCompareText_ ? "vb6_StrCmpT(" : "vb6_StrCmp(") + tempVar + ", " + strLabel193(cv.value.get(), lastExpr_) + ") == 0";
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

        clauseConds.push_back(combinedCond);
    }

    bool first = true;
    for (size_t ci = 0; ci < node.cases.size(); ci++) {
        if (first) {
            c_.emitLine("if (" + clauseConds[ci] + ") {");
            first = false;
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
