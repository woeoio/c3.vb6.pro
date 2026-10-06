#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_control.cpp: For/ForEach 循环语句生成 ---


void CCodeGen::visit(ForStmt& node) {
    // Fix 092n: For 的三个子表达式相互独立 — 进入每个子表达式前清掉上游残留的
    // COM 标记, 生成后立即按整数上下文消费本次产生的标记. 否则 `For i = .Count
    // To 1 Step -1` (With 对象属性读) 会双重出错 (pvSubClass 708/709):
    //   ① start 只得到对象本身 — `i = _vb6_with_2 /* With COM .Count */`;
    //   ② 该标记残留到 Step, 被 UnaryExpr 内的 `if (isComMarker_) resolveComValue()`
    //      消费 → `i_step = (-vb6_ComGetStringProp(_vb6_with_2, L"Count"))` (C2171).
    auto emitForExpr092n = [&](Expr& e) -> std::string {
        isComMarker_ = false;
        emitExpr(e);
        if (isComMarker_) resolveComValue("Long");
        return std::move(lastExpr_);
    };
    std::string start = emitForExpr092n(*node.start);
    std::string end = emitForExpr092n(*node.end);

    std::string step = "1";
    if (node.step) {
        step = emitForExpr092n(*node.step);
    }

    std::string var = cIdent(node.varName);

    // Fix 022: For 循环变量若为类成员必须以 me->var 形式 emit.
    // 原 Fix 010o 无差别将 For 变量注册到 knownLocalVars_, 但当变量实际是类成员
    // (例: cCsv.cls:51 'Dim i As Long, ii As Long' at module level → class field)
    // 时, 注册会让下游 IdentifierExpr (cgen_expr.cpp:425-432) 跳过 me-> 前缀,
    // 循环体内对 ii 的引用变裸标识符 ii, 但函数内无 C 局部 ii 只有 me->ii
    // → C2065 未声明标识符 (cCsv.c 34 处 ii + 13 处 i 同根因).
    //
    // 修复策略:
    //   1. 若 For 变量是类成员 (且未被局部 Dim 遮蔽), 用 me->var 形式 emit loop
    //      自身赋值/比较/累加语句, 且 **不** 注册到 knownLocalVars_ — 让循环体
    //      内 IdentifierExpr 走 me-> 路径 (cgen_expr.cpp 已支持).
    //   2. 否则保持原 Fix 010o 行为 (注册 + 裸 var), 覆盖局部 Dim 变量、
    //      非类模块变量、跨模块变量等情况.
    // 注: 临时变量 var_end / var_step 是 block-scoped int32_t 局部, 无需 me->.
    std::string lower022 = node.varName;
    std::transform(lower022.begin(), lower022.end(), lower022.begin(), ::tolower);
    bool forVarIsClassMember = isClassModule_
        && classMemberVars_.count(lower022)
        && !knownLocalVars_.count(lower022);  // 局部 Dim 优先遮蔽类成员
    std::string varAcc = forVarIsClassMember ? ("me->" + var) : var;
    // Fix 081g: If For-loop var is a ByRef param (C: int32_t*), use *var for assignments/comparisons
    bool forVarIsByRef = !forVarIsClassMember && knownByRefParams_.count(lower022);
    if (forVarIsByRef) {
        varAcc = "(*" + var + ")";
    }
    if (!forVarIsClassMember && !forVarIsByRef) {
        knownLocalVars_.insert(lower022);  // Fix 010o (保留)
    }

    // P14.3.2: 嵌套循环栈 - 支持Exit For跳转到正确层
    std::string exitLabel = "vb6_loop_exit_" + std::to_string(labelCounter_++);
    loopStack_.push_back({ExitKind::For, exitLabel});

    // Fix 084o: For 循环控制变量为 Variant 时, 初始化/比较/步进必须包装.
    // VB6 语义: Variant 循环变量按 Long 语义驱动循环, 循环体内保持实际索引值.
    // 否则生成 var = 0 (int→vb6_VARIANT)、var <= end (vb6_VARIANT vs int32_t)
    // 引发 C2440 (cToolsHttp 等以 Variant 做循环变量的模块).
    bool forVarIsVariant = knownVariantVars_.count(lower022)
        || (forVarIsClassMember && classVariantMembers_.count(lower022))
        || cExprIsVariant(varAcc);
    if (forVarIsVariant) {
        // Fix 090o: 收集本 For body 内定义的标签, 供 GoTo 后缀配对判定
        forSplitLabelStack_.push_back({});
        collectForBodyLabels(node.body, forSplitLabelStack_.back());
        c_.emitLine("{");
        c_.indent();
        c_.emitLine("int32_t " + var + "_end = " + end + ";");
        c_.emitLine("int32_t " + var + "_step = " + step + ";");
        c_.emitLine(varAcc + " = vb6_VariantFromValue(" + start + ");");
        c_.emitLine("if (" + var + "_step > 0) {");
        c_.indent();
        c_.emitLine("for (; vb6_VariantToLong(" + varAcc + ") <= " + var + "_end; "
                    + varAcc + " = vb6_VariantFromValue(vb6_VariantToLong(" + varAcc + ") + " + var + "_step)) {");
        c_.indent();
        labelCopyIdx_ = 0;
        emitStmtList(node.body);
        c_.dedent();
        c_.emitLine("}");
        c_.dedent();
        c_.emitLine("} else {");
        c_.indent();
        c_.emitLine("for (; vb6_VariantToLong(" + varAcc + ") >= " + var + "_end; "
                    + varAcc + " = vb6_VariantFromValue(vb6_VariantToLong(" + varAcc + ") + " + var + "_step)) {");
        c_.indent();
        labelCopyIdx_ = 1;
        emitStmtList(node.body);
        labelCopyIdx_ = 0;
        c_.dedent();
        c_.emitLine("}");
        c_.dedent();
        c_.emitLine("}");
        c_.dedent();
        c_.emitLine("}");
        c_.emitLine(exitLabel + ":;  /* Exit For target */");

        loopStack_.pop_back();
        forSplitLabelStack_.pop_back();
        return;
    }

    // Fix 090o: 收集本 For body 内定义的标签, 供 GoTo 后缀配对判定
    forSplitLabelStack_.push_back({});
    collectForBodyLabels(node.body, forSplitLabelStack_.back());
    c_.emitLine("{");
    c_.indent();
    c_.emitLine("int32_t " + var + "_end = " + end + ";");
    c_.emitLine("int32_t " + var + "_step = " + step + ";");
    c_.emitLine(varAcc + " = " + start + ";");
    c_.emitLine("if (" + var + "_step > 0) {");
    c_.indent();
    c_.emitLine("for (; " + varAcc + " <= " + var + "_end; " + varAcc + " += " + var + "_step) {");
    c_.indent();
    labelCopyIdx_ = 0;
    emitStmtList(node.body);
    c_.dedent();
    c_.emitLine("}");
    c_.dedent();
    c_.emitLine("} else {");
    c_.indent();
    c_.emitLine("for (; " + varAcc + " >= " + var + "_end; " + varAcc + " += " + var + "_step) {");
    c_.indent();
    labelCopyIdx_ = 1;
    emitStmtList(node.body);
    labelCopyIdx_ = 0;
    c_.dedent();
    c_.emitLine("}");
    c_.dedent();
    c_.emitLine("}");
    c_.dedent();
    c_.emitLine("}");
    c_.emitLine(exitLabel + ":;  /* Exit For target */");

    loopStack_.pop_back();
    forSplitLabelStack_.pop_back();
}

void CCodeGen::visit(ForEachStmt& node) {
    // P14.3.2: 嵌套循环栈
    std::string exitLabel = "vb6_loop_exit_" + std::to_string(labelCounter_++);
    loopStack_.push_back({ExitKind::For, exitLabel});
    // P12.4: For Each item In collection
    // 支持: 数组(VB6 SafeArray)迭代
    // 暂不支持: COM集合(IEnumVARIANT)迭代

    // 检测集合是否为数组
    bool isCollArray = false;
    std::string collArrName;
    Vb6Type collElemType = Vb6Type::Variant;
    // Fix 091f: 集合是返回 SafeArray1D* 的表达式 (Split/Filter 等) — 需先物化到
    // 临时数组变量, 再走数组迭代路径. 此前一律走 COM 路径 →
    // vb6_ForEach_Init(vb6_VariantToObjectVal(vb6_Split(...))) C2440
    // (ToolsTlsThunks 754/2913/3182).
    bool collArrIsExpr = false;

    if (node.collection->kind == ASTNodeKind::IdentifierExpr) {
        auto& ident = static_cast<IdentifierExpr&>(*node.collection);
        std::string lower = ident.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);

        // 查已知数组集合
        if (knownArrays_.count(lower)) {
            isCollArray = true;
            collArrName = cIdent(ident.name);
            collElemType = arrayElemTypes_[lower];
        } else {
            // 查符号表
            Symbol* sym = symTab_.lookupModule(ident.name);
            if (sym && sym->kind == SymbolKind::Variable && sym->isArray) {
                isCollArray = true;
                collArrName = cIdent(ident.name);
                collElemType = sym->type;
            }
        }
    } else if (node.collection->kind == ASTNodeKind::IndexOrCallExpr) {
        // Fix 091f: For Each 源为内置返回 String 数组的函数 (Split/Filter) →
        // 数组迭代路径.
        auto& call091f = static_cast<IndexOrCallExpr&>(*node.collection);
        if (call091f.callee && call091f.callee->kind == ASTNodeKind::IdentifierExpr) {
            std::string cn091f = static_cast<IdentifierExpr&>(*call091f.callee).name;
            std::transform(cn091f.begin(), cn091f.end(), cn091f.begin(), ::tolower);
            if (!cn091f.empty() && cn091f.back() == '$') cn091f.pop_back();
            if (cn091f == "split" || cn091f == "filter") {
                isCollArray = true;
                collArrIsExpr = true;
                collElemType = Vb6Type::String;
            } else if (cn091f == "array") {
                // Fix 091g: Array(...) → vb6_ArrayCreate 建 Variant 数组
                // (元素为 vb6_VARIANT). For Each 源为 Array() 时此前走 COM 路径
                // → vb6_VariantToObjectVal(_arr_N) C2440 (ToolsTlsThunks 3178).
                isCollArray = true;
                collArrIsExpr = true;
                collElemType = Vb6Type::Variant;
            }
        }
    }

    // Fix 092q: For Each 遍历当前过程的 ParamArray 参数时, C 侧类型是 SAFEARRAY*
    // (cgen 的参数映射), 而 vb6_LBound/vb6_UBound/VB6_SA_AT 都要求 vb6_SafeArray1D*
    // (`(arr)->data` / `->lBound`) → 对 tagSAFEARRAY 报 C2039 且上下界取值语义错
    // (mdTlsThunks.bas `For Each vElem In a` (a 是 ParamArray), ToolsTlsThunks 5468).
    bool isPA092q = false;
    if (node.collection && node.collection->kind == ASTNodeKind::IdentifierExpr && currentProc_) {
        std::string paLower092q =
            Symbol::toLower(static_cast<IdentifierExpr&>(*node.collection).name);
        for (auto& p092q : currentProc_->params) {
            if (Symbol::toLower(p092q.name) == paLower092q) {
                isPA092q = p092q.isParamArray;
                break;
            }
        }
    }

    std::string var = cIdent(node.varName);
    int tmpIdx = tempCounter_++;

    // Fix 022: ForEach 循环变量若为类成员必须以 me->var 形式 emit (同 ForStmt).
    // 原 Fix 010o 无差别注册循环变量到 knownLocalVars_, 类成员场景会导致 C2065
    // (详见 ForStmt 注释). 此处同步修复.
    std::string lower022fe = node.varName;
    std::transform(lower022fe.begin(), lower022fe.end(), lower022fe.begin(), ::tolower);
    bool forEachVarIsClassMember = isClassModule_
        && classMemberVars_.count(lower022fe)
        && !knownLocalVars_.count(lower022fe);  // 局部 Dim 优先遮蔽
    std::string varAcc = forEachVarIsClassMember ? ("me->" + var) : var;
    if (!forEachVarIsClassMember) {
        knownLocalVars_.insert(lower022fe);  // Fix 010o (保留)
    }

    if (isCollArray) {
        // For Each item In arr -> SafeArray index iteration
        // {
        //     int32_t _fe_i0, _fe_lb0, _fe_ub0;
        //     _fe_lb0 = vb6_LBound(arr, 1);
        //     _fe_ub0 = vb6_UBound(arr, 1);
        //     for (_fe_i0 = _fe_lb0; _fe_i0 <= _fe_ub0; _fe_i0++) {
        //         vb6_item = VB6_SA_AT(VARIANT, arr, _fe_i0);
        //         // body
        //     }
        // }
        c_.emitLine("{");
        c_.indent();
        std::string idxVar = "_fe_i" + std::to_string(tmpIdx);
        std::string lbVar = "_fe_lb" + std::to_string(tmpIdx);
        std::string ubVar = "_fe_ub" + std::to_string(tmpIdx);
        c_.emitLine("int32_t " + idxVar + ", " + lbVar + ", " + ubVar + ";");
        if (isPA092q) {
            // Fix 092q: ParamArray 取值经 vb6_VariantFromStackVARIANT 转成
            // vb6_VARIANT (ToolsTlsThunks 5468 曾因直接赋值报 C2440)。
            // Fix <vbeclipse> rev37: vb6_PA_GetVariant 已改返回 vb6_VARIANT
            // (与生成码的 _Generic 装箱同口径), 槽位随之改声明; 否则
            // `VARIANT = vb6_VARIANT` → C2440 (Charts 2020 Form2.c 的
            // `For Each v In SomeParamArray` 首个命中)。
            c_.emitLine("vb6_VARIANT _pa_v" + std::to_string(tmpIdx) + " = {0};");
        }
        // Fix 091f: 表达式集合物化 — 避免 LBound/UBound/元素取值重复调用
        // Split/Filter (会每次重新分配数组).
        std::string arrRef091f = collArrName;
        if (collArrIsExpr) {
            emitExpr(*node.collection);
            std::string tmpArr091f = "_fe_arr" + std::to_string(tmpIdx);
            c_.emitLine("vb6_SafeArray1D* " + tmpArr091f + " = " + lastExpr_ + ";");
            arrRef091f = tmpArr091f;
        }
        if (isPA092q) {
            // Fix 092q: ParamArray 用 PA 专用上下界 (SafeArrayGetElement 语义)
            c_.emitLine(lbVar + " = vb6_PA_LBound(" + arrRef091f + ");");
            c_.emitLine(ubVar + " = vb6_PA_UBound(" + arrRef091f + ");");
        } else {
            c_.emitLine(lbVar + " = vb6_LBound(" + arrRef091f + ", 1);");
            c_.emitLine(ubVar + " = vb6_UBound(" + arrRef091f + ", 1);");
        }
        c_.emitLine("for (" + idxVar + " = " + lbVar + "; " + idxVar + " <= " + ubVar + "; " + idxVar + "++) {");
        c_.indent();

        // 赋值循环变量: vb6_item = VB6_SA_AT(elemCType, arr, _fe_i0)
        // Fix 084o: 循环变量为 Variant 时需 vb6_VariantFromValue 包装 (String 数组
        // 元素是 BSTR, 直接赋给 vb6_VARIANT 触发 C2440, 如 cToolsHttp 的 Pair As Variant)
        std::string elemCType = mapSaElemCType(collElemType);
        bool feVarIsVariant = knownVariantVars_.count(lower022fe)
            || (forEachVarIsClassMember && classVariantMembers_.count(lower022fe))
            || cExprIsVariant(varAcc);
        if (isPA092q) {
            // Fix 092q: ParamArray 元素用 SafeArrayGetElement 直取 (索引与
            // vb6_PA_LBound/UBound 同域), 按循环变量/元素类型选择解包函数.
            std::string paGet092q = "vb6_PA_GetVariant(" + arrRef091f + ", " + idxVar + ")";
            std::string paTmp092q = "_pa_v" + std::to_string(tmpIdx);
            if (feVarIsVariant) {
                c_.emitLine(paTmp092q + " = " + paGet092q + ";");
                // Fix 176: ParamArray 的 For Each 元素在栈上 (VARIANT 槽位),
                // vb6_VariantFromComResult 尾部会 free(pv) 释放堆宿主 VARIANT,
                // 对栈地址 free → 堆损坏 (0xC0000374). 栈上源用
                // vb6_VariantFromStackVARIANT (P24-03, 不释放源).
                // rev37: 槽位已是 vb6_VARIANT 而该函数收 VARIANT* —— 同布局,
                // 显式转回指针 (改 RTL 签名会波及 vb6rtl_com.c:396 的内部调用).
                c_.emitLine(varAcc + " = vb6_VariantFromStackVARIANT((VARIANT*)&" + paTmp092q + ");");
            } else if (elemCType == "BSTR") {
                c_.emitLine(varAcc + " = vb6_PA_GetBSTR(" + arrRef091f + ", " + idxVar + ");");
            } else if (elemCType == "double") {
                c_.emitLine(varAcc + " = vb6_PA_GetDouble(" + arrRef091f + ", " + idxVar + ");");
            } else {
                c_.emitLine(varAcc + " = vb6_PA_GetLong(" + arrRef091f + ", " + idxVar + ");");
            }
        } else if (feVarIsVariant) {
            c_.emitLine(varAcc + " = vb6_VariantFromValue(VB6_SA_AT(" + elemCType + ", "
                        + arrRef091f + ", " + idxVar + "));");
        } else {
            c_.emitLine(varAcc + " = VB6_SA_AT(" + elemCType + ", " + arrRef091f + ", " + idxVar + ");");
        }

        emitStmtList(node.body);
        c_.dedent();
        c_.emitLine("}");
        c_.dedent();
        c_.emitLine("}");
    } else {
        // P22-11: COM collection For Each via IEnumVARIANT
        // Generate:
        //   {
        //       void* _fe_enum_N = vb6_ForEach_Init(coll);
        //       VARIANT _fe_var_N;
        //       if (_fe_enum_N) {
        //           while (vb6_ForEach_Next(_fe_enum_N, (void*)&_fe_var_N)) {
        //               var = _fe_var_N;
        //               // body
        //               vb6_ComVarClear((void*)&_fe_var_N);
        //           }
        //           vb6_ForEach_Release(_fe_enum_N);
        //       }
        //   }
        c_.emitLine("{");
        c_.indent();
        std::string enumVar = "_fe_enum_" + std::to_string(tmpIdx);
        std::string feVar = "_fe_var_" + std::to_string(tmpIdx);
        emitExpr(*node.collection);
        // P24-05: 物化COM标记 — For Each的集合表达式可能是COM成员访问(如fso.Drives)
        // COM标记设置了comObjExpr_/comMemberName_但lastExpr_只有基础对象名
        if (isComMarker_) {
            resolveComValue("Object");  // 集合必须是Object(IDispatch*), 生成ComGetObjectProp
        }
        std::string collExpr = lastExpr_;
        // Fix 091q: vb6_ComCall/vb6_ComCallObject 已返回对象指针 (void*),
        // vb6_ForEach_Init 直接接收. 不可再按 Variant 提取 — cLang 26:
        // For Each x In me.LangInfo.Item("LangList") (COM Item 调用, 生成
        // vb6_ComCall(...)) 曾被 isDefinitelyVariantExpr 判为 Variant →
        // vb6_VariantToObjectVal(vb6_ComCall(...)) C2440 (void* → vb6_VARIANT).
        bool collIsObjPtr091q = collExpr.rfind("vb6_ComCall(", 0) == 0
                             || collExpr.rfind("vb6_ComCallObject(", 0) == 0;
        // Fix 040c: vb6_ForEach_Init expects void* (IDispatch*). If the collection
        // expression is a Variant (vb6_VARIANT struct), extract the object pointer.
        if (!collIsObjPtr091q && cExprIsVariant(collExpr)) {
            collExpr = "vb6_VariantToObjectVal(" + collExpr + ")";
        } else if (node.collection && node.collection->kind == ASTNodeKind::IdentifierExpr) {
            auto& ident = static_cast<IdentifierExpr&>(*node.collection);
            std::string lower = ident.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (knownVariantVars_.count(lower)) {
                collExpr = "vb6_VariantToObjectVal(" + collExpr + ")";
            }
        }
        // Fix 040c: fallback — project class methods returning Variant (e.g. Dictionary.Keys)
        else if (!collIsObjPtr091q && node.collection && isDefinitelyVariantExpr(*node.collection)) {
            collExpr = "vb6_VariantToObjectVal(" + collExpr + ")";
        }
        c_.emitLine("void* " + enumVar + " = vb6_ForEach_Init(" + collExpr + ");");
        c_.emitLine("VARIANT " + feVar + ";");
        c_.emitLine("if (" + enumVar + ") {");
        c_.indent();
        c_.emitLine("while (vb6_ForEach_Next(" + enumVar + ", (void*)&" + feVar + ")) {");
        c_.indent();
        // P24-05: For Each循环变量类型感知赋值
        // Object类型 → vb6_VariantToObject提取IDispatch*
        // Variant类型 → vb6_VariantFromStackVARIANT转换为vb6_VARIANT
        {
            std::string lower = node.varName;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            // vbeclipse: knownVariantVars_ 必须**先判**. knownObjectVars_ 跨过程不清空
            // (Fix 157f), 而同名变量在不同过程里可以声明成不同类型 — ucTabStrip.ctl
            // 前两个过程 `Dim l_Tab As Object` (243/274) 把 l_tab 记进 knownObjectVars_,
            // 后面 `Dim l_Tab As Variant` (322/373) 的 For Each 于是走 Object 分支发
            // `l_Tab = vb6_ComUnpackObject(&fe_var)` (void*) → 赋给 vb6_VARIANT → C2440
            // (ucTabStrip.c 301/353/394). 本过程的 Dim 已把名字登记进 knownVariantVars_,
            // C 侧声明就是 vb6_VARIANT, 故先判它.
            if (knownVariantVars_.count(lower)) {
                c_.emitLine(varAcc + " = vb6_VariantFromStackVARIANT(&" + feVar + ");  /* P24-05: For Each Variant: VARIANT→vb6_VARIANT */");
            } else if (knownObjectVars_.count(lower) || knownTypedComVars_.count(lower)) {
                c_.emitLine(varAcc + " = vb6_ComUnpackObject(&" + feVar + ");  /* P24-05: For Each Object: VARIANT→IDispatch* */");
            } else if (knownClassVars_.count(lower)) {
                // Fix 090aa: 项目类变量 (vb6_cls_X*) 作 For Each 元素 — 元素是
                // COM VARIANT(IDispatch*), 提取后 cast 到类指针. 此前漏查 knownClassVars_
                // → 走 Variant 分支 oClient = vb6_VARIANT → C2440.
                c_.emitLine(varAcc + " = (vb6_cls_" + cIdent(knownClassVars_[lower])
                            + "*)vb6_ComUnpackObject(&" + feVar + ");  /* P24-05: For Each 项目类: VARIANT→IDispatch*→类指针 */");
            } else {
                c_.emitLine(varAcc + " = vb6_VariantFromStackVARIANT(&" + feVar + ");  /* P24-05: For Each Variant: VARIANT→vb6_VARIANT */");
            }
        }
        emitStmtList(node.body);
        c_.emitLine("VariantClear(&" + feVar + ");  /* P24-03: 栈VARIANT只清不清 */");
        c_.dedent();
        c_.emitLine("}");
        c_.emitLine("vb6_ForEach_Release(" + enumVar + ");");
        c_.dedent();
        c_.emitLine("}");
        c_.dedent();
        c_.emitLine("}");
    }
    c_.emitLine(exitLabel + ":;  /* Exit For target */");

    loopStack_.pop_back();
}



} // namespace vb6c3
