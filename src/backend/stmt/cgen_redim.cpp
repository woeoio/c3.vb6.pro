#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_redim.cpp: ReDim/Erase 数组语句生成 ---

// Fix 180: `ReDim h1.Items(0 To 1)`（owner 不带下标的点号链）此前**两头落空** ——
// 既不进 Fix 100 的 targetExpr 通道（那条只在链上带下标时才有 targetExpr），而
// 下面 :74-79 的裸名回退按 `arrayUdtElemTypes_` 查，键是裸变量名，"h1.items" 永远
// 捡不到 → isUdtArray=false → 发 vb6_SafeArrayReDim1D(vb6_sa_variant, …) 按 16 字节
// 元素分配，读写侧却用 VB6_SA_AT(vb6_type_TOwned, …) → 步长不符越界（现象：整棵
// 子树别名 + 字符串为空）。
// owner 的 C 类型名由 knownUdtVars_ 提供（`Dim h1 As THost` 在 cgen_decl_var.cpp /
// cgen_localdecl.cpp 登记），成员类型复用现成的 udtFieldObjCType —— 它看
// `mi.type == UserDefinedType` + `typeRefName`，不看 isArrayDynamic，正合
// `Items() As TOwned` 的形态（数组性记在 arraySize/isArrayDynamic 上，As 后面是
// 普通 SimpleTypeRef，所以 typeRefName 是有记录的）。
// 边界：owner 是类字段（me->Foo.Items）或 With 成员（.Items）时 knownUdtVars_ 里没有
// owner，返回空串让调用方**按原样回落 Variant**（Variant 元素尺寸最大 → 过分配不会
// 越界），不去猜。
std::string CCodeGen::resolveUdtMemberArrayCType(const std::string& lowerVar) const {
    size_t dot = lowerVar.rfind('.');
    if (dot == std::string::npos || dot == 0 || dot + 1 >= lowerVar.size()) return "";
    auto it = knownUdtVars_.find(lowerVar.substr(0, dot));
    if (it == knownUdtVars_.end()) return "";
    return udtFieldObjCType(it->second, lowerVar.substr(dot + 1));
}

void CCodeGen::visit(ReDimStmt& node) {
    // Fix 100: 复杂目标 (带下标的成员链, 如 ReDim m_Serie(i).PT(n))。
    // node.targetExpr 非空时直接发射左值表达式 (VB6_SA_AT(vb6_type_tSerie, m_Serie, i).PT),
    // 元素类型改由末段成员的 UDT 成员信息解析 — 不走下面的 varName 字符串展开,
    // 因为 cIdent 会把 '.' 替换为 '_' 且无法表达下标。
    if (node.targetExpr) {
        emitReDimComplexTarget(node);
        return;
    }
    // Fix 084y-5: ReDim 目标含成员访问 (ByRef UDT 参数数组字段 uOutput.Buffer,
    // With 块成员 .Field) 时按成员访问展开, 避免 cIdent 把 '.' 替换成 '_'
    std::string cName = resolveArrayTargetIdent(node.varName);
    // Fix 010r: Add me-> prefix for class member arrays
    std::string lowerVar = node.varName;
    std::transform(lowerVar.begin(), lowerVar.end(), lowerVar.begin(), ::tolower);
    // Fix 086: ReDim 目标是当前 Function/PropertyGet 自身名 → 返回值变量
    // (VB6: 数组返回函数内 ReDim Preserve FuncName(...) 重设返回数组.
    //  此前生成裸函数名 → C2065, 如 cMemoryStream.Contents / cImage.PictureToBytes)
    if (currentProc_ && !currentReturnVar_.empty()
        && (currentProc_->kind == SymbolKind::Function
            || currentProc_->kind == SymbolKind::PropertyGet)
        && Symbol::toLower(currentProc_->name) == lowerVar
        && !knownLocalVars_.count(lowerVar)) {
        cName = currentReturnVar_;
        lowerVar = currentReturnVar_;
    }
    if (isClassModule_ && classMemberVars_.count(lowerVar) && !knownLocalVars_.count(lowerVar)) {
        cName = "me->" + cName;
    }
    // Fix 010r-6 rev2: ByRef array param in ReDim needs (*name) since it's vb6_SafeArray1D**
    // Fix 084o-6: ByRef Variant 参数也是 vb6_VARIANT*, 同样需要 (*name)
    // (VB6 允许 As Variant 参数后接 ReDim 变数组, 如 cZipArchive.Extract 的
    // OutputTarget As Variant → ReDim OutputTarget(...) As Byte)
    if (currentProc_) {
        for (auto& param : currentProc_->params) {
            std::string paramLower = param.name;
            std::transform(paramLower.begin(), paramLower.end(), paramLower.begin(), ::tolower);
            if (paramLower == lowerVar && !param.isByVal
                && ((static_cast<uint16_t>(param.type) & static_cast<uint16_t>(Vb6Type::Array))
                    || param.type == Vb6Type::Variant)) {
                cName = "(*" + cName + ")";
                break;
            }
        }
    }
    // Fix 061: With块内 ReDim .Data(...) → _vb6_with_N.Data
    // parser 在 varName 前加了 '.' 前缀表示 With 成员引用 (e.g. ".Data")
    if (node.varName.size() > 1 && node.varName[0] == '.'
        && !withObjectVars_.empty() && !withObjectInfoStack_.empty()) {
        std::string memberName = node.varName.substr(1);  // strip leading '.'
        const auto& info = withObjectInfoStack_.back();
        if (info.kind == WithObjKind::Unknown) {
            // Fix 081j-2: UDT With block 临时变量是指针，用 -> 访问成员
            cName = withObjectVars_.back() + "->" + cIdent(memberName);
        } else if (info.kind == WithObjKind::ClassInstance) {
            // Class With block: _vb6_with_N->member
            cName = withObjectVars_.back() + "->" + cIdent(memberName);
        }
    }
    Vb6Type elemType = resolveArrayElemType(node.asType.get());
    // Fix 192: `ReDim a(1) As Circle` 的元素类名登记 (与声明侧 `Dim a() As C` 对称).
    // 这是本缺口**最初被发现的形状** (`ReDim ... As <类名>`), 但实测静态数组
    // `Dim s(1) As C` 一样坏 —— 根因在访问侧推断, 不在 ReDim。此处登记是让
    // ReDim 首次确立元素类型时也生效 (VB6 允许 ReDim 带 As 覆盖/补声明)。
    // 只认裸变量名: 成员链 (obj.Field) / With 成员 (.Data) 的键形态与访问侧
    // 查表的键对不上, 猜错比不猜更坏 —— 沿用 Fix 180 的"不去猜"纪律。
    {
        std::string cls = resolveArrayClassElemType(node.asType.get());
        if (!cls.empty() && !lowerVar.empty()
            && lowerVar.find('.') == std::string::npos) {
            arrayClassElemTypes_[lowerVar] = cls;
        }
    }
    std::string saElemType = mapSaElemType(elemType);
    // Bug4-Fix: UDT数组需使用vb6_SafeArrayReDim1D_Udt
    std::string udtCType = resolveArrayUdtElemCType(node.asType.get());
    if (udtCType.empty() && !lowerVar.empty()) {
        // ReDim Preserve arr(n) 不带 As 子句时, 元素类型由声明决定
        // (Dim arr() As UDT)。不回退则回落 vb6_sa_empty(4字节) → UDT 越界写入。
        auto it = arrayUdtElemTypes_.find(lowerVar);
        if (it != arrayUdtElemTypes_.end()) udtCType = it->second;
    }
    if (udtCType.empty()) {
        // Fix 180: 裸名回退漏掉的点号链目标 (ReDim h1.Items(0 To 1) 不带 As)
        udtCType = resolveUdtMemberArrayCType(lowerVar);
    }
    bool isUdtArray = !udtCType.empty();

    // Fix 084a/090m: ReDim 目标是否为 Variant 数组 — 逻辑见成员函数 isVariantArrayTarget
    auto isVariantArrayVar = [&](const std::string& nm) -> bool { return isVariantArrayTarget(nm); };

    if (node.dimensions.empty()) return;

    int dimCount = (int)node.dimensions.size();

    if (dimCount == 1) {
        // 涓€缁?ReDim (淇濇寔鍘熸湁1D浠ｇ爜)
        auto& dim = node.dimensions[0];
        std::string lBound = "0";
        std::string uBound = "0";
        if (dim.lower) {
            emitExpr(*dim.lower);
            lBound = std::move(lastExpr_);
        }
        if (dim.upper) {
            emitExpr(*dim.upper);
            uBound = std::move(lastExpr_);
        }

        if (node.preserve) {
            // Fix 084a: Variant 数组 ReDim Preserve — 实参提取 SafeArray, 结果包装回 Variant
            std::string callArg = cName;
            std::string assignVal;
            if (isVariantArrayVar(cName)) {
                callArg = "vb6_VariantToSafeArray1D(" + cName + ")";
                // 显式传 vb6_sa_variant (Fix rev3): 这是本分支的**权威**类型 ——
                // 不传就得靠运行时猜, 而猜错的后果是槽位步长错位(见 else 支注释)。
                assignVal = "vb6_VariantFromValue(vb6_SafeArrayReDimPreserve1D_T(vb6_sa_variant, " + callArg + ", " + lBound + ", " + uBound + "))";
            } else if (isUdtArray) {
                // Bug4-Fix Udt Preserve: 传入 UDT 元素尺寸, 避免 NULL 初值回落 4 字节导致越界
                assignVal = "vb6_SafeArrayReDimPreserve1D_Udt((int32_t)sizeof(" + udtCType + "), " + callArg + ", " + lBound + ", " + uBound + ")";
            } else {
                // Fix <vbeclipse> rev3: **必须把 saElemType 传下去**。旧的
                // 三参入口拿不到元素类型, 只能按 variant 兜底 —— 于是
                // `Dim m_Keys() As String` 首次 ReDim 按 16/24 字节步长分配,
                // 而 VB6_SA_AT(BSTR,…) 按 4 字节步进, 槽里压着别的 BSTR,
                // 销毁时再被当 bstrVal 二次 free (实测 play78.exe 的 List.Contains)。
                // saElemType 来自 mapSaElemType(elemType), 与非 preserve 分支
                // (下方 vb6_SafeArrayReDim1D(saElemType, …)) 同一口径。
                assignVal = "vb6_SafeArrayReDimPreserve1D_T(" + saElemType + ", " + callArg + ", " + lBound + ", " + uBound + ")";
            }
            c_.emitLine(cName + " = " + assignVal + ";");
        } else {
            // Fix 084a: Variant 数组非 preserve ReDim — 先销毁提取出的 SafeArray, 再包装新数组回 Variant
            std::string destroyArg = cName;
            if (isVariantArrayVar(cName)) destroyArg = "vb6_VariantToSafeArray1D(" + cName + ")";
            c_.emitLine("vb6_SafeArrayDestroy1D(" + destroyArg + ");");
            std::string newVal;
            if (isUdtArray) {
                // Bug4-Fix: UDT数组使用_Udt版本，传入sizeof(UDT类型)
                newVal = "vb6_SafeArrayReDim1D_Udt((int32_t)sizeof(" + udtCType + "), " + lBound + ", " + uBound + ")";
            } else {
                newVal = "vb6_SafeArrayReDim1D(" + saElemType + ", " + lBound + ", " + uBound + ")";
            }
            if (isVariantArrayVar(cName)) newVal = "vb6_VariantFromValue(" + newVal + ")";
            c_.emitLine(cName + " = " + newVal + ";");
        }
    } else {
        // P8.1: 多维 ReDim
        // Use raw variable name for boundsVar (must be a valid C identifier)
        std::string rawCName = cIdent(node.varName);
        if (isClassModule_ && classMemberVars_.count(lowerVar) && !knownLocalVars_.count(lowerVar)) {
            rawCName = "me_" + rawCName;
        }
        std::string boundsVar = "_redim_bounds_" + rawCName;
        c_.emitLine("vb6_SafeArrayBound " + boundsVar + "[] = {");
        c_.indent();
        for (int d = 0; d < dimCount; d++) {
            auto& dim = node.dimensions[d];
            std::string lb = "0", ub = "0";
            if (dim.lower) { emitExpr(*dim.lower); lb = std::move(lastExpr_); }
            if (dim.upper) { emitExpr(*dim.upper); ub = std::move(lastExpr_); }
            std::string trailing = (d < dimCount - 1) ? "," : "";
            // vb6_SafeArrayBound = {lLbound, cElements}
            // cElements = uBound - lBound + 1 (VB6 "0 To 3" has 4 elements)
            c_.emitLine("{" + lb + ", (" + ub + " - " + lb + " + 1)}" + trailing);
        }
        c_.dedent();
        c_.emitLine("};");
        // Fix 084a: Variant 多维数组 ReDim — 实参提取 SafeArray, 结果包装回 Variant
        std::string redimArg = cName;
        std::string wrapBack = "";
        if (isVariantArrayVar(cName)) {
            redimArg = "vb6_VariantToSafeArray1D(" + cName + ")";
            wrapBack = "vb6_VariantFromValue(";
        }
        if (node.preserve) {
            std::string res;
            if (isUdtArray) {
                // Bug4-Fix Udt Preserve: 多维 UDT 数组传入元素尺寸
                res = "vb6_SafeArrayReDimPreserveND_Udt((int32_t)sizeof(" + udtCType + "), (vb6_SafeArrayND*)" + redimArg + ", " + std::to_string(dimCount) + ", " + boundsVar + ")";
            } else {
                res = "vb6_SafeArrayReDimPreserveND((vb6_SafeArrayND*)" + redimArg + ", " + std::to_string(dimCount) + ", " + boundsVar + ")";
            }
            if (!wrapBack.empty()) res = wrapBack + res + ")";
            c_.emitLine(cName + " = (vb6_SafeArray1D*)" + res + ";");
        } else {
            std::string destroyArg = redimArg;
            c_.emitLine("vb6_SafeArrayDestroyND((vb6_SafeArrayND*)" + destroyArg + ");");
            std::string newVal;
            if (isUdtArray) {
                // Bug4-Fix: UDT多维数组使用_Udt版本，传入sizeof(UDT类型)
                newVal = "vb6_SafeArrayReDimND_Udt((int32_t)sizeof(" + udtCType + "), " + std::to_string(dimCount) + ", " + boundsVar + ")";
            } else {
                newVal = "vb6_SafeArrayReDimND(" + saElemType + ", " + std::to_string(dimCount) + ", " + boundsVar + ")";
            }
            if (!wrapBack.empty()) newVal = wrapBack + newVal + ")";
            c_.emitLine(cName + " = (vb6_SafeArray1D*)" + newVal + ";");
        }

        // 鏇存柊鏁扮粍缁村害淇℃伅
        std::string lower = node.varName;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        arrayDimCounts_[lower] = dimCount;
    }
}

// Fix 100: 解析 ReDim 复杂目标 (arr(i).Field) 的元素类型。
// targetExpr 的末段必须是 MemberAccessExpr; 其 object 的 UDT C 类型由
// inferUdtTypeOfExpr 推断 (UDT 数组元素 arr(idx) 走 arrayUdtElemTypes_),
// 再在 udtMembers 中查该成员的标量与 typeRefName — 与 cgen_expr_call.cpp
// Pattern B (动态数组成员元素访问 VB6_SA_AT) 使用同一来源, 保证二者一致。
void CCodeGen::resolveReDimComplexElemType(const ReDimStmt& node, Vb6Type& outType,
                                           std::string& outUdtCType) const {
    outType = Vb6Type::Variant;
    outUdtCType.clear();
    if (!node.targetExpr) return;
    if (node.targetExpr->kind != ASTNodeKind::MemberAccessExpr) return;
    auto& ma = static_cast<const MemberAccessExpr&>(*node.targetExpr);
    if (!ma.object) return;

    std::string ownerCType = inferUdtTypeOfExpr(*ma.object);
    const std::string prefix = "vb6_type_";
    if (ownerCType.size() <= prefix.size()
        || ownerCType.compare(0, prefix.size(), prefix) != 0) return;
    std::string udtName = ownerCType.substr(prefix.size());

    Symbol* udtSym = symTab_.lookupModule(udtName);
    if (!udtSym) udtSym = symTab_.lookup(udtName);
    if (!udtSym || udtSym->kind != SymbolKind::UserDefinedType) return;

    std::string memLower = Symbol::toLower(ma.memberName);
    std::string memLowerM = "m_" + memLower;
    for (auto& mi : udtSym->udtMembers) {
        std::string miLower = Symbol::toLower(mi.name);
        if (miLower != memLower && miLower != memLowerM) continue;
        outType = mi.type;
        if (mi.type == Vb6Type::UserDefinedType && !mi.typeRefName.empty()) {
            outUdtCType = "vb6_type_" + cIdent(mi.typeRefName);
        }
        return;
    }
}

// Fix 100: 发射 ReDim 复杂目标 — ReDim arr(i).Field(dims) / obj.List(j).Field(dims)。
// 目标左值直接由 emitExpr(*targetExpr) 发射: UDT 数组元素 arr(i) 生成
// VB6_SA_AT(vb6_type_<UDT>, arr, i), 再串接 .Field。
// UDT 动态数组成员在 C 结构体中就是 vb6_SafeArray1D* (cgen_decl.cpp TypeDecl),
// 因此可直接对它做 Destroy/ReDim 赋值, 无需 Variant 包装分支。
void CCodeGen::emitReDimComplexTarget(ReDimStmt& node) {
    emitExpr(*node.targetExpr);
    std::string cName = std::move(lastExpr_);

    Vb6Type elemType = Vb6Type::Variant;
    std::string udtCType;
    resolveReDimComplexElemType(node, elemType, udtCType);
    // 目标显式写了 As Type 时以 As 子句为准
    if (node.asType) {
        elemType = resolveArrayElemType(node.asType.get());
        udtCType = resolveArrayUdtElemCType(node.asType.get());
    }
    if (udtCType.empty() && !node.varName.empty()) {
        std::string cv = node.varName;
        std::transform(cv.begin(), cv.end(), cv.begin(), ::tolower);
        auto it = arrayUdtElemTypes_.find(cv);
        if (it != arrayUdtElemTypes_.end()) udtCType = it->second;
        // Fix 180: 与 visit(ReDimStmt&) 共用同一套点号链解析，避免第三份实现
        if (udtCType.empty()) udtCType = resolveUdtMemberArrayCType(cv);
    }
    if (udtCType.empty() && elemType == Vb6Type::UserDefinedType) {
        // 推断不到 UDT 的 C 类型名: 回落 Variant 分配 (VB6_SA_AT 只用 data/lBound,
        // Variant 元素尺寸最大 → 过分配不会越界), 避免 sizeof(未定义类型)
        elemType = Vb6Type::Variant;
    }
    std::string saElemType = mapSaElemType(elemType);
    bool isUdtArray = !udtCType.empty();

    if (node.dimensions.empty()) return;
    int dimCount = (int)node.dimensions.size();

    if (dimCount == 1) {
        auto& dim = node.dimensions[0];
        std::string lBound = "0";
        std::string uBound = "0";
        if (dim.lower) { emitExpr(*dim.lower); lBound = std::move(lastExpr_); }
        if (dim.upper) { emitExpr(*dim.upper); uBound = std::move(lastExpr_); }

        if (node.preserve) {
            if (isUdtArray) {
                // Bug4-Fix Udt Preserve: 传入 UDT 元素尺寸
                c_.emitLine(cName + " = vb6_SafeArrayReDimPreserve1D_Udt((int32_t)sizeof(" + udtCType + "), " + cName + ", "
                          + lBound + ", " + uBound + ");");
            } else {
                // Fix <vbeclipse> rev3: 与下方非 preserve 分支 / visit(ReDimStmt&) 的
                // 1D 路径**同一口径** —— 元素类型由 codegen 权威给出 (mapSaElemType),
                // 不让运行时回落 vb6_sa_variant 猜 (猜错 = 槽步长错位 + BSTR 二次 free,
                // 见 vb6rtl_array.c 里 ReDimPreserve1D_T 的注释)。
                c_.emitLine(cName + " = vb6_SafeArrayReDimPreserve1D_T(" + saElemType + ", " + cName + ", "
                          + lBound + ", " + uBound + ");");
            }
        } else {
            c_.emitLine("vb6_SafeArrayDestroy1D(" + cName + ");");
            std::string newVal = isUdtArray
                ? ("vb6_SafeArrayReDim1D_Udt((int32_t)sizeof(" + udtCType + "), "
                   + lBound + ", " + uBound + ")")
                : ("vb6_SafeArrayReDim1D(" + saElemType + ", " + lBound + ", " + uBound + ")");
            c_.emitLine(cName + " = " + newVal + ";");
        }
    } else {
        // 多维 ReDim: bounds 局部数组名由目标左值净化得到 (仅需 C 标识符合法)
        std::string boundsVar = "_redim_bounds";
        for (char ch : cName) {
            if (isalnum((unsigned char)ch) || ch == '_') boundsVar += ch;
        }
        c_.emitLine("vb6_SafeArrayBound " + boundsVar + "[] = {");
        c_.indent();
        for (int d = 0; d < dimCount; d++) {
            auto& dim = node.dimensions[d];
            std::string lb = "0", ub = "0";
            if (dim.lower) { emitExpr(*dim.lower); lb = std::move(lastExpr_); }
            if (dim.upper) { emitExpr(*dim.upper); ub = std::move(lastExpr_); }
            std::string trailing = (d < dimCount - 1) ? "," : "";
            c_.emitLine("{" + lb + ", (" + ub + " - " + lb + " + 1)}" + trailing);
        }
        c_.dedent();
        c_.emitLine("};");
        c_.emitLine("vb6_SafeArrayDestroyND((vb6_SafeArrayND*)" + cName + ");");
        std::string newVal = isUdtArray
            ? ("vb6_SafeArrayReDimND_Udt((int32_t)sizeof(" + udtCType + "), "
               + std::to_string(dimCount) + ", " + boundsVar + ")")
            : ("vb6_SafeArrayReDimND(" + saElemType + ", " + std::to_string(dimCount)
               + ", " + boundsVar + ")");
        c_.emitLine(cName + " = (vb6_SafeArray1D*)" + newVal + ";");
    }
}

void CCodeGen::visit(EraseStmt& node) {
    for (size_t k = 0; k < node.varNames.size(); ++k) {
        auto& name = node.varNames[k];
        // 账 #186: 带下标的成员链目标 (`Erase m_tvFiles(lIndex).bvData`) 由 parser 备好表达式,
        // 这里发左值 —— 与 ReDim 的复杂目标同一套机制 (emitExpr → lastExpr_)。
        Expr* complex = (k < node.targets.size()) ? node.targets[k].get() : nullptr;
        std::string cName;
        if (complex) {
            emitExpr(*complex);
            cName = lastExpr_;
        } else {
            // Fix 084y-5: Erase 目标含成员访问 (With 块成员 .Field, ByRef UDT 参数
            // 数组字段) 时按成员访问展开, 避免 cIdent 把 '.' 替换成 '_' 生成
            // 未声明的单标识符 (_RemoteLegacyNextTrafficKey / uOutput_Buffer → C2065)
            cName = resolveArrayTargetIdent(name);
            // Fix 010r: Add me-> prefix for class member arrays
            std::string lower = name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (isClassModule_ && classMemberVars_.count(lower) && !knownLocalVars_.count(lower)) {
                cName = "me->" + cName;
            }
        }
        // Fix 090p: UDT 的 As Variant 数组字段 (如 (*uFile).BufferArray) Erase —
        // 裸 vb6_SafeArrayDestroy1D((*uFile).BufferArray) 把 VARIANT 当 SafeArray* →
        // C2440; 用 vb6_VariantClear 释放数组并置 VT_EMPTY (cZipArchive pvVfsSetEof).
        // 复杂目标同样按**去下标的点链名**问这一条 (name 里存的就是它) ⇒ 判据只有一处。
        if (isVariantArrayTarget(name)) {
            c_.emitLine("vb6_VariantClear(&" + cName + ");");
            continue;
        }
        if (complex) {
            // UDT 成员的动态数组**恒为一维** (结构体发码单点: cgen_decl.cpp 的
            // `vb6_SafeArray1D* <成员>; /* dynamic array member */`) ⇒ 销毁只走 1D 那一支。
            c_.emitLine("vb6_SafeArrayDestroy1D(" + cName + "); " + cName + " = NULL;");
            continue;
        }
        // P8.1: 根据维度数选择1D/ND销毁
        std::string lowerName = name;
        std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::tolower);
        auto it = arrayDimCounts_.find(lowerName);
        if (it != arrayDimCounts_.end() && it->second > 1) {
            c_.emitLine("vb6_SafeArrayDestroyND((vb6_SafeArrayND*)" + cName + "); " + cName + " = NULL;");
        } else {
            c_.emitLine("vb6_SafeArrayDestroy1D(" + cName + "); " + cName + " = NULL;");
        }
    }
}


} // namespace vb6c3
