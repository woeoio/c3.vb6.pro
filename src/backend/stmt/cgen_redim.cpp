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
    std::string ctype = udtFieldObjCType(it->second, lowerVar.substr(dot + 1));
    // 账 #475 §B220: `Names() As String` 的成员 C 类型是 BSTR, 但 BSTR **不是 UDT** ——
    // 这一格不能交给 _Udt 分配 (那样描述符记成 vb6_sa_udt, Erase/Destroy 不释放元素,
    // VarType 也答 8204 而不是 vb6 的 8200|vtBSTR)。元素档由 resolveReDimElemType
    // 给出的 String → vb6_sa_bstr 才是权威。
    if (ctype == "BSTR") return "";
    return ctype;
}

// 账 #305 §B156 (第四十四刀): 多维 ReDim 的界数组临时量以前**按数组名**拼
// (`_redim_bounds_<名>`)，于是同一过程里第二条指向同一枚数组的多维 ReDim 会把同一条
// `vb6_SafeArrayBound <名>[] = {…};` 在同一个 C 函数里发两遍 ⇒ cl C2374「重定义；多次初始化」
// (两台同形，旧的那台也红 —— 读数 .build/b1395_k43/{n6,b6}/c3-error.log)。
// 这一枚是**唯一拼法**出口：名字自带语句序号，visit(ReDimStmt) 与 emitReDimComplexTarget
// 两条发码路都问它（以前两各路各拼一份，是「同一件事两处写法不同形」的又一例）。
std::string CCodeGen::nextNDBoundsTemp(const std::string& base) {
    return "_redim_bounds_" + base + "_" + std::to_string(ndBoundsSeq_++);
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
    // 账 #305 §B154 (第四十二刀): 元素档改问 resolveReDimElemType —— 写了 As 子句照旧优先,
    // 没写时**问声明** (arrayElemTypes_, 与访问侧 VB6_SA_AT / VB6_SA_ND_AT2 同一张表)。
    // 旧写法只问 asType, 于是裸写 `ReDim a(5)` 的标量/字符串/对象数组一律造出
    // vb6_sa_variant 而访问按声明的 C 类型步进 —— 只有 UDT 那一族有回退 (:109)。
    Vb6Type elemType = resolveReDimElemType(lowerVar, node.asType.get());
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

    // 账 #305 第二十二层 (2026-10-13): ReDim 的边界表达式若是 **COM 成员读**
    // (`ReDim Preserve a(1 To coll.Count)`), emitExpr 只发出**接收者对象**
    // (`me->coll`) 并在 lastExpr_ 里留下待 resolve 的 COM 标记; 此处以前直接取
    // lastExpr_ 当界 ⇒ 把**对象指针**当元素个数。实测 ImcComboItems.Add 的
    // `ReDim Preserve PropIndexPtrArr(1 To PropComboItem.Count)`: 界成了
    // PropComboItem 的堆指针 (≈7.8e7) ⇒ 每次 Add 给 TINDEXPTR 数组 calloc+memcpy
    // ≈600MB, 30 次 Add 拷贝 ≈15GB ⇒ 启动窗口被拖到 8s+ (采样剖析: 84% 时间在
    // CopyUpLargeMov/memmove)。
    // 与 For 循环三表达式同一律 (cgen_control.cpp emitForExpr092n): 取界之前先
    // 消费 COM 标记 —— 清标记 → emitExpr → 若产生标记则 resolveComValue("Long")。
    auto emitRedimBound = [&](Expr& e) -> std::string {
        isComMarker_ = false;
        emitExpr(e);
        if (isComMarker_) resolveComValue("Long");
        return std::move(lastExpr_);
    };

    int dimCount = (int)node.dimensions.size();

    if (dimCount == 1) {
        // 涓€缁?ReDim (淇濇寔鍘熸湁1D浠ｇ爜)
        auto& dim = node.dimensions[0];
        std::string lBound = "0";
        std::string uBound = "0";
        if (dim.lower) lBound = emitRedimBound(*dim.lower);
        if (dim.upper) uBound = emitRedimBound(*dim.upper);

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
                // 账 #305 §B218 (甲) 腿⑤ 第四格 (第一百四十八刀): 元素是**本工程类实例**时换那位
                // 带回调的兄弟出口 —— 收缩丢尾的那几格在搬移之前逐枚交回同一位主人。界仍只算一次:
                // lBound/uBound 是同一对文本, 出厂逻辑在 RTL 那一处, 搬移本身直接转调 _T。
                // 认不出 (As Object = 包装器 / As Variant = 槽 / 外部类 / 点链成员数组) ⇒ 今天那一发
                // 逐字节留着 ⇒ 全语料的 Preserve 形状只有类数组那几格会动。
                const std::string clsPreserve = classOfArrayCarrier(node.varName);
                assignVal = clsPreserve.empty()
                    ? ("vb6_SafeArrayReDimPreserve1D_T(" + saElemType + ", " + callArg + ", " + lBound + ", " + uBound + ")")
                    : ("vb6_SafeArrayReDimPreserve1D_Tr(" + saElemType + ", " + callArg + ", " + lBound + ", " + uBound + ", " + clsPreserve + "_Release)");
            }
            c_.emitLine(cName + " = " + assignVal + ";");
        } else {
            // Fix 084a: Variant 数组非 preserve ReDim — 先销毁提取出的 SafeArray, 再包装新数组回 Variant
            std::string destroyArg = cName;
            if (isVariantArrayVar(cName)) destroyArg = "vb6_VariantToSafeArray1D(" + cName + ")";
            // 账 #305 §B218 (甲) 腿⑤ 第三格 (第一百四十六刀): 不带 Preserve 的那一发把旧载体整块
            // 换掉 ⇒ 旧元素那一枚引用没人放手 = 漏。销毁之前问一处 (认不出 = 一声不出)。
            // Preserve 那一支**不吃**: 元素是被搬走的, 不是被丢的 (丢尾那一档随 RTL 出口另议)。
            emitClassArrayElemHandback(cName, node.varName);
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
        // 账 #305 §B156 (第四十四刀): 界数组临时量走唯一出口，名字带语句序号
        std::string boundsVar = nextNDBoundsTemp(rawCName);
        c_.emitLine("vb6_SafeArrayBound " + boundsVar + "[] = {");
        c_.indent();
        for (int d = 0; d < dimCount; d++) {
            auto& dim = node.dimensions[d];
            std::string lb = "0", ub = "0";
            if (dim.lower) lb = emitRedimBound(*dim.lower);
            if (dim.upper) ub = emitRedimBound(*dim.upper);
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
            // 账 #305 §B154 第二格: 装箱那一层**必须先把 ND 描述符投成 vb6_SafeArray1D***。
            // vb6_VariantFromValue 是 _Generic, 名单里只有 `vb6_SafeArray1D*` 这一档 (没有
            // SafeArrayND 的 arm), 递一枚 vb6_SafeArrayND* 进去会掉进 default = vb6_VariantObject
            // ⇒ 槽的 vt 成了 DISPATCH, 而 vb6_VariantArraySet/GetND 开头只认 vt & vb6_vtArray
            // ⇒ 一声不响地不写也不读 (实测 `Dim x As Variant` + `ReDim x(1, 2)` + x(1,1)="S"
            // 读回**空串**)。仓里对多维描述符的既有约定本来就是「1D 指针当载体, 用时再投回 ND」
            // (见上面 (vb6_SafeArrayND*)redimArg 与下面 DestroyND 那两枚强转), 所以补的是这**一层
            // 投影**, 不是给 _Generic 开新档。
            wrapBack = "vb6_VariantFromValue((vb6_SafeArray1D*)";
        }
        if (node.preserve) {
            std::string res;
            if (isUdtArray) {
                // Bug4-Fix Udt Preserve: 多维 UDT 数组传入元素尺寸
                res = "vb6_SafeArrayReDimPreserveND_Udt((int32_t)sizeof(" + udtCType + "), (vb6_SafeArrayND*)" + redimArg + ", " + std::to_string(dimCount) + ", " + boundsVar + ")";
            } else {
                // 账 #305 §B155 格④ (第四十三刀): 这一枚也要**把档递下去** —— 上面
                // 不带 Preserve 的那一支 (:244) 一直递 saElemType, Preserve 这一支却递了
                // 空手, 于是「首次分配就走 Preserve」的形状在 RTL 里回落 vb6_sa_empty
                // (4 字节一步) 而访问按声明档走 —— 同一件事两处写法不同形, 第三例。
                // 账 #305 §B218 (甲) 腿⑤ 第六格 (第一百五十四刀): 元素是本工程类时换那位带回调的兄弟出口 ——
                // 收缩丢尾的那几格在旧块被 free 之前逐枚交回同一位主人。哪些格出局、每格多宽都由载体自己在
                // 运行期答, 所以这一站不吃 arrayDimCounts_ 那张按代码顺序的表; 认不出 (As Object 包装器 /
                // As Long 裸值 / 外部类 / 点链成员数组) 继续发今天那一发, 产物逐字节留着。
                const std::string clsNDPreserve = classOfArrayCarrier(node.varName);
                res = clsNDPreserve.empty()
                    ? ("vb6_SafeArrayReDimPreserveND(" + saElemType + ", (vb6_SafeArrayND*)" + redimArg + ", " + std::to_string(dimCount) + ", " + boundsVar + ")")
                    : ("vb6_SafeArrayReDimPreserveND_Tr(" + saElemType + ", (vb6_SafeArrayND*)" + redimArg + ", " + std::to_string(dimCount) + ", " + boundsVar + ", " + clsNDPreserve + "_Release)");
            }
            if (!wrapBack.empty()) res = wrapBack + res + ")";
            // 账 #305 §B154 第二格: 那枚 (vb6_SafeArray1D*) 强转**只在该槽是数组指针时**才成立。
            // wrapBack 非空时右值已经是 vb6_VARIANT (Dim x As Variant + ReDim x(1, 2)), 强转
            // 就把装箱盖住了 —— BASE 上这一形实测 cl C2440「无法从 vb6_VARIANT 转换为
            // vb6_SafeArray1D*」, 也就是多维的 Variant 槽那一条从没能编过 (一维那条没有强转,
            // 见上面 dimCount==1 的 newVal, 所以只有多维红 —— 又是一个同一件事两处写法不同形)。
            c_.emitLine(cName + " = " + (wrapBack.empty() ? "(vb6_SafeArray1D*)" : "") + res + ";");
        } else {
            std::string destroyArg = redimArg;
            // 账 #305 §B218 (甲) 腿⑤ 第五格 (第一百四十九刀): 多维载体的元素出厂走同一位主人，
            // 由它自己问「这枚载体是几维」挑 1D / ND 那一发 —— 排在销毁之前同样是硬约束。
            emitClassArrayElemHandback("(vb6_SafeArrayND*)" + destroyArg, node.varName);
            c_.emitLine("vb6_SafeArrayDestroyND((vb6_SafeArrayND*)" + destroyArg + ");");
            std::string newVal;
            if (isUdtArray) {
                // Bug4-Fix: UDT多维数组使用_Udt版本，传入sizeof(UDT类型)
                newVal = "vb6_SafeArrayReDimND_Udt((int32_t)sizeof(" + udtCType + "), " + std::to_string(dimCount) + ", " + boundsVar + ")";
            } else {
                newVal = "vb6_SafeArrayReDimND(" + saElemType + ", " + std::to_string(dimCount) + ", " + boundsVar + ")";
            }
            if (!wrapBack.empty()) newVal = wrapBack + newVal + ")";
            // 与上面 Preserve 那一支同一个决定 (账 #305 §B154 第二格): wrapBack 在场就是
            // vb6_VARIANT, 那枚指针强转会把装箱盖住 (BASE 实测 cl C2440)。
            c_.emitLine(cName + " = " + (wrapBack.empty() ? "(vb6_SafeArray1D*)" : "") + newVal + ";");
        }

        // 鏇存柊鏁扮粍缁村害淇℃伅
        std::string lower = node.varName;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        arrayDimCounts_[lower] = dimCount;
    }
}

// 账 #476 §B190: 这一枚从"自己摸 udtMembers"改成**那一枚出口的薄委托** —— §B190 的口径是
// 「三处同问同一个出口」(Pattern B 元素访问 / ReDim 分配侧 / 类字段类外访问), 而 §B220 那版
// 自己查表, 于是 UDT 腿留下了第二份答案。现在它只负责把 udtCType 翻成出口的两个入参,
// 答案一律由 resolveMemberArrayDecl 给。
// ⚠ 它与 Pattern B 的差别只剩一个: 这里答的是"成员声明档"(定长也答), 而出口只认**动态**
// 数组成员 —— 所以定长那一形照旧由调用点走 Pattern A, 不经本格。
bool CCodeGen::resolveUdtMemberElemInfo(const std::string& udtCType,
                                        const std::string& memberName,
                                        Vb6Type& outType,
                                        std::string& outUdtCType) {
    outType = Vb6Type::Variant;
    outUdtCType.clear();
    MemberArrayDecl mad;
    if (!resolveMemberArrayDecl(udtCType, std::string(), memberName, mad) || !mad.ok) {
        return false;
    }
    outType = mad.elemType;
    if (mad.isUdt) outUdtCType = mad.elemCType;
    return true;
}

// Fix 100: 解析 ReDim 复杂目标 (arr(i).Field) 的元素类型。
// targetExpr 的末段必须是 MemberAccessExpr; 其 object 的 UDT C 类型由
// inferUdtTypeOfExpr 推断 (UDT 数组元素 arr(idx) 走 arrayUdtElemTypes_),
// 再交给 resolveUdtMemberElemInfo 问那唯一一份成员表。
void CCodeGen::resolveReDimComplexElemType(const ReDimStmt& node, Vb6Type& outType,
                                           std::string& outUdtCType) {
    outType = Vb6Type::Variant;
    outUdtCType.clear();
    if (!node.targetExpr) return;
    if (node.targetExpr->kind != ASTNodeKind::MemberAccessExpr) return;
    auto& ma = static_cast<const MemberAccessExpr&>(*node.targetExpr);
    if (!ma.object) return;

    std::string ownerCType = inferUdtTypeOfExpr(*ma.object);
    resolveUdtMemberElemInfo(ownerCType, ma.memberName, outType, outUdtCType);
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
    // 账 #475 §B220: 这一段是**非复杂目标**的裸名回退 (键 = node.varName)。复杂目标
    // (node.targetExpr 非空) 的元素档已由上面 resolveReDimComplexElemType 问过成员表，
    // 而它的 node.varName 常只是基数组名 (`ReDim arr(1).Numbers(...)` → "arr") —— 再拿它
    // 查 arrayUdtElemTypes_ 会命中的是**基数组元素 UDT**，于是标量成员被按 sizeof(UDT)
    // 分配 (实测 Probe5: vb6_SafeArrayReDim1D_Udt(sizeof(vb6_type_TNum)) 而访问是 int32)。
    if (!node.targetExpr && udtCType.empty() && !node.varName.empty()) {
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

    // 账 #305 第二十二层: 与 visit(ReDimStmt&) 同一律 —— ReDim 界若是 COM 成员读,
    // emitExpr 只发对象并留 COM 标记, 必须 resolve 后再当界 (否则对象指针成了元素数)。
    auto emitRedimBound = [&](Expr& e) -> std::string {
        isComMarker_ = false;
        emitExpr(e);
        if (isComMarker_) resolveComValue("Long");
        return std::move(lastExpr_);
    };

    if (dimCount == 1) {
        auto& dim = node.dimensions[0];
        std::string lBound = "0";
        std::string uBound = "0";
        if (dim.lower) lBound = emitRedimBound(*dim.lower);
        if (dim.upper) uBound = emitRedimBound(*dim.upper);

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
        // 多维 ReDim: bounds 局部数组名由目标左值净化得到 (仅需 C 标识符合法)。
        // 账 #305 §B156 (第四十四刀): 拼法改问同一枚出口 —— 以前这里也按名字裸拼
        // (`_redim_bounds` + 净化后的左值)，同一过程里两条指向同一目标的多维 ReDim
        // 一样撞 C2374。序号由出口统一给，两条路不再各写一份。
        std::string boundsSuffix;
        for (char ch : cName) {
            if (isalnum((unsigned char)ch) || ch == '_') boundsSuffix += ch;
        }
        std::string boundsVar = nextNDBoundsTemp(boundsSuffix);
        c_.emitLine("vb6_SafeArrayBound " + boundsVar + "[] = {");
        c_.indent();
        for (int d = 0; d < dimCount; d++) {
            auto& dim = node.dimensions[d];
            std::string lb = "0", ub = "0";
            if (dim.lower) lb = emitRedimBound(*dim.lower);
            if (dim.upper) ub = emitRedimBound(*dim.upper);
            std::string trailing = (d < dimCount - 1) ? "," : "";
            c_.emitLine("{" + lb + ", (" + ub + " - " + lb + " + 1)}" + trailing);
        }
        c_.dedent();
        c_.emitLine("};");
        // 账 #305 §B218 (甲) 腿⑤ 第五格 (第一百四十九刀): 这一支的目标是复杂左值 (成员链 / 下标)，
        // 问句只认裸名 ⇒ 认不出就一声不出；认得出时仍由主人自己判维度挑出口，站点不代替它判。
        emitClassArrayElemHandback("(vb6_SafeArrayND*)" + cName, node.varName);
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
            // 账 #305 §B218 (甲) 腿⑤ 第五格 (第一百四十九刀): 多维载体的 Erase 也走同一位主人。
            emitClassArrayElemHandback("(vb6_SafeArrayND*)" + cName, name);
            c_.emitLine("vb6_SafeArrayDestroyND((vb6_SafeArrayND*)" + cName + "); " + cName + " = NULL;");
        } else {
            // Fix <vbeclipse> 2026-10-06: 1D 数组统一走 vb6_EraseArrayVar —
            // 动态数组释放存储置 NULL, **固定数组保留存储**只把元素重置
            // (数值/BSTR/Variant 按 VB6 文档口径; 此前固定数组也被整块释放,
            //  Erase VarArray 后再访问 UBound/元素直接错误 9, t3.bas 实测)。
            // 账 #305 §B218 (甲) 腿⑤ 第三格 (第一百四十六刀): 类数组的元素由**这一格**握着 ⇒
            // 销毁载体之前逐枚交回同一位主人 (问句与出口只住在 emitClassArrayElemHandback 那一处,
            // 认不出就一声不出 = 维持今天的裸销毁)。固定数组那一支同样吃这一发: RTL 里的
            // vb6_ReleaseObject 认不出工程类裸结构体 (Fix 191), 过去只把槽清成 NULL 而不放手。
            emitClassArrayElemHandback(cName, name);
            c_.emitLine("vb6_EraseArrayVar(&" + cName + ");");
        }
    }
}


} // namespace vb6c3
