#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_util_comwrite.cpp: COM 左值改写 (链式写 / LValue 重写 / Let 参数打包) ---


// ============================================================
// Fix 090ae: 链式 COM 默认属性索引赋值 (P25b helper, 原内联于 AssignmentStmt)
// VB: dic(a)(b) = v 或  Set dic(a)(b) = v — 多层默认属性/Item 索引写.
// emitExpr(target) 只支持链式读 (生成 vb6_VariantFromComResult(vb6_ComCall(...))
// 的 rvalue, 作为 LHS 触发 C2440). 这里逐层把中间层结果解包为对象
// (vb6_ComCallObject = ComCall+UnpackObject), 对最外层用 vb6_ComSetPropArg.
// AssignmentStmt/SetStmt 共用.
// ============================================================
bool CCodeGen::tryEmitChainedComWrite(Expr* targetNode, Expr* valueNode) {
    std::vector<IndexOrCallExpr*> chain;
    Expr* cur = targetNode;
    while (cur && cur->kind == ASTNodeKind::IndexOrCallExpr) {
        chain.push_back(static_cast<IndexOrCallExpr*>(cur));
        cur = static_cast<IndexOrCallExpr*>(cur)->callee.get();
    }
    if (chain.size() < 2 || !cur || cur->kind != ASTNodeKind::IdentifierExpr) {
        return false;
    }
    auto& rootId25 = static_cast<IdentifierExpr&>(*cur);
    std::string rootLower25 = Symbol::toLower(rootId25.name);
    bool rootIsCom25 = knownObjectVars_.count(rootLower25)
                    || knownVariantVars_.count(rootLower25)
                    || knownTypedComVars_.count(rootLower25);
    // Fix 090e: 根也可以是"模块默认成员属性" — PropertyGet 返回 COM 对象
    // (如 cIni.Root 的 VB_UserMemId=0 默认成员, 返回 Dictionary) 时,
    // Root(Section)(Key) = v 与 dic(Section)(Key) = v 同构. emitExpr(Root)
    // 生成 vb6_cIni_prop_get_Root((void*)me) 调用文本作为链起点.
    // 判定: 该返回类型符号带 comDefaultMemberName (Dictionary→Item).
    Symbol* rootRtSym25 = nullptr;  // root 返回的 COM 类型符号 (默认成员名来源)
    if (!rootIsCom25) {
        Symbol* rootSym25 = symTab_.lookupModule(rootId25.name);
        if (rootSym25 && rootSym25->kind == SymbolKind::PropertyGet
            && !rootSym25->variableTypeName.empty()) {
            rootRtSym25 = lookupDotted(rootSym25->variableTypeName);
            if (rootRtSym25 && !rootRtSym25->comDefaultMemberName.empty()) {
                rootIsCom25 = true;
            }
        }
    }
    // Fix 090ae: 根也可以是"当前类的 void* COM 字段" — me->Data 这类 COM 字段
    // (Dim Data As New Dictionary) 的 Data(L)(C) = v 链式写. knownObjectVars_/
    // knownTypedComVars_ 只注册局部/参数, 不含类字段; 类 void* 字段在
    // classVoidFieldMap_ (driver 预扫描, 键=模块名, 集合=字段名小写含 m_ 变体).
    if (!rootIsCom25 && isClassModule_ && classVoidFieldMap_) {
        auto itV90ae = classVoidFieldMap_->find(moduleName_);
        if (itV90ae != classVoidFieldMap_->end()
            && (itV90ae->second.count(rootLower25)
                || itV90ae->second.count("m_" + rootLower25))) {
            rootIsCom25 = true;
        }
    }
    if (!rootIsCom25) return false;

    // 默认成员名: 优先 root 返回类型符号 (Dictionary→Item), 其次类型化 COM
    // 变量注册, 回退 "Item"
    std::string defMem25 = "Item";
    if (rootRtSym25 && !rootRtSym25->comDefaultMemberName.empty()) {
        defMem25 = rootRtSym25->comDefaultMemberRealName;
    } else {
        auto itTyped25 = knownTypedComVars_.find(rootLower25);
        if (itTyped25 != knownTypedComVars_.end()
            && !itTyped25->second->comDefaultMemberName.empty()) {
            defMem25 = itTyped25->second->comDefaultMemberRealName;
        }
    }
    std::string defMemLit25 = "L\"" + defMem25 + "\"";
    // 打包某层索引参数为 (void*[]){pack(a),...} 字符串
    auto packLayer25 = [&](IndexOrCallExpr& ic, std::vector<std::string>& out) -> int32_t {
        for (size_t j = 0; j < ic.positional.size(); j++) {
            std::string packFn25 = comPackExpr(*ic.positional[j]);
            emitExpr(*ic.positional[j]);
            { std::string r25 = resolveComMarkerForPack(packFn25); if (!r25.empty()) lastExpr_ = r25; }
            out.push_back(packFn25 + "(" + lastExpr_ + ")");
        }
        return (int32_t)ic.positional.size();
    };
    emitExpr(*cur);  // 根对象表达式 (COM 变量 / me->Field / PropertyGet 调用)
    std::string accExpr25 = std::move(lastExpr_);
    // chain[0]=最外层索引, chain[last]=最内层索引. 先按最内→外取对象,
    // 即逆序遍历 chain 深度层 (除最外层), 每层取默认成员对象
    for (int32_t i = (int32_t)chain.size() - 1; i >= 1; i--) {
        std::vector<std::string> packedMid;
        int32_t argcMid = packLayer25(*chain[i], packedMid);
        std::string arrMid = "(void*[]){";
        for (size_t k = 0; k < packedMid.size(); k++) {
            if (k > 0) arrMid += ", ";
            arrMid += packedMid[k];
        }
        arrMid += "}";
        accExpr25 = "vb6_ComCallObject(" + accExpr25 + ", " + defMemLit25
                  + ", " + arrMid + ", " + std::to_string(argcMid) + ")";
    }
    // 最外层: 带索引属性 Put
    std::vector<std::string> packedOut25;
    int32_t argcOut25 = packLayer25(*chain[0], packedOut25);
    std::string arrOut25 = "(void*[]){";
    for (size_t k = 0; k < packedOut25.size(); k++) {
        if (k > 0) arrOut25 += ", ";
        arrOut25 += packedOut25[k];
    }
    arrOut25 += "}";
    emitExpr(*valueNode);
    std::string valExpr25 = std::move(lastExpr_);
    std::string packVal25 = comPackExpr(*valueNode);
    // Fix <vbeclipse>: 实参本身是 COM 属性读 (`.RefId` / `.Ratio` 这类 With 块内
    // 成员) 时, emitExpr 只留下接收者对象表达式, 成员名还挂在 comMemberName_ 上。
    // 不消费就整体丢失 —— 生成的 vb6_ComPackBSTR(_vb6_with_1) 把接收者当值打包,
    // 而 frmViewSnapshot.frm:106-107
    //   li.SubItems(1) = .RefId   → vb6_ComPackBSTR(vb6_VariantFromComResult(vb6_ComCall(_vb6_with_1, L"RefId", NULL, 0)))
    // 的实参已是 vb6_VARIANT 结构体, 与 packer 期望的 const wchar_t* / double 不匹配
    // → C2440 (4 处: 181/182/205/206)。
    // 与 cgen_assign_com_prop.inc 同名 Fix 110i 同一处理: 按 packer 反推解封类型后
    // resolveComValue, 得到 vb6_ComGetXxxProp(receiver) 这类与 packer 匹配的标量。
    if (isComMarker_) {
        std::string hint25 = "BSTR";
        if (packVal25 == "vb6_ComPackDouble") hint25 = "Double";
        else if (packVal25 == "vb6_ComPackInt") hint25 = "Long";
        else if (packVal25 == "vb6_ComPackObject") hint25 = "Object";
        else if (packVal25 == "vb6_ComPackValue") hint25 = "Variant";
        resolveComValue(hint25);
        valExpr25 = std::move(lastExpr_);
    }
    c_.emitLine("vb6_ComSetPropArg(" + accExpr25 + ", " + defMemLit25 + ", "
                + arrOut25 + ", " + std::to_string(argcOut25) + ", "
                + packVal25 + "(" + valExpr25 + "));  /* COM chained default-prop assign (P25b) */");
    return true;
}


// ============================================================
// Fix 010r-16: COM/Property-Get 左值重写辅助函数实现
// ============================================================
bool CCodeGen::tryRewriteCOMLvalue(const std::string& target, const std::string& value,
                                   Expr* valueExpr, bool isSet) {
    if (target.empty()) return false;

    // 工具 lambda: 在 target 中查找首个顶层括号 (跳过嵌套括号), 返回 '(' 与 ')'
    // 的位置; 找不到返回 npos.
    auto findCallParens = [](const std::string& s, size_t startAt) -> std::pair<size_t, size_t> {
        size_t openPos = std::string::npos;
        for (size_t i = startAt; i < s.size(); ++i) {
            if (s[i] == '(') { openPos = i; break; }
        }
        if (openPos == std::string::npos) return {std::string::npos, std::string::npos};
        int depth = 0;
        for (size_t i = openPos; i < s.size(); ++i) {
            if (s[i] == '(') depth++;
            else if (s[i] == ')') {
                depth--;
                if (depth == 0) return {openPos, i};
            }
        }
        return {std::string::npos, std::string::npos};
    };
    // 工具 lambda: 按顶层逗号分割调用实参 (跳过括号/方括号/花括号内的逗号), 每段 trim.
    auto splitTopLevelArgs = [](const std::string& s) -> std::vector<std::string> {
        std::vector<std::string> out;
        if (s.empty()) return out;
        int depth = 0;
        size_t start = 0;
        for (size_t i = 0; i < s.size(); ++i) {
            char ch = s[i];
            if (ch == '(' || ch == '[' || ch == '{') { depth++; }
            else if (ch == ')' || ch == ']' || ch == '}') { if (depth > 0) depth--; }
            else if (ch == ',' && depth == 0) {
                out.push_back(s.substr(start, i - start));
                start = i + 1;
            }
        }
        out.push_back(s.substr(start));
        for (auto& seg : out) {
            size_t b = seg.find_first_not_of(" \t");
            if (b == std::string::npos) { seg.clear(); continue; }
            size_t e = seg.find_last_not_of(" \t");
            seg = seg.substr(b, e - b + 1);
        }
        return out;
    };

    // ---- P20-42: SSTab 索引属性的写 ----
    // 读侧 (cgen_expr_call_callee_member.inc) 已把 `SSTab1.TabCaption(0)` 改道成
    // vb6_SSTab_GetTabCaption(hwnd, 0) —— 它既不是左值, 也不带 vb6_ComCall 前缀,
    // 下面的 Pattern A 匹配不到。所以这里把 Get 换回对应的 Set, 原样带上实参。
    // value 不加 comPack: RTL 的 Set 形参是 void* bstr / int32_t, 直接给原表达式
    // (字符串侧此时已经是 vb6_BSTR_FromStr(...))。
    {
        static const char* kTsGet[3] = { "vb6_SSTab_GetTabCaption(", "vb6_SSTab_GetTabVisible(",
                                         "vb6_SSTab_GetTabToolTipText(" };
        static const char* kTsSet[3] = { "vb6_SSTab_SetTabCaption(", "vb6_SSTab_SetTabVisible(",
                                         "vb6_SSTab_SetTabToolTipText(" };
        for (int tsI = 0; tsI < 3; tsI++) {
            std::string getPfx = kTsGet[tsI];
            if (target.compare(0, getPfx.size(), getPfx) != 0) continue;
            size_t tsOpen = target.find('(');
            size_t tsClose = target.rfind(')');
            if (tsOpen == std::string::npos || tsClose == std::string::npos
                || tsClose < tsOpen) continue;
            std::string tsArgs = target.substr(tsOpen + 1, tsClose - tsOpen - 1);
            c_.emitLine(std::string(kTsSet[tsI]) + tsArgs + ", " + value
                        + ");  /* SSTab indexed prop assignment */");
            return true;
        }
    }

    // ---- Pattern A: vb6_ComCall(obj, L"Item", args, n) = value ----
    // vb6_ComCall 返回 VARIANT*, 不是左值. 改走 vb6_ComSetPropArg (内部
    // DISPATCH_PROPERTYPUT|PUTREF).
    if (target.find("vb6_ComCall") == 0) {
        auto parens = findCallParens(target, 0);
        if (parens.first != std::string::npos && parens.second != std::string::npos) {
            std::string callArgs = target.substr(parens.first + 1,
                                                 parens.second - parens.first - 1);
            std::string packFn = valueExpr ? comPackExpr(*valueExpr) : "vb6_ComPackVariant";
            std::string tag = isSet ? "Set" : "Let";
            c_.emitLine("vb6_ComSetPropArg(" + callArgs + ", " + packFn + "(" + value +
                        "));  /* COM Item assignment (Pattern A, " + tag + ") */");
            return true;
        }
    }

    // ---- Pattern F: vb6_ComGetStringProp(obj, L"Prop") = value ----
    // With-block ClassInstance member: 当 Property 在跨模块符号表里找不到时,
    // codegen 会走 "vb6_ComGetStringProp(_vb6_with_X, L\"Prop\")" 路径; 但读路径
    // 产生的 wchar_t* 不是左值, 赋值触发 C2106. 改走 vb6_ComSetProp.
    // 另外也匹配 vb6_ComGetIntProp / vb6_ComGetDoubleProp / vb6_ComGetObjectProp.
    const std::vector<std::string> comGetFns = {
        "vb6_ComGetStringProp", "vb6_ComGetIntProp", "vb6_ComGetDoubleProp",
        "vb6_ComGetObjectProp", "vb6_ComGetProp"
    };
    for (const auto& fnName : comGetFns) {
        if (target.find(fnName) == 0) {
            auto parens = findCallParens(target, 0);
            if (parens.first != std::string::npos && parens.second != std::string::npos) {
                std::string callArgs = target.substr(parens.first + 1,
                                                     parens.second - parens.first - 1);
                // isSet=true => 对象引用语义, 用 PackObject (DISPATCH_PROPERTYPUTREF)
                // isSet=false => Let 语义, 用 comPackExpr 推断
                std::string packFn = isSet ? "vb6_ComPackObject"
                                  : (valueExpr ? comPackExpr(*valueExpr) : "vb6_ComPackVariant");
                std::string tag = isSet ? "Set" : "Let";
                c_.emitLine("vb6_ComSetProp(" + callArgs + ", " + packFn + "(" + value +
                            "));  /* COM prop assignment (Pattern F, " + tag + ") */");
                return true;
            }
        }
    }

    // ---- Pattern C/D2: vb6_X_prop_get_Y(args) = value ----
    // Property Get 用作 LHS. 改写为 prop_let_Y(args, value) (Let) 或
    // prop_set_Y(args, value) (Set).
    // Fix 084i: 也匹配 prop_set_/prop_let_ 前缀 — 当属性只有 PropertySet
    // (无 Get) 时 emitExpr 直接生成 vb6_cX_prop_set_Y(obj) 单参数形式,
    // 追加 value 参数改写成完整调用, 避免 prop_set_(obj) = value 左值错误.
    {
        const char* verbs[] = {"prop_get_", "prop_set_", "prop_let_"};
        size_t pgPos = std::string::npos;
        std::string matchedVerb;
        for (const char* v : verbs) {
            size_t p = target.find(v);
            if (p != std::string::npos) { pgPos = p; matchedVerb = v; break; }
        }
        if (pgPos != std::string::npos) {
            auto parens = findCallParens(target, pgPos);
            if (parens.first != std::string::npos && parens.second != std::string::npos) {
                std::string prefix = target.substr(0, pgPos);              // vb6_cX_
                std::string afterPg = target.substr(pgPos + matchedVerb.size(),
                                                    parens.first - (pgPos + matchedVerb.size()));
                std::string argsStr = target.substr(parens.first + 1,
                                                    parens.second - parens.first - 1);
                std::string newVerbs = (matchedVerb != "prop_get_") ? matchedVerb
                                      : (isSet ? "prop_set_" : "prop_let_");
                // Fix <vbeclipse> rev9: **Property Set 的值槽是 ByRef 对象/类形参**
                // (C 侧 `void**` / `vb6_cls_X**`) 时, 值实参必须交付**槽地址**,
                // 不能交付对象指针本身。
                //
                // 根因: 定义侧 (makeParamCType) 对 `!isByVal` 一律加 `*`, 于是
                //   View.cls:77  Public Property Set View(ByRef NewView As Object)
                //     → void vb6_View_prop_set_View(vb6_cls_View* me, void** NewView)
                //     → 函数体 `Set m_View = NewView` 落成 me->m_View = (*NewView)
                //   ucPerspective.ctl:1398  Public Sub AddView(ByVal ViewId As String,
                //                                        ByRef View As Object)
                //     → void vb6_ucPerspective_AddView(..., void** View)
                // 而 Pattern C/D2 是**字符串级**改写: 它把 RHS 的 C 文本原样拼上去。
                // RHS 标识符 `View` 在 AddView 体内 emitExpr 出的是右值 `(*View)`
                // (读那个 ByRef 槽里的对象指针), 于是拼成
                //     vb6_View_prop_set_View(_vb6_with_32, (*View));
                // 被调方再解一层 → 拿**对象指针**当槽地址去读 → av read target=0x0
                // (x86 实测: `mov edx,[ecx]` / ecx=0), 崩在 Form_Load 里,
                // 整个 IDE 起不来。x64 上同一处更早地变成栈溢出/跳飞 (Long 存指针
                // + 手搓机器码的 MagneticWnd 把它放大成 24 帧重复的垃圾栈)。
                //
                // 为什么不能像 Let 那样"看着像右值就包": 判据必须落在**被调方槽的
                // C 类型**上 (makeParamCType 的 `!isByVal → +*`), 与改写目标是不是
                // prop_let_ 无关 —— `Property Set Icon(ByVal NewIcon As Picture)`
                // (cgen_setlet_set_prop.inc Fix 166 同款) 的槽是单指针, 包了就是
                // 传"指针的指针", 反而坏。
                //
                // 与 cgen_setlet_set_prop.inc:301-342 (Set Property 路径) 同款包装:
                // 那条路径早就包了 `&(void*){...}`, 只有本条 (With 块内 ClassInstance
                // 成员经 prop_get_ 改写) 漏了, 于是同一工程两种写法产出两种码。
                bool wrapSetSlotAddr9 = false;
                if (isSet && prefix.rfind("vb6_", 0) == 0 && !afterPg.empty()
                    && afterPg.find('(') == std::string::npos) {
                    std::string clsRev9 = prefix.substr(4);   // 去 "vb6_" → "cXx_"
                    if (!clsRev9.empty() && clsRev9.back() == '_') clsRev9.pop_back();
                    std::vector<ParameterInfo> setSigRev9;
                    if (!clsRev9.empty()
                        && findClassMemberWriteParams(clsRev9, afterPg, /*isSet=*/true,
                                                       setSigRev9)
                        && !setSigRev9.empty()) {
                        const ParameterInfo& valP9 = setSigRev9.back();
                        // 槽的 C 类型 = mapTypeRef(asType) (+'*' 当 !isByVal, 见
                        // makeParamCType)。只有"对象/类指针的指针"才要槽地址; 标量
                        // ByRef (int32_t* / BSTR*) 与 ByVal 对象 (单指针) 都不在此列。
                        // 走 mapTypeRef 而不是手列类型名 —— 定义侧也是它, 两边不会
                        // 各说一套 (同 cParamClassPtrType 的设计理由)。
                        const std::string vt9 =
                            cTypeForDeclaredTypeName(valP9.typeRefName);
                        const bool isObjPtrSlot9 =
                            !valP9.isByVal
                            && (vt9 == "void*" || vt9.compare(0, 8, "vb6_cls_") == 0
                                || vt9.compare(0, 10, "vb6_ivref_") == 0
                                || vt9.compare(0, 13, "vb6_ComIface_") == 0);
                        if (isObjPtrSlot9) {
                            // RHS 已是槽地址 (&x) / 复合字面量 (&(T){..}) / ComPack
                            // 结果 / NULL ⇒ 原样交付, 别套第二层。
                            const bool alreadyAddr9 =
                                value.empty() || value == "NULL" || value == "0"
                                || (value[0] == '&')
                                || value.find("vb6_ComPack") != std::string::npos;
                            if (!alreadyAddr9) wrapSetSlotAddr9 = true;
                        }
                    }
                }
                // Fix 090w: Property Let 末参 As Variant (cJson.Item Dat As Variant
                // ByRef / cCsv.Value ByVal Dat As Variant / cHttpServerCookieAttr.Expires
                // Let(v As Variant) — Get 无参) — 值实参 (vb6_Now() double / 666 int /
                // BSTR) 需打包成 vb6_VARIANT, 否则 Pattern C/D2 裸拼 → C2440
                // "double/int/BSTR → vb6_VARIANT(/ *)". 仅当改写目标是 prop_let_/
                // prop_set_ 且 **Let/Set 方向**末形参解析为 Variant 时打包 —
                // findClassMemberCallParams 按 Get > Function > Sub > Let > Set 优先,
                // 对 Get 有参/无参而 Let 末参才是 value 的属性 (Item(key),
                // Expires) 会取错方向, 故直接查 PropertyLet/PropertySet 符号.
                std::string valArg = wrapSetSlotAddr9 ? ("&(void*){" + value + "}")
                                                      : value;
                // Fix 159-A: 保存 Property Let 的**写方向**形参表, 供拼接值实参时
                // 定位 Value 的真实槽位 (见本函数尾部 newCall 构造).
                std::vector<ParameterInfo> letSig159;
                if (!isSet && newVerbs.find("prop_get_") == std::string::npos
                    && prefix.rfind("vb6_", 0) == 0 && !afterPg.empty()
                    && afterPg.find('(') == std::string::npos) {
                    std::string cls90w = prefix.substr(4);  // 去 "vb6_" → "cCsv_"
                    if (!cls90w.empty() && cls90w.back() == '_') cls90w.pop_back();
                    // Fix 091a: 写方向参数查找 (Phase A 模块作用域 Let 符号;
                    // Phase B Class 符号 memberLetParams — 覆盖 item$pl 等跨模块
                    // storageKey 冲突场景, 此时读方向 memberParams 只有 Get 的
                    // [key] 会取错方向).
                    const ParameterInfo* lastP90w = nullptr;
                    std::vector<ParameterInfo> writeP90w;
                    bool writeDirResolved90w = false;   // Fix 156-B
                    if (findClassMemberWriteParams(cls90w, afterPg, isSet, writeP90w)
                        && !writeP90w.empty()) {
                        lastP90w = &writeP90w.back();
                        writeDirResolved90w = true;
                        letSig159 = writeP90w;   // Fix 159-A
                    } else {
                        // fallback: 写方向也查不到时沿用旧 findClassMemberCallParams
                        // (Get 优先 — 可能取到 Get 方向参数, 此时保守不打包)
                        std::vector<ParameterInfo> letParams90w;
                        bool letBuiltin90w = false;
                        if (findClassMemberCallParams(cls90w, afterPg, letParams90w,
                                                      letBuiltin90w)
                            && !letParams90w.empty()) {
                            lastP90w = &letParams90w.back();
                        }
                    }
                    if (lastP90w && lastP90w->type == Vb6Type::Variant) {
                        valArg = packLetValueArg(*lastP90w, valueExpr, value);
                    } else if (lastP90w && writeDirResolved90w
                               && lastP90w->type != Vb6Type::Empty) {
                        // Fix 156-B: 反方向 — Property Let 末形参是**具体类型**
                        // (Integer/Long/String/Double/Object...) 而 RHS 是 vb6_VARIANT
                        // → 裸拼 valArg 触发 C2440 "无法从 vb6_VARIANT 转换为 int16_t".
                        // 与 cgen_expr_call_arg_variant.inc 的 Fix 029 (位置实参路径)
                        // 同构, 但 Pattern C/D2 是**字符串级改写**(直接拼 prop_let_ 调用),
                        // 不经过实参发射循环, 因此那条修正覆盖不到这里.
                        //   VBFlexGrid.ctl: Select Case VarType(Value) 把 Variant 分发到
                        //     各类型化属性 (28 处):
                        //       vb6_VBFlexGrid_prop_let_CellTextStyle((void*)me, Value)
                        //         (形参 Integer) → C2440 vb6_VARIANT→int16_t
                        //       vb6_VBFlexGrid_prop_let_Text((void*)me, Value)
                        //         (形参 String)  → C2440 vb6_VARIANT→BSTR
                        //   MainForm/UserEditingForm: prop_let_Cell(r, c, Value) 同理.
                        // 仅在**写方向参数表解析成功** (findClassMemberWriteParams) 时执行:
                        // 090w 的 fallback 走 findClassMemberCallParams (Get 优先), 末形参
                        // 可能是 Get 的索引 (如 Cell 的 Col) 而非 Let 的 value, 按它提取
                        // 会得到语义错误的结果, 故宁可保持原状.
                        // Variant 判定与 toLongIfVariant 一致: C 串级前缀 + 已知 Variant 变量.
                        bool valIsVar156 = cExprIsVariant(value);
                        if (!valIsVar156 && valueExpr
                            && valueExpr->kind == ASTNodeKind::IdentifierExpr) {
                            std::string vl156 = Symbol::toLower(
                                static_cast<IdentifierExpr&>(*valueExpr).name);
                            if (knownVariantVars_.count(vl156)) valIsVar156 = true;
                        }
                        // Fix 169: RHS 是返回 **Windows VARIANT\*** 的 COM 调用
                        // (vb6_ComCall/vb6_ComGetProp/…) 时, 上面两条判定都不认它
                        // (它们返回 void\*, 不是 vb6_VARIANT) → 于是裸拼给具体类型形参,
                        // C 里指针→整数只是告警不是错误, 值被截断成垃圾:
                        //   VBFlexGrid.c:1107 (UserControl_ReadProperties)
                        //     prop_let_OLEDropMode(me, vb6_ComCall(PBag, L"ReadProperty", …))
                        //   → .ctl 的 Select Case 落 Case Else → Err.Raise 380
                        //   → 该链上无错误处理器 → vb6_ErrRaise 里 ExitProcess(380),
                        //     demo 启动即退 (Fix 167 让它进了消息循环后才暴露出来)。
                        // 按既有惯用法先 vb6_VariantFromComResult 归一成 vb6_VARIANT,
                        // 再走同一条 ext156 提取链 (与 ScaleWidth 读取处同构)。
                        std::string valExpr169 = value;
                        {
                            static const std::vector<std::string> comRes169 = {
                                "vb6_ComCall(", "vb6_ComCallByDispid(",
                                "vb6_ComGetProp(", "vb6_ComGetPropArg(",
                            };
                            size_t s169 = valExpr169.find_first_not_of(" \t\r\n(");
                            for (const auto& pr : comRes169) {
                                if (s169 != std::string::npos
                                    && valExpr169.compare(s169, pr.size(), pr) == 0) {
                                    valExpr169 = "vb6_VariantFromComResult(" + valExpr169 + ")";
                                    valIsVar156 = true;
                                    break;
                                }
                            }
                        }
                        if (valIsVar156) {
                            bool pIsArr156 =
                                (static_cast<uint16_t>(lastP90w->type)
                                 & static_cast<uint16_t>(Vb6Type::Array)) != 0;
                            Vb6Type pBase156 = static_cast<Vb6Type>(
                                static_cast<uint16_t>(lastP90w->type)
                                & ~static_cast<uint16_t>(Vb6Type::Array));
                            const char* ext156 = nullptr;
                            if (pIsArr156) {
                                ext156 = "vb6_VariantToSafeArray1D";
                            } else switch (pBase156) {
                                case Vb6Type::String:                          ext156 = "vb6_VariantToString";    break;
                                case Vb6Type::Long:   case Vb6Type::Integer:
                                case Vb6Type::Byte:   case Vb6Type::Boolean:                                     ext156 = "vb6_VariantToLong";     break;
                                case Vb6Type::Double: case Vb6Type::Single:
                                case Vb6Type::Currency: case Vb6Type::Date:                                     ext156 = "vb6_VariantToDouble";   break;
                                case Vb6Type::LongPtr: case Vb6Type::ULong:                                     ext156 = "vb6_VariantToLongPtr";  break;
                                case Vb6Type::Object:                          ext156 = "vb6_VariantToObjectVal"; break;
                                default: break;
                            }
                            // Fix 169: 用归一后的 valExpr169 (COM 结果已包成 vb6_VARIANT)
                            if (ext156) valArg = std::string(ext156) + "(" + valExpr169 + ")";
                        } else if (lastP90w->type == Vb6Type::String && valueExpr
                                   && isScalarImplicitStr159(
                                          inferExprType(*valueExpr))) {
                            // Fix 159-B (与 cgen_assign_com_prop.inc 同族): 写方向末参是
                            // ByVal String 而 RHS 是日期/数值标量 → VB6 隐式 CStr.
                            //   VBFlexGrid.ctl:5118 EditText = DateSerial(...)  →
                            //   VBFlexGrid.c:58027 prop_let_EditText(me, double) C2440.
                            valArg = wrapToBSTR(value, *valueExpr);
                        }
                    }
                }
                // Fix 090ad: 只写属性 (无 Get, 如 Dictionary.key(OldKey)=NewKey) 的 LHS
                // target — emitExpr 按 prop_let_/prop_set_ 完整调用生成时 (读上下文
                // fallback 到 Let/Set), 对缺失的 value 形参也 pad 了默认值
                // (vb6_BSTR_Empty()), 例如 prop_let_key(me->m_Dict, OldKey,
                // vb6_BSTR_Empty()). 若括号实参顶层段数 == 业务形参数+1 (对象+全部
                // 形参含 value), 末段即被 pad 的 value 位 → 丢弃, 由真实 RHS value
                // 拼接补齐, 否则 C2197 参数太多 (Dictionary.c 110/162).
                std::string finalArgs = argsStr;
                bool valuePadDropped159ad = false;   // Fix 159-A: 090ad 弹掉占位 Value 时跳过重排
                if (matchedVerb != "prop_get_" && !argsStr.empty()
                    && prefix.rfind("vb6_", 0) == 0 && !afterPg.empty()
                    && afterPg.find('(') == std::string::npos) {
                    std::string clsCd2 = prefix.substr(4);  // 去 "vb6_" → "cXx_"
                    if (!clsCd2.empty() && clsCd2.back() == '_') clsCd2.pop_back();
                    if (!clsCd2.empty()) {
                        std::vector<ParameterInfo> paramsCd2;
                        bool builtinCd2 = false;
                        if (findClassMemberCallParams(clsCd2, afterPg, paramsCd2, builtinCd2)
                            && !paramsCd2.empty()) {
                            std::vector<std::string> topArgs = splitTopLevelArgs(argsStr);
                            if (topArgs.size() == paramsCd2.size() + 1 && topArgs.size() > 1) {
                                topArgs.pop_back();  // 移除被 pad 的 value 默认值
                                valuePadDropped159ad = true;   // Fix 159-A
                                finalArgs.clear();
                                for (size_t i = 0; i < topArgs.size(); i++) {
                                    if (i) finalArgs += ", ";
                                    finalArgs += topArgs[i];
                                }
                            }
                        }
                    }
                }
                std::string newCall;
                // Fix 159-A: Value 实参槽位修正.
                // VB6 的 Property Let 把被赋的值放在**声明末尾的业务形参** (Cell 的
                // 第 6 个业务参数), 而 C 签名把 Optional 的 `int _has_*` 标志统一追加在
                // 所有业务参数之后 → Value 并非最后一个 C 形参. 本函数此前一律把
                // valArg 直接拼在最后, 于是 Value 落进 _has_Row 槽, 4 个标志整体前移:
                //   VBFlexGrid1.Cell(FlexCellToolTipText, i, j) = "i/j info tip."
                //   → prop_let_Cell(me, 8, i, j, -1, -1, 1, 1, 0, 0, "…")
                //     (签名: me, Setting, Row, Col, RowSel, ColSel, Value, _has_×4)
                //   → MainForm.c 5 处成对 C2440 "int→vb6_VARIANT" + "vb6_VARIANT→int".
                // 修正: 解析到写方向参数表时, 把尾部恰好等于 Optional 形参个数的那几段
                // (即 _has_* 标志) 摘出来, 将 valArg 插到业务参数末尾、标志之前.
                // 严格前置条件 (段数 == 1 + 业务形参数 + Optional 数) 不满足时原样拼接,
                // 保持既有行为 (090ad 已把 pad 出的占位 Value 弹掉的场景即不满足).
                bool spliced159 = false;
                if (!finalArgs.empty() && letSig159.size() >= 2 && !valuePadDropped159ad) {
                    std::vector<std::string> seg159 = splitTopLevelArgs(finalArgs);
                    size_t nBiz159 = letSig159.size();   // 含末位 Value (其槽位待本处填)
                    size_t optCnt159 = 0;
                    for (const auto& p : letSig159) {
                        if (p.isOptional && !p.isParamArray) optCnt159++;
                    }
                    // finalArgs = [obj] + (业务形参 - Value) + _has_* 标志
                    size_t nFlag159 = (seg159.size() > nBiz159 - 1)
                                          ? (seg159.size() - nBiz159) : 0;
                    bool flagsLookSimple159 = nFlag159 > 0;
                    for (size_t i = seg159.size() - nFlag159;
                         i < seg159.size() && flagsLookSimple159; i++) {
                        for (char ch : seg159[i]) {
                            if (!isdigit(static_cast<unsigned char>(ch)) && ch != '-') {
                                flagsLookSimple159 = false;
                                break;
                            }
                        }
                    }
                    if (nFlag159 > 0 && nFlag159 == optCnt159 && flagsLookSimple159
                        && seg159.size() == (nBiz159 - 1) + nFlag159 + 1) {
                        std::string head159, tail159;
                        for (size_t i = 0; i + nFlag159 < seg159.size(); i++) {
                            if (i) head159 += ", ";
                            head159 += seg159[i];
                        }
                        for (size_t i = seg159.size() - nFlag159;
                             i < seg159.size(); i++) {
                            if (!tail159.empty()) tail159 += ", ";
                            tail159 += seg159[i];
                        }
                        newCall = prefix + newVerbs + afterPg + "(" + head159
                                + ", " + valArg + ", " + tail159 + ")";
                        spliced159 = true;
                    }
                }
                if (!spliced159) {
                    if (finalArgs.empty()) {
                        newCall = prefix + newVerbs + afterPg + "(" + valArg + ")";
                    } else {
                        newCall = prefix + newVerbs + afterPg + "(" + finalArgs + ", " + valArg + ")";
                    }
                }
                std::string tag = isSet ? "Set" : "Let";
                c_.emitLine(newCall + ";  /* Property " + tag + " via prop_get_ rewrite (Pattern C/D2) */");
                return true;
            }
        }
    }

    return false;
}


// ============================================================
// Fix 090w/090x: Property Let/Set 值实参打包 (声明见 cgen.hpp)
// ============================================================
std::string CCodeGen::packLetValueArg(const ParameterInfo& lastP, Expr* valueExpr,
                                      const std::string& val) {
    if (lastP.type != Vb6Type::Variant) {
        // Fix 185: ByRef 值类型形参 (Property Let v As String/Boolean/Long/...)
        // 的 C 侧是 <T>* (BSTR* / int16_t* / int32_t*) — VB6 默认 ByRef.
        // 原先实参只传"值", 被调用方把该值当地址解引用: BSTR 值被当 BSTR* 使用
        // 时读到字符串头 4 字节当指针 → 非确定 0xC0000005 / 堆破坏
        // (cHttpServerSvr.c prop_let_WebRoot; cHttpServer.c CookieAttr.Path/HttpOnly
        //  /SameSite, m_oServer_ConnectionRequest). 用 C99 标量复合字面量取临时
        // 对象地址, 同时兼容非常量左值实参 (常量 (-1)、函数返回值
        // vb6_BSTR_FromStr(L"/")、prop_get_xxx()). 未列出的类型 (Object/UDT/
        // 数组/Date/Currency) 保持原样, 不改变既有行为.
        if (lastP.isByVal) return val;
        switch (lastP.type) {
            case Vb6Type::String:
                if (valueExpr && cExprIsVariant(val)) {
                    return "(&(BSTR){(" + wrapToBSTR(val, *valueExpr) + ")})";
                }
                return "(&(BSTR){(" + val + ")})";
            case Vb6Type::Integer:
            case Vb6Type::Boolean:
                return "(&(int16_t){(" + val + ")})";
            case Vb6Type::Long:
                return "(&(int32_t){(" + val + ")})";
            case Vb6Type::Byte:
                return "(&(uint8_t){(" + val + ")})";
            case Vb6Type::Single:
                return "(&(float){(" + val + ")})";
            case Vb6Type::Double:
                return "(&(double){(" + val + ")})";
            default:
                return val;
        }
    }
    // 装箱走 boxToVariant 这一权威 (按 VB 声明类型选档, 未知才让 C 的 _Generic 猜)
    if (lastP.isByVal) return boxToVariant(valueExpr, val);
    // ByRef Variant 形参需要可寻址的 vb6_VARIANT*.
    // (&(vb6_VARIANT){vb6_VariantFromValue(x)}) 是非法 C (结构体复合字面量
    // 不能用另一个结构值初始化 → C2440 "vb6_VARIANT→vb6_vartype"), 必须按实参
    // VB 类型字段式构造 (与 cgen_expr M22 ByRef Variant 参数分支一致).
    Vb6Type vtT = valueExpr ? inferExprType(*valueExpr) : Vb6Type::Variant;
    switch (vtT) {
        case Vb6Type::String:
            return "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=" + val + "})";
        case Vb6Type::Long:
        case Vb6Type::Integer:
            return "(&(vb6_VARIANT){.vt=VT_I4, .lVal=(int32_t)(" + val + ")})";
        case Vb6Type::Byte:
            return "(&(vb6_VARIANT){.vt=VT_UI1, .bVal=(uint8_t)(" + val + ")})";
        case Vb6Type::Double:
        case Vb6Type::Single:
        case Vb6Type::Date:
        case Vb6Type::Currency:
        case Vb6Type::Decimal:
            return "(&(vb6_VARIANT){.vt=VT_R8, .dblVal=(double)(" + val + ")})";
        case Vb6Type::Boolean:
            return "(&(vb6_VARIANT){.vt=VT_BOOL, .boolVal=(int16_t)(" + val + ")})";
        default:
            // 与 M22 default 分支对齐: 变体表达式提取 BSTR, 其余未知按 BSTR 兜底
            if (cExprIsVariant(val)) {
                return "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=vb6_VariantToString("
                       + val + ")})";
            }
            return "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=" + val + "})";
    }
}
} // namespace vb6c3
