#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>
#include "common/host_pseudo.hpp"

namespace vb6c3 {

// --- cgen_util_com.cpp: COM 值解析 + 打包标记 + 类字段名规范化 ---


// ============================================================
// COM辅助 (P6.2)
// ============================================================

// 账 #255: 「一个 COM 调用的结果被当**对象引用**用」这一问的唯一点。
//
// 这些调用交回的是 VARIANT 指针 (calloc 出来的包装)，不是对象本身；直接 (void*) 硬转
// 得到的是 VARIANT 结构体的地址，于是 `With 集合(1) : .Width = …` 的接收者谁都认不出
// (RTL 侧一路 fall through 到 `vb6_ComSetProp: property "Width" not found`)，写在虚空里。
// Set 那一路早就做了这件事 (改名成 vb6_ComCallObject，内部 Unpack+Free)；With 那一路
// 只认「表达式是 VARIANT **值**」那一档 (Fix 090e / 092j)，漏了「VARIANT* 指针」这一档。
//
// vb6_ComCallObject / vb6_ComGetObjectProp 交回的已经是对象，不在清单里 (再套一层就错)。
// ByDispid 那一路没有一体化出口，只能就地 Unpack —— 那条 VARIANT* 的释放缺口与本刀无关，
// 今天同样在漏 (没有把它交给任何 VarFree)。
std::string CCodeGen::comObjectRefFromCallExpr(const std::string& expr) {
    auto top = [&expr](const char* head) {
        return expr.compare(0, std::char_traits<char>::length(head), head) == 0;
    };
    if (top("vb6_ComCall(")) {
        return "vb6_ComCallObject" + expr.substr(std::char_traits<char>::length("vb6_ComCall"));
    }
    if (top("vb6_ComCallByDispid(") || top("vb6_ComGetProp(")) {
        return "vb6_ComUnpackObject(" + expr + ")";
    }
    // Fix 223: 后期绑定 COM 调用/属性读被包成 VARIANT 值 (Fix 219 的 late-bound 读形态
    // `vb6_VariantFromComResult(vb6_ComCall(...))` / `(vb6_ComGetProp(...))`)。With 目标是
    // 具体类指针 (vb6_cls_X*) 或 Set 属性取对象时, 要的是**对象指针**而非 VARIANT 盒,
    // 故剥掉外层盒子、递归处理内层拿到解包成对象指针的形态:
    //   vb6_VariantFromComResult(vb6_ComCall(a, L"Item", ...))
    //     → 内层 vb6_ComCall(...) → vb6_ComCallObject(a, L"Item", ...)   (正确对象指针)
    //   vb6_VariantFromComResult(vb6_ComGetProp(o, L"m"))
    //     → 内层 vb6_ComGetProp(...) → vb6_ComUnpackObject(vb6_ComGetProp(o, L"m"))
    // 修复前: 此分支不识别该前缀 → 原样返回 → 调用方只补 (vb6_cls_X*) 强转, 把
    // vb6_VARIANT 结构体值当指针 → C2440 "无法从 vb6_VARIANT 转换为 vb6_cls_X *"
    // (CoolBar.c:1547 `With Me.Bands(i)` 等, 约 50 处; Set 属性取对象同款)。
    if (top("vb6_VariantFromComResult(")) {
        std::string inner = expr.substr(std::char_traits<char>::length("vb6_VariantFromComResult("));
        // 剥掉开头的 "vb6_VariantFromComResult(" 后, 末尾多一个闭合 VARIANT 盒的右括号需摘除。
        if (!inner.empty() && inner.back() == ')') {
            inner.pop_back();
            // 内层可能仍以 vb6_VariantFromComResult( 嵌套 (多重 late-bound), 递归到底。
            return comObjectRefFromCallExpr(inner);
        }
        // 非 COM 调用形态 (如 vb6_VariantFromComResult(me->x), me->x 已是 VARIANT 值):
        // 直接用 VariantToObjectVal 取对象指针。
        return "vb6_VariantToObjectVal(" + expr + ")";
    }
    return expr;
}

// 账 #245/§B94: comMethods 的唯一读出口 + 两条档位出口（声明见 cgen_helpers.inc）。
CCodeGen::ComSigView CCodeGen::comSigViewOf(const Symbol* comSym,
                                            const std::string& memberName) const {
    ComSigView sv;
    if (!comSym) return sv;
    auto it = comSym->comMethods.find(Symbol::toLower(memberName));
    if (it == comSym->comMethods.end()) return sv;
    sv.found = true;
    sv.isGet = it->second.isPropertyGet;
    sv.retCType = mapType(it->second.returnType);
    sv.sig = &it->second;
    return sv;
}

std::string CCodeGen::comGetterExpr(const ComSigView& sv, const std::string& objExpr,
                                    const std::string& memberName) const {
    if (!sv.found || !sv.isGet) return "";   // put/method/未登记 ⇒ 调用方走晚绑定
    std::string args = objExpr + ", L\"" + memberName + "\"";
    if (sv.retCType == "BSTR") return "vb6_ComGetStringProp(" + args + ")";
    if (sv.retCType == "int32_t" || sv.retCType == "int16_t") return "vb6_ComGetIntProp(" + args + ")";
    if (sv.retCType == "double" || sv.retCType == "float") return "vb6_ComGetDoubleProp(" + args + ")";
    if (sv.retCType == "void*") return "vb6_ComGetObjectProp(" + args + ")";
    // 类型未知 (Enum→UserDefinedType / 签名压根没读到) ⇒ **不猜目标档**: 交回原生 VARIANT,
    // 让运行期按 VARIANT 自己办。这一档以前有两处各自猜 (一处按 unpackType 猜 Long ⇒ BSTR 属性
    // 发成 vb6_ComGetIntProp，读回是指针宽度当整数用; 一处直接发 StringProp ⇒ 未初始化的
    // returnType 字段落的垃圾值被当成「这是个字符串」当真值用)。
    return "vb6_VariantFromComResult(vb6_ComGetProp(" + args + "))";
}

std::string CCodeGen::comTypedCallExpr(const ComSigView& sv, const std::string& callArgs) const {
    if (sv.found) {
        if (sv.retCType == "BSTR") return "vb6_ComCallBSTR(" + callArgs + ")";
        if (sv.retCType == "int32_t" || sv.retCType == "int16_t") return "vb6_ComCallInt(" + callArgs + ")";
        if (sv.retCType == "double" || sv.retCType == "float") return "vb6_ComCallDouble(" + callArgs + ")";
        if (sv.retCType == "void*") return "vb6_ComCallObject(" + callArgs + ")";
    }
    return "vb6_ComCall(" + callArgs + ")";   // 未知返回类型 / 表里没有: 返回 void*
}

std::string CCodeGen::resolveComValue(const std::string& unpackType) {
    // P24-07: 早期绑定 — 利用签名returnType选择正确的解包函数
    if (!isComMarker_) return lastExpr_;

    std::string objExpr = std::move(comObjExpr_);
    std::string memberName = std::move(comMemberName_);

    // 账 #299 第二十一刀: 这里原来还有一份 `Toolbar1.Buttons` 的抄本 (C29-5b 当时只
    // 有这一处)。现在由下面 memberCollectionObjectExpr 那一处统一答 —— 那条同样排在
    // P24-07 的 early-bound 决策之前，顺序与原意一致。

    // P20-44: OLE 拖放的 **DataObject 形参**成员 —— `Data.GetText` 等。
    // 处理器形参 `void** Data` 不是 IDispatch, 掉 COM 派发运行期必炸/答错。
    // **必须放在 early-bound 分支之前**: DataObject 是 COM 类, early-bound 会先
    // 命中并发 ComGetStringProp (实测踩过, 拦截放后面根本到不了)。
    // DataObject 指针 = *Data, 所以 RTL 调用首参是 `(void*)(*Data)`。
    {
        std::string ddLower = Symbol::toLower(objExpr);
        auto* ddSym = symTab_.lookup(ddLower);
        // **形参**的 `As <类型>` 原文记在 srcTypeName (tB B08c 那条), 变量声明才记
        // variableTypeName —— 两边都看, 否则处理器形参永远判不中 (实测踩过)。
        if (ddSym && !ddSym->srcTypeName.empty()
            && Symbol::toLower(ddSym->srcTypeName) == "dataobject") {
            std::string memDD = Symbol::toLower(memberName);
            std::string ddArg = "(void*)(*" + objExpr + ")";
            if (memDD == "gettext")      { lastExpr_ = "vb6_oleDD_GetText(" + ddArg + ")"; isComMarker_ = false; return lastExpr_; }
            if (memDD == "getfilecount") { lastExpr_ = "vb6_oleDD_GetFileCount(" + ddArg + ")"; isComMarker_ = false; return lastExpr_; }
        }
    }

    // 账 #278 §B124: recordset 那两格直译 (认 `vb6_Data_FieldValueStr(` / `vb6_Data_Self(`
    // 前缀) 撤掉 —— 发射者先没了，前缀判据恒假。`.RecordCount` / `.BOF` 这些标量成员
    // 今天由 memberobj 的真 IDispatch 应答，成员类型走既有的 COM 型别表
    // (`vb6_ComGetIntProp` / `vb6_ComGetStringProp`，九形探针量过)。

    // 账 #299 第二十一刀: 成员集合的创建式只由 memberCollectionObjectExpr 一处答
    // (这张表以前在本函数里各写一遍，四族一套、Panels 那一族根本没有)。
    // 位置留在原来那几段的地方: 必须排在 P20-39 / P20-40 两段 recognizer **之前** ——
    // 那两段管链上更外层的成员 (ListImages.Count / Panels(i).Width)，本段先把**集合本身**
    // 立起来；两边按成员名匹配，六个集合名与那两段的成员名互不相交，所以顺序不敏感。
    {
        std::string collObj = memberCollectionObjectExpr(objExpr, memberName);
        if (!collObj.empty()) {
            lastExpr_ = collObj;
            isComMarker_ = false;
            return lastExpr_;
        }
    }
    // C29-OLE: `OLE1.Object` → 嵌入对象的 IDispatch (真 OLE 容器)。
    // 其它属性 (Class/OLEType/SizeMode…) 走属性表 (cgen_util_ctrl.cpp), 不在这拦。
    if (Symbol::toLower(memberName) == "object") {
        std::string ocBare = oleConNameOfExpr(objExpr);
        if (!ocBare.empty()) {
            lastExpr_ = "vb6_OleCon_GetObject((void*)vb6_hwnd_" + ocBare + ")";
            isComMarker_ = false;
            return lastExpr_;
        }
    }

    // P20-39: ImageList 原生复刻 —— 把集合/属性读改道到 RTL。
    // 放在两个 COM 分支**之前**, 否则 Count 会走默认 BSTR 解包、Key 会走 ComCallObject。
    {
        std::string memLower = memberName;
        std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
        std::string slot = imageListSlotVarOfExpr(objExpr);
        if (slot.empty()) { }
        else if (memLower == "count") {
            lastExpr_ = "vb6_GetImageListCount((void*)" + slot + ")";
            isComMarker_ = false;
            return lastExpr_;
        } else {
            // ListImages(i).Key / .Index: 从 Item 的第一个实参里抠下标 (只支持字面量)
            std::string num;
            std::string argTxt;   // Item 的实参原文, 下标/Key 两路都要用, 提到外层留着
            size_t ia = objExpr.find("L\"Item\", (void*[]){");
            if (ia != std::string::npos) {
                size_t open = ia + strlen("L\"Item\", (void*[]){");
                size_t close = objExpr.find('}', open);
                argTxt = objExpr.substr(open, close - open);
                // 形如 "vb6_ComPackInt(1)" 或直接的 "1": 抠 '(' 之后到 ')' / ',' 之前
                size_t lp = argTxt.find('(');
                if (lp != std::string::npos) {
                    size_t end = argTxt.find_first_of("),", lp + 1);
                    num = (end == std::string::npos) ? argTxt.substr(lp + 1)
                                                     : argTxt.substr(lp + 1, end - lp - 1);
                } else {
                    num = argTxt;
                }
                bool isNum = !num.empty();
                for (char c : num) if (!isdigit((unsigned char)c)) { isNum = false; break; }
                if (!isNum) num.clear();
            }
            // `ListImages("SomeKey")` 是 VB6 的按 Key 取项, 实参是宽字符串不是下标。
            // 同一条 Item 实参两种形态都得认, 否则掉回假 IDispatch 路径。
            std::string keyLit;
            if (num.empty()) {
                size_t qs = argTxt.find("L\"");
                if (qs != std::string::npos) {
                    size_t qe = argTxt.find('"', qs + 2);
                    if (qe != std::string::npos) {
                        keyLit = argTxt.substr(qs + 2, qe - qs - 2);
                        // 转义引号收尾 (VB6 `""` 在 C 里就是 `\"`)
                        if (!keyLit.empty() && keyLit.back() == '\\') keyLit.pop_back();
                    }
                }
            }
            if (!num.empty() && memLower == "key") {
                lastExpr_ = "vb6_GetImageListKeyAt((void*)" + slot + ", " + num + ")";
                isComMarker_ = false;
                return lastExpr_;
            }
            if (!num.empty() && memLower == "index") {   // ListImage.Index (1 基)
                lastExpr_ = "vb6_ImageListIndexAt((void*)" + slot + ", " + num + ")";
                isComMarker_ = false;
                return lastExpr_;
            }
            if (!keyLit.empty() && memLower == "key") {
                lastExpr_ = "vb6_GetImageListKeyByKey((void*)" + slot
                          + ", (const wchar_t*)(vb6_BSTR_FromStr(L\"" + keyLit + "\")))";
                isComMarker_ = false;
                return lastExpr_;
            }
            if (!keyLit.empty() && memLower == "index") {
                lastExpr_ = "vb6_ImageListIndexByKey((void*)" + slot
                          + ", (const wchar_t*)(vb6_BSTR_FromStr(L\"" + keyLit + "\")))";
                isComMarker_ = false;
                return lastExpr_;
            }
        }
    }

    // P20-40: StatusBar 原生复刻 —— Panels.Count 与 Panels(i).成员 改道到 RTL。
    // 与上面 ImageList 分支同一套手法, 只是槽是 HWND、下标要从 Item 实参里抠。
    {
        std::string memLower = memberName;
        std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
        std::string hwnd = statusBarHwndVarOfExpr(objExpr);
        if (!hwnd.empty()) {
            if (memLower == "count") {
                lastExpr_ = "vb6_StatusBar_GetPanelsCount((void*)" + hwnd + ")";
                isComMarker_ = false;
                return lastExpr_;
            }
            // 抠 Item 的实参: 字面量下标, 或 `Panels("Key")` 那种宽字符串 Key。
            // 注意没有独立的 `L"Item"` 层 —— Panels 被生成器当**默认成员**, 索引跟着
            // 它那一层的 `(void*[]){…}` 走, 所以按参数包抠再剥 vb6_ComPackInt(...)。
            std::string idx;
            std::string keyLit;
            const std::string kArr = "(void*[]){";
            size_t ia = objExpr.find(kArr);
            if (ia != std::string::npos) {
                size_t open = ia + kArr.size();
                size_t close = objExpr.find('}', open);
                if (close != std::string::npos) {
                    std::string argTxt = objExpr.substr(open, close - open);
                    size_t lp = argTxt.find('(');
                    if (lp != std::string::npos) {
                        size_t end = argTxt.find_first_of("),", lp + 1);
                        idx = (end == std::string::npos) ? argTxt.substr(lp + 1)
                                                         : argTxt.substr(lp + 1, end - lp - 1);
                    } else {
                        idx = argTxt;
                    }
                    bool isNum = !idx.empty();
                    for (char c : idx) if (!isdigit((unsigned char)c)) { isNum = false; break; }
                    if (!isNum) idx.clear();
                    if (idx.empty()) {
                        size_t qs = argTxt.find("L\"");
                        if (qs != std::string::npos) {
                            size_t qe = argTxt.find('"', qs + 2);
                            if (qe != std::string::npos)
                                keyLit = argTxt.substr(qs + 2, qe - qs - 2);
                        }
                    }
                }
            }
            auto sbGet = [&](const char* fn) {
                return std::string(fn) + "((void*)" + hwnd + ", " + idx + ")";
            };
            // Key 形态走 *ByKey 两路之一, 没抠出下标就退回按下标那版 (下标为空串时
            // RTL 会越界, 宁可编译期宁可什么都不发也别发坏代码 —— 这里退化成按 Key 查)。
            auto sbFinishByKey = [&](const char* byIdx, const char* byKey) {
                isComMarker_ = false;
                if (!keyLit.empty())
                    lastExpr_ = std::string(byKey) + "((void*)" + hwnd
                              + ", (const wchar_t*)(vb6_BSTR_FromStr(L\"" + keyLit + "\")))";
                else
                    lastExpr_ = std::string(byIdx) + "((void*)" + hwnd + ", " + idx + ")";
                return lastExpr_;
            };
            if (memLower == "key")   { lastExpr_ = sbFinishByKey("vb6_StatusBar_GetPanelKey",
                                                                  "vb6_StatusBar_GetPanelKeyByKey");
                                       isComMarker_ = false; return lastExpr_; }
            if (memLower == "text")  { lastExpr_ = sbFinishByKey("vb6_StatusBar_GetPanelText",
                                                                  "vb6_StatusBar_GetPanelTextByKey");
                                       isComMarker_ = false; return lastExpr_; }
            // 账 #300: 这两个参数以前写的是**同一个名字**（都由 *ByKey 答）⇒ 数字下标那一形
            // 把整数递进了 `const wchar_t*` 槽, 真跑启动期 AV。数字/键两条各答各的。
            if (memLower == "index") { lastExpr_ = sbFinishByKey("vb6_StatusBar_GetPanelIndex",
                                                                  "vb6_StatusBar_GetPanelIndexByKey");
                                       isComMarker_ = false; return lastExpr_; }
            if (memLower == "width")        { lastExpr_ = sbGet("vb6_StatusBar_GetPanelWidth");
                                              isComMarker_ = false; return lastExpr_; }
            if (memLower == "minwidth")     { lastExpr_ = sbGet("vb6_StatusBar_GetPanelMinWidth");
                                              isComMarker_ = false; return lastExpr_; }
            if (memLower == "autosize")     { lastExpr_ = sbGet("vb6_StatusBar_GetPanelAutoSize");
                                              isComMarker_ = false; return lastExpr_; }
            if (memLower == "style")        { lastExpr_ = sbGet("vb6_StatusBar_GetPanelStyle");
                                              isComMarker_ = false; return lastExpr_; }
            if (memLower == "tooltiptext")  { lastExpr_ = sbGet("vb6_StatusBar_GetPanelToolTip");
                                              isComMarker_ = false; return lastExpr_; }
            // 账 #206 第二格: Panel.Left 与 Width 同源 (都是排版结果的投影)。语料里有两处
            // 按面板定位覆盖层的写法在问它, 而改前全仓 0 个出口。
            if (memLower == "left")         { lastExpr_ = sbGet("vb6_StatusBar_GetPanelLeft");
                                              isComMarker_ = false; return lastExpr_; }
        }
    }

    // P24-07: 早期绑定推断 — 利用TypeLib签名的returnType决策
    if (isEarlyBoundCom_ && earlyBoundSym_) {
        isEarlyBoundCom_ = false;
        const Symbol* comSym = earlyBoundSym_;
        earlyBoundSym_ = nullptr;
        // 账 #245/§B94: 读表与档位选择都收在 comSigViewOf/comGetterExpr —— 只有确认的属性
        // getter 才按签名走带类型出口，put/method 那份签名一律不当属性读来解包。
        std::string getter = comGetterExpr(comSigViewOf(comSym, memberName), objExpr, memberName);
        if (!getter.empty()) {
            lastExpr_ = getter;
            isComMarker_ = false;
            return lastExpr_;
        }
    }
    isEarlyBoundCom_ = false;
    isComMarker_ = false;

    std::string getPropArgs = objExpr + ", L\"" + memberName + "\"";

    if (unpackType == "BSTR") {
        lastExpr_ = "vb6_ComGetStringProp(" + getPropArgs + ")";
    } else if (unpackType == "LongPtr") {
        lastExpr_ = "vb6_ComGetLongPtrProp(" + getPropArgs + ")";
    } else if (unpackType == "Int" || unpackType == "Long" || unpackType == "Boolean") {
        lastExpr_ = "vb6_ComGetIntProp(" + getPropArgs + ")";
    } else if (unpackType == "Double" || unpackType == "Single") {
        lastExpr_ = "vb6_ComGetDoubleProp(" + getPropArgs + ")";
    } else if (unpackType == "Object") {
        lastExpr_ = "vb6_ComGetObjectProp(" + getPropArgs + ")";
    } else if (unpackType == "Variant") {
        // P24-02: COM属性返回原生VARIANT(如dic.Keys/dic.Items返回SAFEARRAY)
        lastExpr_ = "vb6_VariantFromComResult(vb6_ComCall(" + objExpr + ", L\"" + memberName + "\", NULL, 0))";
    } else {
        // 默认: BSTR解封 (最通用, COM VARIANT → BSTR自动转换)
        lastExpr_ = "vb6_ComGetStringProp(" + getPropArgs + ")";
    }
    return lastExpr_;
}


// Fix 092m: 目标类字段类型 → COM 解包类型 hint (见 cgen.hpp 声明注释).
// 从当前模块作用域的 Class 符号 (跨模块 external Class 已由 driver 拷贝该表) 取
// memberFieldTypes[字段名] 并映射; 任何缺失都回退 "BSTR" (comGetStringProp 最通用,
// 与 Fix 092m 之前的默认行为一致, 不引入回归).
std::string CCodeGen::classFieldComUnpackHint(const std::string& className,
                                              const std::string& memberName) const {
    // Fix 110g: RTL 宿主 Font 结构 (vb6_ComIface_Font / vb6_cls_StdFont, 见
    // rtl/core/vb6rtl/vb6rtl_userctl.h) 不在符号表里 —— `With UserControl.Font`
    // 的字段写 (Charts 2020 各 UserControl 的 Property Set Font) 走本 helper 时
    // 全部回退 "BSTR" → `.Size = vb6_ComGetStringProp(...)` 把 wchar_t* 赋给
    // float (C2440), `.Bold` 同理. 按该结构的真实成员类型给出解包类型.
    {
        const std::string cls = Symbol::toLower(className);
        if (cls == "stdfont" || cls == "font" || cls == "vb6_cls_stdfont"
            || cls == "vb6_comiface_font") {
            const std::string fld = Symbol::toLower(memberName);
            if (fld == "name") return "BSTR";
            if (fld == "size") return "Double";
            if (fld == "bold" || fld == "italic" || fld == "underline"
                || fld == "strikethrough" || fld == "weight" || fld == "charset") {
                return "Long";
            }
        }
    }
    if (className.empty() || !symTab_.moduleScope()) return "BSTR";
    const std::string want = Symbol::toLower(className);
    const std::string fld = Symbol::toLower(memberName);
    for (const auto& kv : symTab_.moduleScope()->symbols()) {
        const Symbol* cs = kv.second.get();
        if (!cs || cs->kind != SymbolKind::Class) continue;
        if (Symbol::toLower(cs->name) != want) continue;
        auto it = cs->memberFieldTypes.find(fld);
        if (it == cs->memberFieldTypes.end()) return "BSTR";
        const std::string tn = Symbol::toLower(it->second);
        if (tn == "string") return "BSTR";
        if (tn == "long" || tn == "integer" || tn == "boolean" || tn == "byte") return "Long";
        if (tn == "single" || tn == "double" || tn == "date" || tn == "currency") return "Double";
        if (tn == "object") return "Object";
        if (tn == "variant" || tn == "var") return "Variant";
        // 命名类型 (项目类/UDT) → 对象解包; 未知名字回退 BSTR
        return symTab_.lookup(it->second) ? "Object" : "BSTR";
    }
    return "BSTR";
}


// Fix 092p: 类数据字段名规范化 (见 cgen.hpp 声明注释).
std::string CCodeGen::canonicalClassFieldName(const std::string& className,
                                              const std::string& memberName) const {
    if (className.empty() || !symTab_.moduleScope()) return memberName;
    const std::string want = Symbol::toLower(className);
    const std::string fld = Symbol::toLower(memberName);
    for (const auto& kv : symTab_.moduleScope()->symbols()) {
        const Symbol* cs = kv.second.get();
        if (!cs || cs->kind != SymbolKind::Class) continue;
        if (Symbol::toLower(cs->name) != want
            && Symbol::toLower(cs->sourceModule) != want) continue;
        auto it = cs->memberFieldNames.find(fld);
        return (it != cs->memberFieldNames.end()) ? it->second : memberName;
    }
    return memberName;
}


// 成员名规范化 (旧 Fix <vbeclipse> 的独立小表并入本表): RTL 符号名是手写约定
// (vb6_UserControl_hDC), VB6 源码拼写可以不同 (`UserControl.hDc`), 而 C 大小写
// 敏感 → C2065 (ucButton.ctl:64 / ucCaption.ctl:119 / ucTab.ctl:92 实测写 hDc)。
// 表中无此项 (非宿主伪对象 / 本表刻意不收的成员) 时原样返回。
std::string CCodeGen::canonicalHostPseudoMember(const std::string& pseudoObj,
                                                const std::string& memberName) const {
    const HostPseudoRow* r = hostPseudoFind(pseudoObj, memberName);
    return r ? std::string(r->rtl) : memberName;
}

// 账 #278 §B120: 宿主伪成员的 RTL 全局符号 `vb6_<对象>_<成员>` **只在这里装配一次**。
// 从前这句在三处各拼一遍 (With 块 / 赋值 / 裸名)，每抄一次就多一个权威：第十刀修掉的
// `vb6_PropertyPage_hwnd` 就是其中一份抄错大小写的 (RTL 两枚都定义了，所以它编得过，只是编向
// 另一枚全局)。对象那一段与成员那一段现在都出自那张表；表里没有这一行时按源码成员名发
// (与改前逐字节一致，不新增也不撤销任何 C2065)。名字不是宿主伪对象时交回空串 —— 三个调用点
// 都先过了 hostPseudoIsObject 那道闸，空串意味着接线错了，不是运行期会遇到的输入。
std::string CCodeGen::hostPseudoRtlSymbol(const std::string& pseudoObj,
                                          const std::string& memberName) const {
    std::string objSym;
    if (!hostPseudoObjectSymbol(pseudoObj, objSym)) return std::string();
    const HostPseudoRow* r = hostPseudoFind(pseudoObj, memberName);
    return "vb6_" + objSym + "_" + cIdent(r ? r->rtl : memberName);
}

// Fix 296 / 4a: `Extender.<成员>` 作**右值**时的实例感知读入口 —— 写侧 (3a) 的镜像。
//
// 通道 (vb6_UC_ExtenderField) 已经改指 `vb6_UCRec.ext*`, 而读侧发的一直是裸全局
// `vb6_Extender_X` —— push/pop 把它同步着, 于是 getter 读到的是副本而不是当前实例记录。
// ComCtlsDemo 实测 333 处 Property Get 返回槽走的是这枚副本。
//
// 闸门照 3a 的形状: 判**发射符号的前缀**, 不判 hostPseudoBareName(它不覆盖 Extender
// 成员, 拿它当闸门会静默一个都匹配不到 —— 3a 的第一版就这么踩过, 生成码 bare=9 new=0)。
// 符号后缀正是表里那一行的 rtl 列 (hostPseudoRtlSymbol 拼的就是它), 故按 rtl 反查,
// 而不是按 VB6 拼写反查 —— `vb6_Default` 那行 name/rtl 本来就是两种拼法。
//
// 两个入口而不是一个: 读要参与算术与字符串表达式, void* getter 会逼每个调用点写强制转换。
// 数值/字符串的选择只问表里那一行的类型 (kind 约定 0=BSTR / 2=int16 / 3=int32, 见
// vb6_UC_ExtenderField); `Unknown` 的行 (对象成员 Container/DragIcon、方法行
// Drag/SetFocus/ZOrder) 刻意原样返回 —— 两个入口都不递对象与句柄, 硬换会把 void* 槽
// 接成 int32_t。Container/DragIcon 属另一笔账, 本轮不碰。
std::string CCodeGen::extenderMemberReadExpr(const std::string& rtlSymbol) const {
    if (rtlSymbol.rfind("vb6_Extender_", 0) != 0) return rtlSymbol;
    const std::string suffix = rtlSymbol.substr(std::string("vb6_Extender_").size());
    if (suffix.empty()) return rtlSymbol;
    for (const HostPseudoRow& r : kHostPseudoRows) {
        if (hostPseudoLower(r.obj) != "extender") continue;
        if (hostPseudoLower(r.rtl) != hostPseudoLower(suffix)) continue;
        /* Fix 296 / 3b 最后一格: 表里 type=Unknown 的行分**两种**, 不能只看 type ——
           Drag/SetFocus/ZOrder 是 HPF_METHOD, 是**调用**不是读 (换过去会发成一个没人
           接收的 getter 调用); Container/DragIcon 才是对象成员, 换道到对象入口。
           少了这道方法行闸, 三个方法会跟着对象成员一起被换掉。 */
        if (r.flags & HPF_METHOD) return rtlSymbol;
        const char* dock296;
        if (r.type == Vb6Type::String)          dock296 = "vb6_UC_GetExtenderMemberStr";
        else if (r.type == Vb6Type::Unknown)    dock296 = "vb6_UC_GetExtenderMemberObj";
        else                                    dock296 = "vb6_UC_GetExtenderMemberNum";
        return std::string(dock296) + "(L\"" + escapeWideCString(r.rtl) + "\")";
    }
    return rtlSymbol;
}

// "这个名字是不是宿主伪对象"由那张表答 (这一句从前在发码侧抄了五份)。
bool CCodeGen::hostPseudoIsObject(const std::string& name) const {
    return hostPseudoObjectKnown(name);
}

// 裸名能不能解析成宿主伪成员 (.ctl/.pag 里直接写 hDC / ScaleWidth / Changed …)。
bool CCodeGen::hostPseudoBareName(const std::string& pseudoObj,
                                  const std::string& memberName,
                                  std::string& outRtlName) const {
    const HostPseudoRow* r = hostPseudoFind(pseudoObj, memberName);
    if (!r || !(r->flags & HPF_BARE)) return false;
    outRtlName = r->rtl;
    return true;
}

// 按值读时的类型答案。requireBare=true 用于**裸名**那一路 —— 必须与发射侧同口径
// (都认 HPF_BARE), 否则类型 oracle 会给一个发射时根本解析不出宿主符号的名字,
// 那正是"两条路答不一样"的起源。方法成员与对象成员 (Unknown) 一律不答。
bool CCodeGen::hostPseudoValueType(const std::string& pseudoObj,
                                   const std::string& memberName,
                                   Vb6Type& outType, bool requireBare) const {
    const HostPseudoRow* r = hostPseudoFind(pseudoObj, memberName);
    if (!r) return false;
    if (requireBare && !(r->flags & HPF_BARE)) return false;
    if (r->flags & HPF_METHOD) return false;
    if (r->type == Vb6Type::Unknown) return false;
    outType = r->type;
    return true;
}

// 赋值时 RHS 是 Variant 要不要先拆成数值。口径 = 目标槽是**整数** (Long/Boolean):
// 指针/对象成员 (LongPtr/Unknown) 不拆, 与改动前那份名单一致 —— 今天往 hDC/hWnd
// 这类句柄槽写 Variant 的写法在 VB6 里并不存在, 拆成 int32 反而会把指针截断。
// 字符串成员也不拆 (BSTR 槽直传)。
bool CCodeGen::hostPseudoIsNumeric(const std::string& pseudoObj,
                                   const std::string& memberName) const {
    Vb6Type t = Vb6Type::Unknown;
    if (!hostPseudoValueType(pseudoObj, memberName, t, false)) return false;
    return t == Vb6Type::Long || t == Vb6Type::Boolean
        || t == Vb6Type::Integer || t == Vb6Type::Byte;
}


// Fix 093a: 当前类是否声明了同名成员字段 — 裸标识符赋值 (`field = value`) 时,
// 本类字段优先于从全局符号表捡到的外部同名 Property Let/Set (VB6 里同一类中
// 字段与属性不可能同名). 典型: cClientCallback.cls 的 `recvBuffer = data`
// (recvBuffer 是本类 Public 字段) 被误命中 cWinsock 的 Property Let RecvBuffer,
// 生成 vb6_cWinsock_prop_let_recvBuffer((void*)me, ...) → LNK2019.
bool CCodeGen::isOwnClassField(const std::string& memberName) const {
    if (!isClassModule_ || moduleName_.empty() || !symTab_.moduleScope()) return false;
    const std::string want = Symbol::toLower(moduleName_);
    const std::string fld = Symbol::toLower(memberName);
    for (const auto& kv : symTab_.moduleScope()->symbols()) {
        const Symbol* cs = kv.second.get();
        if (!cs || cs->kind != SymbolKind::Class) continue;
        if (Symbol::toLower(cs->name) != want
            && Symbol::toLower(cs->sourceModule) != want) continue;
        return cs->memberFieldNames.count(fld) > 0;
    }
    return false;
}


// Fix 093a: 类成员(方法/属性)名规范化 — VB6 大小写不敏感, 类内声明与调用点拼写
// 可能不同 (类里声明 `Count` 而调用点写 `count`; 或类里是 `test` 调用点写 `Test`),
// 而 C 符号大小写敏感: 生成的函数名与类定义不一致 → LNK2019. 以 Class 符号
// memberNames 中的声明拼写为准 (与 cIdent 规范化类名同理). 表中无此项时原样返回.
std::string CCodeGen::canonicalClassMemberName(const std::string& className,
                                               const std::string& memberName) const {
    if (className.empty() || !symTab_.moduleScope()) return memberName;
    const std::string want = Symbol::toLower(className);
    const std::string mem = Symbol::toLower(memberName);
    for (const auto& kv : symTab_.moduleScope()->symbols()) {
        const Symbol* cs = kv.second.get();
        if (!cs || cs->kind != SymbolKind::Class) continue;
        if (Symbol::toLower(cs->name) != want
            && Symbol::toLower(cs->sourceModule) != want) continue;
        for (const auto& mn : cs->memberNames) {
            if (Symbol::toLower(mn) == mem) return mn;
        }
        return memberName;
    }
    return memberName;
}


std::string CCodeGen::comPackExpr(Expr& expr) {
    // Fix <vbeclipse>: **UDT (Type ... End Type) 实参**必须走字节数组编组。
    // `m_Rect As RECT` 传给 COM 方法时, inferExprType 推不出标量类型 → 落
    // default 分支 → vb6_ComPackInt(me->m_Rect) → C2440 "无法从 vb6_type_RECT
    // 转换为 int32_t" (ucSplitBar.ctl:149 `.SplitterMouseDown UserControl.hWnd,
    // m_Rect, x, y`)。
    //
    // 判据用 inferUdtTypeOfExpr (而非 inferExprType): 它查 knownUdtVars_ 与 UDT
    // 成员表, 能认出 "这是一个 UDT 结构体值"。
    //
    // ⚠ 这里只返回**函数名** (本函数的契约), 实参由各调用点自己拼 —— 所以走
    // 宏: vb6_ComPackUdt(x) 内部自己做 &x 与 sizeof(x)。先前试图在返回串里嵌
    // 实参 (lastExprOrSelf_) 是错的方向: 那样 10+ 个调用点会各拼一次实参。
    if (inferUdtTypeOfExpr(expr).rfind("vb6_type_", 0) == 0) return "vb6_ComPackUdt";

    // Fix 110p: 标识符的 C 层跟踪集合优先于 inferExprType. 声明为 Collection/Object
    // 的变量 (C 类型 void*) 会被 inferExprType 误判为 Double → 生成
    // vb6_ComPackDouble(void*) → C2440. 实测 Charts 2020 Form2.frm 524:
    //   ucTreeMaps1.AddLineSeries vbNullString, vbBlue, Value, Lables
    //   (Value / Lables As Collection) → vb6_ComPackDouble(Value) 等.
    if (expr.kind == ASTNodeKind::IdentifierExpr) {
        auto& id = static_cast<IdentifierExpr&>(expr);
        std::string lower = id.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        // 账 #256: 裸控件名在发码那条路上会被**折成默认属性读数**（一枚标量/BSTR），
        // 而 inferExprType 与 knownObjectVars_ 对这个名字都答 Object ⇒ 以前一路走到
        // 下面的 ComPackObject：`vb6_ComPackObject(vb6_GetControlText(...))` 把 BSTR 当
        // IDispatch 装进 VT_DISPATCH，栈上 VARIANT 收尾无条件 Release ⇒ 按 BSTR 头几字节
        // 解 vtable = 0xC0000005（实测 `coll.Add txtA, "k1"`，探针 .build/p256 两架构同崩）。
        // 现在打包与折叠问同一条出口 ctrlDefaultPropOf ⇒ 装箱档跟着**实际交出的值**走。
        // 这一问必须排在下面那串"已知变量型"之前：控件名同时登记在 knownObjectVars_ 里
        // （实测：把它放在那串之后，这一格一条都不变）。反过来的担心是"局部变量与控件同名"
        // —— knownFormControls_ 只由 .frm 的控件清单填，永远不含用户变量；真同名的话
        // 发码走的是变量那条路，而打包会按默认属性的型装 ⇒ 编译期 C2440 响，不会静默错值。
        // 刻意只接管能明确映射的那几型：Picture（PictureBox / Image 的默认属性）在 defaultPropType
        // 里答 Unknown ⇒ 继续走今天那条 ComPackObject 路（那本来就是个对象，装箱成对象是对的），
        // 本刀不动它。
        {
            CtrlDefaultProp fold256 = ctrlDefaultPropOf(lower);
            if (fold256.folds) {
                switch (fold256.valueType) {
                    case Vb6Type::String: case Vb6Type::Integer: case Vb6Type::Long:
                    case Vb6Type::Boolean: case Vb6Type::Single: case Vb6Type::Double:
                        return comPackFnForVbType(fold256.valueType);
                    default:
                        break;
                }
            }
        }
        // 账 #305 §B159 (第五十刀): **数组本体当 COM 实参** —— 这一问必须排在下面那串「已知变量型」
        // 之前, 因为模块级数组同时登记在 knownLongVars_/knownBstrVars_ 里, 按它们打包交出去的是
        // 描述符**指针**的标量档。实测 (探针 .build/k50probe, Scripting.Dictionary 写后读回,
        // x64/x86 两台同形):
        //   mL() As Long      -> vt=3 "Long"    (指针折进 int32, x64 连高位都丢)
        //   mBy() As Byte     -> vt=3 "Long"
        //   mFx(1 To 3) As Long -> vt=3 "Long"
        //   mSt() As String   -> vt=8 "String"  (描述符指针被当 BSTR 交给 SysAllocString = 越界读)
        // 而 VB6 交的是 8192|元素码 的一枚真 SAFEARRAY。改交 vb6_ComPackValue(载体指针): _Generic 认
        // vb6_SafeArray1D* 那一档 (第四十六刀的载体出口, 标签问描述符自己), 真数组由
        // vb6_VariantPackBorrow 的数组臂现造。局部动态数组今天已经落 ComPackValue, 实测读数
        // 0/"Empty" —— 它缺的是运行期那一臂, 不是这里的分派。
        // ⚠ 只接管**声明成一维**的载体: vb6_VariantFromValue 那张 _Generic 没有 vb6_SafeArrayND*
        //   那一档, 二维载体进去会落 default = vb6_VariantObject ⇒ 对端 Release 一枚描述符 = 野 Release
        //   堆损坏 (账 #256 §B85 同形)。那一格归 §B151, 补法是给 _Generic 加 ND 档, 不是在这里绕。
        if (isArrayCarrierArgName(lower, id.name)) {
            auto it50 = arrayDimCounts_.find(lower);
            if (it50 != arrayDimCounts_.end() && it50->second <= 1) return "vb6_ComPackValue";
        }
        // Fix 110v: knownVariantVars_ 先于 knownObjectVars_ 判定. `Dim Value As Variant`
        // 的局部变量在 For Each 语境下同时被登记进 knownObjectVars_ (旧代码把
        // For-Each 元素一律当对象), 使 Variant 变量走 vb6_ComPackObject →
        // C2172 "实参不是指针" (Form2.c 375: NewCollection.Add Value).
        // vb6_ComPackValue 是 _Generic 路由宏 (VARIANT→identity / void*→Object),
        // 对两种情况都正确, 因此优先它是安全的.
        if (knownVariantVars_.count(lower)) return "vb6_ComPackValue";
        if (knownObjectVars_.count(lower)) return "vb6_ComPackObject";
        if (knownTypedComVars_.count(lower)) return "vb6_ComPackObject";  // Fix 160-com-byref: 早期绑定COM变量也是对象 (void*/接口指针), 按对象打包
        if (knownBstrVars_.count(lower)) return "vb6_ComPackBSTR";
        if (knownDoubleVars_.count(lower)) return "vb6_ComPackDouble";
        if (knownLongVars_.count(lower)) return "vb6_ComPackInt";
    }
    // ExeComBridge 03: 本工程类的 `New <类>` 实参.
    //   cgen_expr.cpp visit(NewExpr) 对它生成裸实例指针 (vb6_cls_X_New()); 通用
    //   打包 (vb6_ComPackValue → vb6_VariantFromValue) 会把它当 VT_DISPATCH 原样
    //   传出, 对端取值/释放时 AddRef, 结构体首字段被当 vtable → 0xC0000005
    //   (VBMAN_DEMO: .Router.Reg "Demo", New bHello).
    //   改走按类生成的 vb6_ComPack_<类> (先经 __comObj 包装成真 IDispatch 再转
    //   VARIANT, 见 cgen_com.cpp emitClassFactory / RTL vb6_ComPackVB6InstanceRaw).
    //   只作用于本工程类: ComClass 分支走 vb6_NewObject, 拿到的本来就是真
    //   IDispatch, 保持原路径不动 (收窄原则 —— 泛用实参包装改动曾致回归 30→81).
    if (expr.kind == ASTNodeKind::NewExpr) {
        auto& ne = static_cast<NewExpr&>(expr);
        Symbol* neCls = lookupModuleDotted(ne.className);
        if (neCls && neCls->kind == SymbolKind::Class) {
            return "vb6_ComPack_" + cIdent(neCls->name);
        }
    }
    // Fix: Empty/Null 字面量语义上是 Variant 子类型 (VT_EMPTY/VT_NULL),
    // 但 inferExprType 缺省把它们当作 Long (LiteralExpr 分支 return Vb6Type::Long),
    // 于是省略实参 (parser 填 LiteralKind::Empty) 在晚绑定 COM 调用里生成
    // vb6_ComPackInt(vb6_VariantEmpty()) → C2440 (vb6_VARIANT 无法转 int32_t).
    // 例: Charts 2020 ucTreeMaps.ctl `cValues.Add vTemp, , i` 的省略 Key.
    if (expr.kind == ASTNodeKind::LiteralExpr) {
        auto& lit = static_cast<LiteralExpr&>(expr);
        if (lit.literalKind == LiteralKind::Empty || lit.literalKind == LiteralKind::Null)
            return "vb6_ComPackValue";
    }
    // 根据表达式类型推断应该用的VARIANT封装函数
    Vb6Type vt = inferExprType(expr);
    {
        std::string packForVt = comPackFnForVbType(vt);
        if (!packForVt.empty()) return packForVt;
    }
    // ---- 以下是原 switch 的 default 分支（答不了型时的兜底），一字未改 ----
    // Variant/未知: 尝试用BSTR封装 (运行时会处理转换)
    // 更安全的做法: 检查已知变量类型
    if (expr.kind == ASTNodeKind::IdentifierExpr) {
        auto& id = static_cast<IdentifierExpr&>(expr);
        std::string lower = id.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownObjectVars_.count(lower)) return "vb6_ComPackObject";
        if (knownBstrVars_.count(lower)) return "vb6_ComPackBSTR";
        if (knownDoubleVars_.count(lower)) return "vb6_ComPackDouble";
        if (knownLongVars_.count(lower)) return "vb6_ComPackInt";
        if (knownVariantVars_.count(lower)) return "vb6_ComPackVariant";
    }
    if (expr.kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<MemberAccessExpr&>(expr);
        if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& objId = static_cast<IdentifierExpr&>(*ma.object);
            std::string objLower = objId.name;
            std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
            if (knownVariantVars_.count(objLower)) return "vb6_ComPackVariant";
        }
    }
    return "vb6_ComPackInt";  // 默认整数封装
}


// 账 #256: Vb6Type → 打包函数名，comPackExpr 的类型分支与「裸控件名折成默认属性」那条
// 折叠共用这一张映射（折叠那边只问 String/Integer/Long/Boolean/Single/Double 六档）。
// 空串 = 这一型在这里答不了，交回调用方兜底 —— 与改前 switch 的 default 分支同一语义。
std::string CCodeGen::comPackFnForVbType(Vb6Type vt) {
    switch (vt) {
        case Vb6Type::String:  return "vb6_ComPackBSTR";    // BSTR → VARIANT
        case Vb6Type::Integer:
        case Vb6Type::Long:    return "vb6_ComPackInt";     // int32_t → VARIANT
        case Vb6Type::Boolean: return "vb6_ComPackBool";    // VB6 Boolean → VARIANT VT_BOOL
        case Vb6Type::Single:
        case Vb6Type::Double:  return "vb6_ComPackDouble";  // double → VARIANT
        case Vb6Type::Object:  return "vb6_ComPackObject";  // void* → VARIANT
        case Vb6Type::Variant: return "vb6_ComPackValue";   // Fix 030: _Generic 通用打包宏
        default:               return std::string();
    }
}


// Fix 177b: coclass 创建点表达式. 见 comProjectImplClass 字段注释与
// driver_crossmod.cpp Fix 177b 块. 返回空串 = 未被遮蔽, 调用方走原 vb6_NewObject.
// 生成的 vb6_ComPack_<类>(vb6_cls_<类>_New()) 返回 void* (堆上 VARIANT*,
// VT_DISPATCH), 与 vb6_NewObject 的返回表示一致, 下游晚绑定路径零改动.
std::string CCodeGen::comNewExprFor(const Symbol* comSym) {
    if (!comSym) return "";
    // Fix <vbeclipse>-2: 本工程有同名类模块 ⇒ New 出来的必须是**原生**
    // vb6_cls_<工程类>_New(), 与 mapTypeRef 现在对该类型名返回的 vb6_cls_* 变量
    // 类型一致. 判据与 mapTypeRef 同源 (projectClassNameOf, 不是 isExternal:
    // 类型库自动加载注进来的内建 coclass isExternal=false).
    // 例: 工程类 Folder 撞 Shell32 的 coclass Folder → `New Folder` 必须是原生
    // 实例, 否则把原生指针当 VARIANT*/IDispatch 解 → vt=0 → 分发全 "not found".
    const std::string projCls = projectClassNameOf(comSym->name);
    if (!projCls.empty()) {
        return "vb6_cls_" + cIdent(projCls) + "_New()";
    }
    if (comSym->comProjectImplClass.empty()) return "";
    std::string impl = cIdent(comSym->comProjectImplClass);
    return "vb6_ComPack_" + impl + "(vb6_cls_" + impl + "_New())";
}

// P25: 解析COM标记为类型化属性取值, 用于COM调用参数打包
// 当isComMarker_为true时, 根据packFnHint选择对应类型的COM属性取值函数
// 如果isComMarker_为false, 返回空串
std::string CCodeGen::comMarkerValueForWrite(const std::string& packFn,
                                             const std::string& valExpr) {
    if (!isComMarker_) return valExpr;
    std::string hint = "BSTR";
    if (packFn == "vb6_ComPackDouble") hint = "Double";
    else if (packFn == "vb6_ComPackInt") hint = "Long";
    else if (packFn == "vb6_ComPackBool") hint = "Long";
    else if (packFn == "vb6_ComPackBSTR") hint = "BSTR";
    else if (packFn == "vb6_ComPackObject") hint = "Object";
    else if (packFn == "vb6_ComPackValue") hint = "Variant";
    resolveComValue(hint);
    return lastExpr_;
}

// Fix <vbeclipse> 2026-10-12: 写侧值位「右值整枚是 COM 调用结果」的收口。
//
// 起因 (ComCtlsDemo StatusBar_UserControl_ReadProperties, 页堆下 AV):
//   With New PropertyBag
//       .Contents = PropBag.ReadProperty("InitPanels", 0)
//   发射成
//       vb6_ComSetProp(_vb6_with_2, L"Contents",
//                      vb6_ComPackValue(vb6_ComCall((*PropBag), L"ReadProperty", ...)))
//   于是 vb6_ComPackValue(void*) 的 _Generic 落到 **vb6_VariantObject(void*)** ——
//   把 vb6_ComCall 交回的**堆 VARIANT** 当 IDispatch 去 AddRef, 读它的 `vt` 字段
//   (实测 = 3 = VT_I4) 当 vtable ⇒ 崩在 vb6_ComIsDispatchable 里 (av read target=0x8)。
//   最小复现 (3 行, `--emit-c` 秒出): .temp/tof/t3.bas
//       With inner : .CompareMode = bag.Item("z")
//   同一形态全工程 22 处 (17 处 `/* COM SetProp */` + 5 处 `/* With COM SetProp */`)。
//
// 正解: 这三枚出口的契约是**堆 VARIANT\*** (vb6com_invoke.c: `return (void*)result` /
// `(void*)eres`), 要 Variant 的打包器必须先把所有权交给 vb6_VariantFromComResult
// (它消耗掉那枚指针) 换出 vb6_VARIANT 值, 再打包 —— 这正是 resolveComMarkerForPack
// 早就写下的口径, 只是这几个写侧出口没接。判据前缀与它同源。
// 只接管 ComPackValue / ComPackVariant 两档: 标量档要的是数值, 那条路由 packArgValue
// 的 vb6_VariantToXxx 链负责, 在这里提前拆会把型定死。
std::string CCodeGen::comUnpackForPack(const std::string& packFn,
                                       const std::string& valExpr) const {
    if (packFn != "vb6_ComPackValue" && packFn != "vb6_ComPackVariant") return valExpr;
    if (valExpr.rfind("vb6_ComCall(", 0) == 0
        || valExpr.rfind("vb6_ComCallByDispid(", 0) == 0
        || valExpr.rfind("vb6_ComGetProp(", 0) == 0) {
        return "vb6_VariantFromComResult(" + valExpr + ")";
    }
    return valExpr;
}

std::string CCodeGen::packArgValue(const std::string& packFn,
                                   const std::string& expr) const {
    // ⚠ 长度字面量必须与整串长度相等 —— `std::string::compare(pos, count, str)` 拿
    // [pos, pos+count) 去比**整条** str, count 小于 str.size() 时恒不相等 (静默恒假)。
    // 这里直接用 rfind(..., 0) 前缀判定, 不数长度。
    const bool isVarValue = expr.rfind("vb6_VariantFromComResult(", 0) == 0
                         || expr.rfind("vb6_VariantFromValue(", 0) == 0;
    // 账 #305 §B147: 槽还有**第三种来源** —— 第三十六/三十七刀的发码改道交回的
    // (vb6_VarAddV/SubV/MulV/DivV(...))，前缀与上面两条都不同，而 cExprIsVariant 认得它。
    // 这一族不接就是硬错误 (实测 Charts 2020 ClsResizer.c(144)/(150):
    // `vb6_ComPackDouble((vb6_VarDivV(...)))` = C2440「无法从 vb6_VARIANT 转换为 double」)。
    // 修法刻意与上面四条**不同**：这里不拆成标量，而是整只 VARIANT 原样交给对端
    // (vb6_ComPackValue 的 _Generic 自带 vb6_VARIANT→identity 那一档) —— 拆成标量等于
    // 又一次在编译期替运行期选数值档，正是 §B145 要消灭的那件事 (实测拆成 Double 会
    // 静默把 Decimal 的 0.1+0.2 交成 0.29999999999999998，而 3.5 交成 3)。
    // 只接管「打包器要标量」那四档；ComPackValue/ComPackVariant/ComPackObject 原样不动。
    if (cExprIsVariant(expr) && (packFn == "vb6_ComPackDouble" || packFn == "vb6_ComPackInt"
                                 || packFn == "vb6_ComPackBool" || packFn == "vb6_ComPackBSTR")) {
        return "vb6_ComPackValue(" + expr + ")";
    }
    if (!isVarValue) return packFn + "(" + expr + ")";
    if (packFn == "vb6_ComPackDouble") return "vb6_ComPackDouble(vb6_VariantToDouble(" + expr + "))";
    if (packFn == "vb6_ComPackInt")    return "vb6_ComPackInt(vb6_VariantToLong(" + expr + "))";
    if (packFn == "vb6_ComPackBool")   return "vb6_ComPackBool(vb6_VariantToBool(" + expr + "))";
    if (packFn == "vb6_ComPackBSTR")   return "vb6_ComPackBSTR(vb6_VariantToString(" + expr + "))";
    return packFn + "(" + expr + ")";
}

std::string CCodeGen::resolveComMarkerForPack(const std::string& packFnHint) {
    // P26: vb6_ComPackVariant / Fix 030: vb6_ComPackValue 需要把 COM 调用返回的
    // VARIANT* 转成 vb6_VARIANT (即使 isComMarker_ 已被消费, lastExpr_ 仍可能是 COM 调用结果).
    // vb6_ComPackValue(vb6_VariantFromComResult(ComCall)) 经 _Generic VariantIdentity 路径
    // 等价于 vb6_ComPackVariant(vb6_VariantFromComResult(ComCall)).
    if (packFnHint == "vb6_ComPackVariant" || packFnHint == "vb6_ComPackValue") {
        if (lastExpr_.find("vb6_ComCall(") == 0 ||
            lastExpr_.find("vb6_ComCallByDispid(") == 0 ||
            lastExpr_.find("vb6_ComGetProp(") == 0 ||
            lastExpr_.find("vb6_ComGetObjectProp(") == 0 ||
            lastExpr_.find("vb6_ComCallObject(") == 0) {
            return "vb6_VariantFromComResult(" + lastExpr_ + ")";
        }
    }
    if (!isComMarker_) return "";
    isComMarker_ = false;
    std::string objExpr = std::move(comObjExpr_);
    std::string memName = std::move(comMemberName_);

    // P20-44: **DataObject 形参**成员在参数打包路径也要拦 —— Debug.Print 的实参
    // 打包走的就是这里 (resolveComValue 只管"取值语句"那条路)。形参 `As <类型>`
    // 原文记在 srcTypeName, 变量声明才记 variableTypeName。
    {
        std::string ddLower = Symbol::toLower(objExpr);
        auto* ddSym = symTab_.lookup(ddLower);
        if (ddSym && !ddSym->srcTypeName.empty()
            && Symbol::toLower(ddSym->srcTypeName) == "dataobject") {
            std::string memDD = Symbol::toLower(memName);
            std::string ddArg = "(void*)(*" + objExpr + ")";
            if (memDD == "gettext")      return "vb6_oleDD_GetText(" + ddArg + ")";
            if (memDD == "getfilecount") return "vb6_oleDD_GetFileCount(" + ddArg + ")";
        }
    }

    // 账 #299 第二十一刀: 同一枚唯一出口 (本函数只管"被当对象/实参用"那一头)。
    {
        std::string collObj = memberCollectionObjectExpr(objExpr, memName);
        if (!collObj.empty()) return collObj;
    }

    // C29-OLE: `OLE1.Object` (被当实参/整体赋值的场合, 如 `Set o = OLE1.Object`)。
    if (Symbol::toLower(memName) == "object") {
        std::string ocBare = oleConNameOfExpr(objExpr);
        if (!ocBare.empty())
            return "vb6_OleCon_GetObject((void*)vb6_hwnd_" + ocBare + ")";
    }

    // 前期绑定: 利用签名确定返回类型
    if (isEarlyBoundCom_ && earlyBoundSym_) {
        isEarlyBoundCom_ = false;
        const Symbol* comSym = earlyBoundSym_;
        earlyBoundSym_ = nullptr;
        std::string getter = comGetterExpr(comSigViewOf(comSym, memName), objExpr, memName);
        if (!getter.empty()) return getter;
    }

    // 后期绑定: 按 packer 反推解封类型 —— 这张表与 comMarkerValueForWrite 同一份口径
    // （Boolean 走 Int 档是 resolveComMarkerForPack 一直以来的答案）。
    if (packFnHint == "vb6_ComPackObject") {
        return "vb6_ComGetObjectProp(" + objExpr + ", L\"" + memName + "\")";
    } else if (packFnHint == "vb6_ComPackBSTR" || packFnHint.empty()) {
        return "vb6_ComGetStringProp(" + objExpr + ", L\"" + memName + "\")";
    } else if (packFnHint == "vb6_ComPackInt") {
        return "vb6_ComGetIntProp(" + objExpr + ", L\"" + memName + "\")";
    } else if (packFnHint == "vb6_ComPackBool") {
        return "vb6_ComGetIntProp(" + objExpr + ", L\"" + memName + "\")";
    } else if (packFnHint == "vb6_ComPackDouble") {
        return "vb6_ComGetDoubleProp(" + objExpr + ", L\"" + memName + "\")";
    }
    // vb6_ComPackVariant / Fix 030 vb6_ComPackValue: 需要把 COM 返回的 VARIANT* 转成 vb6_VARIANT
    if (packFnHint == "vb6_ComPackVariant" || packFnHint == "vb6_ComPackValue") {
        if (lastExpr_.find("vb6_ComCall(") == 0 ||
            lastExpr_.find("vb6_ComCallByDispid(") == 0 ||
            lastExpr_.find("vb6_ComGetProp(") == 0 ||
            lastExpr_.find("vb6_ComGetObjectProp(") == 0 ||
            lastExpr_.find("vb6_ComCallObject(") == 0) {
            return "vb6_VariantFromComResult(" + lastExpr_ + ")";
        }
        // Fix 143: 裸 marker (实参是 obj.Prop 形态, emitExpr 后 lastExpr_ 停在
        // 对象表达式上) — 不能把对象本身当值打包, 否则
        //   cboThemes.AddItem iTheme.Name → AddItem(iTheme)
        // 列表条目全变成指针值 ("???" 乱码). 按通用路径生成属性读取再转 VARIANT.
        return "vb6_VariantFromComResult(vb6_ComGetProp(" + objExpr + ", L\"" + memName + "\"))";
    }
    return "";
}
} // namespace vb6c3
