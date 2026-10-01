#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_util_com.cpp: COM 值解析 + 打包标记 + 类字段名规范化 ---



// ============================================================
// COM辅助 (P6.2)
// ============================================================

std::string CCodeGen::resolveComValue(const std::string& unpackType) {
    // P24-07: 早期绑定 — 利用签名returnType选择正确的解包函数
    if (!isComMarker_) return lastExpr_;

    std::string objExpr = std::move(comObjExpr_);
    std::string memberName = std::move(comMemberName_);

    // C29-5b: `Toolbar1.Buttons` → **真集合对象** (C29-3 那套成员对象机制，与 8b 的
    // Nodes 同一条 cheapest route)。5a 那只只改道 `Buttons.Count` 的特例钩子就此退休 ——
    // Count 现在由集合自己答 (memberobj → vb6_Toolbar_ButtonCount → TB_BUTTONCOUNT)，
    // 仍然是**控件侧**证据，不是我那张表自说自话。
    // 位置照旧赶在 P24-07 那段 early-bound 决策之前，否则会被当普通 COM 成员发成 vb6_ComGetProp。
    if (Symbol::toLower(memberName) == "buttons") {
        std::string tbName = toolbarNameOfExpr(objExpr);
        if (!tbName.empty()) {
            lastExpr_ = "vb6_Toolbar_Buttons((void*)vb6_hwnd_" + tbName + ")";
            isComMarker_ = false;
            return lastExpr_;
        }
    }

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

    // C29-Data: `Data1.Recordset.<标量成员>` 直译 —— objExpr 是 vb6_Data_Self( 透传形态。
    // (Fields("x")/Move* 是调用不是属性, 在 cgen_expr_call_com_bind.inc 拦。)
    if (objExpr.find("vb6_Data_FieldValueStr(") == 0) {
        // Fields("x").Value: 值就是 FieldValueStr 本身
        std::string dtLower = Symbol::toLower(memberName);
        if (dtLower == "value") { lastExpr_ = objExpr; isComMarker_ = false; return lastExpr_; }
    }
    if (objExpr.find("vb6_Data_Self(") == 0) {
        std::string dtLower = Symbol::toLower(memberName);
        if (dtLower == "bof")         { lastExpr_ = "vb6_Data_BOF(" + objExpr + ")"; isComMarker_ = false; return lastExpr_; }
        if (dtLower == "eof")         { lastExpr_ = "vb6_Data_EOF(" + objExpr + ")"; isComMarker_ = false; return lastExpr_; }
        if (dtLower == "recordcount") { lastExpr_ = "vb6_Data_RecordCount(" + objExpr + ")"; isComMarker_ = false; return lastExpr_; }
        if (dtLower == "fieldcount")  { lastExpr_ = "vb6_Data_FieldCount(" + objExpr + ")"; isComMarker_ = false; return lastExpr_; }
        if (dtLower == "value" || dtLower == "recordset") {
            // Fields("x").Value 的 .Value: FieldValueStr 就是值本身, 透传不包装
            lastExpr_ = objExpr; isComMarker_ = false; return lastExpr_;
        }
    }

    // C29-7: `ListView1.ListItems` / `.ColumnHeaders` → **真集合对象** (与 C29-3 的
    // ImageList.ListImages 同一口径)。ListView 是**真窗口**, 宿主槽是 vb6_hwnd_X,
    // 不是 ImageList 那种 vb6_com_X 实例指针。
    {
        std::string lvLower = Symbol::toLower(memberName);
        if (lvLower == "listitems" || lvLower == "columnheaders") {
            std::string lvHwnd = listViewHwndExprOf(objExpr);
            if (!lvHwnd.empty()) {
                lastExpr_ = (lvLower == "listitems")
                    ? ("vb6_ListView_ListItems((void*)" + lvHwnd + ")")
                    : ("vb6_ListView_ColumnHeaders((void*)" + lvHwnd + ")");
                isComMarker_ = false;
                return lastExpr_;
            }
        }
    }

    // C29-8b: `TreeView1.Nodes` → **真集合对象** (与 C29-7 的 ListView 同一口径、
    // 同一个 vb6forms_memberobj.c 机制)。不拦就发 `vb6_ComGetObjectProp(vb6_hwnd_X,
    // L"Nodes")` —— HWND 不是 IDispatch, 链接过、运行期读数全空。
    if (Symbol::toLower(memberName) == "nodes") {
        std::string tvBare = treeViewNameOfExpr(objExpr);
        if (!tvBare.empty()) {
            lastExpr_ = "vb6_TreeView_Nodes((void*)vb6_hwnd_" + tvBare + ")";
            isComMarker_ = false;
            return lastExpr_;
        }
        // Fix <vbeclipse>: 跨窗体 TreeView (`frmViewViews.tvwViews.Nodes`) ——
        // objExpr 是控件句柄访问器文本, 在 externalTreeViewAccs_ 里登记过。
        auto itExtTv = externalTreeViewAccs_.find(objExpr);
        if (itExtTv != externalTreeViewAccs_.end()) {
            lastExpr_ = "vb6_TreeView_Nodes((void*)" + itExtTv->second + ")";
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

    // C29-3: `ImageList1.ListImages` → **真集合对象** (原生复刻的成员集合, 见
    // vb6forms_memberobj.c)。必须排在下面那条 P20-39 分支之前: 那条管的是链上更外层的
    // 成员 (ListImages.Count / ListImages(i).Key), 这里要先把**集合本身**立起来。
    // 不拦就会发 `vb6_ComGetObjectProp(vb6_com_X, L"ListImages")` —— 而 ImageList 无窗口,
    // vb6_com_X 槽里放的是 HIMAGELIST 实例指针 (不是 IDispatch), 对它做属性读 =
    // 运行期拿垃圾当 vtable 用。
    {
        std::string liLower = Symbol::toLower(memberName);
        if (liLower == "listimages") {
            std::string liBare = imageListNameOfExpr(objExpr);
            if (!liBare.empty()) {
                lastExpr_ = "vb6_ImageList_ListImages((void*)vb6_com_" + liBare + ")";
                isComMarker_ = false;
                return lastExpr_;
            }
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
            if (memLower == "index") { lastExpr_ = sbFinishByKey("vb6_StatusBar_GetPanelIndexByKey",
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
        }
    }

    // P24-07: 早期绑定推断 — 利用TypeLib签名的returnType决策
    if (isEarlyBoundCom_ && earlyBoundSym_) {
        isEarlyBoundCom_ = false;
        const Symbol* comSym = earlyBoundSym_;
        earlyBoundSym_ = nullptr;
        std::string memLower = memberName;
        std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
        auto it = comSym->comMethods.find(memLower);
        if (it != comSym->comMethods.end()) {
            const auto& sig = it->second;
            std::string returnType = mapType(sig.returnType);
            std::string getPropArgs = objExpr + ", L\"" + memberName + "\"";
            if (returnType == "BSTR") {
                lastExpr_ = "vb6_ComGetStringProp(" + getPropArgs + ")";
            } else if (returnType == "int32_t" || returnType == "int16_t") {
                lastExpr_ = "vb6_ComGetIntProp(" + getPropArgs + ")";
            } else if (returnType == "double" || returnType == "float") {
                lastExpr_ = "vb6_ComGetDoubleProp(" + getPropArgs + ")";
            } else if (returnType == "void*") {
                lastExpr_ = "vb6_ComGetObjectProp(" + getPropArgs + ")";
            } else {
                // P24-07: 未知返回类型(如Enum→UserDefinedType) → 按目标变量类型选择
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
                } else {
                    lastExpr_ = "vb6_VariantFromComResult(vb6_ComCall(" + objExpr + ", L\"" + memberName + "\", NULL, 0))";
                }
            }
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


// Fix <vbeclipse>: 宿主伪对象成员名规范化 (见 cgen_helpers.inc 声明注释).
std::string CCodeGen::canonicalHostPseudoMember(const std::string& pseudoObj,
                                                const std::string& memberName) const {
    if (memberName.empty()) return memberName;
    static const std::pair<const char*, const char*> kUserControlCanon[] = {
        {"hdc", "hDC"},
    };
    const std::string pj = Symbol::toLower(pseudoObj);
    if (pj != "usercontrol" && pj != "ambient"
        && pj != "extender" && pj != "propertypage") {
        return memberName;
    }
    const std::string want = Symbol::toLower(memberName);
    for (const auto& kv : kUserControlCanon) {
        if (want == kv.first) return kv.second;
    }
    return memberName;
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
    switch (vt) {
        case Vb6Type::String:
            return "vb6_ComPackBSTR";  // BSTR → VARIANT
        case Vb6Type::Integer:
        case Vb6Type::Long:
            return "vb6_ComPackInt";   // int32_t → VARIANT
        case Vb6Type::Boolean:
            return "vb6_ComPackBool";  // VB6 Boolean → VARIANT VT_BOOL
        case Vb6Type::Single:
        case Vb6Type::Double:
            return "vb6_ComPackDouble"; // double → VARIANT
        case Vb6Type::Object:
            return "vb6_ComPackObject"; // void* → VARIANT
        case Vb6Type::Variant:
            return "vb6_ComPackValue";  // Fix 030: 通用打包宏 — 路由任意 C 类型实参 (inferExprType 回退 Variant 时安全)
        default:
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
std::string CCodeGen::resolveComMarkerForPack(const std::string& packFnHint) {
    // P26: vb6_ComPackVariant / Fix 030: vb6_ComPackValue 需要把 COM 调用返回的
    // VARIANT* 转成 vb6_VARIANT (即使 isComMarker_ 已被消费, lastExpr_ 仍可能是 COM 调用结果).
    // vb6_ComPackValue(vb6_VariantFromComResult(ComCall)) 经 _Generic VariantIdentity 路径
    // 等价于 vb6_ComPackVariant(vb6_VariantFromComResult(ComCall)).
    if (packFnHint == "vb6_ComPackVariant" || packFnHint == "vb6_ComPackValue") {
        if (lastExpr_.find("vb6_ComCall(") == 0 ||
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

    // C29-7: `ListView1.ListItems` / `.ColumnHeaders` → 真集合对象 (同 resolveComValue)。
    // 走这条的是"集合被当实参 / 被整体赋值"的场合, 例如 `Set c = ListView1.ListItems`。
    {
        std::string lvLower = Symbol::toLower(memName);
        if (lvLower == "listitems" || lvLower == "columnheaders") {
            std::string lvHwnd = listViewHwndExprOf(objExpr);
            if (!lvHwnd.empty())
                return (lvLower == "listitems")
                    ? ("vb6_ListView_ListItems((void*)" + lvHwnd + ")")
                    : ("vb6_ListView_ColumnHeaders((void*)" + lvHwnd + ")");
        }
    }

    // C29-8b: `TreeView1.Nodes` → 真集合对象 (同 resolveComValue 那条)。走这条的是
    // "集合被当实参 / 被整体赋值"的场合, 例如 `Set ns = TreeView1.Nodes`。
    if (Symbol::toLower(memName) == "nodes") {
        std::string tvBare = treeViewNameOfExpr(objExpr);
        if (!tvBare.empty())
            return "vb6_TreeView_Nodes((void*)vb6_hwnd_" + tvBare + ")";
        auto itExtTv = externalTreeViewAccs_.find(objExpr);
        if (itExtTv != externalTreeViewAccs_.end())
            return "vb6_TreeView_Nodes((void*)" + itExtTv->second + ")";
    }

    // C29-5b: `Toolbar1.Buttons` → 真集合对象 (同 resolveComValue 那条)。
    if (Symbol::toLower(memName) == "buttons") {
        std::string tbBare = toolbarNameOfExpr(objExpr);
        if (!tbBare.empty())
            return "vb6_Toolbar_Buttons((void*)vb6_hwnd_" + tbBare + ")";
    }

    // C29-OLE: `OLE1.Object` (被当实参/整体赋值的场合, 如 `Set o = OLE1.Object`)。
    if (Symbol::toLower(memName) == "object") {
        std::string ocBare = oleConNameOfExpr(objExpr);
        if (!ocBare.empty())
            return "vb6_OleCon_GetObject((void*)vb6_hwnd_" + ocBare + ")";
    }

    // C29-3: `ImageList1.ListImages` → 真集合对象 (与 resolveComValue 那条同一口径)。
    // 走这条的是"集合被当实参 / 被整体赋值"的场合, 例如 `Set c = ImageList1.ListImages`。
    {
        std::string liLower = Symbol::toLower(memName);
        if (liLower == "listimages") {
            std::string liBare = imageListNameOfExpr(objExpr);
            if (!liBare.empty())
                return "vb6_ImageList_ListImages((void*)vb6_com_" + liBare + ")";
        }
    }

    // 前期绑定: 利用签名确定返回类型
    if (isEarlyBoundCom_ && earlyBoundSym_) {
        isEarlyBoundCom_ = false;
        const Symbol* comSym = earlyBoundSym_;
        earlyBoundSym_ = nullptr;
        std::string memLower = memName;
        std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
        auto it = comSym->comMethods.find(memLower);
        if (it != comSym->comMethods.end() && it->second.isPropertyGet) {
            const auto& sig = it->second;
            std::string returnType = mapType(sig.returnType);
            if (returnType == "int32_t" || returnType == "int16_t") {
                return "vb6_ComGetIntProp(" + objExpr + ", L\"" + memName + "\")";
            } else if (returnType == "BSTR") {
                return "vb6_ComGetStringProp(" + objExpr + ", L\"" + memName + "\")";
            } else if (returnType == "double" || returnType == "float") {
                return "vb6_ComGetDoubleProp(" + objExpr + ", L\"" + memName + "\")";
            } else if (returnType == "void*") {
                return "vb6_ComGetObjectProp(" + objExpr + ", L\"" + memName + "\")";
            }
        }
    }

    // 后期绑定: 根据packFnHint推断所需的属性取值函数
    // packFnHint由comPackExpr根据上下文确定, 代表参数期望的C类型
    if (packFnHint == "vb6_ComPackObject") {
        return "vb6_ComGetObjectProp(" + objExpr + ", L\"" + memName + "\")";
    } else if (packFnHint == "vb6_ComPackBSTR" || packFnHint.empty()) {
        return "vb6_ComGetStringProp(" + objExpr + ", L\"" + memName + "\")";
    } else if (packFnHint == "vb6_ComPackInt") {
        return "vb6_ComGetIntProp(" + objExpr + ", L\"" + memName + "\")";
    } else if (packFnHint == "vb6_ComPackDouble") {
        return "vb6_ComGetDoubleProp(" + objExpr + ", L\"" + memName + "\")";
    } else if (packFnHint == "vb6_ComPackBool") {
        return "vb6_ComGetIntProp(" + objExpr + ", L\"" + memName + "\")";
    }
    // vb6_ComPackVariant / Fix 030 vb6_ComPackValue: 需要把 COM 返回的 VARIANT* 转成 vb6_VARIANT
    if (packFnHint == "vb6_ComPackVariant" || packFnHint == "vb6_ComPackValue") {
        if (lastExpr_.find("vb6_ComCall(") == 0 ||
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
