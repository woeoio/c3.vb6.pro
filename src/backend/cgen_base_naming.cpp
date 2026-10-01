// cgen_base_naming.cpp - C3 代码生成: 默认值 / 标识符转换 / 常量查询 / 过程名
// 2026-09-17 从 src/backend/cgen_base.cpp 纯搬移（原第 1538~1746 行，逐行未改）。
// 函数: defaultValue / cIdent / resolveArrayTargetIdent / lookupConstSym / isStringConstIdent / isConstIdent / constIdentType / wrapConstArgForByRef / wrapByRefVariantParamArg / cProcName

#include "backend/cgen.hpp"
#include <algorithm>
#include <cstdio>
#include <cctype>
#include <iostream>
#include <functional>
#include <unordered_set>

namespace vb6c3 {

std::string CCodeGen::defaultValue(Vb6Type type) const {
    switch (type) {
        case Vb6Type::Integer:
        case Vb6Type::Long:
        case Vb6Type::Byte:
        case Vb6Type::Boolean:
        case Vb6Type::Error:
            return "0";
        case Vb6Type::Single:
        case Vb6Type::Double:
        case Vb6Type::Date:
            return "0.0";
        case Vb6Type::Currency:
            return "0.0";   // Fix 126: Currency 值语义 (double)
        case Vb6Type::String:
            return "vb6_BSTR_Empty()";
        case Vb6Type::Object:
            return "NULL";
        case Vb6Type::Variant:
        case Vb6Type::Empty:
        case Vb6Type::Null:
            return "vb6_VariantEmpty()";
        case Vb6Type::UserDefinedType:
        case Vb6Type::Unknown:
            return "0";  // Fix 010q: Enum types resolve to Unknown, default to 0
        default:
            return "0";  // Fix 010q: safe default instead of vb6_VariantEmpty()
    }
}

// Fix 190: Optional 形参省略时的缺省实参表达式 (按传值方式包装).
// 位置实参补齐路径 (cgen_expr_call_opt_pad.inc / cgen_expr_call_arg_emit.inc 的
// Fix 142 分支) 早已做此包装, 命名实参补齐路径 (cgen_expr_call_named_args.inc)
// 漏了 —— 于 `m_oSocket.Create SocketType:=x` 这类命名调用上爆发:
// Optional ByRef String 形参被填成 vb6_BSTR_Empty() 的 BSTR **值**, 传到声明为
// BSTR* 的 C 形参后, 被调方 (*SocketAddress) 把 BSTR 数据首 4 字节当 BSTR 指针,
// 在 OLEAUT32 里对 0xBAAD0000 (L"" 之后的未初始化堆字节) 解引用 → 0xC0000005.
std::string CCodeGen::defaultArgForParam(Vb6Type type, bool isByVal,
                                        const std::string& explicitDefault) const {
    std::string defVal = (explicitDefault.empty() ? defaultValue(type) : explicitDefault);
    if (isByVal) return defVal;
    std::string cType = mapType(type);
    // P20-36: Variant/struct 类型不能用 {函数调用} 复合字面量
    if (type == Vb6Type::Variant || type == Vb6Type::Empty
        || type == Vb6Type::Null || type == Vb6Type::Object) {
        return "&(" + cType + "){0}";
    }
    return "&(" + cType + "){" + defVal + "}";
}

// ============================================================
// 标识符命名
// ============================================================

std::string CCodeGen::cIdent(const std::string& vb6Name) const {
    // VB6标识符可能含C关键字冲突, 添加前缀
    static const std::unordered_set<std::string> cKeywords = {
        "auto", "break", "case", "char", "const", "continue", "default", "do",
        "double", "else", "enum", "extern", "float", "for", "goto", "if",
        "int", "long", "register", "return", "short", "signed", "sizeof",
        "static", "struct", "switch", "typedef", "union", "unsigned", "void",
        "volatile", "while", "bool", "true", "false", "NULL",
        // C99/C11
        "inline", "restrict", "_Bool", "_Complex", "_Imaginary",
        // MSVC扩展
        "cdecl", "stdcall", "declspec", "dllimport", "dllexport",
    };

    std::string name = vb6Name;

    // 替换VB6方括号标识符 [Name] → Name
    if (!name.empty() && name.front() == '[' && name.back() == ']') {
        name = name.substr(1, name.size() - 2);
    }

    // Fix 010c: 替换VB6标识符中的特殊字符 (如版本号 ucsOsvWin8.1 → ucsOsvWin8_1)
    // C标识符只允许字母、数字、下划线
    for (auto& ch : name) {
        if (!std::isalnum(static_cast<unsigned char>(ch)) && ch != '_') {
            ch = '_';
        }
    }

    // 如果是C关键字, 添加vb6_前缀
    std::string lower = name;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    if (cKeywords.count(lower)) {
        return "vb6_" + name;
    }

    return name;
}

// Fix 084y-5: ReDim/Erase 目标的 C 标识符解析 (成员访问形态).
// 1) ".Field" — With 块成员 → _vb6_with_N->Field (与 Fix 061 一致)
// 2) "obj.Field" — UDT 成员:
//    - obj 是 ByRef UDT/数组/Variant 参数 → (*obj).Field (参数是 vb6_type_X*)
//    - obj 是 UDT 变量 (knownUdtVars_) → obj.Field (值类型)
// 否则返回 cIdent(varName) 原逻辑.
std::string CCodeGen::resolveArrayTargetIdent(const std::string& varName) {
    if (varName.size() > 1 && varName[0] == '.'
        && !withObjectVars_.empty() && !withObjectInfoStack_.empty()) {
        const auto& info = withObjectInfoStack_.back();
        if (info.kind == WithObjKind::Unknown || info.kind == WithObjKind::ClassInstance) {
            // Fix 092k: With 成员可能是多级路径 (.MessBuffer.Data) — cIdent 把
            // '.' 替换为 '_' 会生成 MessBuffer_Data (ToolsTlsThunks 1780:
            //   Erase .MessBuffer.Data → vb6_SafeArrayDestroy1D(_vb6_with_N->MessBuffer_Data)
            //   C2039 "MessBuffer_Data 不是 vb6_type_UcsTlsContext 的成员").
            // 逐段展开: 首段用 -> (With 对象是指针), 后续段保留 '.'.
            std::string path = varName.substr(1);
            std::string out = withObjectVars_.back();
            size_t pos = 0;
            bool firstSeg = true;
            while (pos <= path.size()) {
                size_t d = path.find('.', pos);
                std::string seg = (d == std::string::npos)
                    ? path.substr(pos) : path.substr(pos, d - pos);
                if (!seg.empty()) {
                    out += (firstSeg ? "->" : ".") + cIdent(seg);
                    firstSeg = false;
                }
                if (d == std::string::npos) break;
                pos = d + 1;
            }
            return out;
        }
    }
    size_t dot = varName.find('.');
    if (dot != std::string::npos && dot > 0 && dot + 1 < varName.size()) {
        std::string objName = varName.substr(0, dot);
        std::string memberName = varName.substr(dot + 1);
        // Fix: 多级成员链 (如 VBFlexGridMergeDrawInfo.Row.Cols) 逐段展开,
        // 每段 cIdent, 段间保留 '.', 避免整段 cIdent 把 '.' 替换成 '_'
        // (Row.Cols → Row_Cols → C2039 不是 vb6_type_TMERGEDRAWINFO 的成员).
        auto expandMemberPath = [this](const std::string& path) {
            std::string out;
            size_t pos = 0;
            while (pos <= path.size()) {
                size_t d = path.find('.', pos);
                std::string seg = (d == std::string::npos)
                    ? path.substr(pos) : path.substr(pos, d - pos);
                if (!seg.empty()) {
                    if (!out.empty()) out += ".";
                    out += cIdent(seg);
                }
                if (d == std::string::npos) break;
                pos = d + 1;
            }
            return out;
        };
        std::string objLower = objName;
        std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
        if (currentProc_) {
            for (auto& p : currentProc_->params) {
                std::string pLower = p.name;
                std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);
                if (pLower == objLower && !p.isByVal
                    && (p.type == Vb6Type::UserDefinedType
                        || (static_cast<uint16_t>(p.type) & static_cast<uint16_t>(Vb6Type::Array))
                        || p.type == Vb6Type::Variant)) {
                    return "(*" + objName + ")." + expandMemberPath(memberName);
                }
            }
        }
        if (knownUdtVars_.count(objLower)) {
            // Fix: 类模块 UDT 字段的带点链 ReDim/Erase 目标 — "VBFlexGridCells.Rows"
            // (Private VBFlexGridCells As TROWS) 需生成 me->VBFlexGridCells.Rows,
            // 此前裸名 → C2065 (VBFlexGrid.c 3370/4274/7999/8155). 局部变量/参数同名
            // 时跳过 (局部遮蔽优先, 避免误加 me-> 前缀).
            if (isClassModule_ && classMemberVars_.count(objLower)
                && !knownLocalVars_.count(objLower)) {
                return "me->" + objName + "." + expandMemberPath(memberName);
            }
            return objName + "." + expandMemberPath(memberName);
        }
    }
    return cIdent(varName);
}

// Fix 084aa: 常量符号查找.
// 模块级 Const 被生成为 #define 宏 (如 #define K (vb6_BSTR_FromStr(L"..."))),
// 对宏取址 &K 是非法的 (C2102). 局部变量同名时 lookup 优先命中局部符号,
// 返回 false → 保持 &变量 正常取址.
Symbol* CCodeGen::lookupConstSym(const std::string& name) const {
    Symbol* sym = symTab_.lookup(name);
    if (!sym) sym = symTab_.lookupModule(name);
    if (sym && sym->kind == SymbolKind::Constant) {
        return sym;
    }
    return nullptr;
}

bool CCodeGen::isStringConstIdent(const std::string& name) const {
    Symbol* sym = lookupConstSym(name);
    return sym && sym->constType == Vb6Type::String;
}

bool CCodeGen::isConstIdent(const std::string& name) const {
    if (lookupConstSym(name) != nullptr) return true;
    // Fix 158l: RTL 头文件里的值型常量宏 (vb6rtl_userctl.h 的 vbPicType*/vbAsyncType*)
    // 不在符号表 (VB6 内建常量, 由 RTL 提供 #define). 此前 isConstIdent 返回 false,
    // 比较路径把 vbPicTypeIcon 当左值 → `&vbPicTypeIcon` 展开为 &3 → C2101
    // (Common.c/VisualStyles.c/VBFlexGrid.c 的 `vbPicTypeIcon = .Type` 等).
    // 凡 RTL #define 的值常量都视为"常量宏不可取址", 走复合字面量/临时变量路径.
    static const std::unordered_set<std::string> kRtlConstMacros158l = {
        "vbpictypenone", "vbpictypebitmap", "vbpictypemetafile",
        "vbpictypeicon", "vbpictypeemetafile",
        "vbasynctypepicture", "vbasynctypefile", "vbasynctypebytearray",
        "vbasyncreadsynchronous", "vbasyncreadasynchronous", "vbasyncreadforceupdate",
    };
    std::string lower = name;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    return kRtlConstMacros158l.count(lower) > 0;
}

Vb6Type CCodeGen::constIdentType(const std::string& name) const {
    Symbol* sym = lookupConstSym(name);
    if (sym) return sym->constType;
    return Vb6Type::Variant;
}

// Fix 084aa: 常量宏作为 ByRef 实参 — 生成可寻址复合字面量.
// 常量是 #define 宏 (可能是函数调用如 vb6_BSTR_FromStr(L"...") 或数值字面量),
// 不能 &CONST 取址, 须按形参 C 类型包装为 (&(TYPE){CONST}):
//   - ByRef Variant → (&(vb6_VARIANT){.vt=..., .xxx=CONST}) 按常量自身类型选字段
//   - ByRef String  → (&(BSTR){CONST})
//   - 其他          → (&(cType){CONST})
std::string CCodeGen::wrapConstArgForByRef(const std::string& argVal, Vb6Type paramVb6Type) const {
    if (paramVb6Type == Vb6Type::Variant || paramVb6Type == Vb6Type::Empty ||
        paramVb6Type == Vb6Type::Null || paramVb6Type == Vb6Type::Object) {
        switch (constIdentType(argVal)) {
            case Vb6Type::String:
                return "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=" + argVal + "})";
            case Vb6Type::Double:
            case Vb6Type::Single:
            case Vb6Type::Decimal:
            case Vb6Type::Currency:
                return "(&(vb6_VARIANT){.vt=VT_R8, .dblVal=(double)(" + argVal + ")})";
            case Vb6Type::Boolean:
                return "(&(vb6_VARIANT){.vt=VT_BOOL, .boolVal=(int16_t)(" + argVal + ")})";
            case Vb6Type::Byte:
                return "(&(vb6_VARIANT){.vt=VT_UI1, .bVal=(uint8_t)(" + argVal + ")})";
            default:
                return "(&(vb6_VARIANT){.vt=VT_I4, .lVal=(int32_t)(" + argVal + ")})";
        }
    }
    if (isStringConstIdent(argVal)) {
        return "(&(BSTR){" + argVal + "})";
    }
    std::string cType = mapType(paramVb6Type);
    if (cType.empty()) cType = "int32_t";
    return "(&(" + cType + "){" + argVal + "})";
}

// Fix 189: 当前函数的 ByRef 具体类型形参 (C 形参是 T*, 如 int32_t*/BSTR*) 作为
// **ByRef Variant** 形参的实参时, 不能把 T* 直接当 vb6_VARIANT* 交出去 ——
// 被调方 `vb6_VariantFromValue((*p))` 会把 T 的字节当 VARIANT.vt 解析
// (vt 取到端口号 800 / BSTR 指针低位), 后续按 vt 取字段 → 解引用垃圾 BSTR →
// 0xC0000005 (OLEAUT32). 按当前形参的 VB 类型生成字段式 VARIANT 复合字面量,
// 取址后作为实参. 返回空串 = 该类型不适用 (Variant/Object/UDT/未知),
// 调用方保持原「直传」行为不变.
std::string CCodeGen::wrapByRefVariantParamArg(const std::string& paramName,
                                               Vb6Type curParamType) const {
    const std::string src = "(*" + paramName + ")";

    // Fix <vbeclipse> rev17: 当前形参声明 `As <工程类>` —— C 侧是 `vb6_cls_X*`, ByRef 传进来
    // 就是 `vb6_cls_X**`。把它交给被调方的 `ByRef As Variant` 槽时必须**装箱成真 IDispatch**:
    // 原样交出去等于把 `vb6_cls_X**` 当 `vb6_VARIANT*` 解 (两者都是指针, MSVC 只给 C4133
    // 警告不报错) ⇒ 被调方读到 vt/pdispVal 全是垃圾 ⇒ 槽位取回 Nothing。
    // 实证 play78: ucFolder.AddView 的 `m_FolderViews.Add View.ViewId, View` 生成
    //   vb6_List_Add(me->m_FolderViews, ..., View, -1, 0);
    // m_FolderViews 里存进垃圾 Variant ⇒ ShowView 里 `Set l_View = .Item(i)` 得到 NULL
    // ⇒ `l_View.ViewId` av read 0x4 (崩在 _vb6_ucFolder_ShowView+0xFC)。
    // 判据只用 cParamClassPtrType —— 它是"这个槽收不收类实例指针"的唯一口径
    // (走 cTypeForDeclaredTypeName → mapTypeRef, 与类模块发定义那一份同一映射);
    // 这里自查形参而不让 4 个调用点多传一个参数, 免得同名逻辑各写一份。
    // 裸结构体指针也不能直接充 VT_DISPATCH —— VariantClear 会把首字段 (`__comObj`) 当
    // vtable 解引用 (与 cgen_setlet Fix 179a / 下面 default 档同口径), 所以先
    // vb6_ComObject_FromInstance 包成真 IDispatch (它的 1 个引用直接转移给 VARIANT)。
    // 接口槽 (`vb6_ivref_*`) 不走这条: 那不是 vb6_ComObject 包装, 保持原行为。
    if (currentProc_) {
        for (const auto& p : currentProc_->params) {
            if (Symbol::toLower(p.name) != Symbol::toLower(paramName)) continue;
            if (p.isByVal || p.isParamArray) break;
            // cParamClassPtrType 非 const (它与 mapTypeRef 同族), 本函数是 const —— 只读调用.
            const std::string clsPtr = const_cast<CCodeGen*>(this)->cParamClassPtrType(p);
            if (clsPtr.rfind("vb6_cls_", 0) == 0) {
                std::string cls = clsPtr.substr(8);
                if (!cls.empty() && cls.back() == '*') cls.pop_back();
                if (!cls.empty()) {
                    return "(&(vb6_VARIANT){.vt=VT_DISPATCH, .pdispVal=(IDispatch*)"
                           "vb6_ComObject_FromInstance(vb6_FindCoClassDesc(\"" + cls
                           + "\"), (void*)" + src + ")})";
                }
            }
            break;
        }
    }

    switch (curParamType) {
        case Vb6Type::String:
            return "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=" + src + "})";
        case Vb6Type::Long:
        case Vb6Type::Integer:
            return "(&(vb6_VARIANT){.vt=VT_I4, .lVal=(int32_t)(" + src + ")})";
        case Vb6Type::Byte:
            return "(&(vb6_VARIANT){.vt=VT_UI1, .bVal=(uint8_t)(" + src + ")})";
        case Vb6Type::Boolean:
            return "(&(vb6_VARIANT){.vt=VT_BOOL, .boolVal=(int16_t)(" + src + ")})";
        case Vb6Type::Double:
        case Vb6Type::Single:
        case Vb6Type::Currency:
            return "(&(vb6_VARIANT){.vt=VT_R8, .dblVal=(double)(" + src + ")})";
        case Vb6Type::Date:
            return "(&(vb6_VARIANT){.vt=VT_DATE, .dblVal=(double)(" + src + ")})";
        case Vb6Type::LongPtr:
        case Vb6Type::ULong:
            return "(&(vb6_VARIANT){.vt=VT_I8, .llVal=(int64_t)(" + src + ")})";
        default:
            return std::string();
    }
}

std::string CCodeGen::cProcName(const std::string& procName, AccessLevel access,
                                 const std::string& sourceModule) const {
    // 跨模块函数(外部符号): vb6_<ModuleName>_<ProcName>
    if (!sourceModule.empty()) {
        return "vb6_" + cIdent(sourceModule) + "_" + cIdent(procName);
    }
    // 多模块项目:
    // - 所有函数都包含模块名, 避免与RTL函数名冲突
    //   (如用户定义 Private Sub ErrRaise 会生成 vb6_ErrRaise, 与RTL的 vb6_ErrRaise 冲突)
    // - Private标准模块函数也包含模块名 (虽然为static, 但其名可能被本模块内
    //   Err.Raise等硬编码RTL调用遮蔽, 导致参数不匹配)
    if (isMultiModule_ && !moduleName_.empty()) {
        return "vb6_" + cIdent(moduleName_) + "_" + cIdent(procName);
    }
    return "vb6_" + cIdent(procName);
}

// 重载变体的 C 名后缀 (O2): 非变体/裸键 head 返回空串 —— 零重载工程产出一字不变.
// 变体: "_ov" + 指纹消毒串 (非字母数字→'_') 并截断, 保证组内唯一 (fp 本身组内唯一).
static std::string ovlSanitizeFp(const std::string& fp) {
    std::string out;
    for (char c : fp) out += (isalnum((unsigned char)c) ? c : '_');
    if (out.size() > 24) out.resize(24);
    return out;
}

std::string CCodeGen::ovlCSuffix(const Symbol* sym) const {
    if (!sym || !sym->isOverloadVariant || sym->overloadFp.empty()) return "";
    return "_ov" + ovlSanitizeFp(sym->overloadFp);
}

std::string CCodeGen::ovlCSuffixFromKey(const std::string& suffixKey) const {
    // suffixKey = 语义层记在 AST 上的符号键后缀 "$ov$<fp>" ("" = 非变体)
    const std::string pre = "$ov$";
    if (suffixKey.empty()) return "";
    if (suffixKey.compare(0, pre.size(), pre) != 0) return "";
    return "_ov" + ovlSanitizeFp(suffixKey.substr(pre.size()));
}

} // namespace vb6c3
