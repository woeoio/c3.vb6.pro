// cgen_base_type.cpp - C3 代码生成: 类型映射与常量折叠
// 2026-09-17 从 src/backend/cgen_base.cpp 纯搬移（原第 1155~1536 行，逐行未改）。
// 函数: mapType / mapComType / emitGuidInitializer / mapTypeRef / mapDeclareType / tryEvalConstInt

#include "backend/cgen.hpp"
#include <algorithm>
#include <cstdio>
#include <cctype>
#include <iostream>
#include <functional>
#include <unordered_set>

namespace vb6c3 {

// ============================================================
// 类型映射
// ============================================================

std::string CCodeGen::mapType(Vb6Type type) const {
    // 检查数组标志
    bool isArray = (static_cast<uint16_t>(type) & static_cast<uint16_t>(Vb6Type::Array)) != 0;
    Vb6Type baseType = isArray
        ? static_cast<Vb6Type>(static_cast<uint16_t>(type) & ~static_cast<uint16_t>(Vb6Type::Array))
        : type;

    std::string cType;
    switch (baseType) {
        case Vb6Type::Empty:    cType = "int16_t"; break;   // VB6 Empty = 0
        case Vb6Type::Null:     cType = "int16_t"; break;   // VB6 Null = 1
        case Vb6Type::Integer:  cType = "int16_t"; break;
        case Vb6Type::Long:     cType = "int32_t"; break;
        case Vb6Type::LongPtr: cType = "intptr_t"; break;   // Fix 081e: architecture-width integer
        case Vb6Type::LongLong: cType = "int64_t"; break;   // Fix 084m: 恒 64 位有符号 (不随架构)
        case Vb6Type::Single:   cType = "float"; break;
        case Vb6Type::Double:   cType = "double"; break;
        // Fix 126: Currency = 64bit/10000 (4 位小数) 的**值** —— 与 Date 一样按值语义
        // 映射为 double。此前映射 int64_t(存放大整数), 而表达式与 VARIANT 打包都按
        // 值使用, 于是 `yRange = yRange + CCur(step)` 每步被放大 10000 倍
        // (Charts 2020 坐标轴数字与柱高严重不符)。放大整数只存在于 VT_CY 变体里,
        // 由读取侧除回来 (rtl: vb6_VariantToDouble / vb6_Format 的 vtCurrency 分支)。
        case Vb6Type::Currency: cType = "double"; break;
        case Vb6Type::Date:     cType = "double"; break;    // OLE date
        case Vb6Type::String:   cType = "BSTR"; break;      // wchar_t* wrapper
        case Vb6Type::Object:   cType = "void*"; break;     // IDispatch* → void* for now
        case Vb6Type::Error:    cType = "int32_t"; break;   // SCODE/HRESULT
        case Vb6Type::Boolean:  cType = "int16_t"; break;   // VB6: True=-1, False=0
        case Vb6Type::Variant:  cType = "vb6_VARIANT"; break;   // tagged union
        case Vb6Type::Byte:     cType = "uint8_t"; break;
        case Vb6Type::ULong:    cType = "uint32_t"; break;
        // C3 扩展 (ai/032): 无符号/小整型族。位宽对照 Integer/Long/LongLong,
        // 见 common/types.hpp 的 Vb6Type 注释。恒定位宽, 不随架构。
        case Vb6Type::SByte:    cType = "int8_t"; break;
        case Vb6Type::UInteger: cType = "uint16_t"; break;
        case Vb6Type::ULongLong:cType = "uint64_t"; break;
        case Vb6Type::Void:     cType = "void"; break;
        case Vb6Type::Decimal:  cType = "vb6_VARIANT"; break;   // 用VARIANT兜底
        case Vb6Type::UserDefinedType: cType = "vb6_VARIANT"; break; // 占位, 后续改进
        default:                cType = "vb6_VARIANT"; break;    // 未知/安全兜底
    }

    if (isArray) {
        if (baseType == Vb6Type::Byte) {
            cType = "uint8_t*";  // Byte数组作为原始字节指针
        } else {
            cType = "vb6_SafeArray1D*";  // 数组用SAFEARRAY
        }
    }
    return cType;
}

std::string CCodeGen::mapComType(Vb6Type type) const {
    // COM vtable 方法签名映射 (用于外部COM事件接收器)
    bool isArray = (static_cast<uint16_t>(type) & static_cast<uint16_t>(Vb6Type::Array)) != 0;
    Vb6Type baseType = isArray
        ? static_cast<Vb6Type>(static_cast<uint16_t>(type) & ~static_cast<uint16_t>(Vb6Type::Array))
        : type;
    if (isArray) {
        return "SAFEARRAY*";  // 数组在COM vtable中总是SAFEARRAY指针
    }
    switch (baseType) {
        case Vb6Type::Integer:  return "int16_t";
        case Vb6Type::Long:     return "int32_t";
        case Vb6Type::LongPtr: return "intptr_t";  // Fix 081e
        case Vb6Type::LongLong: return "int64_t";  // Fix 084m: 恒 64 位有符号
        case Vb6Type::Single:   return "float";
        case Vb6Type::Double:   return "double";
        case Vb6Type::Currency: return "double";   // Fix 126: 值语义
        case Vb6Type::Date:     return "double";
        case Vb6Type::String:   return "BSTR";
        case Vb6Type::Object:   return "IUnknown*";
        case Vb6Type::Error:    return "int32_t";
        case Vb6Type::Boolean:  return "int16_t";
        case Vb6Type::Variant:  return "VARIANT";
        case Vb6Type::Byte:     return "uint8_t";
        case Vb6Type::ULong:    return "uint32_t";
        // C3 扩展 (ai/032)
        case Vb6Type::SByte:    return "int8_t";
        case Vb6Type::UInteger: return "uint16_t";
        case Vb6Type::ULongLong:return "uint64_t";
        case Vb6Type::Void:     return "void";
        default:                return "void*";
    }
}

std::string CCodeGen::emitGuidInitializer(const std::string& iidStr) const {
    // 解析 {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} 格式
    if (iidStr.size() < 38 || iidStr.front() != '{' || iidStr.back() != '}') return "";
    std::string s = iidStr.substr(1, iidStr.size() - 2);  // 去掉花括号
    // 分割为5个部分
    std::vector<std::string> parts;
    size_t start = 0;
    for (size_t i = 0; i < s.size(); ++i) {
        if (s[i] == '-') {
            parts.push_back(s.substr(start, i - start));
            start = i + 1;
        }
    }
    parts.push_back(s.substr(start));
    if (parts.size() != 5) return "";
    std::string d1 = parts[0];
    std::string d2 = parts[1];
    std::string d3 = parts[2];
    std::string d4 = parts[3] + parts[4];  // 16个hex字符
    if (d1.size() != 8 || d2.size() != 4 || d3.size() != 4 || d4.size() != 16) return "";
    char buf[128];
    snprintf(buf, sizeof(buf), "{0x%s, 0x%s, 0x%s, {0x%s, 0x%s, 0x%s, 0x%s, 0x%s, 0x%s, 0x%s, 0x%s}}",
             d1.c_str(), d2.c_str(), d3.c_str(),
             d4.substr(0,2).c_str(), d4.substr(2,2).c_str(), d4.substr(4,2).c_str(), d4.substr(6,2).c_str(),
             d4.substr(8,2).c_str(), d4.substr(10,2).c_str(), d4.substr(12,2).c_str(), d4.substr(14,2).c_str());
    return std::string(buf);
}

// Fix 107: 按"类型类别"查找符号, 忽略同名的过程/变量.
// Fix 103 允许类型与过程同名; 冲突时类型符号改存 <name>$ty (见 Scope::define /
// SymbolTable::define), 此时通用 symTab_.lookup() 只命中过程符号, 类型判定
// (Class/UDT/Enum/COM) 全部落空 → 回落 "void*"/Variant.
Symbol* CCodeGen::lookupTypeSymbol(const std::string& name) const {
    auto isTypeKind = [](SymbolKind k) {
        return k == SymbolKind::Class || k == SymbolKind::UserDefinedType ||
               k == SymbolKind::EnumType || k == SymbolKind::ComClass ||
               k == SymbolKind::ComInterface || k == SymbolKind::ComModule ||
               k == SymbolKind::ComGlobalNs || k == SymbolKind::Delegate;
    };
    if (auto* s = symTab_.lookup(name)) {
        if (isTypeKind(s->kind)) return s;
    }
    // 模块作用域优先 (类型声明通常是模块级的), lookupLocalByKind 内含 $ty 回退.
    static const SymbolKind typeKinds[] = {
        SymbolKind::UserDefinedType, SymbolKind::EnumType, SymbolKind::Class,
        SymbolKind::ComClass, SymbolKind::ComInterface,
        SymbolKind::ComModule, SymbolKind::ComGlobalNs,
        SymbolKind::Delegate,
    };
    for (SymbolKind k : typeKinds) {
        if (auto* m = symTab_.lookupModuleByKind(name, k)) return m;
    }
    for (SymbolKind k : typeKinds) {
        if (auto* l = symTab_.lookupLocalByKind(name, k)) return l;
    }
    return nullptr;
}

// <vbeclipse>: 只有**声明名**、拿不到 AST 类型引用的发射器（dll/COM 入口的 extern 原型
// 就是：它手上只有符号表的 ParameterInfo）用这一条。实现是把名字合成一次 SimpleTypeRef
// 再走 mapTypeRef —— 故意的，不复制那份映射。
// 为什么要它：`As DemoShape` 这类工程类形参在符号表里被折成 Vb6Type::Variant（类名丢了），
// 于是 dll 入口的 extern 把它写成 vb6_VARIANT，而类模块自己发的定义是 vb6_cls_DemoShape*：
// 同一方法两份原型。调用点照 extern 那份去装箱 ⇒ C2440（实测 cc_demo / itf_via /
// cls_inh / modulemethod 一系）。Fix 184 当年只补了返回类型，形参是同一处毛病。
std::string CCodeGen::cTypeForDeclaredTypeName(const std::string& name) {
    SimpleTypeRef tr(SourceLocation{}, name);
    return mapTypeRef(&tr);
}

// <vbeclipse> <VBFlexGridDemo>: 形参槽的类型名是不是一个**类型化 COM 接口指针**
// (vb6_ComIface_X*)。这里不能用 cTypeForDeclaredTypeName —— mapTypeRef 内的
// 限定名截短 (line 282) 走 `symTab_.lookupModule(shortName)`，只看当前模块，
// 而 TLB 注入的 IOleControl 通常挂在**全局**符号表；调用点侧当前模块查不到，
// mapTypeRef 兜底给 "void*" —— 于是 "形参 C 类型是 vb6_ComIface_X*" 这个真值
// 在 caller 侧被吞了 (CI t2 一次性 dbg210 dump 实测: tn=OLEGuids.IOleControl
// ct=void*, 而 callee proc 签名同一模块那一份是 vb6_ComIface_IOleControl*)。
// 本函数按 "先看截短名, 再看全名, 都试全局 symTab_.lookup 找 ComClass/ComInterface"
// 走，不受当前模块作用域影响，与 callee proc 签名那一份同源。返回非空的
// "vb6_ComIface_<X>*" 表示这个槽收 typed ptr；空串表示不是接口槽。
std::string CCodeGen::cParamTypedComIfaceCType(const ParameterInfo& p) {
    if (p.typeRefName.empty()) return {};
    std::string shortNm = p.typeRefName;
    size_t dotPos = shortNm.find('.');
    if (dotPos != std::string::npos) shortNm = shortNm.substr(dotPos + 1);
    // 三张网都撒: (1) 全局 symTab_.lookup; (2) 当前模块 symTab_.lookupModule
    // (typelib 注入的 ComInterface 常在当前模块作用域, 见 driver_semantics.cpp
    // line 114-188 每个模块分析阶段都 define 一遍); (3) lookupTypeSymbol 的
    // kind-specific 回退。任一路命中 ComClass/ComInterface 就算接口槽。b5f681ed
    // 的 dbg210b 定位到 raw symTab_.lookup 双 miss, 而 callee proc 签名同一
    // typeName 却解出 vb6_ComIface_X* —— 差就差在 mapTypeRef line 286 用的
    // 是 lookupModule (module scope), 不是 lookup (global)。
    Symbol* clsSym = lookupTypeSymbol(shortNm);
    if (!clsSym) clsSym = lookupTypeSymbol(p.typeRefName);
    if (!clsSym || clsSym->name.empty()) clsSym = symTab_.lookupModule(shortNm);
    if (!clsSym && shortNm != p.typeRefName) clsSym = symTab_.lookupModule(p.typeRefName);
    if (!clsSym) return {};
    if (clsSym->kind != SymbolKind::ComClass
        && clsSym->kind != SymbolKind::ComInterface) return {};
    std::string ifaceNm = (clsSym->kind == SymbolKind::ComClass
        && !clsSym->comDefaultIfaceName.empty())
        ? clsSym->comDefaultIfaceName : clsSym->name;
    return "vb6_ComIface_" + cIdent(ifaceNm) + "*";
}

// <vbeclipse>: "这个形参槽在 C 侧收不收类实例指针" 的**唯一**回答处。判定完全走
// cTypeForDeclaredTypeName → mapTypeRef，也就是类模块发定义时用的那一份映射，所以
// 本条与定义侧不可能各说一套（这正是它要修的毛病）。
// 见 cgen_helpers.inc 声明处的注释：只认 vb6_cls_ / vb6_ivref_，其余一律空串。
std::string CCodeGen::cParamClassPtrType(const ParameterInfo& p) {
    // ⚠ Fix <vbeclipse> rev13: 这里**不能**再拿 `p.type` 当资格。判据只有一条，
    // 就是下面 cTypeForDeclaredTypeName(typeRefName)（= 定义侧那份 mapTypeRef）。
    //
    // 原守卫 `p.type != Vb6Type::Variant → return {}` 看似在挡"真 Variant 槽"，
    // 实际上是把**同一个类型名被折成什么 Vb6Type**当成了第二套判据 —— 而这两者
    // 会分叉：semantic_analyzer_typeref.cpp 的 resolveTypeRef 里，
    //   · 名字命中符号表里的 Class/ComClass（工程 .cls，或 --emit-c 类型库**自动加载**
    //     注进每个模块的 Shell32 `Folder` / ScrRun `Dictionary`）⇒ 返回 Object；
    //   · 名字什么都查不到（跨模块的 .ctl，如 ucView）⇒ 落到末尾兜底 Variant。
    // 于是**同名撞车**时（vbeclipse 的 Folder.cls vs 自动加载的 Shell32 Folder coclass）
    // 形参被折成 Object，本函数认不出类槽 ⇒ 表达式实参走 090d 的 `void*` 档，把
    // `vb6_VariantToObjectVal(...)`（即 vb6_ComObject**包装器**）直接当实例交付：
    //   vb6_ucPerspective_CreateFolder((void*)me,
    //       (&(void*){vb6_VariantToObjectVal(vb6_List_Item(...))}), -1, 0);
    // 而被调方声明是 vb6_cls_Folder** —— `(*Folder)->m_FolderId` 于是取到
    // vb6_ComObject.vb6Instance 这个**指针值**当 BSTR，SysAllocString(乱指针) →
    // av read target=0x1308000 / 0xffffffff，崩点漂移（实测 ucPerspective.CreateFolder
    // → vb6_Folder_prop_get_FolderId → vb6_BSTR_Assign → vb6_BSTR_FromBSTR）。
    // 一个形参槽"收不收类实例指针"只该有一个回答处，就是类型名的映射结果；
    // `p.type` 只是那份映射的一个**可能被符号表污染**的旁证，不能拿来定罪。
    //
    // 安全性：`As Variant` → typeRefName "Variant" → mapTypeRef 给 "vb6_VARIANT"；
    // `As Object`/`As Collection`/`As IPictureDisp`/`As Shell32.Folder` → "void*" /
    // "vb6_ComIface_IFolder*" —— 都不以 vb6_cls_ / vb6_ivref_ 开头，仍然返回空串，
    // 行为与改动前逐字节一致。只有"类型名确实映射成类指针"的槽才会多认出来。
    if (p.typeRefName.empty()) return {};
    const std::string t = cTypeForDeclaredTypeName(p.typeRefName);
    // ⚠ Fix <vbeclipse> rev7: 第二个前缀是 "vb6_ivref_" —— **10** 个字符, 原写 11。
    // std::string::compare(pos, len, str) 会拿 t 的**前 len 个字符**去和 str 比, len=11
    // 时是 11 vs 10 ⇒ 恒不相等 ⇒ **所有接口槽都被判成"非类槽"**。实测
    // `Property Set Scheme(New_Scheme As IScheme)` 打出 cls=0, 于是调用点把它装箱成
    // vb6_VARIANT* 送进本该收 vb6_ivref_IScheme* 的槽。
    // 教训(同 Fix 164z 的"别按默认最通用解锁"一族): **长度字面量必须数**,
    // 写错不报错、静默恒假, 比写错判据更难发现。
    if (t.compare(0, 8, "vb6_cls_") == 0 || t.compare(0, 10, "vb6_ivref_") == 0) {
        return t;
    }
    return {};
}

// <vbeclipse>: 这个名字是不是**本工程的一个类/窗体模块** —— 它的 prop_let_/prop_set_/
// 成员函数真的会被发出来。判法 = 名字命中"本模块 ∪ 驱动下发的其它模块", 且
// mapTypeRef 把它映射成 vb6_cls_* (与 cParamClassPtrType 同一口径)。
// 为什么不能只查符号表: stdole.StdFont / Font / IPicture 这些在 RTL/typelib 里也有
// "类"的痕迹 (RTL 侧就有 `typedef vb6_ComIface_Font vb6_cls_StdFont;`), 但它们**没有**
// 类模块去发 prop_let_ 定义, 按类实例发码等于调一个不存在的函数 ⇒ LNK2019
// (实测 Charts 2020 四个 UserControl 的 `Property Set Font`: vb6_StdFont_prop_let_* 全无定义)。
bool CCodeGen::isProjectClassName(const std::string& name) {
    if (name.empty()) return false;
    const std::string lk = Symbol::toLower(name);
    bool isModule = Symbol::toLower(moduleName_) == lk;
    if (!isModule) {
        for (const auto& m : externalModules_) {
            if (Symbol::toLower(m) == lk) { isModule = true; break; }
        }
    }
    if (!isModule) return false;
    return cTypeForDeclaredTypeName(name).compare(0, 8, "vb6_cls_") == 0;
}

// <vbeclipse>: 把一个实参**交付**给 cParamClassPtrType 认出来的类槽。规则只有一条：
// 实参在 C 侧已经是 vb6_VARIANT 值 (晚绑定取回来的对象、属性读结果…) 就先剥出
// IDispatch 再换实例指针；其余一律原样直传 —— 类变量的 C 值本来就是 vb6_cls_X*，
// Nothing 是 NULL，加一层强制转换只会把真正的类型错配 (接口薄指针喂给类槽) 变成
// 静默的坏指针。不套 boxToVariant：那是"目标槽收 VARIANT"时才有的动作。
std::string CCodeGen::deliverToClassSlot(const std::string& clsPtr,
                                         const std::string& argVal) const {
    if (!clsPtr.empty() && clsPtr.compare(0, 8, "vb6_cls_") == 0 && cExprIsVariant(argVal)) {
        return "(" + clsPtr + ")vb6_ComObject_GetInstance("
               "vb6_VariantToObjectVal(" + argVal + "))";
    }
    return argVal;
}

// <vbeclipse> rev7: 配套 cParamClassPtrType 的**落临时**侧 —— ByRef 类形参收到
// 右值实参 (函数调用/属性读) 时, `&(vb6_ComObject_GetInstance(...))` 是 C2102
// (& 右值), 必须先落一个类型正确的局部再取址。
// 临时名从 C 类型名里**只取标识符字符**并去掉尾随 '*' —— `vb6_ivref_IScheme*`
// 直接截断会得到带 '*' 的变量名 (C2065), 所以逐字符过滤而不是 substr。
std::string CCodeGen::classSlotTempName(const std::string& clsPtr, int idx) const {
    std::string ident;
    for (char ch : clsPtr) {
        if (std::isalnum(static_cast<unsigned char>(ch)) || ch == '_') ident += ch;
    }
    return "vb6_argtmp_" + ident + "_" + std::to_string(idx);
}

std::string CCodeGen::mapTypeRef(ASTNode* typeRef) {
    if (!typeRef) return "vb6_VARIANT";  // 未指定类型 = Variant

    switch (typeRef->kind) {
        case ASTNodeKind::SimpleTypeRef: {
            auto& simple = static_cast<SimpleTypeRef&>(*typeRef);
            // Fix 010: VB6 As Any (Declare语句) → C void*
            // Any在VB6中仅用于Declare语句的参数, 表示"任意类型指针"
            if (simple.name == "Any" || simple.name == "any") {
                return "void*";
            }
            // Fix 010: VBA.库前缀 (如 VBA.ErrObject, VBA.Collection) — 去除前缀
            std::string typeName = simple.name;
            if (typeName.size() > 4 && typeName.compare(0, 4, "VBA.") == 0) {
                typeName = typeName.substr(4);
            }
            // Fix 010: VB6内置对象类型 (Collection, ErrObject等) → void* (对象指针)
            static const std::unordered_set<std::string> vb6BuiltinObjTypes = {
                "Collection", "Forms", "ErrObject", "App", "Screen", "Printer", "Clipboard"
            };
            if (vb6BuiltinObjTypes.count(typeName)) {
                return "void*";
            }
            // 限定类型名 (如 Scripting.Dictionary): 用最后一部分查找符号
            std::string lookupName = typeName;
            size_t dotPos = typeName.find('.');
            if (dotPos != std::string::npos) {
                std::string shortName = typeName.substr(dotPos + 1);
                auto* dotSym = symTab_.lookupModule(shortName);
                if (dotSym) lookupName = shortName;
            }
            // Fix 164z: 项目符号查找必须**先于** typeSys_.resolveTypeName。
            // 原来 resolveTypeName 在前, 其 "vb" 前缀 ⇒ 枚举 ⇒ Long 启发式会抢先
            // 命中项目类名, 如 `Dim This As VBFlexGrid` → resolveTypeName("VBFlexGrid")
            // 首两字母 "vb" 命中 type_system.cpp 的启发式 → Long(int32_t) → 类实例
            // 被声明成 4 字节标量: VBFlexGridBase.bas 的 FlexWindowProc/FlexReaderModeScroll
            // 里 `Dim This As VBFlexGrid` + FlexObjSetAddRef(ObjPtr(Me)) 把 8 字节
            // 对象指针写进 int32_t → 截断 + 后续 AddRef 把类结构体当 COM 解引用 → 0xC0000005.
            // 检查是否是类名 → 映射为类结构体指针
            // Fix 107: 用 lookupTypeSymbol (含 $ty 回退), 否则同名过程会遮蔽类型.
            // tB Interface 契约 (ai/022 B04): 新式接口名 -> 薄指针 vb6_ivref_<I>*.
            // 必须排在 Class 符号分支之前: 头行宿主的 .cls 同时也是一个同名 Class 符号,
            // 走 legacy 分支会得到根本不存在的 vb6_iface_<I> 胖对类型.
            {
                const std::string ivType = ivrefCType(lookupName);
                if (!ivType.empty()) return ivType;
            }
            auto* clsSym = lookupTypeSymbol(lookupName);
            if (clsSym && clsSym->kind == SymbolKind::Class) {
                // P6.4: 接口类 → vb6_iface_<Name> 包装类型 (非指针)
                if (clsSym->isInterface) {
                    usedVb6IfaceTypes_.insert(cIdent(clsSym->name));  // 收集用于前向声明
                    return "vb6_iface_" + cIdent(clsSym->name);
                }
                // Fix 010: 收集类类型名用于前向声明
                usedClassTypes_.insert(cIdent(clsSym->name));
                return "vb6_cls_" + cIdent(clsSym->name) + "*";
            }
            // Fix <vbeclipse> rev7 (回归修复后回位): 工程类名表兜底 —— 只在**符号真值查不到**
            // 时才走。rev7 首版把这条短路放在 lookupTypeSymbol 之前, 出发点是"跨模块工程类
            // (View/Folder/PopupMenu/SplitBar) 在消费模块里常常查不到符号, 掉到末尾兜底
            // `return "void*"` ⇒ ByRef 变 void** 实参装箱错位 → 堆损坏" —— 那个洞是真的,
            // 但位置错了: 短路抢在 Class 分支前面, 连"符号查得到且带 isInterface 标记"
            // (driver 阶段 3.6 对 Implements 宿主逐模块打标) 的名字也一并劫走, 发成裸类指针。
            // 实测 BalloonTooltips: `Dim Subclass As ISubclass` (ISubclass 是 VB_Creatable=True
            // 的接口宿主, 进不了 ivref 表) 被劫成 `vb6_cls_ISubclass*`, 而 Set 侧仍发
            // `vb6_iface_ISubclass_wrap(...)` ⇒ C2440, 整个工程编不过 (GA vbp#2)。
            // 现在的次序: ivrefCType(tB 接口) → Class 符号分支(按 isInterface 真值分流:
            // 宿主发 vb6_iface_<Name> 包装 / 普通类发 vb6_cls_<Name>*) → 本兜底(只救
            // "符号真查不到"的名字, 免得掉 void*) → ComClass 分支。rev17 之后工程类符号
            // 已铺进消费模块, 兜底平时不触发; 判据仍是 projClassNames_ 纯名字表, 查不中
            // 就是真外部类型, 不会把 stdole.Font 之类拉成原生。
            {
                const std::string projCls = projectClassNameOf(typeName);
                if (!projCls.empty()) {
                    usedClassTypes_.insert(cIdent(projCls));
                    return "vb6_cls_" + cIdent(projCls) + "*";
                }
            }
            // P6.3: 检查是否是COM coclass/接口 → 映射为接口指针类型 (前期绑定)
            if (clsSym && (clsSym->kind == SymbolKind::ComClass || clsSym->kind == SymbolKind::ComInterface)) {
                // Fix <vbeclipse>-2 (原) / rev7 (现): 判据是"本工程有没有同名类模块", 不是
                // isExternal. --emit-c 的类型库**自动加载**会把 Shell32 的 coclass `Folder`
                // (默认接口 `IFolder`)、ScrRun 的 `Dictionary`/`FileSystemObject` 等注进**每一个**
                // 模块的符号表, 且这些内建符号 `isExternal == false`(只置了 isBuiltin).
                // vbeclipse 的 Folder.cls 与它同名 —— 在 ucPerspective.ctl 的模块表里
                // `lookupTypeSymbol("Folder")` 会命中**内建 ComClass**, 工程自己的 Class 符号
                // 并不在这一张表里. VB6 规则是"工程内定义优先于引用库"。
                //
                // ⚠ rev7/回归修复: 这条判据现在位于 Class 符号分支**之后** (见上) ——
                // 符号真值优先 (isInterface 宿主发 vb6_iface_ 包装, 普通工程类发 vb6_cls_),
                // 落到本分支的名字只剩"符号查不到而工程名表兜住"和 TLB 内建两类;
                // 工程类名兜底已在上一段抢走, 这里只处理 TLB 符号, 不再重复问那张表。
                // 限定名 (Scripting.Folder) 保留点号 → projectClassNameOf 查不中 → 走 COM,
                // 不会把真·外部同名 coclass 拉成原生。
                //
                // 若误判成 vb6_ComIface_IFolder* 走 COM, 生成的就是
                //   `vb6_cls_List* w = (vb6_cls_List*)(*Folder);`  (ucPerspective.c)
                // —— 把原生 vb6_cls_Folder* 当 IDispatch/List 解引用, 运行期 0xC0000005
                // (vb6_List_Item 读 NULL+0xc).
                // 生成类型化接口指针: vb6_ComIface_<InterfaceName>*
                // 运行时通过vb6_ComQI获取, vtable直接调用
                std::string ifaceName = clsSym->name;
                if (clsSym->kind == SymbolKind::ComClass && !clsSym->comDefaultIfaceName.empty()) {
                    ifaceName = clsSym->comDefaultIfaceName;
                }
                std::string cIfaceName = cIdent(ifaceName);
                usedComIfaceTypes_.insert(cIfaceName);  // 收集接口类型名用于typedef
                return "vb6_ComIface_" + cIfaceName + "*";
            }
            // 检查是否是用户定义类型 (UDT) → vb6_type_<Name>
            // Fix 107: 同 clsSym, 用类型感知查找避免同名过程遮蔽.
            auto* udtSym = lookupTypeSymbol(lookupName);
            if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                // Fix 010: 收集UDT类型名用于前向声明
                usedUdtTypes_.insert(cIdent(simple.name));
                return "vb6_type_" + cIdent(simple.name);
            }
            // 检查是否是枚举类型 → 基础类型int32_t (VB6枚举底层是Long)
            if (udtSym && udtSym->kind == SymbolKind::EnumType) {
                return "int32_t";
            }
            // Delegate 类型 (tB 扩展) → intptr_t: 委托值 = 生成调用桩的地址,
            // 与 LongPtr 位兼容 (语义层 resolveTypeRef 同口径).
            if (udtSym && udtSym->kind == SymbolKind::Delegate) {
                return "intptr_t";
            }
            // Fix 164z: 项目符号 (类/COM/UDT/枚举) 全部查空后才允许走类型系统启发式。
            // 理由见上方 Fix 164z 注释。VB./VBA. 限定的对象末段 (Control/Form 等) 在
            // type_system.cpp 最前单独处理 (返回 Object), 不受本次移位影响。
            Vb6Type t = typeSys_.resolveTypeName(typeName);
            if (t != Vb6Type::Unknown) {
                return mapType(t);
            }
            // Fix 010b: VB6标准库与外部COM库的类型映射
            // 以下类型不在项目符号表中, 但均为VB6/COM标准类型
            // 注意: 能在符号表中找到的用户类型会已在上面被处理

            // VB6语言类型别名
            // Fix 081e: LongPtr now has its own Vb6Type::LongPtr → intptr_t
            // (handled by resolveTypeName + mapType, this fallback is for edge cases)
            // 账 #228: 从这里往下全是**按名字形状/名单**的内在与外部库别名档, 问的必须是
            // 类型本名而不是"库.本名"整串 —— VB6 里 `As OLE_COLOR` 与 `As stdole.OLE_COLOR`
            // 是同一种类型的两种拼法, 库前缀只是出处。上面那几档符号分支 (ivref/Class/UDT/
            // 枚举/ComClass) 各按自己的规则处理限定名 (lookupName 那里"折裸名"要符号真查得到),
            // 能落到这里的名字已经把所有项目符号都查空了, 再拿整串去比对名单就是第二套口径:
            // 实测限定那一形以前一路掉到本函数末尾的兜底 `void*`, 而裸名答 `int32_t`
            // —— 同一枚 VBFlexGrid 事件 (`.ctl` 写 OLE_COLOR、容器写 stdole.OLE_COLOR) 于是
            // 发送侧交 4 字节、处理器收 8 字节指针, x64 上高 32 位是垃圾。
            std::string aliasName = lookupName;
            {
                const size_t dp8 = aliasName.rfind('.');
                if (dp8 != std::string::npos) aliasName = aliasName.substr(dp8 + 1);
            }

            if (aliasName == "LongPtr") {
                return "intptr_t";
            }
            if (aliasName == "LongLong") {
                return "int64_t";   // Fix 084m: 恒 64 位有符号 (与 LongPtr 的架构宽度不同)
            }
            // C3 扩展 (ai/032): 新整型兜底。正规路径是 typeSys_.resolveTypeName
            // 认出名字后返回 mapType(t), 走到这里说明符号表与类型系统都没认出来 ——
            // 与 LongLong 一样再钉一次, 免得同一种类型有两种答案。
            if (aliasName == "SByte")     return "int8_t";
            if (aliasName == "UInteger")  return "uint16_t";
            if (aliasName == "ULong")     return "uint32_t";
            if (aliasName == "ULongLong") return "uint64_t";
            // VB6内置枚举类型 (Vb前缀): VbCompareMethod, VbTriState, VbFileAttribute等
            // VB6枚举底层是Long (int32_t)
            if (aliasName.size() >= 2 && aliasName.compare(0, 2, "Vb") == 0) {
                return "int32_t";
            }
            // COM类型别名 (OLE_前缀): OLE_COLOR, OLE_HANDLE等, 通常为DWORD
            if (aliasName.size() >= 4 && aliasName.compare(0, 4, "OLE_") == 0) {
                return "int32_t";
            }
            // ADODB等外部COM库枚举类型: 名称以Enum结尾
            // 如 EventStatusEnum, ExecuteOptionEnum, CursorTypeEnum等
            if (aliasName.size() >= 4 &&
                aliasName.compare(aliasName.size() - 4, 4, "Enum") == 0) {
                return "int32_t";
            }
            // ADODB等外部COM库对象类型 → void* (COM对象指针)
            static const std::unordered_set<std::string> comObjTypes = {
                "Connection", "Recordset", "Command", "Parameter",
                "Field", "Fields", "Error", "Errors", "Property",
                "Properties", "Stream"
            };
            if (comObjTypes.count(aliasName)) {
                return "void*";
            }
            // Fix 010c: VB6标准枚举类型别名 (不带Vb前缀的常用枚举)
            // 这些类型在VB6中等价于对应的Vb*枚举, 底层都是Long
            static const std::unordered_set<std::string> vb6EnumAliases = {
                "CompareMethod", "TriState", "FirstDayOfWeek", "FirstWeekOfYear",
                "MsgBoxResult", "MsgBoxStyle", "FileAttribute", "DateFormat",
                "Calendar", "DateTimeFormat", "CallType", "VariantType",
                "VarType", "QueryDef", "EditModeEnum", "FieldAttributeEnum"
            };
            if (vb6EnumAliases.count(aliasName)) {
                return "int32_t";
            }
            // 兜底: 未知类型 (如窗体模块名、外部COM类型别名等) → void*
            // Fix 010c: 窗体模块(.frm)未注册为Class符号, 但VB6中可作为类型使用
            // 任何不是内置类型/UDT/枚举/类/COM类型的名称都视为通用对象指针
            return "void*";
        }
        case ASTNodeKind::ArrayTypeRef:
            return "vb6_SafeArray1D*";  // SAFEARRAY指针
        case ASTNodeKind::FixedStringTypeRef:
            return "BSTR";
        default:
            return "vb6_VARIANT";
    }
}

// Fix 081e: Declare函数返回类型映射
// 在VB6 Declare语句中, 返回值Long常用于返回句柄/指针(HDC/HBITMAP/HWND等)。
// x64下int32_t只有4字节,无法容纳8字节指针,导致截断和后续崩溃。
// LongPtr已经通过mapType映射为intptr_t, 这里只需将Long返回值也映射为intptr_t。
// 映射为intptr_t: x86下4字节(兼容), x64下8字节(与指针同大小)。
std::string CCodeGen::mapDeclareType(ASTNode* typeRef) {
    if (!typeRef) return "vb6_VARIANT";
    std::string base = mapTypeRef(typeRef);
    if (base == "int32_t") {
        // 检查是否是Long类型 (LongPtr已通过mapType返回intptr_t)
        if (typeRef->kind == ASTNodeKind::SimpleTypeRef) {
            auto& simple = static_cast<SimpleTypeRef&>(*typeRef);
            if (simple.name == "Long") {
                return "intptr_t";
            }
        }
    }
    return base;
}

// Fix 010b: 常量折叠 — 将VB6 AST表达式求值为int64_t编译期常量
// 用于enum成员值 (如 2^0 → 1, 2^1|2^2 → 6, &H10 → 16)
bool CCodeGen::tryEvalConstInt(ASTNode* expr, int64_t& result) {
    if (!expr) return false;

    switch (expr->kind) {
    case ASTNodeKind::LiteralExpr: {
        auto* lit = static_cast<LiteralExpr*>(expr);
        switch (lit->literalKind) {
        case LiteralKind::Integer:
            result = lit->intValue;
            return true;
        case LiteralKind::Long:
            result = lit->longValue;
            return true;
        case LiteralKind::LongPtr:  // Fix 082: ^ 后缀 (VBA7)
            result = lit->longValue;
            return true;
        case LiteralKind::Boolean:
            result = lit->boolValue ? 1 : 0;
            return true;
        case LiteralKind::Double:
        case LiteralKind::Single:
            // Integer-only folding must not truncate floating operands.
            // E.g. PI / 2 must remain a floating expression, not 3 / 2 == 1.
            return false;
        default:
            return false;
        }
    }
    case ASTNodeKind::UnaryExpr: {
        auto* unary = static_cast<UnaryExpr*>(expr);
        int64_t val;
        if (!tryEvalConstInt(unary->operand.get(), val)) return false;
        switch (unary->op) {
        case UnaryOp::Negate:
            result = -val;
            return true;
        case UnaryOp::Not:
            result = ~val;
            return true;
        }
        return false;
    }
    case ASTNodeKind::BinaryExpr: {
        auto* bin = static_cast<BinaryExpr*>(expr);
        int64_t l, r;
        if (!tryEvalConstInt(bin->left.get(), l)) return false;
        if (!tryEvalConstInt(bin->right.get(), r)) return false;
        switch (bin->op) {
        case BinaryOp::Add: result = l + r; return true;
        case BinaryOp::Sub: result = l - r; return true;
        case BinaryOp::Mul: result = l * r; return true;
        case BinaryOp::Div:
            // VB6 '/' is floating division even when both operands are integers.
            return false;
        case BinaryOp::IntDiv:
            if (r == 0) return false;
            result = l / r; return true;
        case BinaryOp::Mod:
            if (r == 0) return false;
            result = l % r; return true;
        case BinaryOp::Pow: {
            // 整数幂运算
            if (r < 0) return false;
            int64_t base = l, exp = r;
            int64_t pw = 1;
            while (exp > 0) {
                if (exp & 1) pw *= base;
                base *= base;
                exp >>= 1;
            }
            result = pw;
            return true;
        }
        case BinaryOp::Or:  result = l | r; return true;
        case BinaryOp::And: result = l & r; return true;
        case BinaryOp::Xor: result = l ^ r; return true;
        case BinaryOp::Eqv: result = ~(l ^ r); return true;
        case BinaryOp::Imp: result = (~l) | r; return true;
        default:
            return false;  // 比较、连接等不适用于enum常量
        }
    }
    case ASTNodeKind::IndexOrCallExpr: {
        // 处理 vb6_Pow(base, exp) 调用
        auto* call = static_cast<IndexOrCallExpr*>(expr);
        if (!call->callee) return false;
        // 提取被调用者名称
        std::string fnName;
        if (auto* id = dynamic_cast<IdentifierExpr*>(call->callee.get())) {
            fnName = id->name;
        } else {
            return false;
        }
        // 转小写比较
        std::string fnLower = fnName;
        for (auto& c : fnLower) c = (char)tolower(c);
        if (fnLower == "pow" && call->positional.size() == 2) {
            int64_t base, exp;
            if (!tryEvalConstInt(call->positional[0].get(), base)) return false;
            if (!tryEvalConstInt(call->positional[1].get(), exp)) return false;
            if (exp < 0) return false;
            int64_t pw = 1;
            while (exp > 0) {
                if (exp & 1) pw *= base;
                base *= base;
                exp >>= 1;
            }
            result = pw;
            return true;
        }
        return false;
    }
    case ASTNodeKind::IdentifierExpr: {
        auto* id = static_cast<IdentifierExpr*>(expr);
        // 同枚举兄弟成员优先 (VbEclipse 批次): `MSG_BOTH = MSG_AFTER Or MSG_BEFORE` 里
        // 的裸兄弟名在 cgen 这一层用 symTab_ 查不到 (见 cgen_state.inc 注释), 不先查这张
        // 表就会走回退路径吐出裸标识符 → C2065/C2057。
        if (!enumSiblingConsts_.empty()) {
            // 走 cIdent: 引用侧也可能带方括号转义 (`[MSG_AFTER]`), 与建表侧同形才查得到
            std::string low = cIdent(id->name);
            for (auto& c : low) c = (char)tolower(c);
            auto sib = enumSiblingConsts_.find(low);
            if (sib != enumSiblingConsts_.end()) {
                result = sib->second;
                return true;
            }
        }
        // Fix 010c: 查找符号表中的常量 (跨模块Public Const)
        auto* sym = symTab_.lookup(id->name);
        if (sym && sym->kind == SymbolKind::Constant && sym->hasConstValue) {
            result = sym->constIntValue;
            return true;
        }
        // 也检查枚举成员
        if (sym && sym->kind == SymbolKind::EnumMember) {
            result = sym->constIntValue;
            return true;
        }
        return false;
    }
    default:
        return false;
    }
}

} // namespace vb6c3
