#include "semantics/semantic_analyzer.hpp"
#include "common/vba_math.hpp"  // VBA 标准模块名单: 模块不能当类型用 (Math 定案 §10.4)
#include <algorithm>
#include <cctype>
#include <tuple>
#include <initializer_list>
#include "semantics/semantic_analyzer_internal.h"

namespace vb6c3 {

// --- semantic_analyzer_typeref.cpp: 类型引用解析: resolveTypeOrDefault / resolveTypeRef ---


// ============================================================
// 类型引用解析
// ============================================================

Vb6Type SemanticAnalyzer::resolveTypeOrDefault(const std::string& name, ASTNode* typeRef) {
    Vb6Type t = resolveTypeRef(typeRef);
    // P22: If no explicit type and DefType is active, use DefType inference
    if (!typeRef && defTypeActive_ && !name.empty()) {
        char first = toupper(name[0]);
        int idx = first - 'A';
        if (idx >= 0 && idx < 26 && defTypeMap_[idx] != Vb6Type::Variant) {
            return defTypeMap_[idx];
        }
    }
    return t;
}
Vb6Type SemanticAnalyzer::resolveTypeRef(ASTNode* typeRef) {
    if (!typeRef) return Vb6Type::Variant;  // VB6默认: Variant

    switch (typeRef->kind) {
        case ASTNodeKind::SimpleTypeRef: {
            auto& simple = static_cast<SimpleTypeRef&>(*typeRef);
            // VBA 标准模块不能当类型用（真 VB6: 「Visual Basic 中不支持自动化类型」）。
            // 限定写法（`VBA.Math` / `VBA.Strings` …）一律按标准模块判；裸名（`Math`）
            // 只在工程里没有同名类型时才判 —— 用户自己定义的 `Class Math` 仍按用户类型解析。
            // 名单在 common/vba_math.hpp（与发码层同一张表）。
            {
                const std::string lowerType = Symbol::toLower(simple.name);
                const bool vbaQual = lowerType.size() > 4
                                     && lowerType.compare(0, 4, "vba.") == 0;
                if (isVbaStandardModuleName(lowerType)
                    && (vbaQual || !symTab_.lookupModule(lowerType))) {
                    diag_.error(DiagnosticID::SemModuleAsType, typeRef->loc,
                        "Visual Basic 中不支持自动化类型: '" + simple.name +
                        "' (标准模块不能当类型用)");
                    return Vb6Type::Object;
                }
            }
            // ai/023 S04: 非导出包类作类型名 (Dim x As HiddenCls / 参数 / 返回值)
            // → VB7006。不拦的话类型退化成 Object/void*, 与后期绑定 COM 同一结局。
            reportIfPackageClassBlocked(simple.name, typeRef->loc);
            Vb6Type t = typeSys_.resolveTypeName(simple.name);
            if (t == Vb6Type::Unknown) {
                // Fix 069: "As Any" 是 VB6 Declare 语句中故意使用的基础类型,
                // 映射为 Vb6Type::Unknown 且不应被回退为 Variant.
                // 原代码将所有 Unknown 统一回退为 Variant, 导致 As Any 参数
                // 在 calleeParams 中类型为 Variant → 调用点生成 VARIANT 复合字面量
                // 而非 (void*) 指针 → MSVC C2440 类型转换错误.
                std::string lowerName = Symbol::toLower(simple.name);
                if (lowerName == "any") {
                    return Vb6Type::Unknown;
                }
                // Fix 040a: VB6内置对象类型 (Collection, ErrObject等) → Object (void*).
                // 与 cgen_base.cpp mapTypeRef 的 vb6BuiltinObjTypes 集合保持一致.
                // 若返回 Variant, 则 ByVal Collection 参数会被 Fix 024 P2 错误地用
                // vb6_VariantFromValue() 包装 void* 指针 → C2172 (实参不是指针).
                static const std::unordered_set<std::string> vb6BuiltinObjTypes = {
                    "Collection", "Forms", "ErrObject", "App", "Screen", "Printer", "Clipboard",
                    // VB6 内建对象类型: 通用控件/窗体都是对象引用 (void*).
                    // 未识别会被宽松回退为 Variant, 与 cgen_base mapTypeRef 的
                    // "未知类型 → void*" 兜底不一致 → ByRef 对象实参被误包成
                    // VARIANT 复合字面量 (BalloonTooltips cTT.CreateToolTip
                    // ParentControl As Control → *objControl 读到 vt).
                    "Control", "Form",
                    // Fix 157: stdole / OLE 自动化**智能指针式接口类型** (Picture/Font 系列).
                    // 语义层对未识别类型名回退 Variant, 而 codegen 侧 cgen_base_type.cpp
                    // mapTypeRef 的兜底是 "未知类型 → void*" (见其末尾 return) —— 两侧不一致
                    // 就会方向相反地误包装: Fix 086 见形参 Variant 便把实参套
                    // vb6_VariantFromValue (→ vb6_VARIANT 结构体), 再传给 C 形参 void*
                    // → C2172 "实参不是指针".
                    //   VisualStyles.bas 527:
                    //     Private Function CoalescePicture(ByVal Picture As IPictureDisp,
                    //                                      ByVal DefaultPicture As IPictureDisp)
                    //   → static void* vb6_VisualStyles_CoalescePicture(void*, void*)
                    //   → 调用点 CoalescePicture(vb6_VariantFromValue(vb6_VariantFromComResult
                    //     (vb6_ComGetProp(Button, L"DisabledPicture"))), ...) C2172 参数1/2.
                    // 判为 Object 后改走 Fix 029 的 Variant→Object 提取
                    // (vb6_VariantToObjectVal 返回 void*), 与 C 签名同向.
                    // 只收录**带 I/Std/Disp 前后缀**的 COM 接口拼写: 裸 "Picture"/"Font"
                    // 可能是工程自定义类名, 且本集合的检查发生在符号表查找**之前**,
                    // 收录裸名会遮蔽用户类 → 刻意不加.
                    "IPicture", "IPictureDisp", "StdPicture", "PictureDisp",
                    "IFont", "IFontDisp", "StdFont", "FontDisp"
                };
                // Fix 040a: strip VBA. prefix (e.g. VBA.ErrObject → ErrObject)
                std::string typeName = simple.name;
                if (typeName.size() > 4 && typeName.compare(0, 4, "VBA.") == 0) {
                    typeName = typeName.substr(4);
                }
                if (vb6BuiltinObjTypes.count(typeName)) {
                    return Vb6Type::Object;
                }
                // Fix 157: 库限定类型名 (stdole.IPictureDisp / ole.IFontDisp) 用末段再查一次.
                // 下方既有的 dot 分支只处理"末段能在符号表找到类/ComClass"的情形;
                // 未注册进类型库的 stdole 接口会整串落到 Variant 兜底 → 与本集合不一致.
                size_t typeNameDot = typeName.find('.');
                if (typeNameDot != std::string::npos
                    && vb6BuiltinObjTypes.count(typeName.substr(typeNameDot + 1))) {
                    return Vb6Type::Object;
                }
                                // 可能是用户自定义类型 -> 在符号表中查找
                std::string lower = Symbol::toLower(simple.name);
                if (auto* sym = symTab_.lookupModule(lower)) {
                    if (sym->kind == SymbolKind::UserDefinedType)
                        return Vb6Type::UserDefinedType;
                    if (sym->kind == SymbolKind::Delegate)
                        return Vb6Type::LongPtr;  // 委托值与 LongPtr 位兼容
                    if (sym->kind == SymbolKind::EnumType)
                        return Vb6Type::Long;  // Enum成员是Long
                    if (sym->kind == SymbolKind::ComClass || sym->kind == SymbolKind::ComInterface || sym->kind == SymbolKind::ComModule || sym->kind == SymbolKind::ComGlobalNs)
                        return Vb6Type::Object;
                    if (sym->kind == SymbolKind::Class)
                        return Vb6Type::Object;
                }
                // 限定类型名 (如 Scripting.Dictionary): 用最后一部分查找
                size_t dotPos = simple.name.find('.');
                if (dotPos != std::string::npos) {
                    std::string shortName = simple.name.substr(dotPos + 1);
                    std::string shortLower = Symbol::toLower(shortName);
                    if (auto* sym2 = symTab_.lookupModule(shortLower)) {
                        if (sym2->kind == SymbolKind::ComClass || sym2->kind == SymbolKind::ComInterface || sym2->kind == SymbolKind::ComModule)
                            return Vb6Type::Object;
                        if (sym2->kind == SymbolKind::Class)
                            return Vb6Type::Object;
                    }
                }
                // Fix 050: COM 枚举类型 (如 DataTypeEnum, CursorTypeEnum, LockTypeEnum)
                // 来自引用的 COM 类型库 (如 ADODB)，不在项目符号表中。
                // VB6 中所有 Enum 类型都是 Long (32位整数)，因此以 "Enum" 结尾的
                // 未识别类型名应返回 Long 而非 Variant，避免 vb6_VARIANT→int32_t C2440 错误。
                {
                    std::string lowerName = Symbol::toLower(simple.name);
                    if (lowerName.size() > 4 &&
                        lowerName.compare(lowerName.size() - 4, 4, "enum") == 0) {
                        return Vb6Type::Long;
                    }
                }
                // 未识别类型 → Variant (宽松策略)
                return Vb6Type::Variant;
            }
            return t;
        }
        case ASTNodeKind::ArrayTypeRef: {
            auto& arr = static_cast<ArrayTypeRef&>(*typeRef);
            // 数组元素类型
            Vb6Type elemType = resolveTypeRef(arr.elementType.get());
            return static_cast<Vb6Type>(static_cast<uint16_t>(elemType) |
                                         static_cast<uint16_t>(Vb6Type::Array));
        }
        case ASTNodeKind::FixedStringTypeRef:
            return Vb6Type::String;  // 定长字符串仍为String类型
        default:
            return Vb6Type::Variant;
    }
}
} // namespace vb6c3
