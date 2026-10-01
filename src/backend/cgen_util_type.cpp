#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// Fix <VBFlexGridDemo/GA 36766856082>: 内置函数与"控件自己声明的同名 Property"撞名时的
// 类型裁决。
//
// 背景: Task #40 (40960eb) 给 parseIdentifierOrCall 加了"使用点剥类型后缀",
// `Left$(Value, 1)` 到 cgen 时名字已是 `Left`; VBFlexGrid.ctl 自己声明了
// `Public Property Get Left() As Single`, 于是 inferExprType 的符号表支路命中这个
// **属性**符号 ⇒ 答 Single ⇒ `float _vb6_select_81 = vb6_Left(Temp, 1)` 把 BSTR
// 强转 float → C2440 (27 处)。剥后缀前名字是 `Left$`, 查不到该属性符号, 天然不撞。
//
// 口径 (刻意收窄): 只有"符号 kind 是 Property* **且** 该名字在内置函数表里"才启用,
// 其余一切照旧 —— 用户 Function 同名、普通变量、真·属性访问都不受影响。
// 返回类型按 VB6 内置函数语义给: 字符串类答 String, 整数类答 Long, 其余答 Variant
// (Variant 是既有兜底, 不会引入新的硬错)。
#include "backend/detail/expr/cgen_expr_ident_builtin_table.inc"

static Vb6Type builtinFuncReturnTypeForShadowedProp(const std::string& nameLower) {
    static const std::unordered_set<std::string> kStrFuncs = {
        "left", "mid", "right", "trim", "ltrim", "rtrim", "lcase", "ucase",
        "space", "string", "str", "chr", "hex", "oct", "format", "strconv",
        "input", "command", "curdir", "environ", "dir", "date", "time", "monthname",
        "weekdayname", "error"
    };
    if (kStrFuncs.count(nameLower)) return Vb6Type::String;
    static const std::unordered_set<std::string> kLongFuncs = {
        "len", "lenb", "instr", "instrb", "asc", "cint", "clng", "cbyte",
        "freefile", "erl", "hour", "minute", "second", "weekday", "day",
        "month", "year", "abs", "sgn", "int", "fix"
    };
    if (kLongFuncs.count(nameLower)) return Vb6Type::Long;
    return Vb6Type::Variant;
}

// 判据: 符号是 Property* **且** 该名字在内置函数表里 ⇒ 这是"属性遮蔽了内置函数",
// 类型要按内置函数答 (见上注释)。剥 $ 后缀后比, 与 ident_builtin 的 lookupName 同口径。
static bool isPropShadowingBuiltinFunc(const Symbol* sym, const std::string& idName) {
    if (!sym) return false;
    if (sym->kind != SymbolKind::PropertyGet && sym->kind != SymbolKind::PropertyLet
        && sym->kind != SymbolKind::PropertySet) return false;
    std::string n = Symbol::toLower(idName);
    if (!n.empty() && n.back() == '$') n.pop_back();
    return kBuiltinFuncNames.count(n) != 0;
}

static Vb6Type cgenTypeForShadowedBuiltin(const std::string& idName) {
    std::string n = Symbol::toLower(idName);
    if (!n.empty() && n.back() == '$') n.pop_back();
    return builtinFuncReturnTypeForShadowedProp(n);
}

// --- cgen_util_type.cpp: 表达式类型推断 + Variant 判定 + 运行时参数 C 类型 ---


// ============================================================
void CCodeGen::visit(SimpleTypeRef& node) {}
void CCodeGen::visit(ArrayTypeRef& node) {}
void CCodeGen::visit(FixedStringTypeRef& node) {}

// ============================================================
// 模块 visit (generate()已按类别分派, 此处为空)
// ============================================================

void CCodeGen::visit(Module& node) {}

// ============================================================
// 表达式类型推断 (简化版, 用于Select Case等场景)
// ============================================================

Vb6Type CCodeGen::inferExprType(Expr& expr) const {
    switch (expr.kind) {
        case ASTNodeKind::IdentifierExpr: {
            auto& id = static_cast<IdentifierExpr&>(expr);
            // 优先检查已知的变量类型集合 (局部变量在符号表中作用域可能不可达)
            std::string lower = id.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            // Fix 161f: 无 As Type 但整数可折叠的 Const (符号表里 type=Variant)。
            // C 侧已 emit 成 `#define NAME (4127)`, 标识符也内联为数值; 若这里
            // 仍答 Variant, Declare 调用的实参会被包成
            // vb6_VariantToLong(LVM_GETHEADER) → C2440 "无法从 int 转换为
            // vb6_VARIANT" (extlist MListView SendMessage 实测)。按 Long 答。
            if (moduleIntConstValues_.count(lower)) return Vb6Type::Long;
            if (knownBstrVars_.count(lower)) return Vb6Type::String;
            if (knownSingleVars_.count(lower)) return Vb6Type::Single;
            // Fix 175: Date 必须先于 Double 判 (Date 变量同时登记在 knownDoubleVars_
            // 里以复用既有 double 取值路径, 口径同 Fix 117c 的 Single)。
            if (knownDateVars_.count(lower)) return Vb6Type::Date;
            if (knownDoubleVars_.count(lower)) return Vb6Type::Double;
            // ai/022 W1: 必须先于 knownLongVars_ 判 (口径同 Fix 175 的 Date) ——
            // As Boolean 的 C 型与 Integer 同串, 只按 C 类型登记就永远看不见布尔。
            if (knownBoolVars_.count(lower)) return Vb6Type::Boolean;
            // 账 #123: 口径同 Fix 175 的 Date / W1 的 Boolean —— 登记过就必须由这张表答 Byte,
            // 让局部/形参 Byte 与模块级 Byte (走符号表那条支路) 给出同一份答案。
            if (knownByteVars_.count(lower)) return Vb6Type::Byte;
            if (knownLongVars_.count(lower)) return Vb6Type::Long;
            if (knownLongPtrVars_.count(lower)) return Vb6Type::LongPtr;
            if (knownVariantVars_.count(lower)) return Vb6Type::Variant;
            // 检查符号表
            auto* sym = symTab_.lookup(id.name);
            if (!sym) sym = symTab_.lookupModule(id.name);
            // Fix <VBFlexGridDemo>: 属性遮蔽内置函数时按内置函数的类型答 (见文件头注释)。
            // 但**函数体内 `Name = expr` 引用返回槽**这一支例外 —— VB6 里 `Left = Extender.Left`
            // 在 `Public Property Get Left() As Single` 内部指的是 vb6_ret_Left (float),
            // 不是内置 Left$()。判据同 cgen_assign_value_sem.inc:5-12 —— currentProc_ 是
            // Function/PropertyGet 且名字等于本标识符时, 走返回槽 (sym->type)。
            // 修前: inferExprType 答 String ⇒ vb6_BSTR_AssignMove(&vb6_ret_Left, vb6_CStrDbl(...))
            //       ⇒ BSTR_Free 把 float 位当指针 ⇒ MainForm Form_Resize 一调 VBFlexGrid1.Left
            //       即 0xC0000005, 启动看不到界面。
            // 只改 IdentifierExpr 分支; IndexOrCallExpr 那一支 (本文件 178 行附近) 不动 ——
            // 那里 Left(Temp, 1) 就是真·内置函数调用。
            if (sym && isPropShadowingBuiltinFunc(sym, id.name)) {
                const bool isRetSlotSelfRef =
                    currentProc_
                    && (currentProc_->kind == SymbolKind::Function
                        || currentProc_->kind == SymbolKind::PropertyGet)
                    && Symbol::toLower(currentProc_->name) == Symbol::toLower(id.name);
                if (!isRetSlotSelfRef) return cgenTypeForShadowedBuiltin(id.name);
            }
            if (sym) return sym->type;
            break;
        }
        case ASTNodeKind::LiteralExpr: {
            auto& lit = static_cast<LiteralExpr&>(expr);
            if (lit.literalKind == LiteralKind::String) return Vb6Type::String;
            if (lit.literalKind == LiteralKind::Double || lit.literalKind == LiteralKind::Single
                || lit.literalKind == LiteralKind::Currency || lit.literalKind == LiteralKind::Decimal)
                return Vb6Type::Double;
            if (lit.literalKind == LiteralKind::Boolean) return Vb6Type::Boolean;
            if (lit.literalKind == LiteralKind::Date) return Vb6Type::Date;
            // 装不进 32 位的 Long 字面量, cgen 实际发的是 64 位 (LL 后缀), 这里
            // 必须跟着答 LongLong。原先一律答 Long, 于是收窄检查把它当成"装得下"
            // 跳过 —— `l As Long: l = -9223372036854775807` 就从 Error 6 变成
            // 静默截成 1。类型口径要和实际发出的位宽一致, 否则检查等于没做。
            if (lit.literalKind == LiteralKind::Long
                && (lit.longValue < INT32_MIN || lit.longValue > INT32_MAX))
                return Vb6Type::LongLong;
            return Vb6Type::Long;
        }
        case ASTNodeKind::BinaryExpr: {
            auto& bin = static_cast<BinaryExpr&>(expr);
            // 字符串连接运算符 → String
            if (bin.op == BinaryOp::Concat) return Vb6Type::String;
            // 比较运算符 → Boolean
            if (bin.op == BinaryOp::Eq || bin.op == BinaryOp::Neq ||
                bin.op == BinaryOp::Lt || bin.op == BinaryOp::Gt ||
                bin.op == BinaryOp::Le || bin.op == BinaryOp::Ge ||
                bin.op == BinaryOp::Like || bin.op == BinaryOp::Is)
                return Vb6Type::Boolean;
            // 逻辑运算符 → Boolean (VB6中)
            if (bin.op == BinaryOp::And || bin.op == BinaryOp::Or || bin.op == BinaryOp::Xor)
                return Vb6Type::Boolean;
            // 浮点除法 → Double
            if (bin.op == BinaryOp::Div) return Vb6Type::Double;

            // 算术运算符: 提升左右类型
            {
                Vb6Type lt = inferExprType(*bin.left);
                Vb6Type rt = inferExprType(*bin.right);
                // Fix 175: VB6 日期算术 —— `Date ± 数值` 仍是 Date, `Date - Date` 是
                // 天数差 (Double), `*` `\` `Mod` `^` 无日期语义按数值处理。
                // TypeSystem::isNumeric(Date)==false, 交给 promote 会掉出数值阶梯。
                if (lt == Vb6Type::Date || rt == Vb6Type::Date) {
                    if (bin.op == BinaryOp::Add) return Vb6Type::Date;
                    if (bin.op == BinaryOp::Sub)
                        return (lt == Vb6Type::Date && rt == Vb6Type::Date)
                                   ? Vb6Type::Double : Vb6Type::Date;
                    return Vb6Type::Double;
                }
                return TypeSystem::promote(lt, rt);
            }
        }
        case ASTNodeKind::UnaryExpr: {
            auto& un = static_cast<UnaryExpr&>(expr);
            if (un.op == UnaryOp::Not) return Vb6Type::Boolean;
            return inferExprType(*un.operand);
        }
        case ASTNodeKind::IndexOrCallExpr: {
            // 函数调用: 返回函数返回类型
            auto& call = static_cast<IndexOrCallExpr&>(expr);
            if (call.callee && call.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& id = static_cast<IdentifierExpr&>(*call.callee);
                auto* sym = symTab_.lookup(id.name);
                if (!sym) sym = symTab_.lookupModule(id.name);
                // Fix <VBFlexGridDemo>: 同上 —— 属性遮蔽内置函数时按内置函数的类型答,
                // 否则 `float _vb6_select_81 = vb6_Left(Temp, 1)` 把 BSTR 强转 float (C2440)。
                if (sym && isPropShadowingBuiltinFunc(sym, id.name)) {
                    return cgenTypeForShadowedBuiltin(id.name);
                }
                if (sym) return sym->type;
            }
            // P26: 类实例方法调用 a.Method(args) → callee是MemberAccessExpr
            // 需要查询成员函数的返回类型, 而非直接fallback到Variant
            if (call.callee && call.callee->kind == ASTNodeKind::MemberAccessExpr) {
                return inferExprType(*call.callee);
            }
            break;
        }
        case ASTNodeKind::MemberAccessExpr: {
            auto& ma = static_cast<MemberAccessExpr&>(expr);
            // C29-1a: 内建窗体控件的整数型属性必须在这里就判成 Long。`Left` 同时是 VB
            // 内置函数名, 落到下面的符号表查找会被折成 String (Fix 081i 记过同一类),
            // 于是 `If Line1.Left = 600` 生成 vb6_StrCmp(整数值, BSTR) —— 把 600 当
            // 指针解引用, 运行期直接段错误 (实测就是这条)。口径同上面的
            // UserControl.ScaleWidth 分支: 按"对象是内建控件 + 属性名"给类型。
            // 只认内建控件类型: 工程内 .ctl 宿主同名属性 (如 Value) 由它自己的符号表说话。
            static const char* const kNumericFc[] = {
                "left", "top", "width", "height",
                "shape", "fillstyle", "borderwidth", "borderstyle",
                "fillcolor", "bordercolor", "x1", "y1", "x2", "y2",
            };
            std::string fcName;   // 命中的内建控件名 (空 = 这不是内建控件的属性访问)
            if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                fcName = Symbol::toLower(static_cast<IdentifierExpr&>(*ma.object).name);
            } else if (ma.object && ma.object->kind == ASTNodeKind::IndexOrCallExpr) {
                // 控件数组的元素 (`lamp(1).Left`): 对象位是 `名字(下标)`, 同一条规则。
                // 只认确实在 knownFormControls_ 里的名字, 所以函数调用返回对象
                // (`GetWidget(1).Left`) 不会被误判。
                auto& callFc = static_cast<IndexOrCallExpr&>(*ma.object);
                if (callFc.callee && callFc.callee->kind == ASTNodeKind::IdentifierExpr
                    && callFc.positional.size() == 1) {
                    fcName = Symbol::toLower(
                        static_cast<IdentifierExpr&>(*callFc.callee).name);
                }
            }
            if (!fcName.empty()) {
                auto fcIt = knownFormControls_.find(fcName);
                if (fcIt != knownFormControls_.end() && fcIt->second != FrmControlType::Unknown) {
                    std::string memFc = Symbol::toLower(ma.memberName);
                    for (const char* n : kNumericFc) {
                        if (memFc == n) return Vb6Type::Long;
                    }
                    // C29-1b: 文件系统三控件的四个字符串属性。不登记则推断成 Variant,
                    // `File1.FileName = File1.List(0)` 这类比较就走 vb6_VarCmpEq 而不是
                    // vb6_StrCmp —— 右边 (RTL 声明 void*) 装箱成 VT_UNKNOWN, 于是
                    // 同一条读数 x64 为真、x86 为假。VB6 里这四个属性是 String, 类型
                    // 就该在这里落地, 不在用例里绕。
                    if (fcIt->second == FrmControlType::DriveListBox
                        || fcIt->second == FrmControlType::DirListBox
                        || fcIt->second == FrmControlType::FileListBox) {
                        static const char* const kStringFc3[] = {
                            "drive", "path", "pattern", "filename", "list",
                        };
                        for (const char* n : kStringFc3) {
                            if (memFc == n) return Vb6Type::String;
                        }
                    }
                    // D6 / C29-9: CommonDialog 的成员面。不登记 ⇒ 字符串属性被判成
                    // Variant ⇒ 比较/拼接走错箱（C29-1b 那条"x64 真、x86 假"的同族坑），
                    // 而 `CancelError` 这类布尔判成 Variant 还会让 `If CD1.CancelError`
                    // 走 VarCmp 而不是直接真值判断。
                    if (fcIt->second == FrmControlType::CommonDialog) {
                        static const char* const kStrFcCd[] = {
                            "filter", "filename", "filetitle", "dialogtitle",
                            "initdir", "defaultext", "fontname",
                        };
                        static const char* const kNumFcCd[] = {
                            "flags", "cancelerror", "color", "min", "max", "copies", "fontsize",
                        };
                        for (const char* n : kStrFcCd) {
                            if (memFc == n) return Vb6Type::String;
                        }
                        for (const char* n : kNumFcCd) {
                            if (memFc == n) return Vb6Type::Long;
                        }
                    }
                }
            }
            // P24-12: Err对象特殊处理
            if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objId = static_cast<IdentifierExpr&>(*ma.object);
                std::string objLower = objId.name;
                std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
                if (objLower == "err") {
                    std::string memLower = ma.memberName;
                    std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                    if (memLower == "number") return Vb6Type::Long;
                    if (memLower == "description" || memLower == "source") return Vb6Type::String;
                    if (memLower == "helppath" || memLower == "helpfile" || memLower == "helpcontext") return Vb6Type::String;
                    if (memLower == "lastdllerror") return Vb6Type::Long;
                }
            }
            // czUI fix: 宿主伪对象成员类型 — UserControl.ScaleWidth/Height 等
            // 是 RTL int32_t 全局 (vb6rtl_userctl.h)。此前推断为 Variant,
            // 比较时被取地址当 vb6_VARIANT* 读垃圾值 (MouseUp 里
            // X < ScaleWidth 恒假 → RaiseEvent Click 永不触发 → Connect 无响应)。
            if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objIdCz = static_cast<IdentifierExpr&>(*ma.object);
                std::string objLowerCz = objIdCz.name;
                std::transform(objLowerCz.begin(), objLowerCz.end(), objLowerCz.begin(), ::tolower);
                if (objLowerCz == "usercontrol" || objLowerCz == "propertypage") {
                    std::string memLowerCz = ma.memberName;
                    std::transform(memLowerCz.begin(), memLowerCz.end(), memLowerCz.begin(), ::tolower);
                    if (memLowerCz == "scalewidth" || memLowerCz == "scaleheight"
                        || memLowerCz == "left" || memLowerCz == "top"
                        || memLowerCz == "width" || memLowerCz == "height"
                        || memLowerCz == "enabled") {
                        return Vb6Type::Long;
                    }
                }
            }
            // Fix 081i: UDT字段访问 — 先查找UDT成员类型，避免lookupModule
            // 匹配到内置函数(如Left→String)导致类型推断错误
            {
                std::string udtCType = inferUdtTypeOfExpr(*ma.object);
                if (!udtCType.empty()) {
                    const std::string prefix = "vb6_type_";
                    if (udtCType.size() > prefix.size()
                        && udtCType.compare(0, prefix.size(), prefix) == 0) {
                        std::string udtName = udtCType.substr(prefix.size());
                        Symbol* udtSym = symTab_.lookupModule(udtName);
                        if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                            std::string memLower = ma.memberName;
                            std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                            for (auto& mi : udtSym->udtMembers) {
                                std::string miLower = mi.name;
                                std::transform(miLower.begin(), miLower.end(), miLower.begin(), ::tolower);
                                if (miLower == memLower) {
                                    return mi.type;  // 找到UDT字段，返回其Vb6Type
                                }
                            }
                            // Fix 084o-2: UDT 已确认但字段未找到 → 字段类型未知, 返回 Unknown.
                            // 不要回退 lookupModule(memberName): 成员名是字段名而非模块符号,
                            // 可能误匹配模块级同名符号 (如 SourceFile 变量) 导致类型误判为 String.
                            // Fix 090ab: 也不应返回 Variant — 否则 UDT 字段 (如 COMSTAT.fBitFields,
                            // 跨模块 Type 符号缺失) 被 isDefinitelyVariantExpr 误判为 Variant,
                            // 实参生成时套 vb6_VariantToLong(int32 字段) → C2440.
                            return Vb6Type::Unknown;
                        }
                    }
                    // Fix 090ab: UDT C 类型已知 (knownUdtVars_) 但符号不可解析 (跨模块
                    // Public Type 在符号表注册表不可达) → 字段类型未知, 保守返回 Unknown,
                    // 禁止回退到模块级同名符号或默认 Variant (否则 COMSTAT.fBitFields 等
                    // 会被误判为 Variant 而错误套用 vb6_VariantToLong → C2440).
                    return Vb6Type::Unknown;
                }
            }
            // 类实例成员函数返回类型 (tB 泛型类特化高发的跨类同名碰撞):
            // 按 object 推断出的宿主类查 memberReturnTypes, 内置标量名直接映射.
            // 下面的裸名 lookupModule(memberName) 在多个类同名成员时只命中
            // storageKey 去重胜出的第一个模块版本 (如 box_g1_long 的 GetVal→Long),
            // String 版被误判 → Debug.Print 把 BSTR 指针截断成垃圾数 (Fix 176 同源).
            if (ma.object && symTab_.moduleScope()) {
                std::string clsName = inferClassTypeOfExpr(*ma.object);
                if (!clsName.empty()) {
                    const std::string clsLower = Symbol::toLower(clsName);
                    const std::string memLowerCls = Symbol::toLower(ma.memberName);
                    for (const auto& [k, csym] : symTab_.moduleScope()->symbols()) {
                        if (csym->kind != SymbolKind::Class) continue;
                        bool hit = csym->isExternal
                            ? (Symbol::toLower(csym->sourceModule) == clsLower)
                            : (Symbol::toLower(csym->name) == clsLower);
                        if (!hit) continue;
                        auto it = csym->memberReturnTypes.find(memLowerCls);
                        if (it != csym->memberReturnTypes.end()) {
                            std::string rn = Symbol::toLower(it->second);
                            if (rn == "string")  return Vb6Type::String;
                            if (rn == "long")    return Vb6Type::Long;
                            if (rn == "integer") return Vb6Type::Integer;
                            if (rn == "boolean") return Vb6Type::Boolean;
                            if (rn == "single")  return Vb6Type::Single;
                            if (rn == "double")  return Vb6Type::Double;
                            if (rn == "date")    return Vb6Type::Date;
                            if (rn == "byte")    return Vb6Type::Byte;
                            if (rn == "currency")return Vb6Type::Currency;
                            if (rn == "variant") return Vb6Type::Variant;
                            if (rn == "object")  return Vb6Type::Object;
                        }
                        break;
                    }
                }
            }
            // P20-42: 对象是窗体上的已知控件时, 属性类型先问控件属性表。
            // **不能**直接掉到下面的 lookupModule(memberName): 那是按成员**裸名**
            // 查模块级符号, 凡是与模块级/内置符号同名的控件属性都会被顶掉。
            // 实测 `SSTab1.Tab` 撞上返回 BSTR 的内置函数 `Tab` → 判成 String →
            // Debug.Print 拼接不套 vb6_CStr(vb6_VariantFromValue(...)), 而 RTL 的
            // vb6_SSTab_GetTab 返回 int32_t → 整数当 BSTR 指针解引用 → 0xC0000005。
            if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objIdCtl = static_cast<IdentifierExpr&>(*ma.object);
                std::string objLowerCtl = Symbol::toLower(objIdCtl.name);
                auto itCtl = knownFormControls_.find(objLowerCtl);
                if (itCtl != knownFormControls_.end()) {
                    Vb6Type pt = controlPropType(itCtl->second, ma.memberName);
                    if (pt != Vb6Type::Unknown) return pt;
                }
            }
            // 查找成员函数/属性的返回类型
            auto* memSym = symTab_.lookupModule(ma.memberName);
            if (memSym) return memSym->type;
            break;
        }
        // Fix 084o-2: With 块内 UDT 字段访问 (.SourceFile 等) 的类型推断.
        // WithMemberExpr 由 With 语句展开 (_vb6_with_N->field), 必须按 UDT 字段查
        // udtMembers, 且不能回退 lookupModule(memberName) — 与 MemberAccessExpr 同理.
        case ASTNodeKind::WithMemberExpr: {
            auto& wm = static_cast<WithMemberExpr&>(expr);
            if (withObjectInfoStack_.empty() || withObjectVars_.empty()) return Vb6Type::Variant;
            const auto& winfo = withObjectInfoStack_.back();
            if (winfo.kind != WithObjKind::Unknown) {
                // 非 UDT With (类实例/COM 对象): memberName 是属性/方法 → 查成员符号
                auto* memSym2 = symTab_.lookupModule(wm.memberName);
                if (memSym2) return memSym2->type;
                return Vb6Type::Variant;
            }
            std::string tempLower = Symbol::toLower(withObjectVars_.back());
            auto it = knownUdtVars_.find(tempLower);
            if (it == knownUdtVars_.end()) return Vb6Type::Variant;
            const std::string prefix = "vb6_type_";
            const std::string& udtCType = it->second;
            if (udtCType.size() <= prefix.size()
                || udtCType.compare(0, prefix.size(), prefix) != 0) return Vb6Type::Variant;
            std::string udtName = udtCType.substr(prefix.size());
            Symbol* udtSym = symTab_.lookupModule(udtName);
            if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                std::string memLower = Symbol::toLower(wm.memberName);
                for (auto& mi : udtSym->udtMembers) {
                    if (Symbol::toLower(mi.name) == memLower) return mi.type;
                }
            }
            return Vb6Type::Variant;
        }
        default:
            break;
    }
    return Vb6Type::Variant;
}


// Fix 029: 严格 Variant 推断. 仅当表达式明确为 Variant 时返回 true.
// 与 inferExprType 的差异: 内置函数 (sym==null) 与 UDT 字段访问 (lookupModule 失败) 等
// 通过 fallback 返回 Variant 的情形, 此处视为非 Variant, 避免对 int/BSTR 等实参误包装.
bool CCodeGen::isDefinitelyVariantExpr(Expr& expr, bool* isArrOut) const {
    if (isArrOut) *isArrOut = false;
    uint16_t variantArrRaw = static_cast<uint16_t>(Vb6Type::Variant)
                           | static_cast<uint16_t>(Vb6Type::Array);

    // 多态内置函数 denylist: symTab 注册为 Variant, 但 codegen 按上下文
    // 发出类型化版本 (vb6_IIfBSTR/Long/Double, Choose 嵌套三元, Switch 嵌套三元),
    // 实际 C 返回类型不是 vb6_VARIANT. 视为非 Variant 以避免错误包装.
    auto isPolymorphicBuiltin = [](const std::string& name) -> bool {
        std::string lower = name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        return lower == "iif" || lower == "choose" || lower == "switch";
    };

    auto checkSym = [&](Symbol* sym) -> bool {
        if (!sym) return false;
        if (sym->type == Vb6Type::Variant) return true;
        if (static_cast<uint16_t>(sym->type) == variantArrRaw) {
            if (isArrOut) *isArrOut = true;
            return true;
        }
        return false;
    };

    switch (expr.kind) {
        case ASTNodeKind::IdentifierExpr: {
            auto& id = static_cast<IdentifierExpr&>(expr);
            std::string lower = id.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            // 已知数组变量: C 类型已是 vb6_SafeArray1D* (元素为具体类型或 VARIANT),
            // 不是 vb6_VARIANT. 防止 symTab_ lookupModule 回退命中其他模块同名
            // Variant 符号, 导致 UBound(arr) 被错误包装成 vb6_VariantToSafeArray1D(arr)
            // (C2440: 无法从 vb6_SafeArray1D* 转换为 vb6_VARIANT, 如 cDataBase WhereIn).
            if (knownArrays_.count(lower)) {
                return false;
            }
            // 已知 Variant 局部变量集合
            if (knownVariantVars_.count(lower)) {
                // 无法区分 Variant 与 Variant(), 视作普通 Variant
                return true;
            }
            // Fix 161f: 无 As Type 但整数可折叠的 Const — 符号表里 type=Variant,
            // 走下面的回退会被当成 Variant 表达式。C 侧是数值 (见 inferExprType
            // 同 Fix), 不是 Variant (extlist MListView SendMessage 实参 C2440)。
            if (moduleIntConstValues_.count(lower)) return false;
            // Fix 049b: 如果已知为非 Variant 具体类型 (BSTR/Long/Double),
            // 不应回退到符号表查找 (可能命中其他模块的同名 Variant 符号)
            // 账 #123: 补 Byte 那一档。缺它时局部 `Dim bt As Byte` 掉到下面的符号表回退 ⇒
            // 被判成 Variant ⇒ 比较发成 vb6_VarCmpLongEq(&bt, …) (拿 1 字节对象的地址当
            // vb6_VARIANT* 传) ⇒ 实测 `(bt = 65)` 返回 False。
            if (knownBstrVars_.count(lower) || knownLongVars_.count(lower)
                || knownDoubleVars_.count(lower) || knownSingleVars_.count(lower)
                || knownByteVars_.count(lower)) {
                return false;
            }
            // 符号表查询
            auto* sym = symTab_.lookup(id.name);
            if (!sym) sym = symTab_.lookupModule(id.name);
            return checkSym(sym);
        }
        case ASTNodeKind::IndexOrCallExpr: {
            auto& call = static_cast<IndexOrCallExpr&>(expr);
            if (call.callee && call.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& cid = static_cast<IdentifierExpr&>(*call.callee);
                // 多态内置函数: 实际返回类型与 symTab 注册不同, 不视为 Variant
                if (isPolymorphicBuiltin(cid.name)) return false;
                auto* sym = symTab_.lookup(cid.name);
                if (!sym) sym = symTab_.lookupModule(cid.name);
                // 关键差异: sym==null (内置函数) 视为非 Variant
                return checkSym(sym);
            }
            // 类方法调用 a.Method(): 递归推断 callee 类型
            if (call.callee && call.callee->kind == ASTNodeKind::MemberAccessExpr) {
                bool arr = false;
                bool v = isDefinitelyVariantExpr(*call.callee, &arr);
                if (v && isArrOut) *isArrOut = arr;
                return v;
            }
            return false;
        }
        case ASTNodeKind::MemberAccessExpr: {
            auto& ma = static_cast<MemberAccessExpr&>(expr);
            // Err 对象特殊处理 (与 inferExprType 一致)
            if (ma.object && ma.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objId = static_cast<IdentifierExpr&>(*ma.object);
                std::string objLower = objId.name;
                std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
                if (objLower == "err") {
                    // Err.* 的具体类型见 inferExprType, 均为具体类型 (Long/String), 非 Variant
                    return false;
                }
            }
            // Fix 090l: UDT 字段访问 — 字段声明 As Variant (如 cZipArchive
            // ZipVfsType.BufferArray / SourceFileInfo) 是明确 Variant 表达式.
            // 此前 memSym==null (UDT 字段) 被一律视为非 Variant, 导致
            // UBound(vb6_ret_X.BufferArray) 等不包装 vb6_VariantToSafeArray1D →
            // C2440 (无法从 vb6_VARIANT 转 vb6_SafeArray1D*).
            if (!inferUdtTypeOfExpr(*ma.object).empty()) {
                return inferExprType(expr) == Vb6Type::Variant;
            }
            // 符号表查找成员 (Property/Function)
            auto* memSym = symTab_.lookupModule(ma.memberName);
            // 关键差异: memSym==null (UDT 字段访问或外部类成员) 视为非 Variant
            return checkSym(memSym);
        }
        default:
            // 其他表达式 (BinaryExpr/UnaryExpr/LiteralExpr 等) 不会明确返回 Variant,
            // 除非其子表达式明确为 Variant. 此处不递归, 保持严格性.
            return false;
    }
}


// ============================================================
// Fix 038b-1: 基于 C 表达式字符串的 Variant 检测
// ============================================================
// 补充 isDefinitelyVariantExpr 的 AST 级检测. 当 codegen 生成的 C 表达式
// 包含已知返回 vb6_VARIANT 的函数调用时, 判定为 Variant.
// 仅检查顶层表达式 (去除前导括号/空白后), 避免对子表达式误判.

bool CCodeGen::cExprIsVariant(const std::string& cExpr) const {
    // 去除前导空白和括号
    size_t start = 0;
    while (start < cExpr.size()) {
        char c = cExpr[start];
        if (c == '(' || c == ' ' || c == '\t' || c == '\n' || c == '\r') {
            start++;
        } else {
            break;
        }
    }
    if (start >= cExpr.size()) return false;

    // 已知返回 vb6_VARIANT 的函数前缀
    static const std::vector<std::string> variantPrefixes = {
        "vb6_VariantArrayGet(",
        "vb6_VariantFromComResult(",
        "vb6_VariantFromStackVARIANT(",
        "vb6_VariantFromValue(",
        "vb6_VariantEmpty(",
        "vb6_VariantLong(",
        "vb6_VariantDouble(",
        "vb6_VariantString(",
        "vb6_VariantBool(",
        "vb6_VariantArray(",
        "vb6_VariantObject(",
        "vb6_VariantNull(",
        "vb6_VariantNothing(",
        "vb6_VariantFromI2(",
        "vb6_VariantFromI4(",
        "vb6_VariantFromR4(",
        "vb6_VariantFromR8(",
        "vb6_VariantFromBSTR(",
        "vb6_VariantFromBool(",
        "vb6_VariantFromDate(",
        "vb6_VariantFromUI1(",
        "vb6_VariantFromSafeArray(",
        "vb6_VariantFromSafeArray1D(",
        "vb6_LoadResData(",       // Fix 090bz: VBA LoadResData → vb6_VARIANT;
                                 //   Dim D() As Byte: D = LoadResData(...) 赋值
                                 //   需 VariantToSafeArray1D 提取 (cLang LoadData/LoadInfo C2440).
        // vbeclipse: LoadResPicture 与 LoadResData 同族, RTL 签名都是
        //   vb6_VARIANT vb6_LoadResPicture(int32_t, int32_t) (vb6rtl_runtime.h:66 /
        //   vb6rtl_com.c:526). 不登记会让
        //   `Function getResourceIcon(...) As IPictureDisp` 的
        //   `Set getResourceIcon = LoadResPicture(...)` 直接 `vb6_ret_X =
        //   vb6_LoadResPicture(...)` → C2440 (modResources.c 22/24/33). 登记后走
        //   Set 的 Fix 038b-6 分支包 vb6_VariantToObjectVal 提取对象指针.
        "vb6_LoadResPicture(",
        "vb6_DispCallByVtbl(",  // Fix 068: DispCallByVtbl returns Variant
        // Fix 110w: VB6 CallByName 返回 vb6_VARIANT (见 vb6rtl_class_com.h) —
        // 参与算术/关系运算或需 BSTR 时必须按 Variant 处理, 否则 C2088
        // ("*" 对于 struct 非法; Charts 2020 ClsResizer.cls:142/148
        //  CallByName(oCtrl, ..., VbGet) * 100).
        "vb6_CallByName(",
    };
    for (const auto& prefix : variantPrefixes) {
        if (cExpr.compare(start, prefix.size(), prefix) == 0) return true;
    }

    // Fix 040b: VB6_SA_AT(vb6_VARIANT, arr, idx) expands to an array element
    // of type vb6_VARIANT — also a Variant expression.
    if (cExpr.compare(start, 21, "VB6_SA_AT(vb6_VARIANT") == 0) return true;

    // Fix 045: 检查项目函数是否返回 Variant — 通过 driver.cpp 预扫描构建的
    // C 函数名集合. 提取 C 表达式中的函数名 (从 start 到第一个 '(') 并查集合.
    if (variantReturnFuncs_) {
        size_t parenPos = cExpr.find('(', start);
        if (parenPos != std::string::npos) {
            std::string funcName = cExpr.substr(start, parenPos - start);
            if (variantReturnFuncs_->count(funcName)) return true;
        }
    }

    // 检查 (&(vb6_VARIANT){...}) 复合字面量 — 也是 VARIANT 类型
    // 但这种形式通常作为 ByRef 参数传递, 不需要再转换, 故不检测.

    return false;
}

// ============================================================
// <vbeclipse>: C 表达式是否**已经是 SafeArray1D\* 载体**
// ============================================================
// Split/Filter/Array 这些内置函数在 VB6 侧的类型是 Variant, 但 codegen 发的是
// 直接返回 vb6_SafeArray1D\* 的 RTL 调用。数组槽实参 (Join 首参 / UBound/LBound
// 首参) 的按 Variant 提取 (vb6_VariantToSafeArray1D) 若套在它们外面就是
// C2440 (vb6_SafeArray1D\* → vb6_VARIANT) —— 实测这三条形全中:
//   Join(Split(s, ","), "|") / UBound(Split(s, ",")) / Join(Filter(a, "x"), "|")
// 与 Fix 092g 的 _arr_N 特例同源、同解法, 区别只是这里包的是内置函数调用。
bool CCodeGen::cExprIsSafeArrayCarrier(const std::string& cExpr) const {
    size_t start = 0;
    while (start < cExpr.size()) {
        char c = cExpr[start];
        if (c == '(' || c == ' ' || c == '\t' || c == '\n' || c == '\r' || c == '*') {
            start++;
        } else {
            break;
        }
    }
    // 匿名数组临时量 (`Array(...)` → _arr_N, Fix 092g) 本身就是载体。
    if (cExpr.compare(start, 5, "_arr_") == 0) return true;

    static const std::vector<std::string> carrierPrefixes = {
        "vb6_Split(",                  // vb6_SafeArray1D* vb6_Split(...)
        "vb6_Filter(",
        "vb6_ArrayCreate(",
        "vb6_ArrayAssign1D(",          // 整体数组赋值的深拷贝 (Fix 170)
        "vb6_SafeArrayCreate1D(",
        "vb6_SafeArrayReDim1D(",
        "vb6_SafeArrayReDimPreserve1D(",
        // Fix <vbeclipse> rev3: 带 elemType 的新入口。注意**不能**指望上一行前缀
        // 命中 —— `vb6_SafeArrayReDimPreserve1D(` 与 `...1D_T(` 在第 25 个字符处分叉。
        "vb6_SafeArrayReDimPreserve1D_T(",
        "vb6_VariantToSafeArray1D(",   // 已提取过, 再包一层就是双重解引用
        "vb6_VariantToByteArray(",
        "vb6_StringToByteArray(",
        "vb6_StrConvToByteArray(",
        "vb6_ComCallByteArray(",
    };
    for (const auto& prefix : carrierPrefixes) {
        if (cExpr.compare(start, prefix.size(), prefix) == 0) return true;
    }
    return false;
}


// ============================================================
// Fix 038b-5: 运行时函数参数 C 类型查找表
// ============================================================
// 当 calleeParams 为空 (运行时/内置函数) 时, 通过函数名和参数索引查找
// 期望的 C 类型. 返回空字符串表示未知.

std::string CCodeGen::getRuntimeParamCType(const std::string& funcName, size_t paramIdx) {
    static const std::unordered_map<std::string, std::vector<std::string>> table = {
        // Array 创建/设置
        {"vb6_ArraySetLong",    {"vb6_SafeArray1D*", "int32_t", "int32_t"}},
        {"vb6_ArraySetBSTR",    {"vb6_SafeArray1D*", "int32_t", "BSTR"}},
        {"vb6_ArraySetDouble",  {"vb6_SafeArray1D*", "int32_t", "double"}},
        {"vb6_ArraySetVariant", {"vb6_SafeArray1D*", "int32_t", "vb6_VARIANT"}},
        {"vb6_ArrayGetLong",    {"vb6_SafeArray1D*", "int32_t"}},
        {"vb6_ArrayGetBSTR",    {"vb6_SafeArray1D*", "int32_t"}},
        {"vb6_ArrayGetDouble",  {"vb6_SafeArray1D*", "int32_t"}},
        // Fix <vbeclipse>: Unload / LoadPicture — 形参是对象指针 (HWND / IDispatch);
        // 实参是 COM 后期绑定读数 (VARIANT) 时按此表解包 (vb6_UnloadForm(l_View.View))。
        {"vb6_UnloadForm",      {"void*"}},
        {"vb6_LoadPictureEx",   {"BSTR"}},
        {"vb6_ArrayGetVariant", {"vb6_SafeArray1D*", "int32_t"}},
        // BSTR 操作
        {"vb6_BSTR_Assign",     {"BSTR*", "BSTR"}},
        {"vb6_BSTR_Concat",     {"BSTR", "BSTR"}},
        {"vb6_BSTR_ConcatFree", {"BSTR", "BSTR"}},
        {"vb6_BSTR_FromStr",    {"const wchar_t*"}},
        {"vb6_BSTR_Empty",      {}},
        {"vb6_BSTR_ToANSI",     {"BSTR"}},
        {"vb6_BSTR_Free",       {"BSTR*"}},
        {"vb6_BSTR_Clone",      {"BSTR"}},
        // 字符串比较/操作
        {"vb6_StrCmp",          {"BSTR", "BSTR"}},
        {"vb6_StrComp",         {"BSTR", "BSTR", "int32_t"}},
        {"vb6_Val",             {"BSTR"}},
        {"vb6_Trim",            {"BSTR"}},
        {"vb6_LTrim",           {"BSTR"}},
        {"vb6_RTrim",           {"BSTR"}},
        {"vb6_Left",            {"BSTR", "int32_t"}},
        {"vb6_Right",           {"BSTR", "int32_t"}},
        {"vb6_Mid",             {"BSTR", "int32_t", "int32_t"}},
        {"vb6_Len",             {"BSTR"}},
        {"vb6_LenB",            {"BSTR"}},
        {"vb6_InStr",           {"BSTR", "BSTR"}},
        {"vb6_Replace",         {"BSTR", "BSTR", "BSTR"}},
        {"vb6_Split",           {"BSTR", "BSTR"}},
        {"vb6_Join",            {"vb6_SafeArray1D*", "BSTR"}},
        {"vb6_UCase",           {"BSTR"}},
        {"vb6_LCase",           {"BSTR"}},
        {"vb6_Space",           {"int32_t"}},
        // Fix 091e: StrConv(BSTR, int32_t, int32_t) — 实参为 Variant 时需
        // vb6_VariantToString (cAesCBC.c 25 StrConv(LoadResData(...), 64, 0) C2440)
        {"vb6_StrConv",         {"BSTR", "int32_t", "int32_t"}},
        {"vb6_String",          {"int32_t", "int32_t"}},
        {"vb6_Chr",             {"int32_t"}},
        {"vb6_Asc",             {"BSTR"}},
        {"vb6_Hex",             {"int32_t"}},
        {"vb6_Oct",             {"int32_t"}},
        // 类型转换
        {"vb6_CStr",            {"vb6_VARIANT"}},
        {"vb6_CLng",            {"double"}},
        {"vb6_CInt",            {"double"}},
        {"vb6_CDbl",            {"double"}},
        {"vb6_CSng",            {"double"}},
        {"vb6_CBool",           {"vb6_VARIANT"}},
        {"vb6_CByte",           {"vb6_VARIANT"}},
        {"vb6_CDate",           {"vb6_VARIANT"}},
        {"vb6_CCur",            {"vb6_VARIANT"}},
        // 数组操作
        {"vb6_UBound",          {"vb6_SafeArray1D*", "int32_t"}},
        {"vb6_LBound",          {"vb6_SafeArray1D*", "int32_t"}},
        {"vb6_ArrayCreate",     {"int32_t"}},
        // 消息框
        {"vb6_MsgBox",          {"BSTR"}},
        {"vb6_MsgBox1",         {"BSTR"}},
        // 错误处理
        {"vb6_ErrRaise",        {"int32_t", "BSTR", "BSTR"}},
        {"vb6_ErrNumber",       {}},
        {"vb6_ErrDescription",  {}},
        {"vb6_ErrSource",       {}},
        {"vb6_ErrClear",        {}},
        // IsMissing
        // Fix 091k: RTL 实际签名 int32_t vb6_IsMissing(SAFEARRAY* psa) (ParamArray 专用,
        // 判断是否未传实参). 原表项 vb6_VARIANT* 与 RTL 不符, 导致 IsMissing(<Variant>)
        // 时表项无效, 裸传 Variant 值 → C2440.
        {"vb6_IsMissing",       {"SAFEARRAY*"}},
        // Variant 提取
        {"vb6_VariantToLong",      {"vb6_VARIANT"}},
        {"vb6_VariantToDouble",    {"vb6_VARIANT"}},
        {"vb6_VariantToString",    {"vb6_VARIANT"}},
        {"vb6_VariantToBool",      {"vb6_VARIANT"}},
        {"vb6_VariantToSafeArray1D", {"vb6_VARIANT"}},
        {"vb6_VariantToObjectVal", {"vb6_VARIANT"}},
        // Fix 046: IIf family + Variant-aware conversion functions
        {"vb6_IIfBSTR",         {"int32_t", "BSTR", "BSTR"}},
        {"vb6_IIfLong",         {"int32_t", "int32_t", "int32_t"}},
        {"vb6_IIfDouble",       {"int32_t", "double", "double"}},
        {"vb6_IIfVariant",      {"int32_t", "vb6_VARIANT", "vb6_VARIANT"}},
        {"vb6_CLngV",           {"vb6_VARIANT"}},
        {"vb6_CIntV",           {"vb6_VARIANT"}},
        {"vb6_IntDiv",          {"int32_t", "int32_t"}},
        // Debug
        {"vb6_DebugPrint",      {"BSTR"}},
        {"vb6_DebugWriteLong",  {"int32_t"}},
        // 对象操作
        {"vb6_StrPtr",          {"BSTR"}},   // Fix 084o-7: StrPtr(Variant) → vb6_VariantToString 先行
        // Fix 155: ObjPtr 实参签名是 `vb6_ObjPtr(void* obj)` (vb6rtl_builtin.h:288),
        // 此前登记 "uintptr_t" (那是**返回**类型, 非形参). 运行时提取分支按形参类型
        // 匹配 ("void*" → vb6_VariantToObjectVal), "uintptr_t" 无对应分支 → ObjPtr
        // 收到 vb6_VARIANT (如 ObjPtr(ParentControls.Item(0)) 的 COM 结果) 时不做
        // 提取, 把结构体裸传给 void* 形参 → C2172 "实参不是指针". 改为真实形参类型.
        {"vb6_ObjPtr",          {"void*"}},
        {"vb6_ReleaseObject",   {"void**"}},
        {"vb6_NewObject",       {"const wchar_t*"}},
        {"vb6_CallByName",      {"void*", "BSTR", "int32_t"}},
        // Fix 113: UserControl 宿主内建方法 (vb6rtl_userctl.h) — 裸名书写,
        // cgen 映射为 vb6_UserControl_<Member> (cgen_expr_ident_builtin.inc).
        // 此前不在运行时参数表, 实参为 vb6_ComCall(...) 等 COM Variant 时缺 BSTR
        // 提取 → 传入 GetTextExtentPoint32W/字符串 API 崩溃. 注册 BSTR 形参.
        {"vb6_UserControl_TextWidth",      {"BSTR"}},
        {"vb6_UserControl_TextHeight",     {"BSTR"}},
        {"vb6_UserControl_AsyncRead",      {"BSTR", "int32_t", "BSTR", "int32_t"}},
        {"vb6_UserControl_PropertyChanged", {"BSTR"}},
        {"vb6_UserControl_CancelAsyncRead", {"BSTR"}},
    };
    auto it = table.find(funcName);
    if (it != table.end() && paramIdx < it->second.size()) {
        return it->second[paramIdx];
    }
    return "";
}

// ============================================================
// Fix 092w: Byte 数组赋值右侧改写
// ============================================================
// VB6/twinbasic 允许 Dim ba() As Byte = "..." 或 = StrConv(s, 64/128), 语义为
// 生成含字节内容的动态数组. RTL 的 vb6_StrConv 返回 BSTR, 不能直接赋给
// vb6_SafeArray1D*. 此处在编译期改用返回字节数组的 RTL helper.
std::string CCodeGen::rewriteByteArrayValue(const std::string& value) const {
    if (value.empty() || value == "NULL") return value;

    // StrConv(...) 用于构造字节数组 → 改用 vb6_StrConvToByteArray(...)
    if (value.compare(0, 12, "vb6_StrConv(") == 0) {
        return "vb6_StrConvToByteArray" + value.substr(11);
    }
    // 已经是字节数组/数组表达式 → 原样返回
    if (value.find("vb6_StrConvToByteArray(") != std::string::npos ||
        value.find("vb6_StringToByteArray(") != std::string::npos ||
        value.compare(0, 14, "vb6_SafeArray") == 0 ||
        value.compare(0, 22, "vb6_VariantToByteArray") == 0 ||
        value.compare(0, 25, "vb6_VariantToSafeArray1D") == 0) {
        return value;
    }
    // 字符串/BSTR 表达式 → 复制为字节数组 (原始 UTF-16LE 字节)
    // Fix 140: 补充 Ambient.DisplayName (BSTR) — LabelPlus.ctl UserControl_InitProperties
    // 里 `m_Caption = Ambient.DisplayName` 生成 `me->m_Caption = vb6_Ambient_DisplayName`,
    // 是 BSTR 赋给 Byte() 字段, 需同样改写成字节数组.
    bool isBstrExpr = value.compare(0, 9, "vb6_BSTR_") == 0 ||
                      value.compare(0, 20, "vb6_VariantToString(") == 0 ||
                      value == "vb6_Ambient_DisplayName";
    if (isBstrExpr) {
        return "vb6_StringToByteArray(" + value + ")";
    }
    // Fix 140: ByRef 解引用形态 `(*Param)` — 形参 C 类型为 BSTR* (ByRef String),
    // `(*Param)` 即真实 BSTR. 若内层名字是已知 BSTR 变量则按"字符串→字节数组"
    // 改写. 场景: LabelPlus.ctl `Property Let Caption(ByRef New_Caption As String)`
    // 内 `m_Caption = New_Caption` 生成 `me->m_Caption = (*New_Caption)`, 若不改写
    // 会把 BSTR 当 SafeArray1D* 赋给 Byte() 字段 → caption 读不到, 卡片空白.
    if (value.size() > 4 && value[0] == '(' && value[1] == '*'
        && value[value.size() - 1] == ')') {
        std::string inner = value.substr(2, value.size() - 3);
        if (inner.compare(0, 4, "me->") == 0) inner = inner.substr(4);
        if (inner.find_first_of("( .->") == std::string::npos) {
            std::string lower = inner;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (knownBstrVars_.count(lower)) {
                return "vb6_StringToByteArray(" + value + ")";
            }
        }
        return value;
    }
    // 裸变量名: 若为已知 BSTR 变量则包装
    std::string name = value;
    if (name.compare(0, 4, "me->") == 0) name = name.substr(4);
    if (name.find_first_of("( .") == std::string::npos) {
        std::string lower = name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownBstrVars_.count(lower)) {
            return "vb6_StringToByteArray(" + value + ")";
        }
    }
    return value;
}

// Fix 170: VB6 **整体数组引用** `A()` —— IndexOrCallExpr 无实参, callee 是已知数组变量或
// 模块级数组。判定口径与 cgen_expr_call_prelude.inc 的 isArrayAccess 一致 (先 knownArrays_,
// 再查模块符号表 sym->isArray), 因此 `GetTickCount()` / `Command()` 这类零参调用不会误判。
bool CCodeGen::isWholeArrayRef(const Expr* e) const {
    if (!e || e->kind != ASTNodeKind::IndexOrCallExpr) return false;
    auto& n = static_cast<const IndexOrCallExpr&>(*e);
    if (!n.positional.empty() || !n.named.empty()) return false;
    if (!n.callee) return false;
    if (n.callee->kind == ASTNodeKind::IdentifierExpr) {
        auto& id = static_cast<const IdentifierExpr&>(*n.callee);
        if (knownArrays_.count(Symbol::toLower(id.name))) return true;
        Symbol* sym = symTab_.lookupModule(id.name);
        return sym && sym->kind == SymbolKind::Variable && sym->isArray;
    }
    // Fix 178: `.Cols() = VBFlexGridDefaultCols.Cols()` —— UDT/With 的动态数组成员
    // 用空括号整体赋值。原先只认裸数组变量, 这类目标退化成载体指针直赋 → 别名。
    return isDynamicArrayMemberCallee(n.callee.get(), nullptr);
}

// Fix 178: callee 是「UDT 的动态数组成员」(`.Cols` / `Default.Cols`) 时返回 true,
// 并按需带出元素 UDT 的 C 类型 (元素是标量类型时保持原值)。
bool CCodeGen::isDynamicArrayMemberCallee(const Expr* callee, std::string* elemUdt) const {
    if (!callee) return false;
    std::string parentUdt, memName;
    if (callee->kind == ASTNodeKind::MemberAccessExpr) {
        auto& ma = static_cast<const MemberAccessExpr&>(*callee);
        if (!ma.object) return false;
        parentUdt = inferUdtTypeOfExpr(*ma.object);
        memName = ma.memberName;
    } else if (callee->kind == ASTNodeKind::WithMemberExpr) {
        if (withObjectInfoStack_.empty() || withObjectVars_.empty()) return false;
        if (withObjectInfoStack_.back().kind != WithObjKind::Unknown) return false;
        auto it = knownUdtVars_.find(Symbol::toLower(withObjectVars_.back()));
        if (it == knownUdtVars_.end()) return false;
        parentUdt = it->second;
        memName = static_cast<const WithMemberExpr&>(*callee).memberName;
    } else {
        return false;
    }
    const std::string prefix = "vb6_type_";
    if (parentUdt.size() <= prefix.size()
        || parentUdt.compare(0, prefix.size(), prefix) != 0) return false;
    Symbol* sym = symTab_.lookupModule(parentUdt.substr(prefix.size()));
    if (!sym || sym->kind != SymbolKind::UserDefinedType) return false;
    std::string memLower = Symbol::toLower(memName);
    for (const auto& mi : sym->udtMembers) {
        if (Symbol::toLower(mi.name) != memLower) continue;
        if (!mi.isArrayDynamic) return false;
        if (elemUdt && mi.type == Vb6Type::UserDefinedType && !mi.typeRefName.empty())
            *elemUdt = "vb6_type_" + cIdent(mi.typeRefName);
        return true;
    }
    return false;
}

// Fix 170: 整体数组赋值 `A() = expr` 的右侧收口。
// 只放行**必然新建载体**的 helper (vb6_StringToByteArray / StrConvToByteArray /
// Array(...) 物化 / Split / Filter / COM ByteArray 解封)；其余形态一律走
// vb6_ArrayAssign1D 深拷贝：
//   · 裸数组变量 `A() = B()` —— 直接赋是两个名字别名同一载体;
//   · vb6_VariantToByteArray / vb6_VariantToSafeArray1D —— Variant 持数组时**返回
//     v.parray 本身** (见 vb6rtl_compat.c:187)，不拷贝就是与那个 Variant 共用载体。
std::string CCodeGen::wrapWholeArrayAssign(const std::string& target,
                                           const std::string& rhs,
                                           const Expr* targetNode) const {
    static const std::vector<std::string> freshCarrier = {
        "vb6_StringToByteArray(", "vb6_StrConvToByteArray(",
        "vb6_ArrayCreate(", "vb6_ArrayAssign1D(", "vb6_ComCallByteArray(",
        "vb6_VariantArray(", "vb6_Split(", "vb6_Filter(",
    };
    // 跳过空白与左括号/解引用 (`(*Arr)` 形态的 ByRef 参数取的是载体本身)
    size_t s = rhs.find_first_not_of(" \t\r\n(*");
    if (s != std::string::npos) {
        for (const auto& p : freshCarrier)
            if (rhs.compare(s, p.size(), p) == 0) return rhs;
    }
    // Fix 178: 元素是含所有权成员的 UDT 时, 载体克隆还要逐元素深拷贝, 否则两侧元素
    // 共享同一 BSTR / 子数组 (VBFlexGridCells.Rows(i).Cols 的别名就是这么来的)。
    std::string elemUdt;
    if (targetNode && targetNode->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& n = static_cast<const IndexOrCallExpr&>(*targetNode);
        if (n.callee) {
            if (n.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& id = static_cast<const IdentifierExpr&>(*n.callee);
                auto it = arrayUdtElemTypes_.find(Symbol::toLower(id.name));
                if (it != arrayUdtElemTypes_.end()) elemUdt = it->second;
            } else {
                isDynamicArrayMemberCallee(n.callee.get(), &elemUdt);
            }
        }
    }
    if (!elemUdt.empty() && elemUdt.compare(0, 9, "vb6_type_") == 0
        && udtHasOwnedMembers(elemUdt)) {
        requestUdtCopy(elemUdt, true);
        return "vb6_ArrayAssign1D_Cb(" + target + ", " + rhs + ", vb6_udtcpy_"
             + elemUdt.substr(9) + "_v)";
    }
    return "vb6_ArrayAssign1D(" + target + ", " + rhs + ")";
}

// ============================================================
// ai/009 §5.10 (P3) 溢出检查
// ============================================================
// 判"右值类型是否可能越出目标范围"。用**比特宽**比, 而不是逐类型列白名单:
// 目标宽度以外的整型、以及一切浮点 (Single/Double/Currency/Date), 都能越界;
// 目标宽度以内的 (含 Boolean —— 值域只有 -1/0), 永远不可能, 套检查纯属噪声。
// Unknown/Variant 一律按"可能越界"算 —— 宁可多查, 不可漏查, 因为漏查就是静默错编。
static int cgenIntBits(Vb6Type t) {
    switch (t) {
    case Vb6Type::Byte: case Vb6Type::Boolean: return 8;
    case Vb6Type::Integer: return 16;
    case Vb6Type::Long: case Vb6Type::ULong:
    case Vb6Type::Single: case Vb6Type::LongPtr: return 32;
    // Currency/Date 在 C 侧是 double, LongLong/LongPtr 是 64 位整数, 都可能越界
    case Vb6Type::Double: case Vb6Type::Currency: case Vb6Type::Date:
    case Vb6Type::LongLong: return 64;
    default: return 0;   // Unknown / Variant / String / Object ... 由调用点决定
    }
}

std::string CCodeGen::narrowCheckAssign(Expr* target, Expr* value,
                                        const std::string& cValue) const {
    if (!target || cValue.empty()) return cValue;
    // 目标只认裸标量标识符。成员/数组/属性写入各自另有类型解析链, 猜错会把
    // 合法赋值判成越界 (那比不检查更糟), 保持改动前的行为。
    if (target->kind != ASTNodeKind::IdentifierExpr) return cValue;

    // 这里**故意不复用** inferExprType 来定目标类型: 它把 Integer 答成 Long、
    // 把 Byte 答成 Variant/Unknown, 那是几十个消费点共同依赖的既有口径, 动它
    // 等于给 Debug.Print / Variant 装箱等一整条链换答案。这里只为溢出检查
    // 单独查一遍精确的窄整型, 顺带靠这个局部性把影响面关在收窄赋值里。
    Vb6Type tt;
    auto& id = static_cast<IdentifierExpr&>(*target);
    std::string lower = id.name;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    if (knownByteVars_.count(lower))        tt = Vb6Type::Byte;
    else if (knownIntVars_.count(lower))    tt = Vb6Type::Integer;
    else if (knownBoolVars_.count(lower))   return cValue;   // 值域只有 -1/0
    else if (knownLongVars_.count(lower))   tt = Vb6Type::Long;
    else                                    tt = inferExprType(*target);

    const char* fn = nullptr;
    int tgtBits = 0;
    switch (tt) {
    case Vb6Type::Byte:    fn = "vb6_ChkByte"; tgtBits = 8;  break;
    case Vb6Type::Integer: fn = "vb6_ChkInt";  tgtBits = 16; break;
    case Vb6Type::Long:    fn = "vb6_ChkLong"; tgtBits = 32; break;
    default: return cValue;
    }

    int srcBits = value ? cgenIntBits(inferExprType(*value)) : 0;
    if (srcBits != 0 && srcBits <= tgtBits) return cValue;   // 装得下, 不套

    return std::string(fn) + "(" + cValue + ")";
}

} // namespace vb6c3
