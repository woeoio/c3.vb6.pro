#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_localdecl.cpp: 局部声明与提升 (Option/Dim/Const/hoist) ---

void CCodeGen::visit(OptionStmt& node) {
    // P18-C: Option Compare Text/Binary
    if (node.optionKind == OptionKind::CompareText) {
        c_.emitLine("g_vb6_optionCompareText = 1; /* Option Compare Text */");
    } else if (node.optionKind == OptionKind::CompareBinary) {
        c_.emitLine("g_vb6_optionCompareText = 0; /* Option Compare Binary */");
    }
    // Option Explicit / Option Base 不影响C代码生成
}

void CCodeGen::visit(LocalDeclStmt& node) {
    // Fix 086: 已在过程序言处提升声明的节点, 原位置跳过
    if (hoistedLocalDeclSet_.count(&node)) return;
    emitLocalDeclCode(node);
}

void CCodeGen::emitLocalDeclCode(LocalDeclStmt& node) {
    if (!node.decl) return;

    switch (node.decl->kind) {
        case ASTNodeKind::MultiDecl: {
            // Fix 152d: 过程体内逗号声明 (`Static hFindFile As LongPtr,
            // AttributesCache As VbFileAttribute` / `Dim a, b As Long` 等) 由
            // parseVariableDeclList 返回 MultiDecl。模块级 MultiDecl 在 parseModule
            // 已展平, 但体级 LocalDeclStmt 仍携 MultiDecl → 原 switch 无此分支 →
            // "unhandled" → 变量从不声明 → 后续引用 C2065。逐子声明展开。
            auto& md = static_cast<MultiDecl&>(*node.decl);
            for (auto& child : md.declarations) {
                if (!child) continue;
                LocalDeclStmt sub(md.loc, std::move(child));
                emitLocalDeclCode(sub);
            }
            return;
        }
        case ASTNodeKind::VariableDecl: {
            auto& var = static_cast<VariableDecl&>(*node.decl);
            std::string cName = cIdent(var.name);

            // P8.1: 局部数组声明 (支持多维)
            if (!var.dimensions.empty()) {
                Vb6Type elemType = resolveArrayElemType(var.asType.get());
                std::string saElemType = mapSaElemType(elemType);
                // Bug4-Fix: UDT数组
                std::string udtCType = resolveArrayUdtElemCType(var.asType.get());
                bool isUdtArr = !udtCType.empty();
                int dimCount = (int)var.dimensions.size();

                if (dimCount == 1) {
                    // 一维数组: 保持原有1D代码
                    auto& dim = var.dimensions[0];
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
                    std::string initCode;
                    if (isUdtArr) {
                        initCode = "vb6_SafeArrayReDim1D_Udt((int32_t)sizeof(" + udtCType + "), " + lBound + ", " + uBound + ")";
                    } else {
                        initCode = "vb6_SafeArrayCreate1D(" + saElemType + ", " + lBound + ", " + uBound + ")";
                    }
                    c_.emitLine("vb6_SafeArray1D* " + cName + " = " + initCode + ";");
                } else {
                    // 多维数组: 使用ND运行时
                    std::string boundsVar = "_bounds_" + cName;
                    c_.emitLine("vb6_SafeArrayBound " + boundsVar + "[] = {");
                    c_.indent();
                    for (int d = 0; d < dimCount; d++) {
                        auto& dim = var.dimensions[d];
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
                    c_.emitLine("vb6_SafeArrayND* " + cName + " = vb6_SafeArrayCreateND(" + saElemType + ", " + std::to_string(dimCount) + ", " + boundsVar + ");");
                }

                // 注册到已知数组集合
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownArrays_.insert(lower);
                arrayElemTypes_[lower] = elemType;
                arrayDimCounts_[lower] = dimCount;
                if (elemType == Vb6Type::Byte) knownByteArrayVars_.insert(lower);
                // Fix 055: 注册UDT数组元素C类型
                {
                    std::string udtCType = resolveArrayUdtElemCType(var.asType.get());
                    if (!udtCType.empty()) arrayUdtElemTypes_[lower] = udtCType;
                }
                // Fix 192: 注册类数组元素类名 (静态数组 `Dim s(1) As ShapeAct`)
                {
                    std::string cls = resolveArrayClassElemType(var.asType.get());
                    if (!cls.empty()) arrayClassElemTypes_[lower] = cls;
                }
                knownLocalVars_.insert(lower);
                break;
            }

            // P8.1: 动态数组声明: Dim arr() As Long → 默认1D, ReDim时可能升级
            if (var.isDynamicArray) {
                Vb6Type elemType = resolveArrayElemType(var.asType.get());
                // Fix 084aa: #undef 防宏污染 — 模块常量被生成 #define 宏 (如 cStartUp 的
                // #define K (vb6_BSTR_FromStr(...))), 同名局部变量声明会被宏展开破坏.
                // 局部变量总是遮蔽模块常量, #undef 是安全且正确的.
                c_.emitLine("#undef " + cName);
                // Fix 092w: Dim arr() As Byte = <初始化表达式> — twinbasic 兼容.
                // 动态数组初值来自字符串/StrConv 时改用字节数组 helper 生成内容,
                // 否则保持 NULL 待 ReDim/赋值.
                std::string dynInit = "NULL";
                if (var.initializer && elemType == Vb6Type::Byte) {
                    emitExpr(*var.initializer);
                    dynInit = rewriteByteArrayValue(lastExpr_);
                }
                c_.emitLine("vb6_SafeArray1D* " + cName + " = " + dynInit + ";");

                // 注册到已知数组集合
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownArrays_.insert(lower);
                arrayElemTypes_[lower] = elemType;
                arrayDimCounts_[lower] = 1;  // 动态数组默认1D
                if (elemType == Vb6Type::Byte) knownByteArrayVars_.insert(lower);
                // Fix 055: 注册UDT数组元素C类型
                {
                    std::string udtCType = resolveArrayUdtElemCType(var.asType.get());
                    if (!udtCType.empty()) arrayUdtElemTypes_[lower] = udtCType;
                }
                // Fix 192: 同上, 动态数组 `Dim a() As ShapeAct`
                {
                    std::string cls = resolveArrayClassElemType(var.asType.get());
                    if (!cls.empty()) arrayClassElemTypes_[lower] = cls;
                }
                knownLocalVars_.insert(lower);
                break;
            }

            std::string cType = mapTypeRef(var.asType.get());

            // 记录变量类型集合 (用于Debug.Print和COM解封类型推断)
            if (cType == "BSTR") {
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownBstrVars_.insert(lower);
            } else if (cType == "uint8_t") {
                // 账 #123: As Byte 的 C 型就是 uint8_t, 上面任何一支都不匹配 ⇒ 以前谁也没登记它
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownByteVars_.insert(lower);
            } else if (cType == "int32_t" || cType == "int16_t" || cType == "VBABOOL") {
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownLongVars_.insert(lower);
                // ai/022 W1: 另登记一份布尔 (口径同 Fix 175 的 Date), 消费点先判它
                if (resolveArrayElemType(var.asType.get()) == Vb6Type::Boolean)
                    knownBoolVars_.insert(lower);
                // ai/009 5.10: 另登记一份 Integer, 收窄检查才分得出 16 位范围
                if (resolveArrayElemType(var.asType.get()) == Vb6Type::Integer)
                    knownIntVars_.insert(lower);
            } else if (cType == "intptr_t") {
                // Bug #2 fix: LongPtr变量注册到独立集合
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownLongPtrVars_.insert(lower);
            } else if (cType.find("vb6_ComIface_") == 0 || cType.find("vb6_ComIface_") != std::string::npos) {
                // Fix 082: COM interface pointer variables are also pointer-sized on x64
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownLongPtrVars_.insert(lower);
            } else if (cType == "vb6_VARIANT") {
                // P8.4: 记录Variant类型全局变量
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownVariantVars_.insert(lower);
            }

            // 记录Object类型变量名 (COM后期绑定)
            if (cType == "void*") {
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                // 项目类实例 (Dim b As Button / Dim WithEvents b As Button) 的 C 类型同样
                // 是 void*, 但它不是 COM 后期绑定对象, 方法调用必须走直接分发
                // (vb6_Button_DoClick(...)). 若误注册进 knownObjectVars_, cgen_expr 的
                // "优先级1" 会把它当 IDispatch 处理, 生成 vb6_ComCall(b, L"DoClick", ...)
                // → 对纯 C 结构体解引用 vtable → 运行期 0xC0000005.
                // 注意: 本处先于下方 knownClassVars_ 注册, 故直接查类符号判断.
                bool isProjectClassVar = false;
                if (var.asType && var.asType->kind == ASTNodeKind::SimpleTypeRef) {
                    auto& st = static_cast<SimpleTypeRef&>(*var.asType);
                    auto* clsSym = lookupModuleDotted(st.name);
                    if (clsSym && clsSym->kind == SymbolKind::Class) isProjectClassVar = true;
                }
                if (!isProjectClassVar) {
                    knownObjectVars_.insert(lower);
                }
            }

            // P6.3: 记录前期绑定COM变量 (Dim x As FileSystemObject)
            // 查找类型名是否对应ComClass符号
            if (var.asType && var.asType->kind == ASTNodeKind::SimpleTypeRef) {
                auto& simple = static_cast<SimpleTypeRef&>(*var.asType);
                auto* comSym = lookupModuleDotted(simple.name);
                if (comSym && (comSym->kind == SymbolKind::ComClass || comSym->kind == SymbolKind::ComInterface)) {
                    std::string lower = var.name;
                    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                    // Fix <vbeclipse>-2: 本工程有同名类模块 ⇒ 局部变量是**原生**工程类实例
                    // (VB6: 工程内定义优先于引用库; 类型库自动加载注进来的内建 coclass
                    // isExternal=false, 不能当外部 OCX 处理). 判据 = projectClassNameOf,
                    // 与 mapTypeRef / cgen_decl_func.cpp 同源.
                    const std::string projClsLoc = projectClassNameOf(simple.name);
                    if (!projClsLoc.empty()) {
                        knownClassVars_[lower] = projClsLoc;
                        knownObjectVars_.erase(lower);
                        if (var.isNew) knownNewVars_[lower] = cIdent(projClsLoc);
                    } else {
                    knownTypedComVars_[lower] = comSym;
                    // 从后期绑定集合中移除 (优先前期绑定)
                    knownObjectVars_.erase(lower);
                    }
                }
            }

            // 记录double/single类型变量名 (用于Debug.Print浮点输出)
            if (cType == "float") {
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownSingleVars_.insert(lower); knownDoubleVars_.insert(lower);   // Fix 117c: VT_R4
            } else if (cType == "double") {
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownDoubleVars_.insert(lower);
                // Fix 175: `Dim d As Date` 在 C 层同为 double, 只按 C 类型串登记时
                // inferExprType 只会回 Double → Print/拼接走 vb6_CStrDbl → 打出序列号。
                if (resolveArrayElemType(var.asType.get()) == Vb6Type::Date)
                    knownDateVars_.insert(lower);
            }

            // 记录类类型变量名, 默认值用NULL
            bool isLocalClassType = false;
            bool isLocalUdtType = false;
            bool isLocalEnumType = false;  // Fix 010q
            bool isLocalComIfaceType = false;
            bool isLocalVb6IfaceType = false;  // P6.4: VB6接口引用
            bool isLocalIvrefType = false;     // tB Interface 契约 (B04): 薄指针接口变量
            if (var.asType && var.asType->kind == ASTNodeKind::SimpleTypeRef) {
                auto& simple = static_cast<SimpleTypeRef&>(*var.asType);
                auto* clsSym = lookupModuleDotted(simple.name);
                // Fix <vbeclipse> rev7: **工程类名表兜底** —— 与 mapTypeRef 上移同一根因。
                // `Dim l_View As View` (modSubClass.bas:93) / `Dim l_Folder As Folder`
                // (ucPerspective.ctl 20+ 处) 里, 跨模块工程类在当前模块符号表**查不到
                // Class 符号** (driver 只把工程类名注入 projClassNames_), 于是
                // isLocalClassType 假 → 不进 knownClassVars_ → 该变量在成员访问时被
                // 当"模块限定符" → `l_View.View.hWnd` 发成 `vb6_View_prop_get_View.hWnd`
                // (丢了实参) → C2224 ".hWnd 左侧必须具有结构/联合类型" ×44。
                // 判据共用 projectClassNameOf (与 mapTypeRef / cParamClassPtrType 同源),
                // 不新增第二套"算不算工程类"的判断。
                std::string localProjCls;
                if (!clsSym || clsSym->kind != SymbolKind::Class) {
                    localProjCls = projectClassNameOf(simple.name);
                }
                if ((clsSym && clsSym->kind == SymbolKind::Class) || !localProjCls.empty()) {
                    const std::string localClsName =
                        localProjCls.empty() ? clsSym->name : localProjCls;
                    const bool localIsInterface =
                        clsSym && clsSym->kind == SymbolKind::Class && clsSym->isInterface;
                    std::string lower = var.name;
                    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                    // tB Interface 契约 (B04): 新式接口变量 -> knownIvrefVars_ + NULL 初值
                    if (!ivrefCType(simple.name).empty()) {
                        knownIvrefVars_[lower] = simple.name;
                        isLocalIvrefType = true;
                        // B05: 过程级接口变量在正常出口处 Release 自己那一份引用
                        trackIvrefLocalForRelease(cIdent(var.name));
                    } else if (localIsInterface) {
                        // P6.4: 接口类 → knownIfaceVars_ (而非 knownClassVars_)
                        knownIfaceVars_[lower] = localClsName;
                        isLocalVb6IfaceType = true;
                    } else {
                        // Fix 010r-10: map赋值, 存储类名以便方法分发时查找
                        knownClassVars_[lower] = localClsName;
                        isLocalClassType = true;
                        // P14.3.1: Dim As New自动实例化
                        if (var.isNew) {
                            knownNewVars_[lower] = cIdent(localClsName);
                        }
                    }
                }
                // As New 内建宿主类 Collection (无 Class 符号) → 惰性实例化
                if (var.isNew) {
                    std::string tn = simple.name;
                    std::transform(tn.begin(), tn.end(), tn.begin(), ::tolower);
                    if (tn == "collection") {
                        std::string lower = var.name;
                        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                        knownNewVars_[lower] = "Collection";
                    }
                }
                if (clsSym && (clsSym->kind == SymbolKind::ComClass || clsSym->kind == SymbolKind::ComInterface)) {
                    isLocalComIfaceType = true;
                    // Fix 090v-com: 前期绑定COM变量C类型缓存 (供 As New 守卫转型)
                    {
                        std::string lower = var.name;
                        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                        knownTypedComVarCType_[lower] = cType;
                    }
                    // Fix 090v-com: Dim As New ComClass 局部变量自动实例化守卫
                    // (clsSym->kind==Class 分支不含 COM 类, 这里补注册, 否则 json.Version 永为 NULL)
                    if (var.isNew && clsSym->kind == SymbolKind::ComClass) {
                        std::string lower = var.name;
                        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                        knownNewVars_[lower] = cIdent(clsSym->name);
                    }
                }
                // 检查是否是UDT类型
                auto* udtSym = lookupDotted(simple.name);
                if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                    isLocalUdtType = true;
                    // M22-fix: 注册到knownUdtVars_，防止成员访问被误判为模块名限定
                    std::string udtLower = var.name;
                    std::transform(udtLower.begin(), udtLower.end(), udtLower.begin(), ::tolower);
                    knownUdtVars_[udtLower] = "vb6_type_" + cIdent(simple.name);
                }
                // Fix 010q: 检查是否是Enum类型 (mapTypeRef映射为int32_t, 但defaultValue返回vb6_VariantEmpty())
                if (udtSym && udtSym->kind == SymbolKind::EnumType) {
                    isLocalEnumType = true;
                }
            }

            // Fix 110j: cType 已是 UDT 结构类型名 (vb6_type_X) 但符号表 lookupDotted
            // 未命中时, 仍按 UDT 处理. 否则:
            //   初始化 → `vb6_type_CHOOSECOLOR CC = 0;`      C2440 (int → struct)
            //   未注册 knownUdtVars_ → `With CC` 生成
            //                `void* w = (void*)CC;`           C2440 (struct → void*)
            // 实测: ppProgressCircular.pag 的 `Private Type CHOOSECOLOR` + ShowColor 的
            // `Dim CC As CHOOSECOLOR` / `With CC`.
            if (!isLocalUdtType && cType.rfind("vb6_type_", 0) == 0
                && cType.find('*') == std::string::npos) {
                isLocalUdtType = true;
                std::string udtLower = var.name;
                std::transform(udtLower.begin(), udtLower.end(), udtLower.begin(), ::tolower);
                knownUdtVars_[udtLower] = cType;
            }

            // VB6 Static变量: 跨调用持久化 → C static局部变量
            // 包括: 显式Static声明 或 Static Sub/Function内的所有局部变量
            std::string storageClass = (var.isStatic || inStaticProc_) ? "static " : "";

            // Fix 010r-12c: Register non-array local variable to knownLocalVars_
            // so it shadows cross-module external Public symbols with the same name.
            // (Array/dynamic-array cases already insert above; this covers all other types)
            {
                std::string lower = var.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownLocalVars_.insert(lower);
            }

            if (var.initializer) {
                // Fix 084aa: #undef 防宏污染 (见 P8.1 动态数组处注释)
                c_.emitLine("#undef " + cName);
                emitExpr(*var.initializer);
                c_.emitLine(storageClass + cType + " " + cName + " = " + lastExpr_ + ";");
            } else {
                std::string initVal;
                if (isLocalClassType || isLocalComIfaceType || isLocalIvrefType) {
                    initVal = "NULL";  // 薄指针接口变量 (B04) 与类实例一样是裸指针
                } else if (isLocalVb6IfaceType) {
                    initVal = "{0}";  // P6.4: 接口引用 = {vtbl=NULL, obj=NULL}
                } else if (isLocalUdtType) {
                    initVal = "{0}";
                } else if (isLocalEnumType) {  // Fix 010q
                    initVal = "0";
                } else if (var.asType && var.asType->kind == ASTNodeKind::FixedStringTypeRef) {
                    // String * N: 初始化为N个空格的BSTR, LSet/RSet使用固定长度
                    auto& fs = static_cast<FixedStringTypeRef&>(*var.asType);
                    emitExpr(*fs.length);
                    std::string fsLen = lastExpr_;
                    initVal = "vb6_BSTR_FixedSTR(" + fsLen + ")";
                    // 注册定长字符串变量名→长度
                    std::string fsLower = var.name;
                    std::transform(fsLower.begin(), fsLower.end(), fsLower.begin(), ::tolower);
                    knownFixedStringLen_[fsLower] = fsLen;
                } else {
                    initVal = defaultValue(
                        var.asType && var.asType->kind == ASTNodeKind::SimpleTypeRef
                            ? typeSys_.resolveTypeName(static_cast<SimpleTypeRef*>(var.asType.get())->name)
                            : Vb6Type::Variant
                    );
                }
                // Fix 084aa: 静态局部变量初始化必须是编译期常量 (C2099).
                // vb6_VariantEmpty()/vb6_BSTR_Empty() 是函数调用, 静态初始化会报错.
                // {0} (vt=0=VT_EMPTY) 与 vb6_VariantEmpty() 语义一致; NULL 即空BSTR.
                if (!storageClass.empty()) {
                    if (initVal == "vb6_VariantEmpty()") initVal = "{0}";
                    if (initVal == "vb6_BSTR_Empty()") initVal = "NULL";
                }
                // Fix 084aa: #undef 防宏污染 (见 P8.1 动态数组处注释)
                c_.emitLine("#undef " + cName);
                c_.emitLine(storageClass + cType + " " + cName + " = " + initVal + ";");
            }
            break;
        }
        case ASTNodeKind::ConstDecl: {
            auto& con = static_cast<ConstDecl&>(*node.decl);
            std::string cType = mapTypeRef(con.asType.get());
            // Fix 091d: 无 As 类型常量按字面量推断 C 类型. 此前一律 vb6_VARIANT →
            // `const vb6_VARIANT SW_SHOWNORMAL = 1;` 非法初始化 → C2440
            // (cToolsSystem.c 11/13); 且 Variant 常量参与位运算时操作数被包装
            // vb6_VariantToLong(<字面量>) → C2440 (cDialog.c 36 BIF_USENEWUI).
            if (!con.asType && con.value
                && con.value->kind == ASTNodeKind::LiteralExpr) {
                auto& lit091d = static_cast<LiteralExpr&>(*con.value);
                switch (lit091d.literalKind) {
                    case LiteralKind::Integer:
                    case LiteralKind::Long:
                        cType = "int32_t";
                        break;
                    case LiteralKind::LongPtr:  // Fix 082: ^ 后缀 -> 指针宽度
                        cType = "intptr_t";
                        break;
                    case LiteralKind::Single:
                    case LiteralKind::Double:
                        cType = "double";
                        break;
                    case LiteralKind::String:
                        cType = "BSTR";
                        break;
                    case LiteralKind::Boolean:
                        cType = "VBABOOL";
                        break;
                    default:
                        break;
                }
            }
            // Fix 110k: 无 As 类型常量且值是**算术/比较表达式** (非字面量, Fix 091d
            // 不覆盖) 时, cType 仍是 vb6_VARIANT →
            //   `const vb6_VARIANT PItoRAD = ((double)3.141592 / (double)180);`  C2440
            // (Charts 2020 ucPieChart.ctl:1187 `Const PItoRAD = 3.141592 / 180`).
            // VB6 无类型常量取表达式结果类型, 算术式 → Double. 字符串拼接 (&) 与
            // 比较式除外 (后者少见, 维持原 Variant 行为不变).
            if (!con.asType && con.value) {
                bool numeric110k = false;
                if (con.value->kind == ASTNodeKind::BinaryExpr) {
                    auto& be110k = static_cast<BinaryExpr&>(*con.value);
                    numeric110k = (be110k.op != BinaryOp::Concat
                                   && be110k.op != BinaryOp::Eq && be110k.op != BinaryOp::Neq
                                   && be110k.op != BinaryOp::Lt && be110k.op != BinaryOp::Gt
                                   && be110k.op != BinaryOp::Le && be110k.op != BinaryOp::Ge
                                   && be110k.op != BinaryOp::Is && be110k.op != BinaryOp::Like);
                } else if (con.value->kind == ASTNodeKind::UnaryExpr) {
                    numeric110k = true;
                }
                if (numeric110k) cType = "double";
            }
            std::string cName = cIdent(con.name);
            // Fix 010r-12c: Register local constant to knownLocalVars_
            {
                std::string lower = con.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownLocalVars_.insert(lower);
            }
            // Fix 049: Register local constant to type-specific known*Vars_ sets.
            // Same logic as Dim (cgen_decl.cpp:749-767). Without this, inferExprType
            // falls back to Variant for unknown identifiers, causing wrapToBSTR to
            // generate vb6_CStr(BSTR_const) which triggers C2440 (BSTR→VARIANT).
            {
                std::string lower = con.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                if (cType == "BSTR") {
                    knownBstrVars_.insert(lower);
            } else if (cType == "uint8_t") {
                    // 账 #123: 局部 Const As Byte 同 Dim 分支
                    knownByteVars_.insert(lower);
            } else if (cType == "int32_t" || cType == "int16_t" || cType == "VBABOOL") {
                    knownLongVars_.insert(lower);
                    // ai/022 W1: 同 Dim 分支 (Const 也吃这个读数)
                    if (resolveArrayElemType(con.asType.get()) == Vb6Type::Boolean)
                        knownBoolVars_.insert(lower);
                } else if (cType == "intptr_t") {
                    // Bug #2 fix: LongPtr局部const变量注册到独立集合
                    knownLongPtrVars_.insert(lower);
                } else if (cType.find("vb6_ComIface_") != std::string::npos) {
                    // Fix 082: COM interface pointer types are pointer-sized on x64
                    knownLongPtrVars_.insert(lower);
                } else if (cType == "float") {
                    knownSingleVars_.insert(lower); knownDoubleVars_.insert(lower);   // Fix 117c: VT_R4
                } else if (cType == "double") {
                    knownDoubleVars_.insert(lower);
                } else if (cType == "vb6_VARIANT") {
                    knownVariantVars_.insert(lower);
                }
            }
            if (con.value) {
                emitExpr(*con.value);
                // Fix 010r-13: Local Const redefining Windows API macro? #undef first.
                // 必须在声明 const 变量之前 #undef, 防止名字被 <windows.h> 等头文件中的
                // 宏展开 (例: WHITE_BRUSH、MEM_COMMIT、CP_UTF8、SW_SHOWNORMAL 等
                // 都是 windows.h 中的 #define, 否则 `const int32_t WHITE_BRUSH = 0;`
                // 会被宏展开为 `const int32_t 0 = 0;` 引发 C2106).
                // #undef 对没有定义为宏的名字是空操作, 无副作用.
                c_.emitLine("#undef " + cName);
                c_.emitLine("const " + cType + " " + cName + " = " + lastExpr_ + ";");
            }
            break;
        }
        default:
            c_.emitLine("/* unhandled LocalDeclStmt: " + std::string(node.decl->kindName()) + " */");
            break;
    }
}

} // namespace vb6c3
