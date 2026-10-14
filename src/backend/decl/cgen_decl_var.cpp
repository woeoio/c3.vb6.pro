#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_decl_var.cpp: 变量声明生成（VariableDecl，含模块级 / 类成员 / 局部） ---

// Forward declaration from cgen_base.cpp
 FrmControlType controlTypeFromName(const std::string& name);


void CCodeGen::visit(VariableDecl& node) {
    std::string cName = cIdent(node.name);

    // P8.1: 模块级数组声明 (支持多维)
    if (!node.dimensions.empty()) {
        Vb6Type elemType = resolveArrayElemType(node.asType.get());
        std::string saElemType = mapSaElemType(elemType);
        int dimCount = (int)node.dimensions.size();
        std::string cType = (dimCount > 1) ? "vb6_SafeArrayND*" : "vb6_SafeArray1D*";

        // 前向声明 -> .h
        if (!trackOnly_) {
        if (isPublicModuleDecl(node)) {
            h_.emitLine("extern " + cType + " " + cName + ";");
        }

        if (dimCount == 1) {
            // 一维数组
            auto& dim = node.dimensions[0];
            std::string lBound = "0";
            std::string uBound = "0";
            if (dim.lower) { emitExpr(*dim.lower); lBound = std::move(lastExpr_); }
            if (dim.upper) { emitExpr(*dim.upper); uBound = std::move(lastExpr_); }
            std::string initCode = "vb6_SafeArrayCreate1D(" + saElemType + ", " + lBound + ", " + uBound + ")";
            // Fix 054: C语言文件作用域变量必须用常量表达式初始化 (C2099)
            // 改为先声明为NULL, 再在模块初始化函数中赋值
            if (isPublicModuleDecl(node)) {
                c_.emitLine(cType + " " + cName + " = NULL;");
                moduleInitStmts_.push_back(cName + " = " + initCode + ";");
            } else {
                c_.emitLine("static " + cType + " " + cName + " = NULL;");
                moduleInitStmts_.push_back(cName + " = " + initCode + ";");
            }
        } else {
            // 多维数组: 使用ND运行时
            std::string boundsVar = "_bounds_" + cName;
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
            std::string initCode = "vb6_SafeArrayCreateND(" + saElemType + ", " + std::to_string(dimCount) + ", " + boundsVar + ")";
            // Fix 054: C语言文件作用域变量必须用常量表达式初始化 (C2099)
            if (isPublicModuleDecl(node)) {
                c_.emitLine(cType + " " + cName + " = NULL;");
                moduleInitStmts_.push_back(cName + " = " + initCode + ";");
            } else {
                c_.emitLine("static " + cType + " " + cName + " = NULL;");
                moduleInitStmts_.push_back(cName + " = " + initCode + ";");
            }
        }
        } // end if (!trackOnly_)

        // 注册到已知数组集合
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownArrays_.insert(lower);
        arrayElemTypes_[lower] = elemType;
        arrayDimCounts_[lower] = dimCount;
        // Fix 062: Byte 数组变量注册
        if (elemType == Vb6Type::Byte) knownByteArrayVars_.insert(lower);
        // Fix 055: 注册UDT数组元素C类型
        {
            std::string udtCType = resolveArrayUdtElemCType(node.asType.get());
            if (!udtCType.empty()) arrayUdtElemTypes_[lower] = udtCType;
        }
        // Fix 192: 注册类数组元素类名 (`Dim s(1) As ShapeAct` — 静态数组同样踩 C2224,
        // 不是只有 ReDim 的形状坏)。没有它, s(0).Move 的接收者推断不出类。
        {
            std::string cls = resolveArrayClassElemType(node.asType.get());
            if (!cls.empty()) arrayClassElemTypes_[lower] = cls;
        }
        if (!trackOnly_) knownLocalVars_.insert(lower);
        return;
    }

    // P8.1: 动态数组声明: Dim arr() As Long -> 默认1D
    if (node.isDynamicArray) {
        Vb6Type elemType = resolveArrayElemType(node.asType.get());
        std::string cType = "vb6_SafeArray1D*";

        if (!trackOnly_) {
        if (isPublicModuleDecl(node)) {
            h_.emitLine("extern " + cType + " " + cName + ";");
            c_.emitLine(cType + " " + cName + " = NULL;");
        } else {
            c_.emitLine("static " + cType + " " + cName + " = NULL;");
        }
        // Fix 092w: Dim arr() As Byte = <初始化表达式> (twinbasic 兼容) — 文件作用域
        // 必须用常量初始化 (C2099), 故声明为 NULL, 初始化表达式放到模块初始化函数中,
        // 与 Fix 054 静态数组的 moduleInitStmts_ 模式一致.
        if (node.initializer && elemType == Vb6Type::Byte) {
            emitExpr(*node.initializer);
            moduleInitStmts_.push_back(cName + " = " + rewriteByteArrayValue(lastExpr_) + ";");
        }
        } // end if (!trackOnly_)

        // 注册到已知数组集合
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownArrays_.insert(lower);
        arrayElemTypes_[lower] = elemType;
        arrayDimCounts_[lower] = 1;  // 动态数组默认1D
        // Fix 062: Byte 数组变量注册
        if (elemType == Vb6Type::Byte) knownByteArrayVars_.insert(lower);
        // Fix 055: 注册UDT数组元素C类型
        {
            std::string udtCType = resolveArrayUdtElemCType(node.asType.get());
            if (!udtCType.empty()) arrayUdtElemTypes_[lower] = udtCType;
        }
        // Fix 192: 同上, 动态数组 `Dim a() As ShapeAct`
        {
            std::string cls = resolveArrayClassElemType(node.asType.get());
            if (!cls.empty()) arrayClassElemTypes_[lower] = cls;
        }
        if (!trackOnly_) knownLocalVars_.insert(lower);
        return;
    }

    std::string cType = mapTypeRef(node.asType.get());

    // 检查是否是类类型变量 → 注册到 knownClassVars_
    if (node.asType && node.asType->kind == ASTNodeKind::SimpleTypeRef) {
        auto& simple = static_cast<SimpleTypeRef&>(*node.asType);
        auto* clsSym = lookupModuleDotted(simple.name);
        if (clsSym && clsSym->kind == SymbolKind::Class) {
            std::string lower = node.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            // tB Interface 契约 (B04): 新式接口变量 -> knownIvrefVars_
            if (!ivrefCType(simple.name).empty()) {
                knownIvrefVars_[lower] = simple.name;
            } else if (clsSym->isInterface) {
                // P6.4: 接口类 → 注册到 knownIfaceVars_ (而非 knownClassVars_)
                knownIfaceVars_[lower] = clsSym->name;
            } else {
                // Fix 010r-10: map赋值, 存储类名以便方法分发时查找
                knownClassVars_[lower] = clsSym->name;
                // P14.3.1: Dim As New自动实例化 (模块级)
                if (node.isNew) {
                    knownNewVars_[lower] = cIdent(clsSym->name);  // Fix 090v reg
                    moduleNewVars_[lower] = cIdent(clsSym->name);  // Fix 090v
                }
            }
            // P6.5: WithEvents变量 → 注册到 knownWithEventsVars_
            if (node.isWithEvents) {
                knownWithEventsVars_[lower] = clsSym->name;
            }
        }
        // As New 内建宿主类 Collection: 无 Class 符号, 但需按 VB6 语义惰性实例化
        // (否则 colTooltips 恒为 NULL, On Error Resume Next + Err 的
        //  "集合项是否存在" 判断全部走错分支)。
        if (node.isNew) {
            std::string tn = simple.name;
            std::transform(tn.begin(), tn.end(), tn.begin(), ::tolower);
            if (tn == "collection") {
                std::string lower = node.name;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                knownNewVars_[lower] = "Collection";
                moduleNewVars_[lower] = "Collection";
            }
        }
    }

    // 检查是否是UDT类型变量 → 注册到 knownUdtVars_
    if (node.asType && node.asType->kind == ASTNodeKind::SimpleTypeRef) {
        auto& simpleUdt = static_cast<SimpleTypeRef&>(*node.asType);
        auto* udtSymDecl = lookupDotted(simpleUdt.name);
        if (udtSymDecl && udtSymDecl->kind == SymbolKind::UserDefinedType) {
            std::string udtLower = node.name;
            std::transform(udtLower.begin(), udtLower.end(), udtLower.begin(), ::tolower);
            knownUdtVars_[udtLower] = "vb6_type_" + cIdent(simpleUdt.name);
            // Fix 233: 与 UDT 互斥的那几张表全程不清空 (全后端无一处 .clear()), 同名
            // 条目会从别的模块/过程泄漏过来。变量名复用极常见 —— 实测 ComCtlsDemo
            // ToolBar.ctl 里 `NewButton` 既是 `Friend Sub FButtonsAdd(ByVal NewButton
            // As TbrButton)` 的类形参, 又是 `Friend Property Let FButtonStyle` 里的
            // `Dim NewButton As ShadowButtonStruct`。残留的
            // knownClassVars_["newbutton"]="TbrButton" 让 `With NewButton` (cgen_with
            // 的类分支排在 UDT 分支之前) 被当类实例, `.TBB / .fsState / .fsStyle`
            // 全走类成员分发 ⇒ C2039×33 + C2440(vb6_type_ShadowButtonStruct →
            // vb6_cls_CheckBoxW*)。局部 Dim 是**本过程**的作用域事实, 故照
            // cgen_decl_func.cpp 的 UDT 形参守卫同一口径, 把互斥条目一并擦掉。
            knownClassVars_.erase(udtLower);
            knownTypedComVars_.erase(udtLower);
            knownIfaceVars_.erase(udtLower);
            knownIvrefVars_.erase(udtLower);
            knownObjectVars_.erase(udtLower);
            knownVariantVars_.erase(udtLower);
        }
    }

    // 检查是否是定长字符串变量 → 注册到 knownFixedStringLen_
    if (node.asType && node.asType->kind == ASTNodeKind::FixedStringTypeRef) {
        auto& fs = static_cast<FixedStringTypeRef&>(*node.asType);
        std::string fsLower = node.name;
        std::transform(fsLower.begin(), fsLower.end(), fsLower.begin(), ::tolower);
        // 评估长度表达式(必须是编译期常量)
        emitExpr(*fs.length);
        knownFixedStringLen_[fsLower] = lastExpr_;
    }

    // vbeclipse: 声明类型名是**工程类**但被本模块同名成员遮蔽时的补注册.
    //
    // 上面那段 Class 分支用 lookupModuleDotted(simple.name) 判工程类, 而每模块一张
    // SymbolTable —— 同名成员会遮蔽跨模块类名. vbeclipse 实测两种形态:
    //   ucSplitBar.ctl `Private WithEvents SplitBar As SplitBar`
    //     → 类型名被字段自身 (SymbolKind::Variable) 占位;
    //   ucCaption.ctl `Private WithEvents m_PopupMenu As PopupMenu`
    //     → 类型名被同名 Sub 占位 (SymbolKind::Sub).
    // Class 分支判不中 → mapTypeRef 回落 void* → 下面注册进 knownObjectVars_ →
    // `With SplitBar` 走 COM 后期绑定, .SplitterMouseDown 的 RECT 结构体实参被
    // vb6_ComPackInt 打包 (C2440), 运行期还会拿 C 结构体当 IDispatch 解 vtable.
    // 工程级类名表不受遮蔽, 是这类判定的唯一正确来源 (与 driver 侧
    // projectClassNames 同源, 见 setProjectClassNames 注释).
    if (node.asType && node.asType->kind == ASTNodeKind::SimpleTypeRef && cType == "void*") {
        auto& shadowed = static_cast<SimpleTypeRef&>(*node.asType);
        std::string projCls = projectClassNameOf(shadowed.name);
        if (!projCls.empty() &&
            knownClassVars_.find(Symbol::toLower(node.name)) == knownClassVars_.end() &&
            knownIfaceVars_.find(Symbol::toLower(node.name)) == knownIfaceVars_.end()) {
            std::string lower = node.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            knownClassVars_[lower] = projCls;
            knownObjectVars_.erase(lower);
            if (node.isWithEvents) knownWithEventsVars_[lower] = projCls;
            if (node.isNew) {
                knownNewVars_[lower] = cIdent(projCls);
                moduleNewVars_[lower] = cIdent(projCls);
            }
        }
    }

    // 检查是否是Object类型变量 → 注册到 knownObjectVars_ (COM后期绑定)
    if (cType == "void*") {  // Object类型映射为void*
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        // 项目类实例 (Dim b As Button / Dim WithEvents b As Button) 的 C 类型同样是 void*,
        // 但它并非 COM 后期绑定对象, 方法调用必须走直接分发 (vb6_Button_DoClick(...)).
        // 若误注册进 knownObjectVars_, cgen_expr 的"优先级1"会把它当 IDispatch 处理,
        // 生成 vb6_ComCall(b, L"DoClick", ...) → 对纯 C 结构体解引用 vtable → 运行期 0xC0000005.
        if (knownClassVars_.find(lower) == knownClassVars_.end() &&
            knownIfaceVars_.find(lower) == knownIfaceVars_.end()) {
            knownObjectVars_.insert(lower);
        }
    }

    // Fix 082: COM interface pointer types (vb6_ComIface_*) are pointer-sized on x64
    if (cType.find("vb6_ComIface_") != std::string::npos) {
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownLongPtrVars_.insert(lower);
    }

    // P6.3: 检查是否是前期绑定COM变量 → 注册到 knownTypedComVars_
    if (node.asType && node.asType->kind == ASTNodeKind::SimpleTypeRef) {
        auto& simple = static_cast<SimpleTypeRef&>(*node.asType);
        auto* comSym = lookupModuleDotted(simple.name);
        if (comSym && (comSym->kind == SymbolKind::ComClass || comSym->kind == SymbolKind::ComInterface)) {
            std::string lower = node.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            // Fix <vbeclipse>-2: 本工程有同名类模块 ⇒ 该类型名指工程类, 走**原生**
            // (VB6: 工程内定义优先于引用库). 类型库自动加载把 Shell32 的 coclass
            // Folder / ScrRun 的 Dictionary 注进每个模块, 且 isExternal=false ——
            // 若登记进 knownTypedComVars_ 就会把原生 vb6_cls_<Name>* 当 IDispatch
            // 解 vtable (ucPerspective CreateFolder 的 `With Folder.Views` → 0xC0000005).
            const std::string projClsVar = projectClassNameOf(simple.name);
            if (!projClsVar.empty()) {
                knownClassVars_[lower] = projClsVar;
                knownObjectVars_.erase(lower);
                if (node.isNew) {
                    knownNewVars_[lower] = cIdent(projClsVar);   // Dim As New 工程类
                    moduleNewVars_[lower] = cIdent(projClsVar);
                }
                if (node.isWithEvents) knownWithEventsVars_[lower] = projClsVar;
            } else {
            knownTypedComVars_[lower] = comSym;
            knownTypedComVarCType_[lower] = cType;  // Fix 090v-com: 供 As New 守卫转型
            knownObjectVars_.erase(lower);  // 优先前期绑定
            // Dim As New ComClass 自动实例化 (P14.3.1扩展)
            if (node.isNew && comSym->kind == SymbolKind::ComClass) {
                knownNewVars_[lower] = cIdent(comSym->name);  // Fix 090v reg2
                moduleNewVars_[lower] = cIdent(comSym->name);  // Fix 090v
            }
            // P13.23: ComClass WithEvents -> knownWithEventsVars_
            if (node.isWithEvents && comSym->kind == SymbolKind::ComClass && comSym->comHasSourceIface) {
                knownWithEventsVars_[lower] = comSym->name;
            }
            }
        }
    }

        // P16: WithEvents控件类型检测 → 注册到 knownWithEventsCtrlVars_
    // Dim WithEvents cmd As CommandButton → knownWithEventsCtrlVars_["cmd"] = CommandButton
    if (node.isWithEvents && node.asType && node.asType->kind == ASTNodeKind::SimpleTypeRef) {
        auto& simple16 = static_cast<SimpleTypeRef&>(*node.asType);
        FrmControlType ctrlType = controlTypeFromName(simple16.name);
        if (ctrlType != FrmControlType::Unknown) {
            std::string lower16 = node.name;
            std::transform(lower16.begin(), lower16.end(), lower16.begin(), ::tolower);
            knownWithEventsCtrlVars_[lower16] = ctrlType;
            knownWithEventsCtrlOrigNames_[lower16] = cName;  // 保留原始变量名(大小写)
            cType = "HWND";  // 控件WithEvents变量存储HWND
            knownObjectVars_.erase(lower16);  // 移除可能的void*标记
            knownVariantVars_.erase(lower16);  // 移除可能的Variant标记
        } else {
            // Fix 178: 非标准控件的 WithEvents 变量若解析为工程类 (如 cHttpServer 的
            // Private WithEvents m_oServer As cTlsReMaster), 必须保持 knownClassVars_
            // 早绑定路径, 不能按 Fix 056a 强转 void* + knownObjectVars_ — 否则
            // With m_oServer 在 cgen_with 的检测序 (knownObjectVars_ 先于
            // knownClassVars_) 中被判 COMObject, 块内 .Protocol = 0 / .Bind 生成
            // vb6_ComSetProp/vb6_ComCall(裸结构体, ...) → vb6_getDispid 把
            // __comObj 首字段当 vtable 解引用 → 运行期 0xC0000005
            // (VBMAN_DEMO Form_Load → cHttpServer.Start, 2026-09-21 CI 实证)。
            // 工程类实例不是 COM 对象; 未解析类型 (VBControlExtender 等) 不受影响。
            // 守卫条件与上方 void* 分支 (knownClassVars_/knownIfaceVars_ 排除) 对齐。
            bool weProjectClass = false;
            if (node.asType->kind == ASTNodeKind::SimpleTypeRef) {
                auto* weClsSym = lookupModuleDotted(simple16.name);
                if (weClsSym && weClsSym->kind == SymbolKind::Class
                    && !weClsSym->isInterface) {
                    weProjectClass = true;
                }
            }
            if (!weProjectClass) {
                // Fix 056a: 非标准控件的WithEvents变量(如VBControlExtender)当作COM对象
                // cType可能是int32_t(mapTypeRef默认值), 必须改为void*
                std::string lower16 = node.name;
                std::transform(lower16.begin(), lower16.end(), lower16.begin(), ::tolower);
                if (cType != "void*") {
                    cType = "void*";
                    knownObjectVars_.insert(lower16);
                    knownVariantVars_.erase(lower16);
                    knownLongVars_.erase(lower16);
                }
            }
        }
    }
// 记录变量类型集合 (用于Debug.Print和COM解封类型推断)
    if (cType == "float") {
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownSingleVars_.insert(lower); knownDoubleVars_.insert(lower);   // Fix 117c: VT_R4
    } else if (cType == "double") {
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownDoubleVars_.insert(lower);
        // Fix 175: 模块级 `Private d As Date` 与 Double 同型, 需另登记才能被
        // inferExprType 认成 Date (口径同 cgen_localdecl.cpp 的 Dim 分支)。
        if (resolveArrayElemType(node.asType.get()) == Vb6Type::Date)
            knownDateVars_.insert(lower);
    } else if (cType == "BSTR") {
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownBstrVars_.insert(lower);
    } else if (cType == "uint8_t") {
        // ai/009 §5.10: 模块级 Byte 走独立集合 (口径同 cgen_localdecl.cpp 的 Dim 分支)
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownByteVars_.insert(lower);
    } else if (cType == "int32_t" || cType == "int16_t" || cType == "VBABOOL") {
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownLongVars_.insert(lower);
        // ai/022 W1: 模块级 As Boolean 同型, 需另登记 (口径同上处 Dim 分支)
        if (resolveArrayElemType(node.asType.get()) == Vb6Type::Boolean)
            knownBoolVars_.insert(lower);
        // ai/009 5.10: 模块级 As Integer 同型, 收窄检查需分出 16 位范围
        if (resolveArrayElemType(node.asType.get()) == Vb6Type::Integer)
            knownIntVars_.insert(lower);
    } else if (cType == "vb6_VARIANT") {
        // P8.4: 记录Variant类型局部变量
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownVariantVars_.insert(lower);
        // Fix 091m: 模块级 (过程外) Variant 变量额外登记, 供过程体内回灌.
        // 类模块/窗体的成员变量不属此列 (方法体内须写 me->name, 回灌裸名会 C2065).
        if (currentProc_ == nullptr) {
            if (isClassModule_) {
                // Fix 091p: 类模块/窗体字段单独登记 — 供 me->field 形态判定
                // (cWinsock.m_vUserData As Variant: Set m_vUserData = Value 需
                // 走 Variant 容器分支, 而非 ToObjectVal 提取).
                classVariantFields_.insert(lower);
            } else {
                moduleVariantVars_.insert(lower);
            }
        }
    }

    // Fix 010: 类模块tracking-only模式, 跳过变量声明生成(已在结构体中)
    if (trackOnly_) return;

    // Fix 010o: 注册局部变量到 knownLocalVars_ (非trackOnly模式 = 过程内局部Dim)
    {
        std::string lower = node.name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        knownLocalVars_.insert(lower);
    }

    // 前向声明 → .h, 定义 → .c
    if (isPublicModuleDecl(node)) {
        h_.emitLine("extern " + cType + " " + cName + ";");
    }

    // 变量定义 → .c
    // 类类型变量的默认值是NULL
    bool isClassType = false;
    bool isComIfaceType = false;
    bool isVb6IfaceType = false;  // P6.4: VB6接口引用类型
    if (node.asType && node.asType->kind == ASTNodeKind::SimpleTypeRef) {
        auto& simple = static_cast<SimpleTypeRef&>(*node.asType);
        auto* clsSym = lookupModuleDotted(simple.name);
        isClassType = (clsSym && clsSym->kind == SymbolKind::Class && !clsSym->isInterface);
        isVb6IfaceType = (clsSym && clsSym->kind == SymbolKind::Class && clsSym->isInterface);
        isComIfaceType = (clsSym && (clsSym->kind == SymbolKind::ComClass || clsSym->kind == SymbolKind::ComInterface));
    }

    if (node.initializer) {
        emitExpr(*node.initializer);
        if (isPublicModuleDecl(node)) {
            c_.emitLine(cType + " " + cName + " = " + lastExpr_ + ";");
        } else {
            c_.emitLine("static " + cType + " " + cName + " = " + lastExpr_ + ";");
        }
    } else {
        // 判断是否是UDT/Enum类型 → 用 {0} 或 0 初始化
        bool isUdtType = false;
        bool isEnumType = false;  // Fix 010q
        if (node.asType && node.asType->kind == ASTNodeKind::SimpleTypeRef) {
            auto& simple = static_cast<SimpleTypeRef&>(*node.asType);
            auto* sym = lookupDotted(simple.name);
            isUdtType = (sym && sym->kind == SymbolKind::UserDefinedType);
            isEnumType = (sym && sym->kind == SymbolKind::EnumType);  // Fix 010q
        }
        std::string initVal;
        if (isClassType || isComIfaceType || cType == "HWND") {
            initVal = "NULL";
        } else if (isVb6IfaceType) {
            initVal = "{0}";  // P6.4: 接口引用 = {vtbl=NULL, obj=NULL}
        } else if (isUdtType) {
            initVal = "{0}";
        } else if (isEnumType) {  // Fix 010q
            initVal = "0";
        } else {
            // 账 #116: 这一句原先无条件 `static_cast<SimpleTypeRef*>(node.asType.get())->name`,
            // 而定长串 `As String * N` 的 typeRef 是 **FixedStringTypeRef** —— 把它的
            // `ExprPtr length` 当成 `std::string` 读, 那个"长度"其实是一个堆指针, 于是拷贝
            // 字符串时张口就要几十 GB: operator new 失败 → std::bad_alloc → 无人接住 →
            // abort() (退出码 3, 零诊断, 调试版 CRT 还弹模态框)。
            // 只有 kinds 判过才转。**非 SimpleTypeRef 一律回 Unknown 而不是 Variant**:
            // 修好之前那条瞎读路径的"实际效果"就是 Unknown (`resolveTypeName` 查不到那个乱码名),
            // 而它印出来是 `0` —— 数组那类 `vb6_SafeArray1D*` 要的正是这个空指针; 改成 Variant
            // 会发 `vb6_VariantEmpty()`, 两个 GUI 存量工程立刻 C2440 (VARIANT ↔ SafeArray1D*)。
            Vb6Type asType = Vb6Type::Variant;   // 无 As 子句
            if (node.asType) {
                asType = Vb6Type::Unknown;
                if (node.asType->kind == ASTNodeKind::SimpleTypeRef) {
                    asType = typeSys_.resolveTypeName(
                        static_cast<SimpleTypeRef*>(node.asType.get())->name);
                } else if (node.asType->kind == ASTNodeKind::FixedStringTypeRef) {
                    asType = Vb6Type::String;
                }
            }
            initVal = defaultValue(asType);
        }
        // M22: 文件作用域BSTR初始化不能用函数调用(vb6_BSTR_Empty), 用NULL替代
        if (initVal == "vb6_BSTR_Empty()") initVal = "NULL";
        // Fix 084aa: 文件作用域Variant初始化不能用函数调用(vb6_VariantEmpty), 用{0}替代
        // ({0} 即 vt=0=VT_EMPTY, 与 vb6_VariantEmpty() 语义一致)
        if (initVal == "vb6_VariantEmpty()") initVal = "{0}";
        if (isPublicModuleDecl(node)) {
            c_.emitLine(cType + " " + cName + " = " + initVal + ";");
        } else {
            c_.emitLine("static " + cType + " " + cName + " = " + initVal + ";");
        }
    }
}

} // namespace vb6c3
