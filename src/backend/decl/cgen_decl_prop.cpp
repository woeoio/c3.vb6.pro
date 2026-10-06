#include "backend/cgen.hpp"
#include "semantics/interface_sig.hpp"  // ifaceLower (Fix <vbeclipse>: 接口桩槽键)
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_decl_prop.cpp: 属性与事件声明生成（PropertyDecl / EventDecl + 属性签名） ---


void CCodeGen::visit(PropertyDecl& node) {
    // 泛型模板 (tB, G2/G3): 模板本体不发码 (泛型器注入特化副本)
    if (!node.typeParams.empty()) return;
    // Property Get/Let/Set → C函数
    // 类模块: 第一个参数为 me 指针
    std::string sig = makePropertySignature(node);
    c_.emitLine(sig + " {");

    // P6.6修复: 设置currentProc_ (与SubDecl/FunctionDecl相同)
    // 这确保IdentifierExpr中的类模块变量加me->前缀, PropertyGet返回值赋值正确
    // Fix 032: 必须按属性种类精确查找符号. lookupModule() 默认返回 PropertyGet
    // ($pg 后缀, 优先级最高), 导致 Property Let/Set 体内 currentProc_ 被错误设为
    // 同名 PropertyGet 的符号 — currentProc_->params 缺失 Let/Set 的最后一个
    // 形参 (赋值 RHS, 如 Property Let SockOpt 中的 Value), 致使 IdentifierExpr
    // 的 ByRef 形参短路 (line 311-332) 无法匹配 → 形参被误解析为跨模块
    // PropertyGet 调用 (如 cAsyncSocket Value 形参 → vb6_cCsv_prop_get_Value).
    // Fix 032 改 args 发射期间 asCallCallee_=false 后, 该引用从"函数名裸引用"
    // 变为"实际函数调用 vb6_cCsv_prop_get_Value((void*)me)", 返回 vb6_VARIANT,
    // 触发 100+ 个 C2440. 与语义分析 Pass2 (semantic_analyzer.cpp:775) 一致使用
    // lookupModuleByKind 精确定位属性符号.
    SymbolKind propSk;
    switch (node.propKind) {
        case ProcKind::PropertyGet:  propSk = SymbolKind::PropertyGet; break;
        case ProcKind::PropertyLet:  propSk = SymbolKind::PropertyLet; break;
        case ProcKind::PropertySet:  propSk = SymbolKind::PropertySet; break;
        default:                     propSk = SymbolKind::PropertyGet; break;
    }
    auto* propSym = symTab_.lookupModuleByKind(node.name, propSk);
    currentProc_ = propSym;

    // Fix 056b: 清理局部数组注册 (模块级/类成员数组跨过程保留)
    clearProcArrayTracking();
    ansiTempsToFree_.clear();
    ansiOutParams_.clear();
    ivrefLocalsToRelease_.clear();  // tB Interface B05
    ansiCounter_ = 0;
    asmMixedBlockCounter_ = 0;   // ai/vb-asm-extension-spec 项2: 混排片段序号按过程重置
    knownBstrVars_.clear();
    knownDoubleVars_.clear();
    knownSingleVars_.clear();
    knownDateVars_.clear();   // Fix 175
    knownBoolVars_.clear();     // ai/022 W1
    knownByteVars_.clear();     // 账 #123
    knownIntVars_.clear();       // ai/009 5.10
    knownLongVars_.clear();
    knownLongPtrVars_.clear();  // Bug #2 fix: 也清空LongPtr集合
    knownVariantVars_.clear();
    // M22-fix: 只清空UDT变量map(旧条目会冲突), set类型不清空(WithEvents等模块级条目需跨过程保留)
    knownUdtVars_.clear();
    knownFixedStringLen_.clear();
    // Fix 010o: 清空局部变量集合
    knownLocalVars_.clear();
    knownNewVars_.clear();
    knownNewVars_.insert(moduleNewVars_.begin(), moduleNewVars_.end());  // Fix 090v
    // Fix 091m: 回灌模块级 Variant 变量 (knownVariantVars_ 已被 clear)
    knownVariantVars_.insert(moduleVariantVars_.begin(), moduleVariantVars_.end());
    knownByRefParams_.clear();  // Fix 081g
    // Fix 160w: 过程入口重置 COM 属性派发标记与 With 栈 — 上一个过程的 WithMemberExpr
    // 值 (如 `.Result`, `PropCellFont.Size`) 只把对象表达式留在 lastExpr_, 成员名挂
    // 在 comObjExpr_/comMemberName_/isComMarker_… 若不清, 会泄漏到下一个过程被其
    // BinaryExpr/赋值误当作 COM 属性读取 → C2065 (_vb6_with_14/_vb6_with_15,
    // Command5_Click -> Command10_Click) / C2440 (vb6_ComIface_Font*→float).
    isComMarker_ = false;
    comObjExpr_.clear();
    comMemberName_.clear();
    isEarlyBoundCom_ = false;
    earlyBoundSym_ = nullptr;
    withObjectVars_.clear();
    withObjectInfoStack_.clear();
    // Fix 010r/010r-10: 类模块中注册me到knownClassVars_ (使Me.Method()正确分发)
    // 改为map赋值: me → 当前模块名(类名)
    if (isClassModule_) knownClassVars_["me"] = moduleName_;
    // P6.11: 恢复类模块成员变量类型 (clear后从持久化集合恢复)
    knownBstrVars_.insert(classBstrMembers_.begin(), classBstrMembers_.end());
    knownDoubleVars_.insert(classDoubleMembers_.begin(), classDoubleMembers_.end());
    knownLongVars_.insert(classLongMembers_.begin(), classLongMembers_.end());
    // Fix 010n: 恢复类模块UDT成员变量 (knownUdtVars_被clear后需要从classUdtMembers_恢复)
    knownUdtVars_.insert(classUdtMembers_.begin(), classUdtMembers_.end());
    // Fix 010n (扩展): 恢复普通模块模块级UDT变量 (同 classUdtMembers_ 机制)
    knownUdtVars_.insert(moduleUdtMembers_.begin(), moduleUdtMembers_.end());

    // M22-fix: 注册参数中的UDT/类/接口变量到跟踪集合
    for (auto& p : node.params) {
        // Fix 081g: Register ByRef params for For-loop dereference fix
        if (!p->isByVal) {
            std::string brKey = p->name;
            std::transform(brKey.begin(), brKey.end(), brKey.begin(), ::tolower);
            knownByRefParams_.insert(brKey);
        }
        // Fix 084m: 无类型子句的 Optional 参数 (如 Optional RecordsAffected) 默认是
        // Variant, C 类型 vb6_VARIANT*; 必须注册到 knownVariantVars_,
        // 否则 `RecordsAffected = 123` 生成裸赋值 → C2440 (cDataBase Exec).
        if (!p->asType) {
            std::string pLower = p->name;
            std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);
            knownVariantVars_.insert(pLower);
        }
        if (p->asType && p->asType->kind == ASTNodeKind::SimpleTypeRef) {
            auto& simpleP = static_cast<SimpleTypeRef&>(*p->asType);
            std::string pLower = p->name;
            std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);
            // Fix 161f: 同 cgen_decl_func.cpp — 限定名 (ComctlLib.ColumnHeader) 用
            // lookupTypeSymbol 按全名→末段查, 与 mapTypeRef 同口径。
            auto* pSym = lookupTypeSymbol(simpleP.name);
            if (!pSym) {
                size_t pDot161f = simpleP.name.find('.');
                if (pDot161f != std::string::npos)
                    pSym = lookupTypeSymbol(simpleP.name.substr(pDot161f + 1));
            }
            // Fix <vbeclipse> rev7: 工程类名表兜底 —— 同 cgen_decl_proc.cpp 同款。
            if (!pSym || pSym->kind != SymbolKind::Class) {
                const std::string projClsP092r = projectClassNameOf(simpleP.name);
                if (!projClsP092r.empty()) knownClassVars_[pLower] = projClsP092r;
            }
            if (pSym && pSym->kind == SymbolKind::UserDefinedType) {
                knownUdtVars_[pLower] = "vb6_type_" + cIdent(simpleP.name);
            } else if (pSym && pSym->kind == SymbolKind::Class) {
                knownClassVars_[pLower] = pSym->name;
            } else if (pSym && (pSym->kind == SymbolKind::ComClass || pSym->kind == SymbolKind::ComInterface)) {
                // Fix <vbeclipse>-2: 同 cgen_decl_func.cpp —— 本工程有同名类模块时走原生
                // (判据 projectClassNameOf, 不是 isExternal; 类型库自动加载注进来的
                // 内建 coclass isExternal=false).
                const std::string projCls2 = projectClassNameOf(simpleP.name);
                if (!projCls2.empty()) {
                    knownClassVars_[pLower] = projCls2;
                } else {
                    knownTypedComVars_[pLower] = pSym;
                }
            }
            auto* pSym2 = symTab_.lookup(simpleP.name);
            if (pSym2 && pSym2->kind == SymbolKind::Class && pSym2->isInterface) {
                knownIfaceVars_[pLower] = pSym2->name;
            }
            // Fix 161f: `lv As ListView` 形参槽就是 HWND —— 同 cgen_decl_func.cpp。
            if (Symbol::toLower(simpleP.name).find("listview") != std::string::npos)
                listViewSlotVars_.insert(pLower);
            // 注册BSTR/Double/Long类型参数到类型跟踪集合
            Vb6Type paramType = typeSys_.resolveTypeName(simpleP.name);
            if (paramType == Vb6Type::String) knownBstrVars_.insert(pLower);
            // Fix 123b: Currency/Single/Date 形参同属 C double 组 — 此前漏注册,
            // inferExprType(形参) 落符号表查找 (129 模块工程里同名符号撞车) 或
            // Variant → RaiseEvent 实参打包误走 BSTR 分支 (cZipArchive
            // frFireProgress 的 Current/Total As Currency → vb6_BSTR_FromStr(double)
            // C2440, cZipArchive.c 2341/2343). 与 cgen_decl_func.cpp /
            // cgen_decl_proc.cpp 同步 (三处副本必须一致).
            else if (paramType == Vb6Type::Double || paramType == Vb6Type::Currency
                     || paramType == Vb6Type::Single || paramType == Vb6Type::Date) {
                knownDoubleVars_.insert(pLower);
                // Fix 175: Date 形参额外登记 —— C 层 Date/Double 同型, 只按
                // C 类型串分派会让 inferExprType 看不见 Date (打出序列号)。
                if (paramType == Vb6Type::Date) knownDateVars_.insert(pLower);
            }
            else if (paramType == Vb6Type::Long || paramType == Vb6Type::Integer || paramType == Vb6Type::Boolean
                     || paramType == Vb6Type::Byte) {
                knownLongVars_.insert(pLower);
                // ai/022 W1: 布尔形参另登记一份 (口径同 Fix 175 的 Date 形参)
                if (paramType == Vb6Type::Boolean) knownBoolVars_.insert(pLower);
                // 账 #123: Byte 形参一并进这一支 (口径同 ai/022 W1 的 Boolean —— C 型不同串
                // 就永远看不见 ⇒ 比较被当 Variant 取地址)。另登记一份到 knownByteVars_,
                // inferExprType 先判 Byte 那张表, 所以这里进 knownLongVars_ 不会把它读成 Long。
                if (paramType == Vb6Type::Byte) knownByteVars_.insert(pLower);
            }
            // Bug #2 fix: LongPtr 参数注册到独立集合
            else if (paramType == Vb6Type::LongPtr || paramType == Vb6Type::LongLong) knownLongPtrVars_.insert(pLower);   // Fix 084m
            // Fix 035: Variant 参数也要注册, 否则 `(*X) = concrete` 赋值不会触发
            // wrapVariantValue 包装, 导致 C2440 (ByRef Variant 参数写穿透场景).
            else if (paramType == Vb6Type::Variant) knownVariantVars_.insert(pLower);

            // Fix 023c: 注册 void* 参数 (As Object / As Collection / 外部 COM 类型如
            // ADODB.Recordset / Scripting.Dictionary 等) 到 knownObjectVars_ —
            // 让成员访问走 COM dispatch (vb6_ComCall / vb6_ComGet*Prop),
            // 而非直接 obj.member 字段访问, 避免 C2224 (void* 上 .member).
            // 仅当参数未被前面分支精确注册为 Class / ComClass / Interface / UDT 时
            // 才查 C 类型, 避免对 vb6_cls_* / vb6_ComIface_* 等 C 类型参数的错误
            // 注册. 与 visit(VariableDecl) line 651-656 行为一致 (局部 void* 同样注册).
            // Fix <VBFlexGridDemo>: UDT 形参**优先**登记 (口径同 cgen_decl_func.cpp 那处)。
            // knownClassVars_ 全程不清空, 别的模块 `Dim This As <工程类>` 会把同名条目
            // 泄漏过来 ⇒ UDT 形参被当类实例 ⇒ 成员访问回落 COM 后期绑定
            // vb6_ComGetObjectProp((*This), …) C2172。UDT 与类互斥, 故擦掉残留类条目。
            {
                std::string udtCTypeF = mapTypeRef(p->asType.get());
                while (!udtCTypeF.empty() && (udtCTypeF.back() == '*' || udtCTypeF.back() == ' '))
                    udtCTypeF.pop_back();
                if (udtCTypeF.compare(0, 9, "vb6_type_") == 0) {
                    knownUdtVars_[pLower] = udtCTypeF;
                    // 同 cgen_decl_func.cpp: 擦掉互斥表里的同名残留 (它们都全程不清空,
                    // 且判定分支排在 obj_dispatch 的 UDT 字段分支之前)。
                    knownClassVars_.erase(pLower);
                    knownTypedComVars_.erase(pLower);
                    knownIfaceVars_.erase(pLower);
                    knownIvrefVars_.erase(pLower);
                    knownObjectVars_.erase(pLower);
                    knownVariantVars_.erase(pLower);
                }
            }
            if (!knownClassVars_.count(pLower) && !knownTypedComVars_.count(pLower)
                && !knownIfaceVars_.count(pLower) && !knownUdtVars_.count(pLower)) {
                std::string paramCType = mapTypeRef(p->asType.get());
                if (paramCType == "void*") {
                    knownObjectVars_.insert(pLower);
                }
                // Bug #2 fix: Enum等未知类型参数, C类型为int32_t时注册为Long
                else if (paramCType == "int32_t" || paramCType == "int16_t" || paramCType == "VBABOOL") {
                    if (!knownLongVars_.count(pLower)) knownLongVars_.insert(pLower);
                }
                // Bug #2 fix: C类型为intptr_t时注册为LongPtr
                else if (paramCType == "intptr_t") {
                    if (!knownLongPtrVars_.count(pLower)) knownLongPtrVars_.insert(pLower);
                }
                // Fix 082: COM interface pointer types (vb6_ComIface_*) are pointer-sized on x64
                else if (paramCType.find("vb6_ComIface_") != std::string::npos) {
                    if (!knownLongPtrVars_.count(pLower)) knownLongPtrVars_.insert(pLower);
                }
            }
        } else if (p->asType && p->asType->kind == ASTNodeKind::ArrayTypeRef) {
            // Fix 010r-6: Register array parameters so arr(idx) generates VB6_SA_AT instead of (*arr)(idx)
            Vb6Type elemType = resolveArrayElemType(p->asType.get());
            std::string pLower = p->name;
            std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);
            knownArrays_.insert(pLower);
            arrayElemTypes_[pLower] = elemType;
            arrayDimCounts_[pLower] = 1;
            // Fix 055: 注册UDT数组元素C类型
            std::string udtCType = resolveArrayUdtElemCType(p->asType.get());
            if (!udtCType.empty()) arrayUdtElemTypes_[pLower] = udtCType;
            knownLocalVars_.insert(pLower);
        }
    }

        // P12.3: 恢复调用者的错误处理状态
    if (hasOnError_) {
        c_.emitLine("vb6_RestoreErrState();");
    }

    // Property Get: 设置返回值变量 (与Function相同语义)
    if (node.propKind == ProcKind::PropertyGet) {
        currentReturnVar_ = "vb6_ret_" + cIdent(node.name);
        if (node.returnType) {
            std::string retType = mapTypeRef(node.returnType.get());
            // Fix 160-F: 与 cgen_decl_func.cpp:168 对齐 — Property Get 与 Function
            // 的"给函数名赋值"语义相同, 但此字段从未设置, 使 Fix 092f 的
            // 按返回类型解包 Variant 分支 (cgen_assign_value_sem.inc:265) 在
            // Property Get 内永不生效:
            //   Prop Get CellFontSize As Single 内
            //     vb6_ret_CellFontSize = vb6_VariantFromComResult(vb6_ComCall(...))
            //   → VBFlexGrid.c 32210/32212 两条 C2440 (vb6_VARIANT→float)。
            currentReturnCType_ = retType;
            // 账 #116 同族 (与 cgen_decl_func.cpp 那处一字一样): 定长串返回类型的节点是
            // FixedStringTypeRef, 按 SimpleTypeRef 读 name 就是把指针当字符串 ⇒ 天文数字的分配。
            Vb6Type retVb6Type = Vb6Type::Unknown;   // 数组等复合形: '0' 就是它的空值
            if (node.returnType->kind == ASTNodeKind::SimpleTypeRef) {
                retVb6Type = typeSys_.resolveTypeName(
                    static_cast<SimpleTypeRef*>(node.returnType.get())->name);
            } else if (node.returnType->kind == ASTNodeKind::FixedStringTypeRef) {
                retVb6Type = Vb6Type::String;
            }
            // Fix 038/054: UDT 返回值不能用 = 0 初始化 (C2440), 改用 {0}
            // 修复: 仅检查 C 类型名前缀即可 (typeSys 可能将 UDT 解析为 Unknown/Variant)
            std::string initVal = defaultValue(retVb6Type);
            if (retType.rfind("vb6_type_", 0) == 0) {
                initVal = "{0}";
            }
            c_.emitLine(retType + " " + currentReturnVar_ + " = " + initVal + ";");
            // P6.11: 注册返回值变量类型 (用于BSTR安全赋值)
            // Property Get 的 Prefix = me->m_Prefix 会被替换为 vb6_ret_Prefix = me->m_Prefix
            // 如果返回类型是String, 必须使用 vb6_BSTR_Assign 确保 deep copy,
            // 否则返回浅引用会导致 COM 调用者 SysFreeString 与 me->m_Prefix 双重释放
            std::string retLower = currentReturnVar_;
            std::transform(retLower.begin(), retLower.end(), retLower.begin(), ::tolower);
            if (retVb6Type == Vb6Type::String) knownBstrVars_.insert(retLower);
            else if (retVb6Type == Vb6Type::Double) knownDoubleVars_.insert(retLower);
            else if (retVb6Type == Vb6Type::Long || retVb6Type == Vb6Type::Integer || retVb6Type == Vb6Type::Boolean) knownLongVars_.insert(retLower);
            else if (retVb6Type == Vb6Type::Variant) knownVariantVars_.insert(retLower);
        }
    }

    // Bug #1 fix (082h): 预扫描UBound/LBound(arr,N>1)收集ND数组名
    scanNDArraysInStmts(node.body);

    c_.indent();
    // P12.3: 检测On Error并声明局部错误处理
    hasOnError_ = hasOnErrorInStmts(node.body);
    procExitLabelUsed_ = false;
    if (hasOnError_) {
        c_.emitLine("jmp_buf vb6_local_err_jmp;");
        c_.emitLine("vb6_SaveErrState();");
    }
    // Fix 086: 先将块内 Dim/Const 提升到过程顶部 (VB6 局部声明是过程级作用域)
    hoistLocalDecls(node.body);
    // Fix <vbeclipse>: VB6 隐式变量 (无 Option Explicit 时未声明即使用) ——
    // 语义层登记的名字在此预声明为 Variant C 局部并注册 knownVariantVars_。
    if (currentProc_) {
        auto* impl = symTab_.implicitVarsFor(Symbol::toLower(moduleName_),
                                             Symbol::toLower(currentProc_->name));
        if (impl) {
            for (const auto& n : *impl) {
                std::string nLower = Symbol::toLower(n);
                if (!knownLocalVars_.count(nLower)) {
                    knownLocalVars_.insert(nLower);
                    knownVariantVars_.insert(nLower);
                    c_.emitLine("vb6_VARIANT " + cIdent(n) + " = vb6_VariantEmpty();  /* 隐式变量 */");
                }
            }
        }
    }
    // 账 #179: 控件代码运行在自己的宿主上下文里 (与下面的 PopInstance 成对)。
    if (ucCtxScoped()) c_.emitLine("vb6_UC_PushInstance((void*)me);");
    emitStmtList(node.body);

    // Fix <vbeclipse>: Property 的统一出口 + **缺失的错误状态恢复**。
    // 原先 Property 只在中段发 vb6_SaveErrState() 而**从不** RestoreErrState,
    // 于是每次调用都泄漏一层错误状态, 且 vb6_error_jmp_ptr 停留在本 Property
    // 已失效的 vb6_local_err_jmp 上; 之后任何 vb6_ErrRaise 都会 longjmp 进死帧
    // (野读崩)。Exit Property 现在发 `goto vb6_proc_exit;` 落到这里。
    if (procExitLabelUsed_) {
        c_.emitLine("vb6_proc_exit:;");
    }
    if (hasOnError_) {
        c_.emitLine("vb6_RestoreErrState();");
    }

    // M22: 释放ANSI临时变量
    for (auto& ansiVar : ansiTempsToFree_) {
        c_.emitLine("vb6_FreeANSI(" + ansiVar + ");");
    }
    ansiTempsToFree_.clear();
    ansiOutParams_.clear();

    // tB Interface B05: 接口变量持有引用, 正常出口处经槽 Release
    emitIvrefScopeRelease();

    // 账 #179: 与体首 PushInstance 成对 (Property Get/Let/Set 三条路都落到这里)
    if (ucCtxScoped()) c_.emitLine("vb6_UC_PopInstance();");

    // Property Get: 隐式返回 vb6_ret_<propName>
    if (node.propKind == ProcKind::PropertyGet && node.returnType) {
        c_.emitLine("return " + currentReturnVar_ + ";");
    } else if (procExitLabelUsed_) {
        // Property Let/Set: 无返回值。给 vb6_proc_exit 一个后继语句 (标签后必须有语句)。
        c_.emitLine("return;");
    }
    c_.dedent();

    // Fix <vbeclipse>: 复原过程内被 #undef 的 Win32 宏 (隐式变量名隔离)
    for (const auto& n : implicitMacroNames_) {
        c_.emitLine("#pragma pop_macro(\"" + n + "\")");
    }
    implicitMacroNames_.clear();
    // 清理返回值变量和currentProc_
    if (node.propKind == ProcKind::PropertyGet) {
        currentReturnVar_ = "";
    }
    currentProc_ = nullptr;
    hasOnError_ = false;
    procExitLabelUsed_ = false;

    c_.emitLine("}");
    c_.emitBlank();
}

void CCodeGen::visit(EventDecl& node) {
    // P6.5: Event声明 → 回调函数指针typedef在emitEventSink中统一生成
    // 此处仅生成注释标记
    c_.emitLine("/* Event " + node.name + " — callback typedef in event sink table */");
}

// ============================================================
// Property签名生成
// ============================================================

std::string CCodeGen::makePropertySignature(PropertyDecl& node) {
    // Property Get/Let/Set使用不同前缀: prop_get_/prop_let_/prop_set_
    // 避免同名Property在C层面链接冲突
    std::string prefix;
    switch (node.propKind) {
        case ProcKind::PropertyGet:  prefix = "prop_get_"; break;
        case ProcKind::PropertyLet:  prefix = "prop_let_"; break;
        case ProcKind::PropertySet:  prefix = "prop_set_"; break;
        default:                     prefix = "prop_get_"; break;
    }
    // 类模块属性始终带 vb6_<ClassName>_ 前缀 (与 dll_entry.c / resolveClassMemberCall 调用一致)
    std::string propName = cProcName(prefix + node.name, node.access, isClassModule_ ? moduleName_ : "");
    std::string params;

    // 类模块: 第一个参数为 me 指针
    if (isClassModule_) {
        params = classMeParam();
        if (!node.params.empty()) params += ", ";
    }

    // 空参数列表: 类模块已有me参数时不需要"void"
    if (node.params.empty() && isClassModule_) {
        // params已经有me, 不追加
    } else {
        params += makeParamList(node.params);
    }

    switch (node.propKind) {
        case ProcKind::PropertyGet: {
            std::string retType = node.returnType ? mapTypeRef(node.returnType.get()) : "vb6_VARIANT";
            return retType + " " + propName + "(" + params + ")";
        }
        case ProcKind::PropertyLet: {
            // VB6: Property Let Name(v) — 最后一个参数v就是赋值值
            // 不追加额外的vb6_let_value参数
            return "void " + propName + "(" + params + ")";
        }
        case ProcKind::PropertySet: {
            // VB6: Property Set Name(v) — 最后一个参数v就是对象引用
            return "void " + propName + "(" + params + ")";
        }
        default:
            return "void " + propName + "(" + params + ")";
    }
}

// ============================================================
// 接口类成员的默认实现 (Fix <vbeclipse>)
// ============================================================

// Fix <vbeclipse>: VB6 接口 (`Attribute VB_Exposed = True`) 的成员声明没有实现体,
// 但 `Public m_Scheme As New IScheme` (modPublic.bas:12) 会**实例化接口本身** ——
// VB6 语义下, 接口的 `Public Property Get X()` 就是它的默认实现. 生成端此前按
// "接口成员是抽象的" 跳过 (P6.4), 前提是接口不可实例化; `As New` 打破该前提:
// 调用方 (ucTab/ucButton/ucPerspective) 生成的 `vb6_IScheme_prop_get_BackColor(m_Scheme)`
// 既无原型也无实体 → 链接期 LNK2001 "无法解析的外部符号" ×26.
//
// 这里补默认实现: Getter/Function 返回类型零值, Sub/Property Let/Set 空操作.
// 真实实现由实现类 (SchemeWinXP 等, 经 vb6_ivtbl_IScheme_for_<C> 槽表) 提供, 两者
// 名字不同 (vb6_SchemeWinXP_prop_get_BackColor vs vb6_IScheme_prop_get_BackColor),
// 不冲突 —— 走的是"接口自带默认实现"这条独立通道, 与实现类无关.
// 与 vb6rtl_userctl.h:56 既有口径一致: 无容器/无实现时返回空值即可满足编译链接.
//
// Fix <vbeclipse>: 上面那条"与实现类无关"是**错的** —— 调用点
// `vb6_IScheme_prop_get_BackColor(m_Scheme)` 传进来的 `me` 通常**就是**某个实现类
// 实例 (ucPerspective 里 `Set m_Scheme = New SchemeWinXP`, 全局共享同一个变量)。
// 桩若恒返回零值, `m_Scheme.BackColor` / `.EditorAreaBackColor` / `.FrameColor` 一律
// 读到 0 = 纯黑, 而 SchemeWinXP 里写的是 RGB(145,155,156) 这类值。实测 play78:
// 24 次 SetBackColor 里 17 次 color=000000, 左侧停靠面板 (UC宿主/ViewArea/Edit)
// 整片纯黑 99%, 右边同类控件却是白的 —— 差别只在配色读到 0 还是读到真值。
//
// VB6 语义: 接口的 `Property Get` 既是"接口自带默认实现", 也是实现类覆写的**入口**
// —— 通过接口引用调用时必须动态分派到实际实现类。所以桩要先问"me 是谁",
// 是实现类实例就转调它的槽; 否则才用零值(接口自身实例 / Nothing)。
//
// 怎么问: 不能读 `me->__iv_<Iface>` —— 接口自身实例 (vb6_cls_IScheme 只有
// __comObj + _placeholder) 根本没有那个字段, 读它会越界。也不能用
// `vb6_iv_from_iv_<Cls>`: 它收的是**薄指针** (&me->__iv_<Iface>, 首字段才是 vt),
// 而这里拿到的是**实例指针** (首字段是 __comObj), 喂进去永远比不中槽表地址。
// 用每个实现类新增的 `vb6_iv_thin_of_<Cls>`: 从实例出发按 `me->__iv_<I>.vt` 认类,
// 非本类实例返回 NULL。
void CCodeGen::emitIfaceMemberStub(const SubDecl&, const std::string& sig, bool /*returnsValue*/) {
    c_.emitLine(sig + " {");
    c_.emitLine("    (void)me;");
    c_.emitLine("}");
    c_.emitBlank();
}

void CCodeGen::emitIfaceMemberStub(const FunctionDecl& fn, const std::string& sig, bool) {
    const std::string ret = fn.returnType ? mapTypeRef(fn.returnType.get()) : "vb6_VARIANT";
    // Function 的槽键就是裸成员名 (小写), 不带 get_/put_ 前缀 (见 interfaces_registry.hpp
    // 的槽键规范)。有形参则不分派 —— 见 emitIfaceStubBody 里的理由。
    const bool takesArgs = !fn.params.empty();
    emitIfaceStubBody(sig, ret, ifaceLower(fn.name), takesArgs);
}

void CCodeGen::emitIfaceMemberStub(const PropertyDecl& pn, const std::string& sig, bool) {
    const char* prefix = pn.propKind == ProcKind::PropertyGet ? "get_"
                        : pn.propKind == ProcKind::PropertyLet ? "put_"
                                                                : "putref_";
    const std::string slot = std::string(prefix) + ifaceLower(pn.name);
    // Property Get 的形参在 VB6 里只能是索引 (Variant/变参), 本次不分派;
    // Let/Set 恒有一个 New_Value 形参, 同样不分派 (落零值/空操作, 与修复前一致)。
    const bool takesArgs = pn.propKind != ProcKind::PropertyGet;
    if (pn.propKind != ProcKind::PropertyGet) {  // Let/Set 无返回值 → 分派后空操作
        emitIfaceStubBody(sig, "void", slot, takesArgs);
        return;
    }
    const std::string ret = pn.returnType ? mapTypeRef(pn.returnType.get()) : "vb6_VARIANT";
    emitIfaceStubBody(sig, ret, slot, takesArgs);
}

// 返回类型零值: 标量/指针统一 `return 0;` (C 里 0 是合法空指针常量);
// 结构体 (Variant 等) 不能用 0 → 零初始化复合字面量.
// Fix <vbeclipse>: 零值不再是唯一出路 —— 先按"me 是哪个实现类的实例"分派到它的槽,
// 全部实现类都不是它 (接口自身实例 / Nothing) 才落回零值。
//
// 生成的形状 (以 IScheme.BackColor / 实现类 SchemeWinXP 为例):
//     vb6_ivref_IScheme* h_ = NULL;
//     void* t_ = vb6_iv_thin_of_SchemeWinXP((void*)me);
//     if (t_) h_ = (vb6_ivref_IScheme*)t_;
//     if (h_ && h_->vt && h_->vt->get_backcolor) return h_->vt->get_backcolor(h_);
//     return 0;
// 三个守卫缺一不可: thin_of_* 对非本类实例返回 NULL; 槽位可能是 NULL 占位 (契约
// 缺失时发码留的 NULL, 见 cgen_iface_vtbl.cpp 的 slotFns.push_back("NULL"))。
void CCodeGen::emitIfaceStubBody(const std::string& sig, const std::string& retType,
                                 const std::string& slotKey, bool takesArgs) {
    const bool isVoid = (retType == "void");
    const bool isStruct = (retType == "vb6_VARIANT" || retType.compare(0, 7, "struct ") == 0);
    const std::string ifaceId = cIdent(moduleName_);

    c_.emitLine(sig + " {");
    c_.emitLine("    (void)me;");

    // 实现类名单由 driver 预扫描注入 (接口小写名 → 全部实现类)。空 = 该接口没有
    // 实现类 (或表未注入) → 直接落零值, 与本次修复前的行为完全一致。
    //
    // 只对**无参**槽分派: 槽签名除 self 外还有形参时, 转发要按槽声明逐个重建实参
    // (ByRef 指针、Optional 的 _has_ 尾标记), 那是 ivForwardArgs 的活。本次针对的
    // 是 Property Get (IScheme 全是这类 —— 实测 m_Scheme.BackColor /
    // .EditorAreaBackColor / .FrameColor 都读成 0), 它们都无参; 有参槽仍走零值,
    // 与修复前一致, 不引入新的错误转发。
    const std::vector<std::string> impls =
        takesArgs ? std::vector<std::string>() : allImplementationClasses(moduleName_);
    if (!slotKey.empty() && !impls.empty()) {
        // 槽键 → C 标识符 (槽表字段名, 如 get_backcolor)。契约里没有这一席就不分派:
        // 编出来的 vt->get_xxx 名字不存在, 会在链接期炸。
        const IfaceView* iv = ivLookupIface(moduleName_);
        std::string field;
        if (iv) {
            const std::string k = ivSlotKeyForMember(*iv, slotKey);
            if (!k.empty()) field = cIdent(k);
        }
        if (!field.empty()) {
            c_.emitLine("    {");
            c_.emitLine("        vb6_ivref_" + ifaceId + "* h_ = NULL;");
            for (const std::string& cls : impls) {
                c_.emitLine("        {");
                c_.emitLine("            void* t_ = vb6_iv_thin_of_" + cIdent(cls) + "((void*)me);");
                c_.emitLine("            if (t_ && !h_) h_ = (vb6_ivref_" + ifaceId + "*)t_;");
                c_.emitLine("        }");
            }
            c_.emitLine("        if (h_ && h_->vt && h_->vt->" + field + ") {");
            if (isVoid) {
                c_.emitLine("            h_->vt->" + field + "(h_);");
                c_.emitLine("            return;");
            } else {
                c_.emitLine("            return h_->vt->" + field + "(h_);");
            }
            c_.emitLine("        }");
            c_.emitLine("    }");
        }
    }

    if (!isVoid) {
        if (isStruct) {
            c_.emitLine("    " + retType + " vb6_iface_stub_z_ = {0};");
            c_.emitLine("    return vb6_iface_stub_z_;");
        } else {
            c_.emitLine("    return 0;");
        }
    }
    c_.emitLine("}");
    c_.emitBlank();
}

// ============================================================
// 类工厂函数生成
// ============================================================

std::string CCodeGen::classMeParam() const {
    std::string clsStruct = "vb6_cls_" + cIdent(moduleName_);  // Fix 013: 用 moduleName_ (VB_Name)
    return clsStruct + "* me";
}

// Fix 019: 在事件包装函数体内, handler (void*) 需要被强制转换为类指针类型
// 以正确调用 vb6_<Mod>_<Handler>(cls* me, ...) 形式的类方法。
// 历史上这里误用了 classMeParam() ("vb6_cls_X* me" — 参数声明) 作为函数调用实参,
// 导致生成伪 C: vb6_<Mod>_<Hand>(vb6_cls_X* me/* from handler */, ...) 报 C2065 'me' 未声明。
std::string CCodeGen::classHandlerCast() const {
    std::string clsStruct = "vb6_cls_" + cIdent(moduleName_);
    return "((" + clsStruct + "*)handler)";
}


} // namespace vb6c3
