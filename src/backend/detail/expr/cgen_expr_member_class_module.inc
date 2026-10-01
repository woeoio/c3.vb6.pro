// cgen_expr_member_class_module.inc - 函数体片段：类实例成员分派与模块名.方法、模块前缀解析
// 由 src/backend/expr/cgen_expr_member.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 553~774 行，逐行未改。

            // 优先级2: 类实例成员访问: obj.Method → vb6_ClassName_MethodName(obj)
            // Fix 010r-10: 使用map查找, 可获取类名用于方法分发
            // Fix 011r-1: 当obj在knownClassVars_中时, 优先用resolveClassMemberCall
            // 精确解析该类的方法/属性, 避免memSym捡错模块的同类同名方法/属性
            // P6.4+: 类模块默认实例 (VB_PredeclaredId=True) 的裸类名 (cTT.CreateToolTip
            // / cTT.EnableTooltips(...) 在 frm 里) — 注册为"类实例变量"并按类精确
            // 解析成员 (注入 me 首参与类型化形参), 对象表达式用 vb6_cls_X_Default().
            std::string defaultInstCls = registerDefaultInstanceClass(objLower);
            auto itClassVar = knownClassVars_.find(objLower);
            if (itClassVar == knownClassVars_.end() && node.object) {
                // Fix 090al: 链式对象 (Db.Sql(s).Param(p).QueryParam) 或局部类变量
                // 未注册 knownClassVars_ → AST 推断类名再分发 (含值上下文补参),
                // 否则无括号裸调用只发 this → C2198 参数太少.
                std::string inferredCls = inferClassTypeOfExpr(*node.object);
                if (!inferredCls.empty()
                    && !resolveClassMemberCall(inferredCls, node.memberName).empty()) {
                    itClassVar = knownClassVars_.emplace(objLower, inferredCls).first;
                }
            }
            if (itClassVar != knownClassVars_.end()) {
                // 用对象的真实类名查找方法, 防止跨模块同名冲突
                std::string resolvedFn = resolveClassMemberCall(itClassVar->second, node.memberName);
                if (!resolvedFn.empty()) {
                    std::string objExpr;
                    if (!defaultInstCls.empty()) {
                        objExpr = "vb6_cls_" + cIdent(defaultInstCls) + "_Default()";
                    } else {
                        // Fix 180: 接收者是值上下文, 必须在值语境发射. 外层 IndexOrCallExpr
                        // / CallStmt 会把 asCallCallee_ 置 true 后才 emit 本 MemberAccessExpr,
                        // 若直接泄漏到接收者求值, 接收者是 PropertyGet (如 cTlsRemaster 内
                        // pvSocket.Bind(...) 的 pvSocket) 时 IdentifierExpr 只发裸函数名
                        // (本意"我是callee, 外层补参"), 外层却把它当对象表达式前置 →
                        // 生成 vb6_cTlsSocket_Bind(vb6_cTlsReMaster_prop_get_pvSocket, ...)
                        // 把函数指针当实例传 → 运行期解 vtable 0xC0000005 (VBMAN CI 实证).
                        bool savedObjCtx180 = asCallCallee_;
                        asCallCallee_ = false;
                        emitExpr(*node.object);
                        asCallCallee_ = savedObjCtx180;
                        objExpr = std::move(lastExpr_);
                    }
                    // 类虚表 (tB, ai/022 B08d): 静态类型是"链上有虚槽的类" → 改按 __cvtbl 索引。
                    // 覆盖 `Me.M()` 与 `obj.M()` 两种写法, 后者才真正修掉"基类型变量持有派生实例
                    // 却静默绑基类实现"那个 B07b 遗留的切片问题。directFn 留着给下面的 _prop_ 判据。
                    const std::string directFn = resolvedFn;
                    const std::string dispatchFn = virtDispatchCallee(
                        itClassVar->second, node.memberName, objExpr,
                        /*mustDispatch=*/defaultInstCls.empty(), node.loc);
                    if (!dispatchFn.empty()) resolvedFn = dispatchFn;
                    // Fix 089h: 值上下文无括号方法引用 (Trim(FileStream.ReadLine)
                    // → vb6_cToolsStream_ReadLine(me->FileStream) 只发 this) —
                    // C2198 参数太少. 补默认参数 (Optional/必选), 与链式值上下文
                    // / With 分支 (Fix 044a) 一致. asCallCallee_ 由外层补参.
                    // 属性 (prop_get_/prop_let_/prop_set_) 不 pad — findClassMemberCallParams
                    // 对 Get+Let 并存属性返回 Let 参数, 无括号属性读只发 me 即正确
                    // (prop_get_State(oClient)), pad 反而 C2197 参数过多.
                    if (!asCallCallee_
                        && directFn.find("_prop_") == std::string::npos) {
                        std::vector<ParameterInfo> params89h;
                        bool isBuiltin89h = false;
                        if (findClassMemberCallParams(itClassVar->second, node.memberName,
                                                       params89h, isBuiltin89h)
                            && !params89h.empty() && !isBuiltin89h) {
                            std::string argList89h = objExpr;
                            for (size_t i = 0; i < params89h.size(); i++) {
                                const auto& param = params89h[i];
                                argList89h += ", ";
                                std::string defVal89h;
                                if (param.hasDefaultValue && !param.defaultValueExpr.empty()) {
                                    defVal89h = param.defaultValueExpr;
                                } else {
                                    defVal89h = defaultValue(param.type);
                                }
                                if (param.isByVal) {
                                    argList89h += defVal89h;
                                } else {
                                    std::string cType89h = mapType(param.type);
                                    if (param.type == Vb6Type::Variant
                                        || param.type == Vb6Type::Empty
                                        || param.type == Vb6Type::Null
                                        || param.type == Vb6Type::Object) {
                                        argList89h += "&(" + cType89h + "){0}";
                                    } else {
                                        argList89h += "&(" + cType89h + "){" + defVal89h + "}";
                                    }
                                }
                            }
                            for (size_t i = 0; i < params89h.size(); i++) {
                                const auto& param = params89h[i];
                                if (param.isOptional && !param.isParamArray) {
                                    argList89h += ", 0";
                                }
                            }
                            lastExpr_ = resolvedFn + "(" + argList89h + ")";
                        } else {
                            lastExpr_ = resolvedFn + "(" + objExpr + ")";
                        }
                    } else {
                        lastExpr_ = resolvedFn + "(" + objExpr + ")";
                    }
                    return;
                }
                // resolveClassMemberCall返回空 → 该类中无此方法/属性 → 数据字段访问
                // 落入下方cross-module fallback分支处理 obj->member
            }

            // 优先级3: 模块名.方法名: MathUtils.Add → vb6_MathUtils_Add
            //    object名称不是已知变量, 但成员是函数 → 视为模块限定调用
            bool isVarName = false;
            auto* objSym = symTab_.lookup(objIdent.name);
            if (!objSym) objSym = symTab_.lookupModule(objIdent.name);
            if (objSym && (objSym->kind == SymbolKind::Variable || objSym->kind == SymbolKind::Parameter)) {
                isVarName = true;
            }
            // Fix 084x: 补充检查cgen层跟踪集合 (与M22分支一致), 防止过程级
            // 局部类变量/UDT变量被误判为"模块名", 落入优先级3 "模块名.方法名"
            // 用 memSym->sourceModule 生成 vb6_<他类>_<member> 函数名 (C2065).
            // 例: oCallback.Socket.Accept 中 oCallback 是 Dim As cClientCallback 局部
            // 变量, 不在模块作用域符号表 → 原判定 isVarName=false, memSym 命中
            // cTlsSocket.Socket 属性 → 生成 vb6_cTlsSocket_Socket (错误).
            if (!isVarName && knownUdtVars_.count(objLower)) isVarName = true;
            if (!isVarName && knownClassVars_.find(objLower) != knownClassVars_.end()) isVarName = true;
            if (!isVarName && knownNewVars_.count(objLower)) isVarName = true;
            if (!isVarName && knownObjectVars_.count(objLower)) isVarName = true;
            if (!isVarName && knownTypedComVars_.count(objLower)) isVarName = true;
            if (!isVarName && knownIfaceVars_.count(objLower)) isVarName = true;

            // Fix 083c: obj 是当前模块的属性(返回类对象) → pvSocket.GetLocalHost(...)
            // 不能按"模块名.方法"处理(会丢失对象实例导致 me 参数缺失 C2198),
            // 生成: vb6_<ObjClass>_<Method>((void*)vb6_<Mod>_prop_get_pvSocket(me), args...)
            if (!isVarName && objSym && objSym->kind == SymbolKind::PropertyGet &&
                !objSym->variableTypeName.empty()) {
                std::string propMod = objSym->isExternal ? objSym->sourceModule
                                    : (isClassModule_ ? moduleName_ : "");
                if (!propMod.empty()) {  // 仅类模块属性带 me
                    std::string propClass = objSym->variableTypeName;
                    std::string resolvedFn = resolveClassMemberCall(propClass, node.memberName);
                    if (!resolvedFn.empty()) {
                        std::string propFn = cProcName("prop_get_" + objIdent.name, objSym->access, propMod);
                        std::string objExpr = propFn + "((void*)me)";
                        lastExpr_ = resolvedFn + "((void*)" + objExpr + ")";
                        return;
                    }
                }
            }

            // Fix 086: 成员命中本模块符号但限定符是另一个模块名时, 优先在该模块
            // 中解析成员常量 (modSerialPortAPI.SetDTR — 本模块恰有同名 Sub SetDTR,
            // 误生成 vb6_modSerialPortAPI_SetDTR → C2065; 实际取的是模块常量 SETDTR=5).
            if (!isVarName && !memSym->isExternal) {
                std::string qModLower = Symbol::toLower(objIdent.name);
                std::string curModLower86 = Symbol::toLower(moduleName_);
                if (!qModLower.empty() && qModLower != curModLower86 && symTab_.moduleScope()) {
                    const std::string want86 = Symbol::toLower(node.memberName);
                    bool qResolved = false;
                    for (const auto& [qkey, qsym] : symTab_.moduleScope()->symbols()) {
                        if (!qsym || !qsym->isExternal) continue;
                        if (qsym->lowerName != want86 && Symbol::toLower(qsym->name) != want86) continue;
                        if (Symbol::toLower(qsym->sourceModule) != qModLower) continue;
                        if (qsym->kind == SymbolKind::Constant && qsym->hasConstValue) {
                            if (qsym->constType == Vb6Type::Long || qsym->constType == Vb6Type::Integer
                                || qsym->constType == Vb6Type::Boolean || qsym->constType == Vb6Type::Byte
                                || qsym->constType == Vb6Type::Error) {
                                lastExpr_ = std::to_string(qsym->constIntValue);
                                return;
                            }
                            if (qsym->constType == Vb6Type::Single || qsym->constType == Vb6Type::Double) {
                                lastExpr_ = std::to_string(qsym->constFloatValue);
                                return;
                            }
                        }
                        break;
                    }
                    // 注入被本地同名符号阻止时, 查 driver 预扫描的模块公共常量表
                    if (!qResolved && modulePublicConsts_) {
                        auto itMod = modulePublicConsts_->find(qModLower);
                        if (itMod != modulePublicConsts_->end()) {
                            auto itVal = itMod->second.find(want86);
                            if (itVal != itMod->second.end()) {
                                lastExpr_ = std::to_string(itVal->second);
                                return;
                            }
                        }
                    }
                }
            }

            if (!isVarName) {
                // Fix 159-F: 对象是**过程级变量** (symTab_ 只看模块作用域, 形参/局部
                // 不可见 → isVarName 误判为 false) 而成员是 hWnd 时, 直接取该变量的值.
                // 运行时约定"窗体/控件对象值本身就是它的 HWND"
                //   (vb6_Forms_Register(void* hwnd)、Me→vb6_hwnd_MainForm、
                //    控件实例经 vb6_UC_InstanceOf(vb6_hwnd_X) 反查),
                // 故 VB6 的 Form.hWnd / RefForm.hWnd / CurrControl.hWnd 即该变量本身.
                // 此前落到下面的"模块.成员"分支: memSym 命中 VBFlexGrid.ctl 的 hWnd 属性
                // → 发出 vb6_VBFlexGrid_hWnd 全局 → C2065×17
                // (Common.bas:1172/1175…, VisualStyles.bas:230/236/239…).
                // 类实例变量 (knownClassVars_) 不介入 —— 它的 hWnd 应走上面的类成员分派.
                {
                    const std::string memL159F = Symbol::toLower(node.memberName);
                    const std::string objL159F = Symbol::toLower(objIdent.name);
                    bool objIsProcVar159F =
                        !knownClassVars_.count(objL159F)
                        && (knownLongVars_.count(objL159F)
                            || knownObjectVars_.count(objL159F)
                            || knownVariantVars_.count(objL159F)
                            || knownLongPtrVars_.count(objL159F));
                    if (memL159F == "hwnd" && objIsProcVar159F) {
                        emitExpr(*node.object);
                        std::string objE159F = std::move(lastExpr_);
                        lastExpr_ = "(intptr_t)(" + objE159F + ")";
                        if (std::getenv("C3_DBG159")) {
                            std::fprintf(stderr,
                                "[DBG159] 159F %s.hWnd -> %s (mod=%s)\n",
                                objIdent.name.c_str(), lastExpr_.c_str(),
                                moduleName_.c_str());
                        }
                        return;
                    }
                }
                // 确定函数名的模块前缀
                std::string sourceMod;
                // Fix 092v: `Common.Version()` — 限定符是**模块名**时以限定模块为准.
                // Common4DLL.bas 与 cVBMAN.cls 都有 Version: lookupModule("Version")
                // 命中 cVBMAN.Version → 前缀取 memSym->sourceModule = "cVBMAN" →
                // vb6_cVBMAN_Version(...) 少了 me 参数 → C2198
                // (cHttpServerResponse.cls 440; 实参已按 Common 版形参正确生成
                // vb6_Common_Version(void**, int)). 这里改用限定符对应的规范模块名.
                std::string modCanon092v;
                {
                    const std::string objL092v = Symbol::toLower(objIdent.name);
                    if (symTab_.moduleScope()) {
                        for (const auto& [k092v, s092v] : symTab_.moduleScope()->symbols()) {
                            if (!s092v || !s092v->isExternal) continue;
                            if (Symbol::toLower(s092v->sourceModule) == objL092v) {
                                modCanon092v = s092v->sourceModule;
                                break;
                            }
                        }
                    }
                    // Fix 092y: 兜底 — 限定符本身就是工程内的**标准模块**名时, 直接
                    // 采用其规范名. 当该模块的公开过程与其它类的成员**全部**同键冲突时
                    // (Common 的 Version/Path 与 cVBMAN 的 Version/Path 同键,
                    // storageKey(Function)=lowerName → defineExternal 先到先得, Common
                    // 的两个成员全被丢弃), 本模块作用域内不存在任何
                    // sourceModule=="Common" 的符号 → 上面的扫描落空 → sourceMod 误取
                    // memSym->sourceModule ("cVBMAN") → vb6_cVBMAN_Version 少 me 参数 →
                    // C2198 (cHttpServerResponse.cls 440 Common.Version()).
                    // 仅对非类模块生效, 不改变 类名.成员 / 窗体名.成员 的现有语义.
                    if (modCanon092v.empty()) {
                        for (const auto& extMod092y : externalModules_) {
                            if (Symbol::toLower(extMod092y) != objL092v) continue;
                            Symbol* extSym092y = symTab_.lookup(extMod092y);
                            if (extSym092y && extSym092y->kind == SymbolKind::Class) break;
                            modCanon092v = extMod092y;
                            break;
                        }
                    }
                }
                if (!modCanon092v.empty()) {
                    sourceMod = (Symbol::toLower(modCanon092v) == Symbol::toLower(moduleName_))
                                ? "" : modCanon092v;
                } else if (memSym->isExternal) {
                    sourceMod = memSym->sourceModule;
                } else {
                    // 同模块调用也允许 Module.Method 语法
                    // 检查object名是否匹配当前模块名
                    std::string modLower = moduleName_;
                    std::transform(modLower.begin(), modLower.end(), modLower.begin(), ::tolower);
                    if (objLower == modLower) {
                        sourceMod = "";  // 同模块, 不需要前缀
                    } else {
                        sourceMod = objIdent.name;  // 假设object名就是模块名
                    }
                }
                // Fix <vbeclipse> rev7: 模块限定路径必须按 **memSym->kind** 决定
                // prop_get_/prop_let_/prop_set_ 前缀, 不能一律用裸成员名。
                //
                // 为什么: 这条路径的前提是"obj 查不到局部变量 ⇒ 它是模块名"。但
                // **过程级局部类变量**恰好两边都查不到 (symTab_ 只有模块作用域,
                // knownClassVars_ 登记在 emit 声明处, 但本行在它之前判过一轮) ——
                // 于是 `l_View.View.hWnd` (l_View 是 `Dim l_View As View`) 里的
                // `.View` 段被当成"View 模块的 View 成员", memSym 命中 View.cls 的
                // `Property Get View() As Object`(kind=PropertyGet), 却不加前缀, 发出
                // `vb6_View_View` → C2065 "未声明的标识符" + C2224 ".hWnd 左侧必须是
                // 结构/联合" (实测 SubClass.c / ucPerspective.c / modSubClass 全线, 200+ 错误)。
                // 属性访问函数在 C 侧一律叫 `vb6_<Cls>_prop_get_<Name>` (见
                // cgen_expr_with.cpp ClassInstance 分支同款判定), 少前缀必不匹配。
                std::string memberProcName = node.memberName;
                if (memSym->kind == SymbolKind::PropertyGet)      memberProcName = "prop_get_" + node.memberName;
                else if (memSym->kind == SymbolKind::PropertyLet) memberProcName = "prop_let_" + node.memberName;
                else if (memSym->kind == SymbolKind::PropertySet) memberProcName = "prop_set_" + node.memberName;
                std::string funcName = cProcName(memberProcName, memSym->access, sourceMod);
                lastExpr_ = funcName;  // 仅输出函数名, 参数由IndexOrCallExpr添加
                return;
            }
        }
    }
