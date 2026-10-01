// cgen_base_generate_body_pass.inc - generate 函数体片段：第二遍 = 过程体与类模块工厂函数
// 由 src/backend/cgen_base.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 673~795 行，逐行未改。

    // === 第二遍: 过程体 → .c ===
    c_.emitBlank();
    c_.emitLine("// === 过程实现 ===");
    c_.emitBlank();

    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::SubDecl) {
            auto& sub = static_cast<SubDecl&>(*decl);
            if (currentClassIsInterface) { emitIfaceMemberStub(sub, makeProcSignature(sub), false); continue; }
            visit(sub);
        } else if (decl->kind == ASTNodeKind::FunctionDecl) {
            auto& func = static_cast<FunctionDecl&>(*decl);
            if (currentClassIsInterface) { emitIfaceMemberStub(func, makeProcSignature(func), true); continue; }
            visit(func);
        } else if (decl->kind == ASTNodeKind::PropertyDecl) {
            auto& prop = static_cast<PropertyDecl&>(*decl);
            if (currentClassIsInterface) { emitIfaceMemberStub(prop, makePropertySignature(prop), prop.propKind == ProcKind::PropertyGet); continue; }
            visit(prop);
        }
    }

    // === Inherits (tB B07b): 祖先成员的转发桩定义 (排在过程体之后、工厂函数之前) ===
    if (isClassModule_ && !currentClassIsInterface) emitInheritedProcDefs(module);

    // === 类模块: 生成工厂函数 ===
    if (isClassModule_ && !currentClassIsInterface) {  // P6.4: 接口类不生成工厂/析构函数
        emitClassFactory(module);
        emitClassDefaultInstance(module);  // P6.4+: VB_PredeclaredId=True 默认实例单例
        // Fix 099: Public 字段的 COM 读写访问器 (必须生成在类模块自身 .c:
        // dll_entry.c 只有结构体前向声明, 不完整类型无法访问字段)
        emitClassFieldAccessors(module);
    }

    // === P6.5: 生成事件处理器包装函数 ===
    // 查找 WithEvents 变量，为 obj_EventName 形式的处理器生成C回调包装
    if (!knownWithEventsVars_.empty()) {
        c_.emitBlank();
        c_.emitLine("// === P6.5: Event handler wrappers ===");
        for (auto& [varLower, srcClassName] : knownWithEventsVars_) {
            auto* srcClsSym = symTab_.lookup(srcClassName);
            if (!srcClsSym || srcClsSym->kind != SymbolKind::Class) continue;

            for (auto& evtName : srcClsSym->eventNames) {
                std::string handlerName;
                // 查找原始变量名（保留大小写）
                std::string varName = varLower;
                // 从符号表查找保留大小写的变量名 (用lookup跨模块查找)
                auto* varSym = symTab_.lookup(varLower);
                if (varSym) varName = varSym->name;
                handlerName = varName + "_" + evtName;

                auto* handlerSym = symTab_.lookup(handlerName);
                if (!handlerSym) continue;

                // 生成包装函数: void vb6_evt_wrap_<var>_<evt>(void* handler, params...)
                std::string wrapperName = evtWrapperName(varLower, evtName);
                std::string sinkName = "vb6_events_" + cIdent(srcClassName);

                // 查找Event声明的参数
                std::vector<ParameterInfo> evtParams;
                std::vector<std::string> evtParamCTypes;
                std::vector<std::string> evtParamCTypeOnly;
                std::vector<bool> evtParamIsClass;
                for (auto& decl2 : module.declarations) {
                    // 事件处理器参数与源类Event声明参数相同
                    // 从处理器Sub的参数获取
                    if (decl2->kind == ASTNodeKind::SubDecl) {
                        auto& sub = static_cast<SubDecl&>(*decl2);
                        if (Symbol::toLower(sub.name) == Symbol::toLower(handlerName)) {
                            for (auto& param : sub.params) {
                                ParameterInfo pi;
                                pi.name = param->name;
                                pi.type = (param->asType && param->asType->kind == ASTNodeKind::SimpleTypeRef)
                                    ? typeSys_.resolveTypeName(static_cast<SimpleTypeRef*>(param->asType.get())->name)
                                    : Vb6Type::Variant;
                                if (pi.type == Vb6Type::Unknown || pi.type == Vb6Type::Empty)
                                    pi.type = Vb6Type::Variant;
                                pi.isByVal = true;
                                evtParams.push_back(pi);
                                // Fix 084k: 类名参数用 mapTypeRef → vb6_cls_X* (同 .h 前向声明),
                                // 不能 resolveTypeName 解析失败回退为 vb6_VARIANT
                                std::string wType = mapTypeRef(param->asType.get());
                                evtParamCTypes.push_back(wType + " " + cIdent(param->name));
                                evtParamCTypeOnly.push_back(wType);
                                evtParamIsClass.push_back(wType.find("vb6_cls_") == 0);
                            }
                            break;
                        }
                    }
                }

                // 构建包装函数签名
                std::string sig = "void " + wrapperName + "(void* handler";
                for (auto& pt : evtParamCTypes) {
                    sig += ", " + pt;
                }
                sig += ")";

                // .h前向声明已在上方（#endif之前）输出，此处只生成.c实现

                // 生成.c实现
                c_.emitLine(sig + " {");
                c_.indent();
                // 调用实际的VB6处理器函数
                // 如果是类模块，处理器是 vb6_<Mod>_<HandlerName>(me, params...)
                // 如果是标准模块，处理器是 vb6_<HandlerName>(params...)
                // Fix 019: 类模块必须将 handler (void*) 转换为 cls* 再作为 me 传入,
                //          不可使用 classMeParam() (那是参数声明)
                std::string procCall = cProcName(handlerName, handlerSym->access, isClassModule_ ? moduleName_ : "");
                std::string callArgs = "(";
                if (isClassModule_) {
                    callArgs += classHandlerCast();
                }
                for (size_t i = 0; i < evtParams.size(); i++) {
                    if (i > 0 || isClassModule_) {
                        callArgs += ", ";
                    }
                    if (evtParamIsClass[i]) {
                        // Fix 084k: 处理器Sub形参是 ByRef 具体类 (vb6_cls_X**),
                        // 包装器形参是 ByVal 单指针 (vb6_cls_X*) — 用临时地址适配,
                        // 否则 vb6_VARIANT→vb6_cls_X** 触发 C2440
                        callArgs += "&(" + evtParamCTypeOnly[i] + "){" + cIdent(evtParams[i].name) + "}";
                    } else {
                        callArgs += cIdent(evtParams[i].name);
                    }
                }
                callArgs += ");";
                c_.emitLine(procCall + callArgs);
                c_.dedent();
                c_.emitLine("}");
                c_.emitBlank();
            }
        }
    }

