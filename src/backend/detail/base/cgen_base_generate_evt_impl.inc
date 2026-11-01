// cgen_base_generate_evt_impl.inc - generate 函数体片段：COM 事件回调、vtable 接收器实现与模块级变量延迟初始化
// 由 src/backend/cgen_base.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 796~954 行，逐行未改。

        // P13.23: External COM WithEvents event callback generation
        for (auto& [varLower, srcClassName] : knownWithEventsVars_) {
            auto* srcClsSym = symTab_.lookup(srcClassName);
            if (!srcClsSym || srcClsSym->kind != SymbolKind::ComClass) continue;

            for (auto& evtName : srcClsSym->eventNames) {
                std::string varName = varLower;
                auto* varSym = symTab_.lookup(varLower);
                if (varSym) varName = varSym->name;
                std::string handlerName = varName + "_" + evtName;

                auto* handlerSym = symTab_.lookup(handlerName);
                if (!handlerSym) continue;

                // Generate: void vb6_com_evt_<var>_<evt>(void* handler, VARIANT* args, int argc, VARIANT* result)
                std::string wrapperName = comEvtWrapperName(varLower, evtName);

                // Get event handler parameters from Sub declaration
                // 每个参数记录: Vb6Type / ByRef标志 / mapTypeRef 值类型 (对象→vb6_ComIface_X*)
                std::vector<ParameterInfo> evtParams;
                std::vector<std::string> evtParamRefCTypes;  // mapTypeRef 值类型
                std::vector<bool> evtParamByRef;
                for (auto& decl2 : module.declarations) {
                    if (decl2->kind == ASTNodeKind::SubDecl) {
                        auto& sub = static_cast<SubDecl&>(*decl2);
                        if (Symbol::toLower(sub.name) == Symbol::toLower(handlerName)) {
                            for (auto& param : sub.params) {
                                ParameterInfo pi;
                                pi.name = param->name;
                                pi.isByVal = param->isByVal;
                                if (param->asType && param->asType->kind == ASTNodeKind::ArrayTypeRef) {
                                    auto& arr = static_cast<ArrayTypeRef&>(*param->asType);
                                    Vb6Type elemType = Vb6Type::Variant;
                                    if (arr.elementType && arr.elementType->kind == ASTNodeKind::SimpleTypeRef) {
                                        elemType = typeSys_.resolveTypeName(static_cast<SimpleTypeRef*>(arr.elementType.get())->name);
                                    }
                                    if (elemType == Vb6Type::Unknown || elemType == Vb6Type::Empty)
                                        elemType = Vb6Type::Variant;
                                    pi.type = static_cast<Vb6Type>(static_cast<uint16_t>(elemType) | static_cast<uint16_t>(Vb6Type::Array));
                                } else if (param->asType && param->asType->kind == ASTNodeKind::SimpleTypeRef) {
                                    pi.type = typeSys_.resolveTypeName(static_cast<SimpleTypeRef*>(param->asType.get())->name);
                                    if (pi.type == Vb6Type::Unknown || pi.type == Vb6Type::Empty)
                                        pi.type = Vb6Type::Variant;
                                } else {
                                    pi.type = Vb6Type::Variant;
                                }
                                evtParams.push_back(pi);
                                evtParamRefCTypes.push_back(mapTypeRef(param->asType.get()));
                                evtParamByRef.push_back(!param->isByVal);
                            }
                            break;
                        }
                    }
                }

                c_.emitLine("void " + wrapperName + "(void* handler, VARIANT* args, int argc, VARIANT* result) {");
                c_.indent();

                // Extract parameters from VARIANT args
                for (size_t pi = 0; pi < evtParams.size(); pi++) {
                    std::string pName = cIdent(evtParams[pi].name);
                    std::string pType = mapType(evtParams[pi].type);
                    std::string idx = std::to_string(pi);
                    bool isArray = (static_cast<uint16_t>(evtParams[pi].type) & static_cast<uint16_t>(Vb6Type::Array)) != 0;
                    Vb6Type baseType = isArray
                        ? static_cast<Vb6Type>(static_cast<uint16_t>(evtParams[pi].type) & ~static_cast<uint16_t>(Vb6Type::Array))
                        : evtParams[pi].type;
                    // Fix 084m: COM对象参数 (As ADODB.Error等) — 值类型用 mapTypeRef
                    // (vb6_ComIface_X* / void*), 不能解析为 vb6_VARIANT 否则 C2440
                    std::string refType = evtParamRefCTypes[pi];
                    bool isObj = refType.find("vb6_ComIface_") != std::string::npos
                        || refType.find("vb6_cls_") != std::string::npos
                        || refType.find("vb6_iface_") != std::string::npos
                        || refType == "void*";
                    // 枚举形参: resolveTypeName 只认简单名, 限定名 (XtremeDockingPane.DockingPaneAction /
                    // XtremeCalendarControl.CalendarSelectionChanged) 落到 Object, 而 mapTypeRef 的类型感知
                    // 查找能认成 EnumType → int32_t。照 Object 发 `void*` 再赋 vb6_VariantFromStackVARIANT
                    // 就是 C2440 (枚举实参本来就是 VT_I4, 不是 VT_DISPATCH):
                    //   Outlook2013 frmMain.c(2355) DockingPaneManager_Action / frmInbox.c(1309) SelectionChanged
                    bool refIsInt = !isObj && refType != pType
                        && (refType == "int32_t" || refType == "int16_t" || refType == "uint32_t"
                            || refType == "uint16_t" || refType == "intptr_t");
                    std::string localType = pType;
                    if (!isArray && (isObj || refIsInt)) localType = refType;
                    c_.emitLine(localType + " " + pName + ";");
                    if (isArray) {
                        // P13.23: COM event arrays come as VARIANT containing SAFEARRAY (VT_ARRAY | VT_xxx).
                        // Extract raw pointer to the first element for 1D Byte arrays.
                        if (baseType == Vb6Type::Byte) {
                            c_.emitLine(pName + " = (uint8_t*)((args[" + idx + "].vt & VT_ARRAY) ? args[" + idx + "].parray->pvData : NULL);");
                        } else {
                            c_.emitLine(pName + " = (" + pType + ")((args[" + idx + "].vt & VT_ARRAY) ? args[" + idx + "].parray->pvData : NULL);");
                        }
                    } else if (evtParams[pi].type == Vb6Type::String) {
                        c_.emitLine(pName + " = (args[" + idx + "].vt == VT_BSTR) ? args[" + idx + "].bstrVal : NULL;");
                    } else if (isObj) {
                        // COM对象: 从VARIANT中取 IDispatch/IUnknown 指针, 与处理器形参类型一致
                        c_.emitLine(pName + " = (" + refType + ")((args[" + idx + "].vt & VT_DISPATCH) ? args[" + idx + "].pdispVal : ((args[" + idx + "].vt & VT_UNKNOWN) ? args[" + idx + "].punkVal : NULL));");
                    } else if (evtParams[pi].type == Vb6Type::Long || evtParams[pi].type == Vb6Type::Integer) {
                        c_.emitLine(pName + " = (" + pType + ")V_I4(&args[" + idx + "]);");
                    } else if (refIsInt) {
                        // 见上方 refIsInt: 枚举形参 (Object 兜底 + mapTypeRef 认出枚举) 按 VT_I4 取
                        c_.emitLine(pName + " = (" + refType + ")V_I4(&args[" + idx + "]);");
                    } else if (evtParams[pi].type == Vb6Type::Boolean) {
                        c_.emitLine(pName + " = (args[" + idx + "].vt == VT_BOOL) ? (V_I4(&args[" + idx + "]) != 0) : 0;");
                    } else if (evtParams[pi].type == Vb6Type::Single || evtParams[pi].type == Vb6Type::Double) {
                        c_.emitLine(pName + " = (" + pType + ")V_R8(&args[" + idx + "]);");
                    } else {
                        // Variant: 转换COM VARIANT → vb6_VARIANT (栈上, 不释放)
                        c_.emitLine(pName + " = vb6_VariantFromStackVARIANT(&args[" + idx + "]);");
                    }
                }

                // Call the VB6 handler function
                // Fix 019: 类模块必须将 handler (void*) 转换为 cls* 再作为 me 传入,
                //          不可使用 classMeParam() (那是参数声明)
                // Fix 084m: ByRef 参数取地址 (&name), 与处理器形参 (int32_t* 等) 匹配
                std::string procCall = cProcName(handlerName, handlerSym->access, isClassModule_ ? moduleName_ : "");
                std::string callArgs = "(";
                if (isClassModule_) {
                    callArgs += classHandlerCast();
                }
                for (size_t i = 0; i < evtParams.size(); i++) {
                    if (i > 0 || isClassModule_) {
                        callArgs += ", ";
                    }
                    if (evtParamByRef[i]) {
                        callArgs += "&";
                    }
                    callArgs += cIdent(evtParams[i].name);
                }
                callArgs += ");";
                c_.emitLine(procCall + callArgs);

                c_.dedent();
                c_.emitLine("}");
                c_.emitBlank();
            }
        }

    // P13.23: 生成外部COM vtable source interface 事件接收器实现
    emitComVtableSinks();

    // Fix 054/080: 模块级变量延迟初始化函数 (C2099 workaround)
    // C语言文件作用域变量不能用运行时函数调用初始化, 所以先声明为 NULL/0,
    // 再在模块初始化函数中执行实际的 SafeArrayCreate 等运行时初始化.
    // Fix 080: 多模块工程中, 每个模块都生成初始化函数(即使为空),
    // 非入口模块的初始化函数需要被入口模块调用,
    // 所以去掉 static 改为 extern 可见, 并在 .h 中添加声明.
    std::string modInitFuncName = "vb6_mod_" + cIdent(baseName_) + "_init";
    bool needModInitFunc = isMultiModule_ || !moduleInitStmts_.empty();
    if (needModInitFunc) {
        c_.emitBlank();
        c_.emitLine("// Fix 054/080: Module-level variable deferred initialization");
        if (isMultiModule_) {
            c_.emitLine("void " + modInitFuncName + "(void) {");
            h_.emitLine("extern void " + modInitFuncName + "(void);");
        } else {
            c_.emitLine("static void " + modInitFuncName + "(void) {");
        }
        c_.indent();
        for (auto& stmt : moduleInitStmts_) {
            c_.emitLine(stmt);
        }
        c_.dedent();
        c_.emitLine("}");
        c_.emitBlank();
    }
    if (!needModInitFunc) modInitFuncName.clear();

