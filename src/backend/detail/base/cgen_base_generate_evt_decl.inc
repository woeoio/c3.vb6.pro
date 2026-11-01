// cgen_base_generate_evt_decl.inc - generate 函数体片段：事件处理器包装声明与 vtable 前向声明
// 由 src/backend/cgen_base.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 597~672 行，逐行未改。

    // P6.5/P13.23: 先生成事件处理器包装函数的.h声明（在#endif之前）
    // 收集wrapper函数签名，稍后在#endif之前输出到.h
    std::vector<std::string> evtWrapperSigs;
    if (!knownWithEventsVars_.empty()) {
        for (auto& [varLower, srcClassName] : knownWithEventsVars_) {
            auto* srcClsSym = symTab_.lookup(srcClassName);
            if (!srcClsSym) continue;
            if (srcClsSym->kind == SymbolKind::Class) {
                for (size_t evtIdx256 = 0; evtIdx256 < srcClsSym->eventNames.size(); ++evtIdx256) {
                    const std::string& evtName = srcClsSym->eventNames[evtIdx256];
                    // 账 #305 256: .h 的前向声明与 .c 的定义必须同形 —— 两处都读同一张
                    // Symbol::eventParamByAddr (语义层由 src/common/event_param_abi.hpp 算一次)。
                    const std::vector<char>* evtSlots256 =
                        evtIdx256 < srcClsSym->eventParamByAddr.size()
                            ? &srcClsSym->eventParamByAddr[evtIdx256] : nullptr;
                    std::string varName = varLower;
                    auto* varSym = symTab_.lookup(varLower);
                    if (varSym) varName = varSym->name;
                    std::string handlerName = varName + "_" + evtName;
                    auto* handlerSym = symTab_.lookup(handlerName);
                    if (!handlerSym) continue;
                    std::string wrapperName = evtWrapperName(varLower, evtName);
                    // 查找Event声明的参数
                    std::vector<ParameterInfo> evtParams;
                    std::vector<std::string> evtParamCTypes;
                    for (auto& decl2 : module.declarations) {
                        if (decl2->kind == ASTNodeKind::SubDecl) {
                            auto& sub = static_cast<SubDecl&>(*decl2);
                            if (Symbol::toLower(sub.name) == Symbol::toLower(handlerName)) {
                                size_t pidx256 = 0;
                                for (auto& param : sub.params) {
                                    const bool byAddr256 =
                                        evtSlots256 && pidx256 < evtSlots256->size() &&
                                        (*evtSlots256)[pidx256] != 0;
                                    pidx256++;
                                    ParameterInfo pi;
                                    pi.name = param->name;
                                    pi.type = (param->asType && param->asType->kind == ASTNodeKind::SimpleTypeRef)
                                        ? typeSys_.resolveTypeName(static_cast<SimpleTypeRef*>(param->asType.get())->name)
                                        : Vb6Type::Variant;
                                    if (pi.type == Vb6Type::Unknown || pi.type == Vb6Type::Empty)
                                        pi.type = Vb6Type::Variant;
                                    pi.isByVal = !byAddr256;
                                    evtParams.push_back(pi);
                                    // Fix 084k: 类名参数用 mapTypeRef → vb6_cls_X*, 不能
                                    // resolveTypeName 解析失败回退为 vb6_VARIANT (导致 C2440)
                                    evtParamCTypes.push_back(eventParamCDecl(
                                        mapTypeRef(param->asType.get()), byAddr256, cIdent(param->name)));
                                }
                                break;
                            }
                        }
                    }
                    std::string sig = "void " + wrapperName + "(void* handler";
                    for (auto& pt : evtParamCTypes) {
                        sig += ", " + pt;
                    }
                    sig += ")";
                    evtWrapperSigs.push_back(sig);
                }
            } else if (srcClsSym->kind == SymbolKind::ComClass && srcClsSym->comHasSourceIface) {
                // P13.23: COM WithEvents 包装函数签名统一为 (void* handler, VARIANT* args, int argc, VARIANT* result)
                for (auto& evtName : srcClsSym->eventNames) {
                    std::string varName = varLower;
                    auto* varSym = symTab_.lookup(varLower);
                    if (varSym) varName = varSym->name;
                    std::string handlerName = varName + "_" + evtName;
                    auto* handlerSym = symTab_.lookup(handlerName);
                    if (!handlerSym) continue;
                    std::string wrapperName = comEvtWrapperName(varLower, evtName);
                    std::string sig = "void " + wrapperName + "(void* handler, VARIANT* args, int argc, VARIANT* result)";
                    evtWrapperSigs.push_back(sig);
                }
            }
        }
    }
    // 输出wrapper声明到.h（在#endif之前）
    for (auto& sig : evtWrapperSigs) {
        h_.emitLine(sig + ";");
        h_.emitBlank();
    }

    // P13.23: vtable source interface 事件接收器创建函数前向声明
    emitComVtableSinkDecls();

    h_.emitBlank();
    h_.emitLine("#endif /* " + guard + " */");

