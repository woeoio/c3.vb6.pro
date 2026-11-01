// driver_codegen_variant_funcs.inc - runCodeGeneration 函数体片段：Fix 045 返回 Variant 的项目函数集合
// 由 src/driver/driver.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 272~341 行，逐行未改。

    // Fix 045: 预扫描所有模块, 构建返回 Variant 的项目函数 C 名集合.
    // 遍历所有模块的 FunctionDecl 和 PropertyDecl(PropertyGet),
    // 检查 returnType 是否为 Variant (nullptr=隐式Variant, 或 SimpleTypeRef name="Variant").
    // 构建 C 函数名: vb6_<moduleName>_<procName> (与 cProcName 多模块模式一致).
    // 供 cExprIsVariant() 检测项目类函数返回 Variant 的表达式.
    std::unordered_set<std::string> variantReturnFuncs;
    // 账 #305 第318刀 (账 #594): 第三形 —— cProcName 在「本工程只有这一枚发码模块」时发
    //   **裸名** vb6_<proc>（cgen_base_naming.cpp:372 那条闸反过来读），上面登记的
    //   vb6_<模块>_<proc> / vb6_pv_<proc> 两形都够不着它。先按模块记下候选，裸名要不要补
    //   由模块循环按 cProcName 的同一条条件答（isClassModule 的方法恒带模块名，不进候选）。
    std::unordered_map<std::string, std::vector<std::string>> variantBareProcNames;
    for (size_t i = 0; i < modules_.size(); i++) {
        const auto& module = modules_[i];
        const std::string& modName = module->moduleName;
        for (const auto& decl : module->declarations) {
            std::string procName;
            bool isVariant = false;

            if (decl->kind == ASTNodeKind::FunctionDecl) {
                auto& fn = static_cast<const FunctionDecl&>(*decl);
                procName = fn.name;
                if (!fn.returnType) {
                    isVariant = true;  // 隐式 Variant
                } else if (fn.returnType->kind == ASTNodeKind::SimpleTypeRef) {
                    auto& tr = static_cast<const SimpleTypeRef&>(*fn.returnType);
                    std::string lower = tr.name;
                    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                    if (lower == "variant") isVariant = true;
                }
            } else if (decl->kind == ASTNodeKind::PropertyDecl) {
                auto& prop = static_cast<const PropertyDecl&>(*decl);
                if (prop.propKind != ProcKind::PropertyGet) continue;
                procName = "prop_get_" + prop.name;
                if (!prop.returnType) {
                    isVariant = true;
                } else if (prop.returnType->kind == ASTNodeKind::SimpleTypeRef) {
                    auto& tr = static_cast<const SimpleTypeRef&>(*prop.returnType);
                    std::string lower = tr.name;
                    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                    if (lower == "variant") isVariant = true;
                }
            }

            if (isVariant && !procName.empty()) {
                // cIdent 等价: 替换特殊字符为 _
                auto cIdentS = [](const std::string& vb6Name) -> std::string {
                    std::string name = vb6Name;
                    if (!name.empty() && name.front() == '[' && name.back() == ']') {
                        name = name.substr(1, name.size() - 2);
                    }
                    for (auto& ch : name) {
                        if (!std::isalnum(static_cast<unsigned char>(ch)) && ch != '_') {
                            ch = '_';
                        }
                    }
                    return name;
                };
                std::string cName = "vb6_" + cIdentS(modName) + "_" + cIdentS(procName);
                variantReturnFuncs.insert(cName);
                if (!module->isClassModule) {
                    variantBareProcNames[modName].push_back(cIdentS(procName));
                }
                // Fix 066: 同模块调用 Private 函数时 C 函数名使用 vb6_pv_ 前缀,
                // 也需加入 variantReturnFuncs 以便 cExprIsVariant 正确检测.
                bool isPrivate = false;
                if (decl->kind == ASTNodeKind::FunctionDecl) {
                    isPrivate = (static_cast<const FunctionDecl&>(*decl).access == AccessLevel::Private);
                } else if (decl->kind == ASTNodeKind::PropertyDecl) {
                    isPrivate = (static_cast<const PropertyDecl&>(*decl).access == AccessLevel::Private);
                }
                if (isPrivate) {
                    std::string privateCName = "vb6_pv_" + cIdentS(procName);
                    variantReturnFuncs.insert(privateCName);
                }
            }
        }
    }

