// cgen_base_generate_entry.inc - generate 函数体片段：入口点生成
// 由 src/backend/cgen_base.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 955~1101 行，逐行未改。

    // 生成入口点 (类模块不生成main; 多模块工程中仅有Sub Main的模块生成main)
    // P6.6: ActiveX DLL入口点统一由dll_entry.c生成, 不在各模块.c中生成
    if (!isClassModule_) {
        if (isDll_) {
            // P6.6: DLL模式下, 入口点由driver额外生成的dll_entry.c提供
            // 这里只生成vb6_Init/vb6_Exit的调用桩 (供dll_entry.c中的DllMain使用)
            // 不再调用emitActiveXDll()以避免符号重复定义
        } else {
            c_.emitBlank();
            c_.emitLine("// === 入口点 ===");
            bool hasMain = false;
            // Fix 160: ComLib= 免注册组件表烘焙 — 在 vb6_Init() 之前发出.
            // vb6_ComLibRegister 只填一张静态表, 与 vb6_Init 无顺序依赖, 但必须在任何
            // 用户代码 (vb6_CreateObject) 之前完成; 放入口点最前保证首次调用即就绪.
            // 表空 (未声明 ComLib=) 时一行不发 → 产物与改动前逐字节一致, 零回归.
            auto emitComLibReg = [&]() {
                if (!comLibRefs_ || comLibRefs_->empty()) return;
                c_.emitBlank();
                c_.emitLine("/* Fix 160: ComLib= 免注册组件表 (vbp 声明; 路径相对 exe) */");
                c_.emitLine("static const Vb6ComLib c3_com_libs[] = {");
                for (const auto& e : *comLibRefs_) {
                    // Vb6ComLib 四字段全是 const wchar_t* → 必须发宽字符串字面量
                    c_.emitLine("    { L\"" + escapeWideCString(e[0]) + "\", L\""
                                + escapeWideCString(e[1]) + "\", L\"" + escapeWideCString(e[2])
                                + "\", L\"" + escapeWideCString(e[3]) + "\" },");
                }
                c_.emitLine("};");
                c_.emitLine("vb6_ComLibRegister(c3_com_libs, "
                            + std::to_string(comLibRefs_->size()) + ");");
            };
            // Fix <vbeclipse> 2026-10-06: Object= OCX 免注册表烘焙 — 与 ComLib 同构,
            // 入口点最前发出, vb6_Form_ControlsAdd 按 ProgID 命中即绕开注册表.
            // 表空 (未声明 Object=) 时一行不发 → 产物与改动前逐字节一致, 零回归.
            auto emitOcxRefReg = [&]() {
                if (!ocxLibRefs_ || ocxLibRefs_->empty()) return;
                c_.emitBlank();
                c_.emitLine("/* Fix <vbeclipse>: Object= 免注册 OCX 控件表 (vbp 声明; 路径相对 exe) */");
                c_.emitLine("static const Vb6OcxRef c3_ocx_libs[] = {");
                for (const auto& e : *ocxLibRefs_) {
                    // Vb6OcxRef 四字段全是 const wchar_t* → 必须发宽字符串字面量
                    c_.emitLine("    { L\"" + escapeWideCString(e[0]) + "\", L\""
                                + escapeWideCString(e[1]) + "\", L\"" + escapeWideCString(e[2])
                                + "\", L\"" + escapeWideCString(e[3]) + "\" },");
                }
                c_.emitLine("};");
                c_.emitLine("vb6_OcxRefRegister(c3_ocx_libs, "
                            + std::to_string(ocxLibRefs_->size()) + ");");
            };
            // 账 §B97: 「入口点里调用哪些模块的 *_init()」以前是 5 份同形循环，各自直接
            // 迭代一枚 unordered_set ⇒ 各模块全局的初始化先后是编译器二进制的函数（实测
            // tests/cc_id/Id.vbp 两台各交一种序）。收成这一处：名单本身由 driver 排成
            // modules_ 的下标序交下来（extModulesInModuleVectorOrder；那枚序**不是** .vbp 的字面行序，
            // 见 driver_frontend.cpp:327-339 的类模块前移），这里只照序逐行发。
            // ⚠ 五份里三份发在「本模块自己的 init」之后、两份发在其**之前**（157 / 205 那两格），
            //   这个 own-vs-others 的相对位置是改动前就有的形状，本刀**刻意不动**（动了就是
            //   第二件事：改变运行期初始化序，需要另立的口径与另一次的判据）。
            auto emitExtInits = [&]() {
                if (!isMultiModule_) return;
                for (const auto& extMod : externalModules_) {
                    c_.emitLine("vb6_mod_" + cIdent(extMod) + "_init();");
                }
            };
        for (auto& decl : module.declarations) {
            if (decl->kind == ASTNodeKind::SubDecl) {
                auto& sub = static_cast<SubDecl&>(*decl);
                if (sub.name == "Main" && sub.access == AccessLevel::Public) {
                    hasMain = true;
                    break;
                }
            }
        }
        // Fix 142: 多模块工程中只有启动对象所在模块能生成进程入口点.
        // 此前四个窗体模块各生成一个 WinMain → LNK2005 "WinMain 已经在 X.obj 中定义".
        // 判定:
        //   Startup="Sub Main"  → 只有含 Public Sub Main 的模块 (hasMain) 生成;
        //   Startup="frmMain"   → 只有该窗体模块生成 (moduleName 或窗体名匹配);
        //   无 Startup 信息     → 单模块照旧; 多模块由 driver 兜底取第一个窗体.
        bool isStartupModule = true;
        if (isMultiModule_ && !startupObject_.empty()) {
            std::string so = startupObject_;
            std::transform(so.begin(), so.end(), so.begin(), ::tolower);
            if (so == "sub main") {
                isStartupModule = hasMain;
            } else {
                std::string modLower = module.moduleName;
                std::transform(modLower.begin(), modLower.end(), modLower.begin(), ::tolower);
                std::string frmLower = frmDesc ? frmDesc->formName : std::string();
                std::transform(frmLower.begin(), frmLower.end(), frmLower.begin(), ::tolower);
                isStartupModule = (so == modLower || (!frmLower.empty() && so == frmLower));
            }
        }

        // 单模块: 允许用第一个Public Sub作为入口
        // 多模块: 只有启动模块 (Sub Main / 启动窗体) 才生成main
        bool shouldGenMain = hasMain && isStartupModule;
        if (!shouldGenMain && (!isMultiModule_ || isStartupModule)) {
            // 单模块模式: 无Sub Main时自动找首个Public Sub
    for (auto& decl : module.declarations) {
        if (currentClassIsInterface) continue;  // P6.4: 接口类不生成方法实现体
        if (decl->kind == ASTNodeKind::SubDecl) {
                    auto& sub = static_cast<SubDecl&>(*decl);
                    if (sub.access == AccessLevel::Public) {
                        shouldGenMain = true;
                        // P7: 窗体模块生成WinMain, 标准模块生成main
                        if (module.isFormModule && frmDesc) {
                            // 窗体模块: WinMain → 显示窗体 → 消息循环
                            std::string formName = frmDesc->formName;
                            c_.emitLine("int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrevInst, LPSTR lpCmdLine, int nCmdShow) {");
                            c_.indent();
                            c_.emitLine("(void)hPrevInst; (void)lpCmdLine; (void)nCmdShow;");
            emitComLibReg();
            emitOcxRefReg();
            c_.emitLine("vb6_Init();");
            if (!modInitFuncName.empty()) c_.emitLine(modInitFuncName + "();");
            // Fix 080: 多模块工程中, 调用所有外部模块的初始化函数
            emitExtInits();
            c_.emitLine("vb6_SetAppInstance((void*)hInst);");
            c_.emitLine("vb6_form_show_" + cIdent(formName) + "(NULL, 0);  /* Show form modeless, NULL=hMDIClient */");
                            c_.emitLine("int ret = vb6_MessageLoop();");
                            c_.emitLine("vb6_Exit();");
                            c_.emitLine("return ret;");
                            c_.dedent();
                            c_.emitLine("}");
                        } else {
                            c_.emitLine("int main(int argc, char* argv[]) {");
                            c_.indent();
            emitComLibReg();
            emitOcxRefReg();
            c_.emitLine("vb6_Init();");
            if (!modInitFuncName.empty()) c_.emitLine(modInitFuncName + "();");
            // Fix 080: 多模块工程中, 调用所有外部模块的初始化函数
            emitExtInits();
            c_.emitLine(cProcName(sub.name, sub.access) + "();");
            // Fix 167: VB6 的 Sub Main 返回后进程要继续驻留 (运行时自己泵消息),
            // 直到所有窗体关闭。此处不补循环 → modeless Load 的窗体一返回就退出。
            c_.emitLine("if (vb6_AnyThreadWindowVisible()) vb6_MessageLoop();");
            c_.emitLine("vb6_Exit();");
                            c_.emitLine("return 0;");
                            c_.dedent();
                            c_.emitLine("}");
                        }
                        break;
                    }
                }
            }
        }
        if (shouldGenMain && hasMain) {
            // P7: 窗体模块生成WinMain, 标准模块生成main
            if (module.isFormModule) {
                c_.emitLine("int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrevInst, LPSTR lpCmdLine, int nCmdShow) {");
                c_.indent();
                c_.emitLine("(void)hPrevInst; (void)lpCmdLine; (void)nCmdShow;");
                emitComLibReg();
            emitOcxRefReg();
                c_.emitLine("vb6_Init();");
                // Fix 080: 多模块工程中, 调用所有外部模块的初始化函数
                emitExtInits();
                if (!modInitFuncName.empty()) c_.emitLine(modInitFuncName + "();");
                c_.emitLine("vb6_SetAppInstance((void*)hInst);");
                c_.emitLine(cProcName("Main", AccessLevel::Public) + "();");
                c_.emitLine("int ret = vb6_MessageLoop();");
                c_.emitLine("vb6_Exit();");
                c_.emitLine("return ret;");
                c_.dedent();
                c_.emitLine("}");
            } else {
                c_.emitLine("int main(int argc, char* argv[]) {");
                c_.indent();
            emitComLibReg();
            emitOcxRefReg();
            c_.emitLine("vb6_Init();");
            if (!modInitFuncName.empty()) c_.emitLine(modInitFuncName + "();");
            // Fix 080: 多模块工程中, 调用所有外部模块的初始化函数
            emitExtInits();
            c_.emitLine(cProcName("Main", AccessLevel::Public) + "();");
                // Fix 167: 同上方 main 模板 — Sub Main 返回后仍要驻留 (VB6 语义)
                c_.emitLine("if (vb6_AnyThreadWindowVisible()) vb6_MessageLoop();");
                c_.emitLine("vb6_Exit();");
                c_.emitLine("return 0;");
                c_.dedent();
                c_.emitLine("}");
            }
        } else if (!shouldGenMain) {
            // Fix 142: 非启动模块不生成入口点 (否则多窗体工程 LNK2005)
            if (isStartupModule && module.isFormModule && frmDesc) {
                // P7: 空窗体也要生成WinMain, 自动显示窗体
                std::string formName = frmDesc->formName;
                c_.emitLine("int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrevInst, LPSTR lpCmdLine, int nCmdShow) {");
                c_.indent();
                c_.emitLine("(void)hPrevInst; (void)lpCmdLine; (void)nCmdShow;");
                emitComLibReg();
            emitOcxRefReg();
                c_.emitLine("vb6_Init();");
                if (!modInitFuncName.empty()) c_.emitLine(modInitFuncName + "();");
                // Fix 080: 多模块工程中, 调用所有外部模块的初始化函数
                emitExtInits();
                c_.emitLine("vb6_SetAppInstance((void*)hInst);");
                c_.emitLine("vb6_form_show_" + cIdent(formName) + "(NULL, 0);  /* Show form modeless, NULL=hMDIClient */");
                c_.emitLine("int ret = vb6_MessageLoop();");
                c_.emitLine("vb6_Exit();");
                c_.emitLine("return ret;");
                c_.dedent();
                c_.emitLine("}");
            } else {
                c_.emitLine("// No entry point (library module)");
            }
        }
        }  // end else (EXE mode)
    }

