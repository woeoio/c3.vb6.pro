// cgen_util_dllentry_exports.inc - generateDllEntry 函数体片段：DLL 导出函数
// 由 src/backend/cgen_util.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 624~696 行，逐行未改。

    entry.emitLine("const int g_vb6_coclassCount = " + std::to_string(coClasses.size() + extraGroupRows) + ";");
    entry.emitBlank();

    // ExeComBridge 01: includeDllExports=false (EXE 工程的 com_entry.c) 时只生成
    //   coclass 表, 跳过 DLL 导出函数段 —— EXE 无 .def 导出表, 也不需要 DllMain.
    //   g_vb6_coclassCount / g_vb6_coclasses 必须无条件生成: vb6comserver_obj.c
    //   的 vb6_FindCoClassDesc 引用它们 (DLL 下由 dll_entry.c 提供, EXE 下由
    //   com_entry.c 提供).
    if (includeDllExports) {
    // 4. DLL导出函数
    entry.emitLine("// === DLL Export Functions ===");
    entry.emitBlank();

    entry.emitLine("HRESULT WINAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv) {");
    entry.indent();
    entry.emitLine("return vb6_GetClassFactory(rclsid, riid, ppv, g_vb6_coclasses, g_vb6_coclassCount);");
    entry.dedent();
    entry.emitLine("}");
    entry.emitBlank();

    entry.emitLine("HRESULT WINAPI DllCanUnloadNow(void) {");
    entry.indent();
    entry.emitLine("return vb6_DllCanUnloadNow();");
    entry.dedent();
    entry.emitLine("}");
    entry.emitBlank();

    entry.emitLine("HRESULT WINAPI DllRegisterServer(void) {");
    entry.indent();
    entry.emitLine("wchar_t dllPath[MAX_PATH];");
    entry.emitLine("HRESULT hr = vb6_GetDllPath(dllPath, MAX_PATH);");
    entry.emitLine("if (FAILED(hr)) return hr;");
    // Fix 098: 单个 coclass 注册失败不得中断整轮注册 (原先首个失败即 return,
    // 实测 104 个 coclass 只注册进 13 个, 且 vb6_RegisterTypeLib 从未执行 ——
    // TypeLib 完全缺失, 客户端 CLSIDFromProgID / 早绑定全部失效).
    // 改为累计失败数, 循环跑完, TypeLib 注册无条件执行.
    entry.emitLine("int failed = 0;");
    entry.emitLine("HRESULT firstFail = S_OK;");
    entry.emitLine("for (int i = 0; i < g_vb6_coclassCount; i++) {");
    entry.indent();
    entry.emitLine("hr = vb6_RegisterCoClass(&g_vb6_coclasses[i], dllPath);");
    entry.emitLine("if (FAILED(hr)) { if (failed == 0) firstFail = hr; failed++; }");
    entry.dedent();
    entry.emitLine("}");
    entry.emitLine("// P6.13: Register embedded TypeLib (无论 coclass 注册是否有失败都要执行)");
    entry.emitLine("vb6_RegisterTypeLib(dllPath);");
    entry.emitLine("return failed ? firstFail : S_OK;");
    entry.dedent();
    entry.emitLine("}");
    entry.emitBlank();

    entry.emitLine("HRESULT WINAPI DllUnregisterServer(void) {");
    entry.indent();
    entry.emitLine("wchar_t dllPath[MAX_PATH];");
    entry.emitLine("vb6_GetDllPath(dllPath, MAX_PATH);");
    entry.emitLine("for (int i = 0; i < g_vb6_coclassCount; i++) {");
    entry.indent();
    entry.emitLine("vb6_UnregisterCoClass(&g_vb6_coclasses[i]);");
    entry.dedent();
    entry.emitLine("}");
    entry.emitLine("// P6.13: Unregister embedded TypeLib");
    entry.emitLine("vb6_UnregisterTypeLib(dllPath);");
    entry.emitLine("return S_OK;");
    entry.dedent();
    entry.emitLine("}");
    entry.emitBlank();

    entry.emitLine("BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved) {");
    entry.indent();
    entry.emitLine("if (fdwReason == DLL_PROCESS_ATTACH) {");
    entry.indent();
    entry.emitLine("DisableThreadLibraryCalls(hinstDLL);");
    entry.emitLine("vb6_Init();");
    // 账 #305 第330刀 (账 #605): ActiveX DLL 也得自己播种 App.Title。VB6 实测 (宿主 = 32 位
    // cscript，见 ai/probes/app_title_dll/client_out.txt)：编译好的 DLL 里读 App.Title 答的是
    // **本工程**的 Title (P1)，运行期写它 err=0 (P5) —— 既不是宿主进程的标题，也没有「只读」这一档。
    // 第320刀只把播种发在 Exe 的五个入口模板里，Type=OleDll 工程一条都不进 ⇒ 真语料 48 份 DLL
    // 工程 (37 份带非空 Title) 里那 82 行 App.Title 全落「没播种」那一档，交回宿主 EXE 的名字。
    // 形状与 Exe 那一份逐字节同 (`vb6_SetAppTitle(L"…")`) —— 一位主人、两个入口，不另设第二枚 setter。
    if (!projectTitle_.empty()) {
        entry.emitLine("vb6_SetAppTitle(L\"" + escapeWideCString(projectTitle_)
                    + "\");  /* 账 #592/#605: App.Title = .vbp 的 Title */");
    }
    entry.dedent();
    entry.emitLine("} else if (fdwReason == DLL_PROCESS_DETACH) {");
    entry.indent();
    entry.emitLine("vb6_Exit();");
    entry.dedent();
    entry.emitLine("}");
    entry.emitLine("return TRUE;");
    entry.dedent();
    entry.emitLine("}");
    entry.emitBlank();
    }  // ExeComBridge 01: end if (includeDllExports)

    return entry.str();