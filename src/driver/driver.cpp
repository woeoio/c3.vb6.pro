// driver.cpp - C3 编译器驱动: Driver 构造/析构 + 代码生成阶段（伞文件）
// 2026-09-17 两步拆分：原 2494 行单文件 → 6 个新编译单元（driver_args / driver_compile /
//   driver_frontend / driver_semantics / driver_crossmod / driver_link，均已登记 CMakeLists）
//   + 本文件（class ExternalRefCollector + runCodeGeneration 函数骨架）
//   + detail/ 下 10 个「函数体片段」（在 runCodeGeneration() 函数体内被 #include）：
//     driver_codegen_prelude.inc             —— 模块数与分析器数校验（原 95~99 行）
//     driver_codegen_voidfield_scan.inc      —— Fix 023: className → void* 字段集预扫描（原 100~187 行）
//     driver_codegen_typedfield_scan.inc     —— Fix 037b/086: typed 字段表 + 窗体模块名 + Public 常量表（原 188~271 行）
//     driver_codegen_fieldkind_scan.inc      —— 账 #305 §B168/§B169: className → (字段名 → 声明档) 预扫描
//     driver_codegen_variant_funcs.inc       —— Fix 045: 返回 Variant 的项目函数集合（原 272~341 行）
//     driver_codegen_dup_module_vars.inc     —— Fix 093b: 跨模块同名模块级变量（原 342~370 行）
//     driver_codegen_module_loop.inc         —— 逐模块代码生成主循环（写 .h/.c，原 371~462 行）
//     driver_codegen_dll_typelib.inc         —— P6.6/P9: ActiveX DLL 模式 TypeLib 构建（原 463~698 行）
//     driver_codegen_dll_sync.inc            —— P6.6.6/M29: IID/CLSID/DISPID 符号同步（原 699~753 行）
//     driver_codegen_dll_entry.inc           —— generateDllEntry 与 dll_entry.c 写出 + 返回（原 754~772 行）
// 10 个 .inc 不是独立编译单元（不登记 CMakeLists），单独 include 会编译不过；片段内容逐行未改，零行为改动。

#include "driver/driver.hpp"
#include "driver/detail/phase_timing.hpp"   // ai/030 续: C3_TIMING=1 的相位计时
#include "common/diagnostics.hpp"
#include "common/encoding.hpp"
#include "common/source_manager.hpp"
#include "lexer/lexer.hpp"
#include "lexer/token.hpp"
#include "preprocessor/preprocessor.hpp"
#include "parser/parser.hpp"
#include "ast/ast_printer.hpp"
#include "ast/ast_visitor.hpp"
#include "semantics/semantic_analyzer.hpp"
#include "backend/cgen.hpp"
#include "backend/msvc_driver.hpp"
#include "typelib/typelib_builder.hpp"
#include "driver/rtl_embedded.hpp"
#include "project/vbp_parser.hpp"
#include "project/frm_parser.hpp"

#include <iostream>
#include <fstream>
#include <algorithm>
#include <filesystem>
#include <cstdlib>
#include <set>
#include <map>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <cctype>
#include <cstring>
#include <functional>

#ifdef _WIN32
#include <windows.h>
#include "semantics/interface_sig.hpp"  // 账 #305 新账 #489: ifaceSlotKeyNamed (legacy 槽键与 D2 同源的位主人)
#endif

namespace vb6c3 {

// 账 #611: 工程内 .cls 默认成员的解析只住在 backend/cgen_util_ctrl.cpp
// (check_uc_default_member.ps1 的 R1 单一出口), driver 侧只调它填表。
bool findClassDefaultMember(const Module& mod, std::string& propOut);

// pathToUtf8/utf8ToWide/utf8ToPath moved to common/encoding.hpp

Driver::Driver() : diag_(std::make_unique<Diagnostics>()) {}
Driver::~Driver() = default;

// 账 §56.52: 原来这里的 `ExternalRefCollector` (AST 引用收集器) 已**撤掉** ——
// 它只覆写 visit(IdentifierExpr&) 与 visit(SimpleTypeRef&), 而 AST 遍历根本走不到
// VariableDecl / ParameterDecl / TypeDecl / TypeMember 里的类型引用 ⇒ 那份名单是
// **残缺**的, 拿来裁跨模块 include 会把还需要的头裁掉 ⇒ 编译断。
// 判据现在只有一处: 发码器在发出外部引用那一刻登记 (见 cgen_state.inc 的
// usedExternalModules_), 外加发码文本反查兜底 —— 见 src/backend/cgen_base.cpp 的
// trimCrossModuleIncludes()。

bool Driver::runCodeGeneration(const CompileOptions& options, const std::string& outputDir) {
#include "driver/detail/driver_codegen_prelude.inc"
#include "driver/detail/driver_codegen_voidfield_scan.inc"
#include "driver/detail/driver_codegen_typedfield_scan.inc"
#include "driver/detail/driver_codegen_fieldkind_scan.inc"
#include "driver/detail/driver_codegen_variant_funcs.inc"
#include "driver/detail/driver_codegen_dup_module_vars.inc"
#include "driver/detail/driver_codegen_module_loop.inc"
#include "driver/detail/driver_codegen_dll_typelib.inc"
#include "driver/detail/driver_codegen_dll_sync.inc"
#include "driver/detail/driver_codegen_dll_entry.inc"
}

} // namespace vb6c3