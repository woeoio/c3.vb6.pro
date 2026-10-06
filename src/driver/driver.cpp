// driver.cpp - C3 编译器驱动: Driver 构造/析构 + 代码生成阶段（伞文件）
// 2026-09-17 两步拆分：原 2494 行单文件 → 6 个新编译单元（driver_args / driver_compile /
//   driver_frontend / driver_semantics / driver_crossmod / driver_link，均已登记 CMakeLists）
//   + 本文件（class ExternalRefCollector + runCodeGeneration 函数骨架）
//   + detail/ 下 9 个「函数体片段」（在 runCodeGeneration() 函数体内被 #include）：
//     driver_codegen_prelude.inc             —— 模块数与分析器数校验（原 95~99 行）
//     driver_codegen_voidfield_scan.inc      —— Fix 023: className → void* 字段集预扫描（原 100~187 行）
//     driver_codegen_typedfield_scan.inc     —— Fix 037b/086: typed 字段表 + 窗体模块名 + Public 常量表（原 188~271 行）
//     driver_codegen_variant_funcs.inc       —— Fix 045: 返回 Variant 的项目函数集合（原 272~341 行）
//     driver_codegen_dup_module_vars.inc     —— Fix 093b: 跨模块同名模块级变量（原 342~370 行）
//     driver_codegen_module_loop.inc         —— 逐模块代码生成主循环（写 .h/.c，原 371~462 行）
//     driver_codegen_dll_typelib.inc         —— P6.6/P9: ActiveX DLL 模式 TypeLib 构建（原 463~698 行）
//     driver_codegen_dll_sync.inc            —— P6.6.6/M29: IID/CLSID/DISPID 符号同步（原 699~753 行）
//     driver_codegen_dll_entry.inc           —— generateDllEntry 与 dll_entry.c 写出 + 返回（原 754~772 行）
// 9 个 .inc 不是独立编译单元（不登记 CMakeLists），单独 include 会编译不过；片段内容逐行未改，零行为改动。

#include "driver/driver.hpp"
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

#ifdef _WIN32
#include <windows.h>
#endif

namespace vb6c3 {

// pathToUtf8/utf8ToWide/utf8ToPath moved to common/encoding.hpp

Driver::Driver() : diag_(std::make_unique<Diagnostics>()) {}
Driver::~Driver() = default;

// opt4: AST 引用收集器 — 扫描模块 AST, 收集"实际引用的外部模块"(小写)。
// 注: 符号表经 runCrossModuleResolution 全量注入, getExternalModuleNames()
// 返回所有其他模块(而非实际引用), 因此裁剪 include 需基于 AST 实际引用。
class ExternalRefCollector : public ASTVisitor {
public:
    SymbolTable& symTab;
    const std::unordered_set<std::string>& externalModules;
    std::unordered_set<std::string>& refd;  // 输出: 小写模块名

    ExternalRefCollector(SymbolTable& st, const std::unordered_set<std::string>& ext,
                         std::unordered_set<std::string>& out)
        : symTab(st), externalModules(ext), refd(out) {}

    void visit(IdentifierExpr& node) override {
        // 点号限定形式 Mod.var / Mod.UDT
        size_t dot = node.name.find('.');
        if (dot != std::string::npos) {
            std::string modPart = node.name.substr(0, dot);
            for (const auto& em : externalModules) {
                if (Symbol::toLower(em) == Symbol::toLower(modPart)) {
                    refd.insert(Symbol::toLower(em));
                    return;
                }
            }
            return;
        }
        // 普通名称: 查符号表定位归属模块
        if (Symbol* s = symTab.lookup(node.name)) {
            if (s->isExternal && !s->sourceModule.empty()) {
                refd.insert(Symbol::toLower(s->sourceModule));
            }
        }
    }

    void visit(SimpleTypeRef& node) override {
        size_t dot = node.name.find('.');
        if (dot != std::string::npos) {
            std::string modPart = node.name.substr(0, dot);
            for (const auto& em : externalModules) {
                if (Symbol::toLower(em) == Symbol::toLower(modPart)) {
                    refd.insert(Symbol::toLower(em));
                    return;
                }
            }
            return;
        }
        if (Symbol* s = symTab.lookup(node.name)) {
            if (s->isExternal && !s->sourceModule.empty()) {
                refd.insert(Symbol::toLower(s->sourceModule));
            }
        }
    }
};

bool Driver::runCodeGeneration(const CompileOptions& options, const std::string& outputDir) {
#include "driver/detail/driver_codegen_prelude.inc"
#include "driver/detail/driver_codegen_voidfield_scan.inc"
#include "driver/detail/driver_codegen_typedfield_scan.inc"
#include "driver/detail/driver_codegen_variant_funcs.inc"
#include "driver/detail/driver_codegen_dup_module_vars.inc"
#include "driver/detail/driver_codegen_module_loop.inc"
#include "driver/detail/driver_codegen_dll_typelib.inc"
#include "driver/detail/driver_codegen_dll_sync.inc"
#include "driver/detail/driver_codegen_dll_entry.inc"
}

} // namespace vb6c3