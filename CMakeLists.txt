cmake_minimum_required(VERSION 3.20)
# cache-bust 2026-10-01-19 (VBFlexGridDemo / a280e00f emit-shape 排疑, 一次票)

# Version single source of truth: VERSION file at the repo root (stripped).
file(READ "${CMAKE_CURRENT_LIST_DIR}/VERSION" VB6C3_VERSION_RAW)
string(STRIP "${VB6C3_VERSION_RAW}" VB6C3_VERSION)
if(VB6C3_VERSION STREQUAL "")
    message(FATAL_ERROR "VERSION file is empty: ${CMAKE_CURRENT_LIST_DIR}/VERSION")
endif()

project(vb6c3 VERSION ${VB6C3_VERSION} LANGUAGES CXX)

# Re-run CMake automatically when VERSION changes, so the injected macro stays in sync.
set_property(DIRECTORY APPEND PROPERTY CMAKE_CONFIGURE_DEPENDS "${CMAKE_CURRENT_LIST_DIR}/VERSION")

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_EXPORT_COMPILE_COMMANDS ON)

# MSVC: use UTF-8 source charset (files are UTF-8 without BOM)
if(MSVC)
    add_compile_options(/utf-8)
endif()

# === LLVM backend (disabled by default, decision 006: C code gen + MSVC) ===
option(VB6C3_ENABLE_LLVM "Enable LLVM backend" OFF)
set(VB6C3_LLVM_DIR "" CACHE PATH "LLVM installation directory (for tool mode)")

if(VB6C3_ENABLE_LLVM)
    set(VB6C3_LLVM_API_MODE OFF)
    find_package(LLVM QUIET CONFIG)
    if(LLVM_FOUND)
        include_directories(${LLVM_INCLUDE_DIRS})
        separate_arguments(LLVM_DEFINITIONS_LIST NATIVE_COMMAND ${LLVM_DEFINITIONS})
        add_definitions(${LLVM_DEFINITIONS_LIST})
        llvm_map_components_to_libnames(llvm_libs core support orcjit native)
        set(VB6C3_LLVM_API_MODE ON)
        message(STATUS "LLVM API mode: version ${LLVM_PACKAGE_VERSION}")
    else()
        if(VB6C3_LLVM_DIR)
            set(LLVM_CLANG "${VB6C3_LLVM_DIR}/bin/clang.exe")
        else()
            find_program(LLVM_CLANG clang)
        endif()
        if(EXISTS "${LLVM_CLANG}")
            get_filename_component(LLVM_BIN_DIR "${LLVM_CLANG}" DIRECTORY)
            get_filename_component(VB6C3_LLVM_DIR_CALC "${LLVM_BIN_DIR}" DIRECTORY)
            message(STATUS "LLVM tool mode: clang found at ${LLVM_CLANG}")
            if(NOT VB6C3_LLVM_DIR)
                set(VB6C3_LLVM_DIR "${VB6C3_LLVM_DIR_CALC}")
            endif()
        else()
            message(WARNING "LLVM enabled but neither CMake config nor clang found")
            set(VB6C3_ENABLE_LLVM OFF)
        endif()
    endif()
else()
    message(STATUS "LLVM backend disabled - using C code generation + MSVC")
endif()

include_directories(${CMAKE_SOURCE_DIR}/src)

# ============================================================
# Decision 008: Library-first architecture
# vb6c3-core = lexer/parser/AST/semantics/symbols/types/diagnostics
# c3.exe = thin CLI driver
# ============================================================

# --- Core library vb6c3-core ---
add_library(vb6c3-core STATIC
    src/common/diagnostics.cpp
    src/common/source_manager.cpp
    src/common/encoding.cpp
    src/common/sha1.cpp
    src/lexer/lexer.cpp
    src/lexer/lexer_keywords.cpp
    src/lexer/lexer_number.cpp
    src/lexer/lexer_literal.cpp
    src/lexer/token.cpp
    src/preprocessor/preprocessor.cpp
    src/preprocessor/preprocessor_cond.cpp
    src/preprocessor/preprocessor_expr.cpp
    src/parser/parser.cpp
    src/parser/parser_helpers.cpp
    src/parser/parser_module.cpp
    src/parser/parser_expr.cpp
    src/parser/parser_expr_postfix.cpp
    src/parser/parser_decl.cpp
    src/parser/parser_decl_var.cpp
    src/parser/parser_interface.cpp   # tB 扩展: Interface 契约块 (ai/022 B01)
    # 语句解析家族
    src/parser/stmt/parser_stmt.cpp
    src/parser/stmt/parser_stmt_control.cpp
    src/parser/stmt/parser_stmt_jump.cpp
    src/parser/stmt/parser_stmt_assign.cpp
    src/parser/stmt/parser_stmt_io.cpp
    src/ast/ast.cpp
    src/ast/ast_visitor.cpp
    src/ast/ast_clone.cpp
    src/ast/ast_printer.cpp
    src/semantics/symbol_table.cpp
    src/semantics/coclass_identity.cpp    # tB 扩展: CoClass 身份求解唯一函数 (ai/026 三节, B11/C02)
    src/semantics/type_system.cpp
    src/semantics/semantic_analyzer.cpp
    src/semantics/semantic_analyzer_dispatch.cpp
    src/semantics/semantic_analyzer_register.cpp
    src/semantics/semantic_analyzer_typeref.cpp
    src/semantics/semantic_analyzer_iface.cpp
    src/semantics/semantic_analyzer_decl.cpp
    src/semantics/semantic_analyzer_decl_type.cpp
    src/semantics/semantic_analyzer_stmt.cpp
    src/semantics/semantic_analyzer_expr.cpp
    src/semantics/semantic_analyzer_util.cpp
    src/semantics/semantic_analyzer_builtin.cpp
    src/project/vbp_parser.cpp
    src/driver/driver_pack.cpp
    src/driver/res_inventory.cpp  # ai/029 C29-M: .res 资源目录解析（清单让位判据）
    src/project/package_manifest.cpp   # ai/023 S01: package.c3d 清单解析
    src/project/frm_parser.cpp
    src/project/frm_parser_util.cpp
    src/project/frx_reader.cpp
    src/project/frx_extract.cpp        # Fix 195: .frx 设计期取值 -> VB 代码 (--extract-frx)
    src/com/typelib_parser.cpp
    src/com/typelib_parser_parse.cpp
    src/com/typelib_parser_desc.cpp
    src/typelib/typelib_builder.cpp
    src/typelib/typelib_builder_coclass.cpp
)
target_compile_definitions(vb6c3-core PRIVATE UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
# TypeLib parser needs ole32/oleaut32 (COM API)
target_link_libraries(vb6c3-core PRIVATE ole32 oleaut32 uuid)

# --- C code generation backend ---
add_library(vb6c3-cgen STATIC
    src/backend/cgen_base.cpp
    # base 家族（2026-09-17 纯搬移自 cgen_base.cpp）
    src/backend/cgen_base_type.cpp
    src/backend/cgen_base_naming.cpp
    # 表达式家族
    src/backend/expr/cgen_expr.cpp
    src/backend/expr/cgen_expr_ident.cpp
    src/backend/expr/cgen_expr_member.cpp
    src/backend/expr/cgen_expr_call.cpp
    src/backend/expr/cgen_expr_binary.cpp
    src/backend/expr/cgen_expr_binary_util.cpp
    src/backend/expr/cgen_expr_array.cpp
    src/backend/expr/cgen_expr_with.cpp
    # 语句家族
    src/backend/stmt/cgen_stmt.cpp
    src/backend/stmt/cgen_stmt_asm.cpp
    src/backend/stmt/cgen_assign.cpp
    src/backend/stmt/cgen_setlet.cpp
    src/backend/stmt/cgen_control.cpp
    src/backend/stmt/cgen_loop.cpp
    src/backend/stmt/cgen_select.cpp
    src/backend/stmt/cgen_with.cpp
    src/backend/stmt/cgen_jumps.cpp
    src/backend/stmt/cgen_call.cpp
    src/backend/stmt/cgen_redim.cpp
    src/backend/stmt/cgen_file_io.cpp
    # 声明家族
    src/backend/decl/cgen_decl.cpp
    src/backend/decl/cgen_decl_proc.cpp
    src/backend/decl/cgen_decl_func.cpp
    src/backend/decl/cgen_decl_var.cpp
    src/backend/decl/cgen_decl_api.cpp
    src/backend/decl/cgen_decl_prop.cpp
    src/backend/decl/cgen_localdecl.cpp
    src/backend/decl/cgen_localdecl_hoist.cpp
    # 模块级产物（COM 服务器 / 窗体）
    src/backend/module/cgen_com.cpp
    src/backend/module/cgen_iface_vtbl.cpp
    src/backend/module/cgen_inherit.cpp
    src/backend/module/cgen_com_events.cpp
    src/backend/module/cgen_delegate.cpp
    src/backend/module/cgen_form.cpp
    src/backend/module/cgen_form_menu.cpp
    src/backend/cgen_util.cpp
    src/backend/cgen_util_type.cpp
    src/backend/cgen_util_scan.cpp
    src/backend/cgen_util_com.cpp
    src/backend/cgen_util_comwrite.cpp
    src/backend/cgen_util_classcall.cpp
    src/backend/cgen_util_classtype.cpp
    src/backend/cgen_util_ctrl.cpp
    src/backend/msvc_driver.cpp
    src/backend/msvc_driver_discovery.cpp
    src/backend/msvc_driver_incremental.cpp
)
target_compile_definitions(vb6c3-cgen PRIVATE UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
target_link_libraries(vb6c3-cgen PRIVATE vb6c3-core)

# --- LLVM backend (optional) ---
if(VB6C3_ENABLE_LLVM)
    add_library(vb6c3-llvm STATIC
        src/backend/llvm_backend.cpp
        src/backend/pe_writer.cpp
    )
    if(VB6C3_LLVM_API_MODE)
        target_link_libraries(vb6c3-llvm PRIVATE ${llvm_libs})
    else()
        target_compile_definitions(vb6c3-llvm PRIVATE VB6C3_LLVM_TOOL_MODE VB6C3_LLVM_DIR="${VB6C3_LLVM_DIR}")
    endif()
    target_link_libraries(vb6c3-llvm PRIVATE vb6c3-core)
endif()

# --- IR layer (preserved) ---
add_library(vb6c3-ir STATIC
    src/ir/ir.cpp
    src/ir/ir_emitter.cpp
)
target_compile_definitions(vb6c3-ir PRIVATE UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
target_link_libraries(vb6c3-ir PRIVATE vb6c3-core)

# --- Compiler CLI: C3.exe ---
# P10 (restored): c3rtl.rc embeds RTL .h/.c SOURCE files as RCDATA resources.
# RTL is compiled by MSVC at user-program build time (no pre-compiled .lib).
add_executable(c3
    src/driver/main.cpp
    src/driver/driver.cpp
    src/driver/driver_args.cpp
    src/driver/driver_compile.cpp
    src/driver/driver_frontend.cpp
    src/driver/driver_semantics.cpp
    src/driver/driver_generics.cpp
    src/driver/driver_interface.cpp
    src/driver/coclass_activate.cpp
    src/driver/driver_classchain.cpp
    src/driver/driver_crossmod.cpp
    src/driver/driver_link.cpp
    src/driver/driver_staticlib.cpp
    src/driver/static_lib.cpp
    src/driver/coff_archive.cpp
    src/driver/rtl_embedded.cpp
    src/driver/c3rtl.rc
)
set_target_properties(c3 PROPERTIES OUTPUT_NAME C3)

# Fix 092z-2: CMake 只把 .rc 自身当依赖 —— 改了 src/rtl/core 下的嵌入源之后
# .res 不会重编, C3.exe 会继续嵌入旧文件。这里显式声明嵌入文件为 .rc 的依赖
# (99 个资源: 25 头 + 61 源 + 13 片段, 全部架构无关；与 c3rtl.rc 一一对应)。
set(C3RTL_EMBEDDED_FILES
    # vb6rtl
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_base.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_bstr.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_variant.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_builtin.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_array.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_class_com.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_runtime.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_crash.h
    # vb6rtl.c 按家族拆分 (2026-09-17): 13 个族实现 + format 的 6 个函数体片段
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_string.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_format.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/detail/vb6rtl_format_extract.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/detail/vb6rtl_format_parse.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/detail/vb6rtl_format_numeric_pre.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/detail/vb6rtl_format_numeric_body.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/detail/vb6rtl_format_numeric_tail.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/detail/vb6rtl_format_string.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_conv.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_misc.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_system.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_compat.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_date.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_array.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_file.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_paramarray.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_financial.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_com.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_registry.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6rtl/vb6rtl_userctl.h   # Fix 105 (补登记)
    # vb6com
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com_internal.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com_invoke.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com_pack.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com_wrap.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com_sink.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com_foreach.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com_collection.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6com/vb6com_collection_enum.c
    # vb6comserver
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6comserver/vb6comserver.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6comserver/vb6comserver.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6comserver/vb6comserver_internal.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6comserver/vb6comserver_obj.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6comserver/vb6comserver_factory.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6comserver/vb6comserver_cp.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6comserver/vb6comserver_pci.c
    # vb6forms
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_uc.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_internal.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_ctrl.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_list.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_style.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_scroll.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_picture.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_picture_prop.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_ctrlarr.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_webview.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_widget.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_widget_prop.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_shape.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_progress.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_imagelist.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_statusbar.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_sstab.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_oledd.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_listview.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_memberobj.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_treeview.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_olecon.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_data.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_toolbar.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_dtpicker.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_monthview.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_richtextbox.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_winsock.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_slider.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_axsite.c
${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_axcontainer.c   # Fix 148/160 补登记 (遗漏, 改它不会触发 .rc 重编)
    # vb6forms_axsite 按族细分 (2026-09-20): 内部头 + axsite/ 下 5 个族编译单元
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_axsite_internal.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/axsite/ax_site.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/axsite/ax_site_ext.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/axsite/ax_propbag.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/axsite/ax_load.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/axsite/ax_host.c
    # vb6forms_uc 按族细分 (2026-09-19): 内部头 + 6 族编译单元 + 7 个函数体片段
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_uc_internal.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/uc_host.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/uc_host_window.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/uc_hostmodel.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/uc_controls.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/uc_collection.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/uc_debug.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/uc_propbag.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/detail/uc_host_create.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/detail/uc_hostmodel_getprop.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/detail/uc_hostmodel_setprop.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/detail/uc_hostmodel_call.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/detail/uc_collection_api.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/detail/uc_debug_dib.inc
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/uc/detail/uc_debug_composite.inc
    # vb6forms.h 按功能域拆分 (2026-09-17): 伞头 + 9 个子头
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_window.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_prop.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_prop_ctrl.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_prop_pic.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_prop_form.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_ctrlarr.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_mdi.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_webview.h
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/vb6forms/vb6forms_controls.h
    # DI 转发桩家族（手写桩 + 按 Lib 家族生成的转发桩）
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/vb6_di_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/vb6_di_win32_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/vb6_di_user32_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/gdiplus/vb6_di_gdiplus_draw_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/gdiplus/vb6_di_gdiplus_text_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/gdiplus/vb6_di_gdiplus_image_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/gdiplus/vb6_di_gdiplus_brush_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/gdiplus/vb6_di_gdiplus_path_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/vb6_di_crypto_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/vb6_di_com_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/vb6_di_net_stubs.c
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/vb6_di_shell_stubs.c
    # Fix 164: 此列表是 c3rtl.rc 的 OBJECT_DEPENDS —— 漏项 = 改该 RTL 源文件**不会**重新
    # 内嵌, ninja 报 "no work to do" 却仍打印 [OK] Build succeeded, 于是一边改桩体一边
    # 用着旧 exe。unknown 族 (c3rtl.rc 资源 199) 是 160y 新增时只加了 .rc 条目忘了这里,
    # 现共 100 个 RCDATA / 99 个依赖。
    ${CMAKE_CURRENT_SOURCE_DIR}/src/rtl/core/di/vb6_di_unknown_stubs.c
)
set_source_files_properties(src/driver/c3rtl.rc PROPERTIES OBJECT_DEPENDS "${C3RTL_EMBEDDED_FILES}")
target_link_libraries(c3 PRIVATE vb6c3-core vb6c3-cgen)
if(WIN32)
    # 崩溃栈追踪 (C3_CRASH_TRACE=1) 需要 dbghelp
    target_link_libraries(c3 PRIVATE dbghelp)
endif()
if(WIN32)
    # Fix 212: 二元表达式的 codegen 是**按运算符个数递归**的 (emitExpr → visit(BinaryExpr)
    # → emitExpr(左))，所以栈深 ∝ 表达式里的运算符个数。MSVC 默认栈预留只有 1MB，实测
    # 约 49 层二元嵌套就爆 (0xC00000FD STATUS_STACK_OVERFLOW，实测 48 层过、50 层崩)。
    # 真实工程会超过这个数: ComCtlsDemo 的 Builds/RichTextBox/RichTextBox.ctl:4664 是一张
    # 本地化字符串表 (`Text = VBA.Choose(i, "&" & ChrW(&H391&) & … )`)，单条表达式 326 个
    # `&`; 该工程另有 248 / 175 / 131 / 116 个运算符的几条。1MB 栈下 C3 在 stage4 codegen
    # 的第 62 个模块 (RichTextBox) 直接崩，连 cl.exe 都到不了 —— 表现为"编译到一半无声中断"。
    # 预留只占虚拟地址空间、提交按需增长，故取 64MB (≈3000 层余量，覆盖语料实测最大值 326)。
    target_link_options(c3 PRIVATE "/STACK:67108864")
endif()
if(VB6C3_ENABLE_LLVM)
    target_link_libraries(c3 PRIVATE vb6c3-llvm)
else()
    target_compile_definitions(c3 PRIVATE VB6C3_NO_LLVM)
endif()
target_compile_definitions(c3 PRIVATE UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
# Version string injected into driver.cpp (source: VERSION file, do not hardcode here).
target_compile_definitions(c3 PRIVATE VB6C3_VERSION_STRING="${VB6C3_VERSION}")
