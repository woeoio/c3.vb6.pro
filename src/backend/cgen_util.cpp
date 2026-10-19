#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_util.cpp: ActiveX DLL 入口代码生成 (generateDllEntry) ---

// 2026-09-17 拆分：原 699 行单文件（函数体 685 行）拆为
//   本文件（伞文件）                        —— 前置说明 + generateDllEntry 函数骨架
//   detail/cgen_util_dllentry_prelude.inc   —— 局部结构体 CoClassInfo + 哈希/格式化辅助 lambda（原 12~72 行）
//   detail/cgen_util_dllentry_collect.inc   —— coclass 收集、空类桩、前向声明与 IDispatch 桥接（原 73~397 行）
//   detail/cgen_util_dllentry_tables.inc    —— 方法描述表、IID/事件常量表与 coclass 描述表（原 398~623 行）
//   detail/cgen_util_dllentry_exports.inc   —— DLL 导出函数（原 624~696 行）
// 四个 .inc 是「函数体片段」，在 generateDllEntry() 函数体内被 #include（C++ 允许），故不用 .cpp/.hpp 后缀 ——
// 它们不是独立编译单元，单独 include 会编译不过。片段局部 lambda 与 CoClassInfo 原样不动，逐行未改 → 零行为改动。

// P6.4+: 默认实例类注册 (见 cgen.hpp 声明). 若 name 是 VB_PredeclaredId=True 的类
// 模块名, 把 lower 注册到 knownClassVars_ 使其按"类实例"精确解析成员, 返回类
// 规范名供生成 vb6_cls_X_Default() 对象表达式; 非默认实例类返回空串.
std::string CCodeGen::registerDefaultInstanceClass(const std::string& name) {
    std::string cls = defaultInstanceClassName(name);
    if (cls.empty()) return "";
    std::string lower = name;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    knownClassVars_[lower] = cls;
    return cls;
}

// 账 #305 §B218 (甲): 工程类形参登记收成一处 —— 按名那张表 (成员读要它) 与
// 「这一档槽在过程出口/覆写/置空都不许动计数」的名册必须同时命中, 分两处写就会长漏。
void CCodeGen::registerClassFormal(const std::string& pLower, const std::string& clsName) {
    knownClassVars_[pLower] = clsName;
    procClassFormals_.insert(pLower);
}

std::string CCodeGen::generateDllEntry(const std::string& progId, const std::vector<SymbolTable*>& allSymTabs, bool includeDllExports) {
#include "backend/detail/util/cgen_util_dllentry_prelude.inc"
#include "backend/detail/util/cgen_util_dllentry_collect.inc"
#include "backend/detail/util/cgen_util_dllentry_tables.inc"
#include "backend/detail/util/cgen_util_dllentry_exports.inc"
}
} // namespace vb6c3
