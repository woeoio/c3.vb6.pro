#include "semantics/semantic_analyzer.hpp"
#include <algorithm>
#include <cctype>
#include <tuple>
#include <initializer_list>
#include "semantics/semantic_analyzer_internal.h"

namespace vb6c3 {

// --- semantic_analyzer_decl.cpp: 声明 Visitor (Pass1 注册, Pass2 分析体) ---


// ============================================================
// 声明 Visitor (Pass1注册, Pass2分析体)
// ============================================================

void SemanticAnalyzer::visit(Module& node) {
    // 不在这里遍历, analyze()负责两遍逻辑
}

void SemanticAnalyzer::visit(SubDecl& node) {
    // 泛型模板 (tB, G2/G3): 模板本体不进符号表 (泛型器注入特化副本)
    if (!node.typeParams.empty()) return;
    if (pass_ == 1) {
        // Pass1: 注册Sub符号
        auto sym = std::make_unique<Symbol>(
            SymbolKind::Sub, node.name,
            Vb6Type::Void, node.loc, node.access
        );
        sym->isStatic = node.isStatic;

        // 注册参数
        for (auto& param : node.params) {
            ParameterInfo pi;
            pi.name = param->name;
            pi.type = resolveTypeOrDefault(param->name, param->asType.get());
            pi.typeRefName = simpleTypeRefName(param->asType.get());  // <vbeclipse>: 类名别丢
            pi.isByVal = param->isByVal;
            pi.isOptional = param->isOptional;
            pi.isParamArray = param->isParamArray;
            // P14.1.4: 传播Optional默认值
            if (param->isOptional) {
                pi.hasDefaultValue = (param->defaultValue != nullptr);
                pi.defaultValueExpr = evalOptionalDefault(param->defaultValue.get(), pi.type);
            }
            sym->params.push_back(std::move(pi));
        }

        // Fix 033: Pass1 填充 Class 符号的 memberParams (类型解析正确, 避免 class init phase 的 Variant 回归)
        if (currentModule_->isClassModule) {
            auto* classSym = symTab_.lookupModule(Symbol::toLower(currentModule_->moduleName));
            if (classSym && classSym->kind == SymbolKind::Class) {
                // Sub: 覆盖 (与 memberProcKinds Sub 行为一致; VB6 不允许 Sub 与同名 Property 共存)
                classSym->memberParams[Symbol::toLower(node.name)] = sym->params;
            }
        }

        // 重载分组资格与指纹 (O1); 无资格 (类模块/ParamArray) 时为空, 走旧路径
        sym->overloadFp = computeOverloadFp(*sym);

        symTab_.define(std::move(sym));
    } else {
        // Pass2: 分析过程体
        if (verbose_) std::cerr << "[Sem]   Sub: " << node.name << std::endl;
        // 重载组内按声明位置找回本变体 (裸键 head 可能是别的签名)
        auto* sym = symTab_.lookupModuleOverloadByLoc(node.name, node.loc);
        if (!sym) return;  // 注册失败则跳过

        // Fix 047: Re-resolve parameter types in Pass 2 after cross-module resolution.
        // Enum/UDT types from other modules are not visible during Pass 1, so params
        // declared as As SomeEnum get resolved to Variant in Pass 1. In Pass 2, after
        // runCrossModuleResolution() has injected external symbols, re-resolve to get
        // the correct type (e.g., Enum → Long). Also update classSym->memberParams.
        if (sym->params.size() == node.params.size()) {
            bool paramsChanged = false;
            for (size_t i = 0; i < node.params.size(); i++) {
                Vb6Type newType = resolveTypeOrDefault(node.params[i]->name, node.params[i]->asType.get());
                if (newType != sym->params[i].type) {
                    sym->params[i].type = newType;
                    paramsChanged = true;
                }
            }
            if (paramsChanged && currentModule_->isClassModule) {
                auto* classSym = symTab_.lookupModule(Symbol::toLower(currentModule_->moduleName));
                if (classSym && classSym->kind == SymbolKind::Class) {
                    classSym->memberParams[Symbol::toLower(node.name)] = sym->params;
                }
            }
        }

        currentProc_ = sym;
        symTab_.pushScope(ScopeKind::Procedure);

        // 注册参数到过程作用域
        for (auto& param : node.params) {
            auto paramSym = std::make_unique<Symbol>(
                SymbolKind::Parameter, param->name,
                resolveTypeOrDefault(param->name, param->asType.get()),
                param->loc, AccessLevel::Private
            );
            // tB B08c: 记一条 `As <类型>` 原文, 供 Protected 越权判定认接收者类
            if (param->asType && param->asType->kind == ASTNodeKind::SimpleTypeRef) {
                paramSym->srcTypeName = static_cast<SimpleTypeRef*>(param->asType.get())->name;
            }
            symTab_.define(std::move(paramSym));
        }

        // 分析过程体
        analyzeStmtList(node.body);

        symTab_.popScope();
        // P12.5: GoSub标签边界验证（GoSub目标必须在当前过程内）
        for (auto& [gosubLabel, gosubLoc] : gosubTargetLabels_) {
            bool found = false;
            for (auto& declLabel : declaredLabels_) {
                if (Symbol::toLower(declLabel) == Symbol::toLower(gosubLabel)) {
                    found = true; break;
                }
            }
            if (!found) {
                diag_.error(DiagnosticID::SemUndeclaredIdentifier, gosubLoc,
                    "GoSub target label '" + gosubLabel + "' not found in current procedure");
            }
        }
        declaredLabels_.clear();
        gosubTargetLabels_.clear();
        currentProc_ = nullptr;
    }
}

void SemanticAnalyzer::visit(FunctionDecl& node) {
    // 泛型模板 (tB, G2/G3): 见 visit(SubDecl) 同注释
    if (!node.typeParams.empty()) return;
    if (pass_ == 1) {
        // Pass1: 注册Function符号
        Vb6Type retType = resolveTypeOrDefault(node.name, node.returnType.get());
        auto sym = std::make_unique<Symbol>(
            SymbolKind::Function, node.name,
            retType, node.loc, node.access
        );
        sym->isStatic = node.isStatic;

        // Fix 015: 记录Function的返回类型名(若返回类/UDT等命名类型)
        // 用于cgen解析 method chaining: db.Sql(s).Exec(...) 链式调用时
        // 需要根据 Sql 的返回类名(cDataBase)分发 .Exec → vb6_cDataBase_Exec
        if (node.returnType && node.returnType->kind == ASTNodeKind::SimpleTypeRef) {
            sym->variableTypeName = static_cast<SimpleTypeRef*>(node.returnType.get())->name;
        }

        // 注册参数
        for (auto& param : node.params) {
            ParameterInfo pi;
            pi.name = param->name;
            pi.type = resolveTypeOrDefault(param->name, param->asType.get());
            pi.typeRefName = simpleTypeRefName(param->asType.get());  // <vbeclipse>: 类名别丢
            pi.isByVal = param->isByVal;
            pi.isOptional = param->isOptional;
            pi.isParamArray = param->isParamArray;
            // P14.1.4: 传播Optional默认值
            if (param->isOptional) {
                pi.hasDefaultValue = (param->defaultValue != nullptr);
                pi.defaultValueExpr = evalOptionalDefault(param->defaultValue.get(), pi.type);
            }
            sym->params.push_back(std::move(pi));
        }

        // Fix 033: Pass1 填充 Class 符号的 memberParams (类型解析正确)
        if (currentModule_->isClassModule) {
            auto* classSym = symTab_.lookupModule(Symbol::toLower(currentModule_->moduleName));
            if (classSym && classSym->kind == SymbolKind::Class) {
                // Function: 覆盖 (与 memberProcKinds Function 行为一致)
                classSym->memberParams[Symbol::toLower(node.name)] = sym->params;
            }
        }

        sym->overloadFp = computeOverloadFp(*sym);

        symTab_.define(std::move(sym));
    } else {
        // Pass2: 分析过程体
        if (verbose_) std::cerr << "[Sem]   Function: " << node.name << std::endl;
        auto* sym = symTab_.lookupModuleOverloadByLoc(node.name, node.loc);
        if (!sym) return;

        // Fix 047: Re-resolve parameter types AND return type in Pass 2.
        // Enum/UDT types from other modules are not visible during Pass 1.
        if (sym->params.size() == node.params.size()) {
            bool paramsChanged = false;
            for (size_t i = 0; i < node.params.size(); i++) {
                Vb6Type newType = resolveTypeOrDefault(node.params[i]->name, node.params[i]->asType.get());
                if (newType != sym->params[i].type) {
                    sym->params[i].type = newType;
                    paramsChanged = true;
                }
            }
            // Also re-resolve return type
            Vb6Type newRetType = resolveTypeOrDefault(node.name, node.returnType.get());
            if (newRetType != sym->type) {
                sym->type = newRetType;
                paramsChanged = true;
            }
            if (paramsChanged && currentModule_->isClassModule) {
                auto* classSym = symTab_.lookupModule(Symbol::toLower(currentModule_->moduleName));
                if (classSym && classSym->kind == SymbolKind::Class) {
                    classSym->memberParams[Symbol::toLower(node.name)] = sym->params;
                }
            }
        }

        currentProc_ = sym;
        symTab_.pushScope(ScopeKind::Procedure);

        // 注册参数
        for (auto& param : node.params) {
            auto paramSym = std::make_unique<Symbol>(
                SymbolKind::Parameter, param->name,
                resolveTypeOrDefault(param->name, param->asType.get()),
                param->loc, AccessLevel::Private
            );
            // tB B08c: 记一条 `As <类型>` 原文, 供 Protected 越权判定认接收者类
            if (param->asType && param->asType->kind == ASTNodeKind::SimpleTypeRef) {
                paramSym->srcTypeName = static_cast<SimpleTypeRef*>(param->asType.get())->name;
            }
            symTab_.define(std::move(paramSym));
        }

        analyzeStmtList(node.body);

        symTab_.popScope();
        // P12.5: GoSub标签边界验证（GoSub目标必须在当前过程内）
        for (auto& [gosubLabel, gosubLoc] : gosubTargetLabels_) {
            bool found = false;
            for (auto& declLabel : declaredLabels_) {
                if (Symbol::toLower(declLabel) == Symbol::toLower(gosubLabel)) {
                    found = true; break;
                }
            }
            if (!found) {
                diag_.error(DiagnosticID::SemUndeclaredIdentifier, gosubLoc,
                    "GoSub target label '" + gosubLabel + "' not found in current procedure");
            }
        }
        declaredLabels_.clear();
        gosubTargetLabels_.clear();
        currentProc_ = nullptr;
    }
}

void SemanticAnalyzer::visit(PropertyDecl& node) {
    // 泛型模板 (tB, G2/G3): 见 visit(SubDecl) 同注释
    if (!node.typeParams.empty()) return;
    if (pass_ == 1) {
        SymbolKind sk;
        switch (node.propKind) {
            case ProcKind::PropertyGet:  sk = SymbolKind::PropertyGet; break;
            case ProcKind::PropertyLet:  sk = SymbolKind::PropertyLet; break;
            case ProcKind::PropertySet:  sk = SymbolKind::PropertySet; break;
            default:                     sk = SymbolKind::PropertyGet; break;
        }

        Vb6Type retType = resolveTypeOrDefault(node.name, node.returnType.get());
        auto sym = std::make_unique<Symbol>(sk, node.name, retType, node.loc, node.access);

        // Fix 015: 记录 Property Get 的返回类型名(若返回类/UDT等命名类型)
        // 用于 cgen 解析 method chaining: obj.GetContainer().Method() 链式调用
        if (sk == SymbolKind::PropertyGet
            && node.returnType && node.returnType->kind == ASTNodeKind::SimpleTypeRef) {
            sym->variableTypeName = static_cast<SimpleTypeRef*>(node.returnType.get())->name;
        }

        for (auto& param : node.params) {
            ParameterInfo pi;
            pi.name = param->name;
            pi.type = resolveTypeOrDefault(param->name, param->asType.get());
            pi.typeRefName = simpleTypeRefName(param->asType.get());  // <vbeclipse>: 类名别丢
            pi.isByVal = param->isByVal;
            pi.isOptional = param->isOptional;
            pi.isParamArray = param->isParamArray;
            // P14.1.4: 传播Optional默认值
            if (param->isOptional) {
                pi.hasDefaultValue = (param->defaultValue != nullptr);
                pi.defaultValueExpr = evalOptionalDefault(param->defaultValue.get(), pi.type);
            }
            sym->params.push_back(std::move(pi));
        }

        // Fix 033: Pass1 填充 Class 符号的 memberParams (类型解析正确, 按 Get > Function > Sub > Let > Set 优先级)
        if (currentModule_->isClassModule) {
            auto* classSym = symTab_.lookupModule(Symbol::toLower(currentModule_->moduleName));
            if (classSym && classSym->kind == SymbolKind::Class) {
                std::string lower = Symbol::toLower(node.name);
                // memberProcKinds 在 class init 阶段已确定最终胜出者, 据此判定当前 Property 是否为胜出者
                auto it = classSym->memberProcKinds.find(lower);
                bool wins = false;
                if (node.propKind == ProcKind::PropertyGet) {
                    wins = true;  // Get 总是胜出
                } else if (node.propKind == ProcKind::PropertyLet) {
                    wins = (it != classSym->memberProcKinds.end()
                            && it->second == ProcKind::PropertyLet);
                } else if (node.propKind == ProcKind::PropertySet) {
                    wins = (it != classSym->memberProcKinds.end()
                            && it->second == ProcKind::PropertySet);
                }
                if (wins) {
                    classSym->memberParams[lower] = sym->params;
                }
                // Fix 091a: 写方向参数表独立填充 — 同类内 Let/Set 各自唯一, 不受
                // 读上下文优先级 (Get 有参遮蔽 Let 末参 value) 影响.
                if (node.propKind == ProcKind::PropertyLet) {
                    classSym->memberLetParams[lower] = sym->params;
                } else if (node.propKind == ProcKind::PropertySet) {
                    classSym->memberSetParams[lower] = sym->params;
                }
            }
        }

        symTab_.define(std::move(sym));
    } else {
        // Pass2: 分析过程体 — 使用lookupModuleByKind精确查找同名Property
        SymbolKind sk;
        switch (node.propKind) {
            case ProcKind::PropertyGet:  sk = SymbolKind::PropertyGet; break;
            case ProcKind::PropertyLet:  sk = SymbolKind::PropertyLet; break;
            case ProcKind::PropertySet:  sk = SymbolKind::PropertySet; break;
            default:                     sk = SymbolKind::PropertyGet; break;
        }
        auto* sym = symTab_.lookupModuleByKind(node.name, sk);
        if (!sym) return;

        // Fix 047: Re-resolve parameter types AND return type in Pass 2.
        // Enum/UDT types from other modules are not visible during Pass 1.
        if (sym->params.size() == node.params.size()) {
            bool paramsChanged = false;
            for (size_t i = 0; i < node.params.size(); i++) {
                Vb6Type newType = resolveTypeOrDefault(node.params[i]->name, node.params[i]->asType.get());
                if (newType != sym->params[i].type) {
                    sym->params[i].type = newType;
                    paramsChanged = true;
                }
            }
            // Also re-resolve return type (for PropertyGet)
            if (sk == SymbolKind::PropertyGet) {
                Vb6Type newRetType = resolveTypeOrDefault(node.name, node.returnType.get());
                if (newRetType != sym->type) {
                    sym->type = newRetType;
                    paramsChanged = true;
                }
            }
            if (paramsChanged && currentModule_->isClassModule) {
                auto* classSym = symTab_.lookupModule(Symbol::toLower(currentModule_->moduleName));
                if (classSym && classSym->kind == SymbolKind::Class) {
                    std::string lower = Symbol::toLower(node.name);
                    auto it = classSym->memberProcKinds.find(lower);
                    bool wins = false;
                    if (node.propKind == ProcKind::PropertyGet) {
                        wins = true;
                    } else if (node.propKind == ProcKind::PropertyLet) {
                        wins = (it != classSym->memberProcKinds.end()
                                && it->second == ProcKind::PropertyLet);
                    } else if (node.propKind == ProcKind::PropertySet) {
                        wins = (it != classSym->memberProcKinds.end()
                                && it->second == ProcKind::PropertySet);
                    }
                    if (wins) {
                        classSym->memberParams[lower] = sym->params;
                    }
                    // Fix 091a: 写方向参数表同步 (Pass2 重解析后类型修正)
                    if (node.propKind == ProcKind::PropertyLet) {
                        classSym->memberLetParams[lower] = sym->params;
                    } else if (node.propKind == ProcKind::PropertySet) {
                        classSym->memberSetParams[lower] = sym->params;
                    }
                }
            }
        }

        currentProc_ = sym;
        symTab_.pushScope(ScopeKind::Procedure);

        for (auto& param : node.params) {
            auto paramSym = std::make_unique<Symbol>(
                SymbolKind::Parameter, param->name,
                resolveTypeOrDefault(param->name, param->asType.get()),
                param->loc, AccessLevel::Private
            );
            // tB B08c: 记一条 `As <类型>` 原文, 供 Protected 越权判定认接收者类
            if (param->asType && param->asType->kind == ASTNodeKind::SimpleTypeRef) {
                paramSym->srcTypeName = static_cast<SimpleTypeRef*>(param->asType.get())->name;
            }
            symTab_.define(std::move(paramSym));
        }

        analyzeStmtList(node.body);

        symTab_.popScope();
        // P12.5: GoSub标签边界验证（GoSub目标必须在当前过程内）
        for (auto& [gosubLabel, gosubLoc] : gosubTargetLabels_) {
            bool found = false;
            for (auto& declLabel : declaredLabels_) {
                if (Symbol::toLower(declLabel) == Symbol::toLower(gosubLabel)) {
                    found = true; break;
                }
            }
            if (!found) {
                diag_.error(DiagnosticID::SemUndeclaredIdentifier, gosubLoc,
                    "GoSub target label '" + gosubLabel + "' not found in current procedure");
            }
        }
        declaredLabels_.clear();
        gosubTargetLabels_.clear();
        currentProc_ = nullptr;
    }
}
} // namespace vb6c3
