#include "semantics/semantic_analyzer.hpp"
#include <algorithm>
#include <cctype>
#include <tuple>
#include <initializer_list>
#include "semantics/semantic_analyzer_internal.h"

namespace vb6c3 {

// --- semantic_analyzer_decl_type.cpp: 类型/枚举/Declare/事件/常量/变量 声明 Visitor ---
// 由 src/semantics/semantic_analyzer_decl.cpp 拆出（2026-09-17），纯搬移、零行为改动。

// <vbeclipse>: 见 semantic_analyzer_internal.h 的声明处注释。类型引用的载体就是
// ASTNode（没有 TypeRef 基类）, 只有 SimpleTypeRef 带名字, 其余形态返回空串。
std::string simpleTypeRefName(const ASTNode* typeRef) {
    if (!typeRef || typeRef->kind != ASTNodeKind::SimpleTypeRef) return std::string();
    return static_cast<const SimpleTypeRef*>(typeRef)->name;
}


// Fix 191: 枚举成员值的整数常量求值.
// 原实现只处理「字面量」与「-字面量」两种形态, 其余表达式一律静默忽略, 成员值
// 退化成「上一个成员值 + 1」. 典型受害写法 `ucsSfdAll = 2 ^ 6 - 1` (= 63):
//   ucsSfdRead=2^0, Write=2^1, Oob=2^2, Accept=2^3, Connect=2^4, Close=2^5,
//   ucsSfdAll=2^6-1 → 依次退化为 0,1,2,3,4,5,6 → ucsSfdAll 变成 6 (缺 FD_ACCEPT 位).
// 后果: `Optional ByVal EventMask As ... = ucsSfdAll` 的缺省值被写成 6,
// WSAAsyncSelect 拿不到 FD_ACCEPT → 服务器永不 accept 连接.
// 运算语义与后端 CCodeGen::tryEvalConstInt 保持一致 ('/' 为浮点除法故不求值).
static bool evalEnumMemberConstInt(ASTNode* expr, int64_t& result, const Scope* scope = nullptr,
                                   const SymbolTable* st = nullptr,
                                   const std::map<std::string, int64_t>* localConsts = nullptr) {
    if (!expr) return false;

    switch (expr->kind) {
    case ASTNodeKind::LiteralExpr: {
        auto* lit = static_cast<LiteralExpr*>(expr);
        switch (lit->literalKind) {
        case LiteralKind::Integer: result = lit->intValue;  return true;
        case LiteralKind::Long:    result = lit->longValue; return true;
        case LiteralKind::LongPtr: result = lit->longValue; return true;  // Fix 082: ^ 后缀
        case LiteralKind::Boolean: result = lit->boolValue ? 1 : 0; return true;
        default: return false;   // 浮点/字符串等不参与整数折叠
        }
    }
    // 账 #612: 枚举成员值里引用 **模块级 Const / 其它枚举成员** 时必须真求值。
    // 例: `Private Const LVS_EX_FULLROWSELECT = &H20` 之后
    //     `Public Enum LVStylesEx : FullRowSelect = LVS_EX_FULLROWSELECT`
    // 原先走不到这里 → 求值失败 → 回落「上一成员 + 1」，整块枚举静默变成 0,1,2,3,4,5，
    // 于是 LVStyles 勾 HighlightRow(`FullRowSelect`) 实际只置了 1(GridLines) 的位，
    // 行选中永远不生效。
    case ASTNodeKind::IdentifierExpr: {
        auto* id = static_cast<IdentifierExpr*>(expr);
        if (scope) {
            if (const Symbol* s = scope->lookup(id->name)) {
                if (s->hasConstValue) { result = s->constIntValue; return true; }
            }
            // 方括号转义引用 `[sGoogle]` ↔ 声明 `sGoogle`: 与声明侧同口径地两头都试。
            const std::string bracketed = "[" + id->name + "]";
            if (const Symbol* s2 = scope->lookup(bracketed)) {
                if (s2->hasConstValue) { result = s2->constIntValue; return true; }
            }
        }
        // 跨模块 (Fix 197 预注册) 那一趟没有作用域链: 先查调用方给的常量表 (定义模块里
        // 的 Private Const, 此时还没进符号表), 再退到模块级符号表。
        if (localConsts) {
            auto it = localConsts->find(Symbol::toLower(id->name));
            if (it != localConsts->end()) { result = it->second; return true; }
        }
        if (st) {
            if (const Symbol* s3 = st->lookupModule(Symbol::toLower(id->name))) {
                if (s3->hasConstValue) { result = s3->constIntValue; return true; }
            }
        }
        return false;
    }
    case ASTNodeKind::UnaryExpr: {
        auto* unary = static_cast<UnaryExpr*>(expr);
        int64_t v = 0;
        if (!evalEnumMemberConstInt(unary->operand.get(), v, scope, st, localConsts)) return false;
        switch (unary->op) {
        case UnaryOp::Negate: result = -v;  return true;
        case UnaryOp::Not:    result = ~v;  return true;
        }
        return false;
    }
    case ASTNodeKind::BinaryExpr: {
        auto* bin = static_cast<BinaryExpr*>(expr);
        int64_t l = 0, r = 0;
        if (!evalEnumMemberConstInt(bin->left.get(), l, scope, st, localConsts)) return false;
        if (!evalEnumMemberConstInt(bin->right.get(), r, scope, st, localConsts)) return false;
        switch (bin->op) {
        case BinaryOp::Add:    result = l + r; return true;
        case BinaryOp::Sub:    result = l - r; return true;
        case BinaryOp::Mul:    result = l * r; return true;
        case BinaryOp::Div:    return false;  // VB6 '/' 是浮点除法, 不求整
        case BinaryOp::IntDiv: if (r == 0) return false; result = l / r; return true;
        case BinaryOp::Mod:    if (r == 0) return false; result = l % r; return true;
        case BinaryOp::Pow: {
            if (r < 0) return false;
            int64_t base = l, exp = r, pw = 1;
            while (exp > 0) {
                if (exp & 1) pw *= base;
                base *= base;
                exp >>= 1;
            }
            result = pw;
            return true;
        }
        case BinaryOp::Or:  result = l | r; return true;
        case BinaryOp::And: result = l & r; return true;
        case BinaryOp::Xor: result = l ^ r; return true;
        default: return false;
        }
    }
    default:
        return false;
    }
}


void SemanticAnalyzer::visit(TypeDecl& node) {
    // 泛型模板 (tB, G2): 不进符号表 —— 泛型器已按使用点注入特化副本,
    // 模板本体对下游不存在 (未实例化即被引用会在符号查找处自然失败).
    if (!node.typeParams.empty()) return;
    if (pass_ == 1) {
        auto sym = std::make_unique<Symbol>(
            SymbolKind::UserDefinedType, node.name,
            Vb6Type::UserDefinedType, node.loc, node.access
        );
        // P20-21: 注册UDT成员信息
        for (const auto& memberPtr : node.members) {
            Symbol::UdtMemberInfo mi;
            mi.name = memberPtr->name;
            if (memberPtr->type) {
                mi.type = resolveTypeRef(memberPtr->type.get());
                // 若类型是UDT/Enum等命名类型, 保存类型引用名
                if (auto* stRef = dynamic_cast<SimpleTypeRef*>(memberPtr->type.get())) {
                    if (mi.type == Vb6Type::UserDefinedType) {
                        mi.typeRefName = stRef->name;
                    }
                    // Fix 085: 对象字段 (As 项目类/Collection/COM接口, 均解析为 Object)
                    // 也保存类型引用名 — 供 cgen 区分"项目类对象字段"与"COM/Collection
                    // 对象字段", 从而把 obj.Field.Method(...) 生成到正确的类方法调用/
                    // COM dispatch 通道 (否则 UDT 字段直接 obj.Field.Method 触发 C2039/
                    // C2224). 现有 typeRefName 消费者均限定 UserDefinedType, 安全.
                    else if (mi.type == Vb6Type::Object) {
                        mi.typeRefName = stRef->name;
                    }
                    // ai/022 B08f-1 (D37): `As <项目类>` 在**跨模块**引用时 resolveTypeRef
                    // 认不出来 (Class 符号要到 stage 3.5 才注入到本模块作用域) → 回退成
                    // Variant，于是类名也跟着丢，cgen 侧 `udtFieldObjCType` 再也认不出这是
                    // 对象字段 (Set 被当 Variant 容器、成员调用落 COM 晚绑定)。
                    // **只补名字、不动 `mi.type`**：名字是纯元数据，`typeRefName` 的读侧逐处
                    // 核过 —— 要么限定在 `mi.type == UserDefinedType`/`== Object` 分支内，要么
                    // 查到名字后还要 `kind == UserDefinedType` 才认，所以塞一个 Class 名它们
                    // 全都看不见（清单见 ai/022 D37）。
                    else if (mi.type == Vb6Type::Variant) {
                        mi.typeRefName = stRef->name;
                    }
                }
            }
            if (memberPtr->arraySize) {
                mi.arraySize = 1;
                // 账 #262: 只传**秩** (维数)。各维格数不在这里算 —— 那是发码侧
                // tryEvalConstInt 的活，这里留一份就成了第二张真相。
                mi.arrayRank = 1 + (int32_t)memberPtr->moreDims.size();
            }
            // Fix 037: 传播动态数组标记 (`() As Type`) — UdtMemberInfo.isArrayDynamic
            if (memberPtr->isArrayDynamic) {
                mi.isArrayDynamic = true;
            }
            sym->udtMembers.push_back(std::move(mi));
        }
        symTab_.define(std::move(sym));
    }
}

void SemanticAnalyzer::visit(TypeMember& node) {
    // P20-21: UDT成员已在TypeDecl中统一处理
}

void SemanticAnalyzer::visit(EnumDecl& node) {
    if (pass_ == 1) {
        auto sym = std::make_unique<Symbol>(
            SymbolKind::EnumType, node.name,
            Vb6Type::Long, node.loc, node.access
        );
        symTab_.define(std::move(sym));

        // 注册Enum成员为常量
        int64_t nextValue = 0;
        for (auto& member : node.members) {
            auto memberSym = std::make_unique<Symbol>(
                SymbolKind::EnumMember, member->name,
                Vb6Type::Long, member->loc, node.access
            );
            memberSym->hasConstValue = true;
            memberSym->constType = Vb6Type::Long;

            // 如果有显式值
            if (member->value) {
                // Fix 191: 统一走整数常量求值 (字面量 / -字面量 / 算术与位运算 / 幂),
                // 求值失败才沿用「上一成员值 + 1」的 VB6 递增语义.
                // Fix 082 合并决议: LongPtr (^ 后缀) 字面量纳入求值器.
                int64_t evaluated = 0;
                if (evalEnumMemberConstInt(member->value.get(), evaluated, symTab_.currentScope())) {
                    nextValue = evaluated;
                }
            }
            memberSym->constIntValue = nextValue;
            nextValue++;

            // 方括号转义名 `[sGoogle]` 与裸名 `sGoogle` 在 VB6 是同一个常量。声明侧 AST 名带
            // 方括号, 引用侧裸名查不到该符号 ⇒ 静默退化成隐式 Variant 变量 (ucTreeView
            // ucScrollbar.c:2619 实证)。两种形态都登记, 值相同。
            const std::string rawName = member->name;
            if (rawName.size() >= 2 && rawName.front() == '[' && rawName.back() == ']') {
                const int64_t value = memberSym->constIntValue;
                auto alias = std::make_unique<Symbol>(
                    SymbolKind::EnumMember, rawName.substr(1, rawName.size() - 2),
                    Vb6Type::Long, member->loc, node.access
                );
                alias->hasConstValue = true;
                alias->constType = Vb6Type::Long;
                alias->constIntValue = value;
                symTab_.define(std::move(alias));
            }

            symTab_.define(std::move(memberSym));
        }
    } else if (pass_ == 2) {
        // 账 #612: pass 1 时模块级 Const 不一定已进表 (声明顺序 / 子趟顺序),
        // 于是 `FullRowSelect = LVS_EX_FULLROWSELECT` 求值失败 → 回落到「上一成员+1」的
        // 序号。pass 2 符号已齐, 重算一遍直接覆写成员符号的常量值。
        int64_t nextValue = 0;
        for (auto& member : node.members) {
            if (member->value) {
                int64_t evaluated = 0;
                if (evalEnumMemberConstInt(member->value.get(), evaluated, symTab_.currentScope())) {
                    nextValue = evaluated;
                }
            }
            if (Symbol* s = symTab_.lookup(member->name)) {
                s->hasConstValue = true;
                s->constType = Vb6Type::Long;
                s->constIntValue = nextValue;
                const std::string raw = member->name;
                if (raw.size() >= 2 && raw.front() == '[' && raw.back() == ']') {
                    if (Symbol* a = symTab_.lookup(raw.substr(1, raw.size() - 2))) {
                        a->hasConstValue = true;
                        a->constIntValue = nextValue;
                    }
                }
            }
            nextValue++;
        }
    }
}

void SemanticAnalyzer::visit(EnumMember& node) {
    // 在EnumDecl中处理
}

void SemanticAnalyzer::visit(DeclareDecl& node) {
    if (pass_ == 1) {
        SymbolKind sk = (node.procKind == ProcKind::Sub)
                        ? SymbolKind::DeclareSub : SymbolKind::DeclareFunc;

        Vb6Type retType = Vb6Type::Void;
        if (node.procKind == ProcKind::Function) {
            retType = resolveTypeRef(node.returnType.get());
        }

        auto sym = std::make_unique<Symbol>(sk, node.name, retType, node.loc, node.access);

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
                pi.defaultValueExpr = evalOptionalDefault(param->defaultValue.get(), pi.type, pi.isByVal);
            }
            sym->W().params.push_back(std::move(pi));
        }

        symTab_.define(std::move(sym));
    }
    // 外部声明无过程体
}

void SemanticAnalyzer::visit(EventDecl& node) {
    if (pass_ == 1) {
        auto sym = std::make_unique<Symbol>(
            SymbolKind::Event, node.name,
            Vb6Type::Void, node.loc, node.access
        );
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
                pi.defaultValueExpr = evalOptionalDefault(param->defaultValue.get(), pi.type, pi.isByVal);
            }
            sym->W().params.push_back(std::move(pi));
        }
        symTab_.define(std::move(sym));
    }
}

void SemanticAnalyzer::visit(DelegateDecl& node) {
    if (pass_ == 1) {
        // 委托符号是一个"类型": 其值按 LongPtr 表示 (与指针位兼容),
        // 签名 (procKind/callConv/params/returnType) 挂在符号上供检查与桩生成消费.
        auto sym = std::make_unique<Symbol>(
            SymbolKind::Delegate, node.name,
            Vb6Type::LongPtr, node.loc, node.access
        );
        sym->W().delegateProcKind = node.procKind;
        sym->W().delegateCallConv = node.callingConv;
        if (node.procKind == ProcKind::Function && node.returnType) {
            sym->W().delegateReturnType = resolveTypeRef(node.returnType.get());
        }
        for (auto& param : node.params) {
            ParameterInfo pi;
            pi.name = param->name;
            pi.type = resolveTypeOrDefault(param->name, param->asType.get());
            pi.typeRefName = simpleTypeRefName(param->asType.get());  // <vbeclipse>: 类名别丢
            pi.isByVal = param->isByVal;
            pi.isOptional = param->isOptional;
            pi.isParamArray = param->isParamArray;
            sym->W().params.push_back(std::move(pi));
        }
        symTab_.define(std::move(sym));
    }
    // 委托声明无过程体
}

void SemanticAnalyzer::visit(ConstDecl& node) {
    if (pass_ == 1) {
        registerConstant(node);
    }
}

void SemanticAnalyzer::visit(VariableDecl& node) {
    if (pass_ == 1) {
        registerVariable(node);
    } else if (pass_ == 2 && node.initializer && node.asType &&
               node.asType->kind == ASTNodeKind::SimpleTypeRef) {
        // 模块级 Dim x As Operation = AddressOf Proc — pass1 时过程符号尚未齐,
        // 绑定放在 pass2.
        bindDelegateAddressOf(static_cast<SimpleTypeRef*>(node.asType.get())->name,
                              *node.initializer, node.loc);
    }
}

void SemanticAnalyzer::visit(ParameterDecl& node) {
    // 参数在SubDecl/FunctionDecl中处理
}

// Fix 197: 对外入口 — Driver 在 Pass 1 前预注册跨模块 Public 枚举成员时
// 复用同一求值器 (evalOptionalDefault 查符号表发生在 Pass 1 期间).
bool SemanticAnalyzer::evalEnumConstIntForDriver(ASTNode* expr, int64_t& result,
                                                const SymbolTable* st,
                                                const std::map<std::string, int64_t>* localConsts) {
    return evalEnumMemberConstInt(expr, result, nullptr, st, localConsts);
}

} // namespace vb6c3
