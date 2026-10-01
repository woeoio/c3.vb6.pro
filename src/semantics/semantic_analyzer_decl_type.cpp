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
static bool evalEnumMemberConstInt(ASTNode* expr, int64_t& result) {
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
    case ASTNodeKind::UnaryExpr: {
        auto* unary = static_cast<UnaryExpr*>(expr);
        int64_t v = 0;
        if (!evalEnumMemberConstInt(unary->operand.get(), v)) return false;
        switch (unary->op) {
        case UnaryOp::Negate: result = -v;  return true;
        case UnaryOp::Not:    result = ~v;  return true;
        }
        return false;
    }
    case ASTNodeKind::BinaryExpr: {
        auto* bin = static_cast<BinaryExpr*>(expr);
        int64_t l = 0, r = 0;
        if (!evalEnumMemberConstInt(bin->left.get(), l)) return false;
        if (!evalEnumMemberConstInt(bin->right.get(), r)) return false;
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
                if (evalEnumMemberConstInt(member->value.get(), evaluated)) {
                    nextValue = evaluated;
                }
            }
            memberSym->constIntValue = nextValue;
            nextValue++;

            symTab_.define(std::move(memberSym));
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
                pi.defaultValueExpr = evalOptionalDefault(param->defaultValue.get(), pi.type);
            }
            sym->params.push_back(std::move(pi));
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
                pi.defaultValueExpr = evalOptionalDefault(param->defaultValue.get(), pi.type);
            }
            sym->params.push_back(std::move(pi));
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
        sym->delegateProcKind = node.procKind;
        sym->delegateCallConv = node.callingConv;
        if (node.procKind == ProcKind::Function && node.returnType) {
            sym->delegateReturnType = resolveTypeRef(node.returnType.get());
        }
        for (auto& param : node.params) {
            ParameterInfo pi;
            pi.name = param->name;
            pi.type = resolveTypeOrDefault(param->name, param->asType.get());
            pi.typeRefName = simpleTypeRefName(param->asType.get());  // <vbeclipse>: 类名别丢
            pi.isByVal = param->isByVal;
            pi.isOptional = param->isOptional;
            pi.isParamArray = param->isParamArray;
            sym->params.push_back(std::move(pi));
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
bool SemanticAnalyzer::evalEnumConstIntForDriver(ASTNode* expr, int64_t& result) {
    return evalEnumMemberConstInt(expr, result);
}

} // namespace vb6c3
