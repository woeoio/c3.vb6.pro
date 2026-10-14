#include "lexer/token.hpp"
#include <cstring>

namespace vb6c3 {

bool Token::isKeyword() const {
    return (kind >= TokenKind::TrueKeyword && kind <= TokenKind::GetObject)
        || kind == TokenKind::If || kind == TokenKind::Then
        || kind == TokenKind::Else || kind == TokenKind::ElseIf
        || kind == TokenKind::For || kind == TokenKind::To
        || kind == TokenKind::Step || kind == TokenKind::Next
        || kind == TokenKind::Do || kind == TokenKind::Loop
        || kind == TokenKind::While || kind == TokenKind::Until
        || kind == TokenKind::Wend || kind == TokenKind::Select
        || kind == TokenKind::Case || kind == TokenKind::With
        || kind == TokenKind::End || kind == TokenKind::Dim
        || kind == TokenKind::ReDim || kind == TokenKind::Const
        || kind == TokenKind::Sub || kind == TokenKind::Function
        || kind == TokenKind::Property || kind == TokenKind::Get
        || kind == TokenKind::Let || kind == TokenKind::Set
        || kind == TokenKind::Call || kind == TokenKind::Declare
        || kind == TokenKind::Type || kind == TokenKind::Enum
        || kind == TokenKind::As || kind == TokenKind::New
        || kind == TokenKind::Public || kind == TokenKind::Private
        || kind == TokenKind::Static || kind == TokenKind::Friend
        || kind == TokenKind::Protected
        || kind == TokenKind::ByVal || kind == TokenKind::ByRef
        || kind == TokenKind::Optional || kind == TokenKind::ParamArray
        || kind == TokenKind::Implements || kind == TokenKind::Event
        || kind == TokenKind::Delegate
        || kind == TokenKind::RaiseEvent || kind == TokenKind::WithEvents
        || kind == TokenKind::Interface || kind == TokenKind::Extends
        || kind == TokenKind::Inherits
        || kind == TokenKind::Via
        || kind == TokenKind::CoClass
        || kind == TokenKind::Overridable || kind == TokenKind::Overrides
        || kind == TokenKind::NotOverridable
        || kind == TokenKind::OnError || kind == TokenKind::GoTo
        || kind == TokenKind::GoSub || kind == TokenKind::Return
        || kind == TokenKind::Resume || kind == TokenKind::Exit
        || kind == TokenKind::On;
    // 不完全枚举, 后续补全
}

bool Token::isOperator() const {
    switch (kind) {
        case TokenKind::Plus: case TokenKind::Minus:
        case TokenKind::Star: case TokenKind::Slash:
        case TokenKind::BackSlash: case TokenKind::Mod:
        case TokenKind::Caret: case TokenKind::Ampersand:
        case TokenKind::Equals: case TokenKind::NotEquals:
        case TokenKind::LessThan: case TokenKind::GreaterThan:
        case TokenKind::LessEqual: case TokenKind::GreaterEqual:
        case TokenKind::And: case TokenKind::Or:
        case TokenKind::Xor: case TokenKind::Not:
        case TokenKind::Eqv: case TokenKind::Imp:
        case TokenKind::Is: case TokenKind::Like:
        case TokenKind::AddressOf: case TokenKind::TypeOf:
        // C3 扩展 (ai/032)
        case TokenKind::Shl: case TokenKind::Shr: case TokenKind::IsNot:
        case TokenKind::PlusEq: case TokenKind::MinusEq:
        case TokenKind::StarEq: case TokenKind::SlashEq:
        case TokenKind::BackSlashEq: case TokenKind::CaretEq:
        case TokenKind::AmpEq: case TokenKind::ShlEq:
        case TokenKind::ShrEq: case TokenKind::ModEq:
            return true;
        default:
            return false;
    }
}

bool Token::isLiteral() const {
    switch (kind) {
        case TokenKind::IntegerLiteral:
        case TokenKind::LongLiteral:
        case TokenKind::LongPtrLiteral:
        case TokenKind::FloatLiteral:
        case TokenKind::DecimalLiteral:
        case TokenKind::StringLiteral:
        case TokenKind::DateLiteral:
        case TokenKind::TrueKeyword:
        case TokenKind::FalseKeyword:
        case TokenKind::NothingKeyword:
        case TokenKind::EmptyKeyword:
        case TokenKind::NullKeyword:
            return true;
        default:
            return false;
    }
}

bool Token::isTypeSuffix() const {
    switch (kind) {
        case TokenKind::Percent:   // % Integer
        case TokenKind::Ampersand: // & Long (但也是字符串连接符)
        case TokenKind::Exclamation: // ! Single (也是字典访问)
        case TokenKind::Hash:     // # Double (也是Date定界符)
        case TokenKind::AtSign:   // @ Currency
        case TokenKind::Dollar:   // $ String
            return true;
        default:
            return false;
    }
}

bool Token::isStatementStart() const {
    switch (kind) {
        case TokenKind::Dim: case TokenKind::ReDim:
        case TokenKind::Const: case TokenKind::Public:
        case TokenKind::Private: case TokenKind::Static:
        case TokenKind::Sub: case TokenKind::Function:
        case TokenKind::Property: case TokenKind::Declare:
        case TokenKind::Type: case TokenKind::Enum:
        case TokenKind::Event: case TokenKind::Delegate:
        case TokenKind::Implements:
        case TokenKind::If: case TokenKind::For:
        case TokenKind::Do: case TokenKind::While:
        case TokenKind::Select: case TokenKind::With:
        case TokenKind::Asm:  // ai/vb-asm-extension-spec
        case TokenKind::Set: case TokenKind::Let:
        case TokenKind::Call: case TokenKind::GoTo:
        case TokenKind::GoSub: case TokenKind::Return:
        case TokenKind::OnError: case TokenKind::Resume:
        case TokenKind::Exit: case TokenKind::RaiseEvent:
        case TokenKind::On: case TokenKind::Open:
        case TokenKind::Close: case TokenKind::Name:
        case TokenKind::Beep: case TokenKind::Stop:
        case TokenKind::End: case TokenKind::Attribute:
        case TokenKind::Option: case TokenKind::Mid:
        case TokenKind::LSet: case TokenKind::RSet:
        case TokenKind::Load: case TokenKind::Unload:
        case TokenKind::SavePicture:
        case TokenKind::MsgBox: case TokenKind::InputBox:
        case TokenKind::CreateObject: case TokenKind::GetObject:
        case TokenKind::Identifier: // 赋值/过程调用
            return true;
        default:
            return false;
    }
}

std::string Token::toString() const {
    std::string result = kindToString(kind);
    if (!text.empty() && text != result) {
        result += "('" + text + "')";
    }
    result += " @ " + std::to_string(line) + ":" + std::to_string(column);
    return result;
}

const char* Token::kindToString(TokenKind kind) {
    switch (kind) {
        // 文件级
        case TokenKind::EndOfFile:       return "EOF";
        case TokenKind::NewLine:         return "NewLine";
        // 标识符/字面量
        case TokenKind::Identifier:      return "Identifier";
        case TokenKind::IntegerLiteral:  return "IntLit";
        case TokenKind::LongLiteral:     return "LongLit";
        case TokenKind::LongPtrLiteral:  return "LongPtrLit";
        case TokenKind::FloatLiteral:    return "FloatLit";
        case TokenKind::DecimalLiteral:  return "DecLit";
        case TokenKind::StringLiteral:   return "StrLit";
        case TokenKind::DateLiteral:     return "DateLit";
        case TokenKind::TrueKeyword:     return "True";
        case TokenKind::FalseKeyword:    return "False";
        case TokenKind::NothingKeyword:  return "Nothing";
        case TokenKind::EmptyKeyword:    return "Empty";
        case TokenKind::NullKeyword:     return "Null";
        case TokenKind::MeKeyword:       return "Me";
        // 运算符
        case TokenKind::Plus:            return "+";
        case TokenKind::Minus:           return "-";
        case TokenKind::Star:            return "*";
        case TokenKind::Slash:           return "/";
        case TokenKind::BackSlash:       return "\\";
        case TokenKind::Mod:             return "Mod";
        case TokenKind::Caret:           return "^";
        case TokenKind::Ampersand:       return "&";
        case TokenKind::Equals:          return "=";
        case TokenKind::NotEquals:       return "<>";
        case TokenKind::LessThan:        return "<";
        case TokenKind::GreaterThan:     return ">";
        case TokenKind::LessEqual:       return "<=";
        case TokenKind::GreaterEqual:    return ">=";
        case TokenKind::And:             return "And";
        case TokenKind::Or:              return "Or";
        case TokenKind::Xor:             return "Xor";
        case TokenKind::Not:             return "Not";
        case TokenKind::Eqv:             return "Eqv";
        case TokenKind::Imp:             return "Imp";
        case TokenKind::Is:              return "Is";
        case TokenKind::IsNot:           return "IsNot";
        case TokenKind::Shl:             return "<<";
        case TokenKind::Shr:             return ">>";
        case TokenKind::Like:            return "Like";
        case TokenKind::AddressOf:       return "AddressOf";
        case TokenKind::TypeOf:          return "TypeOf";
        // 复合赋值 (C3 扩展, ai/032)
        case TokenKind::PlusEq:          return "+=";
        case TokenKind::MinusEq:         return "-=";
        case TokenKind::StarEq:          return "*=";
        case TokenKind::SlashEq:         return "/=";
        case TokenKind::BackSlashEq:     return "\\=";
        case TokenKind::CaretEq:         return "^=";
        case TokenKind::AmpEq:           return "&=";
        case TokenKind::ShlEq:           return "<<=";
        case TokenKind::ShrEq:           return ">>=";
        case TokenKind::ModEq:           return "Mod=";
        // 分隔符
        case TokenKind::LeftParen:       return "(";
        case TokenKind::RightParen:      return ")";
        case TokenKind::LeftBracket:     return "[";
        case TokenKind::RightBracket:    return "]";
        case TokenKind::Dot:             return ".";
        case TokenKind::Comma:           return ",";
        case TokenKind::Colon:           return ":";
        case TokenKind::Semicolon:       return ";";
        case TokenKind::Exclamation:     return "!";
        case TokenKind::Hash:            return "#";
        case TokenKind::Percent:         return "%";
        case TokenKind::AtSign:          return "@";
        case TokenKind::Dollar:          return "$";
        case TokenKind::Assign:          return ":=";
        // 注释
        case TokenKind::Comment:         return "Comment";
        case TokenKind::LineContinuation:return "LineCont";
        // 条件编译
        case TokenKind::HashIf:          return "#If";
        case TokenKind::HashElseIf:      return "#ElseIf";
        case TokenKind::HashElse:        return "#Else";
        case TokenKind::HashEnd:         return "#End";
        case TokenKind::HashConst:       return "#Const";
        // 错误
        case TokenKind::Invalid:         return "Invalid";
        default:                         return "Token";
    }
}

} // namespace vb6c3
