#include "parser/parser.hpp"
#include <algorithm>
#include <cctype>

namespace vb6c3 {

// ============================================================
// 软关键字与标识符名
// ============================================================

bool Parser::isSoftKeyword(TokenKind kind) const {
    // VB6 软关键字: 可以在标识符位置使用的保留字
    // 参考: VB6 语言规范 + 常见实践 (Name/String/Integer 等均可作变量名)
    switch (kind) {
        // 内置类型名 (常作变量/参数名)
        case TokenKind::Boolean:   case TokenKind::Byte:     case TokenKind::Integer:
        case TokenKind::Long:      case TokenKind::LongLong: case TokenKind::LongPtr:
        case TokenKind::Single:    case TokenKind::Double:   case TokenKind::Currency:
        case TokenKind::Decimal:   case TokenKind::Date:     case TokenKind::Object:
        case TokenKind::String:    case TokenKind::Variant:  case TokenKind::Any:
        // 软语句关键字
        case TokenKind::Name:      case TokenKind::Step:     case TokenKind::Error:
        case TokenKind::Width:     case TokenKind::Access:
        // 文件 I/O 模式关键字
        case TokenKind::Input:     case TokenKind::Output:   case TokenKind::Append:
        case TokenKind::Binary:    case TokenKind::Random:
        case TokenKind::Read:      case TokenKind::Write:    case TokenKind::ReadWrite:
        case TokenKind::Shared:    case TokenKind::Lock:     case TokenKind::Unlock: case TokenKind::Reset:
        case TokenKind::Put:       case TokenKind::Seek:     case TokenKind::Close:
        case TokenKind::Open:      case TokenKind::Print:    case TokenKind::Line:
        case TokenKind::FreeFile:  case TokenKind::EOF_keyword:
        case TokenKind::Get:
        // 文件系统关键字
        case TokenKind::ChDir:     case TokenKind::ChDrive:  case TokenKind::MkDir:
        case TokenKind::RmDir:     case TokenKind::CurDir:   case TokenKind::Dir:
        case TokenKind::FileCopy:  case TokenKind::Kill:     case TokenKind::SetAttr:
        case TokenKind::GetAttr:   case TokenKind::FileLen:  case TokenKind::FileDateTime:
        // 内置函数关键字
        case TokenKind::MsgBox:    case TokenKind::InputBox: case TokenKind::RGB:
        case TokenKind::QBColor:   case TokenKind::Format:   case TokenKind::Mid:
        case TokenKind::Load:      case TokenKind::Unload:   case TokenKind::LSet:
        case TokenKind::RSet:      case TokenKind::SavePicture: case TokenKind::LoadPicture:
        case TokenKind::CreateObject: case TokenKind::GetObject:
        // 杂项语句关键字
        case TokenKind::Beep:      case TokenKind::DoEvents: case TokenKind::SendKeys:
        case TokenKind::AppActivate: case TokenKind::Shell:  case TokenKind::Environ:
        case TokenKind::Command:   case TokenKind::Randomize: case TokenKind::Timer:
        // 日期/时间关键字
        case TokenKind::DateValue: case TokenKind::TimeValue: case TokenKind::DateAdd:
        case TokenKind::DateDiff:  case TokenKind::DatePart: case TokenKind::DateSerial:
        case TokenKind::TimeSerial:
        // 上下文关键字
        case TokenKind::Compare:   case TokenKind::Base:     case TokenKind::Text:
        case TokenKind::Binary2:   case TokenKind::Explicit: case TokenKind::Private2:
        case TokenKind::Attribute: case TokenKind::Begin: case TokenKind::Default:
        // tB 扩展接口关键字: 登记为软关键字, 存量代码里同名标识符 (变量/成员名) 不受影响
        case TokenKind::Interface: case TokenKind::Extends: case TokenKind::Inherits:
        case TokenKind::Protected:
        case TokenKind::Overridable: case TokenKind::Overrides:
        case TokenKind::NotOverridable:  // tB 扩展 (ai/022 B08b): 虚方法修饰符同为软关键字
        case TokenKind::Via:             // tB 扩展 (ai/022 B10): `Implements I Via m_h` 的 Via
        case TokenKind::CoClass:         // tB 扩展 (ai/026, B11/C01): CoClass 块名可作普通标识符
        // 其他
        case TokenKind::Resume:    case TokenKind::Stop:
        case TokenKind::Let:       case TokenKind::Set:
            return true;
        default:
            return false;
    }
}

bool Parser::canBeName(TokenKind kind) const {
    return kind == TokenKind::Identifier || isSoftKeyword(kind);
}

Token Parser::expectName(const std::string& msg) {
    if (canBeName(cur_.kind)) {
        return advance();
    }
    diag_.error(DiagnosticID::ParseExpectedToken, currentLoc(),
        msg + " (got " + std::string(Token::kindToString(cur_.kind)) + ")");
    return Token{cur_.kind, cur_.text, cur_.line, cur_.column, cur_.length, {0}};
}

// VB6 的行号就是标签名: `100: ...` 声明, `GoTo 100` / `GoSub 100` / `Resume 100` /
// `On Error GoTo 100` / `On x GoTo 100, 200` 引用。标签名统一存为文本, 后端
// `vb6_label_ + cIdent(name)` 与语义层的字符串比对都无需为数字另开分支。
std::string Parser::expectLabelTarget(const std::string& msg) {
    if (cur_.kind == TokenKind::IntegerLiteral) {
        return advance().text;
    }
    return expectName(msg).text;
}

// Fix 028: 见 parser.hpp 注释。剥离 VB6 标识符末尾的类型后缀, 返回剥离后的名字和类型名。
Parser::TypeSuffixStrip Parser::stripTypeSuffix(const std::string& text) const {
    TypeSuffixStrip result;
    result.name = text;
    if (text.size() < 2) return result;  // 单字符标识符不含后缀
    const char last = text.back();
    switch (last) {
        case '$': result.typeName = "String";   break;
        case '%': result.typeName = "Integer";  break;
        case '&': result.typeName = "Long";     break;
        case '!': result.typeName = "Single";   break;
        case '#': result.typeName = "Double";   break;
        case '@': result.typeName = "Currency"; break;
        default:
            return result;  // 无类型后缀, name 保留原文
    }
    result.name = text.substr(0, text.size() - 1);
    return result;
}

// ============================================================
// 运算符优先级查找
// ============================================================

BindingPower Parser::getBindingPower(TokenKind kind) const {
    auto it = bpTable_.find(static_cast<int>(kind));
    if (it != bpTable_.end()) {
        return it->second;
    }
    return {0, 0};  // 非中缀运算符
}

BinaryOp Parser::tokenToBinaryOp(TokenKind kind) const {
    switch (kind) {
        case TokenKind::Or:    return BinaryOp::Or;
        case TokenKind::Xor:   return BinaryOp::Xor;
        case TokenKind::And:   return BinaryOp::And;
        case TokenKind::Eqv:   return BinaryOp::Eqv;
        case TokenKind::Imp:   return BinaryOp::Imp;
        case TokenKind::Equals:      return BinaryOp::Eq;
        case TokenKind::NotEquals:   return BinaryOp::Neq;
        case TokenKind::LessThan:    return BinaryOp::Lt;
        case TokenKind::GreaterThan: return BinaryOp::Gt;
        case TokenKind::LessEqual:   return BinaryOp::Le;
        case TokenKind::GreaterEqual:return BinaryOp::Ge;
        case TokenKind::Ampersand:   return BinaryOp::Concat;
        case TokenKind::Plus:        return BinaryOp::Add;
        case TokenKind::Minus:       return BinaryOp::Sub;
        case TokenKind::Mod:         return BinaryOp::Mod;
        case TokenKind::BackSlash:   return BinaryOp::IntDiv;
        case TokenKind::Star:        return BinaryOp::Mul;
        case TokenKind::Slash:       return BinaryOp::Div;
        case TokenKind::Caret:       return BinaryOp::Pow;
        case TokenKind::Like:        return BinaryOp::Like;
        case TokenKind::Is:          return BinaryOp::Is;
        case TokenKind::IsKeyword:   return BinaryOp::Is;
        default:
            return BinaryOp::Add;  // 不应到达
    }
}

bool Parser::isInfixOperator(TokenKind kind) const {
    return bpTable_.find(static_cast<int>(kind)) != bpTable_.end();
}

bool Parser::isPrefixOperator(TokenKind kind) const {
    return kind == TokenKind::Minus || kind == TokenKind::Not;
}

// ============================================================
// 辅助
// ============================================================

bool Parser::identifierEquals(const std::string& a, const std::string& b) const {
    if (a.size() != b.size()) return false;
    for (size_t i = 0; i < a.size(); i++) {
        if (std::toupper(static_cast<unsigned char>(a[i])) !=
            std::toupper(static_cast<unsigned char>(b[i])))
            return false;
    }
    return true;
}

std::string Parser::toLower(const std::string& s) const {
    std::string result = s;
    std::transform(result.begin(), result.end(), result.begin(),
        [](unsigned char c) { return std::tolower(c); });
    return result;
}

bool Parser::isEndBlock() const {
    if (cur_.kind != TokenKind::End) return false;
    // next_ 是 End 后面的 token
    switch (next_.kind) {
        case TokenKind::If:
        case TokenKind::Sub:
        case TokenKind::Function:
        case TokenKind::Property:
        case TokenKind::Type:
        case TokenKind::Enum:
        case TokenKind::Select:
        case TokenKind::With:
        case TokenKind::Do:  // 不存在 End Do, 但容错
            return true;
        default:
            return false;
    }
}

// ============================================================
// 账 #172: 日期字面量 → OLE 自动化日期
// ============================================================

namespace {

// days from civil (Howard Hinnant 的算法，无循环、跨年正确)
int64_t daysFromCivil(int64_t y, int64_t m, int64_t d) {
    y -= m <= 2;
    const int64_t era = (y >= 0 ? y : y - 399) / 400;
    const uint64_t yoe = static_cast<uint64_t>(y - era * 400);
    const uint64_t mp = static_cast<uint64_t>(m + (m > 2 ? -3 : 9));
    const uint64_t doy = (153 * mp + 2) / 5 + static_cast<uint64_t>(d) - 1;
    const uint64_t doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return era * 146097 + static_cast<int64_t>(doe) - 719468;
}

bool allDigits(const std::string& s) {
    if (s.empty()) return false;
    for (char c : s) {
        if (c < '0' || c > '9') return false;
    }
    return true;
}

bool parseInt32(const std::string& s, int64_t& out) {
    if (!allDigits(s) || s.size() > 9) return false;
    out = 0;
    for (char c : s) out = out * 10 + (c - '0');
    return true;
}

bool isLeap(int64_t y) {
    return (y % 4 == 0 && y % 100 != 0) || y % 400 == 0;
}

int64_t daysInMonth(int64_t y, int64_t m) {
    static const int64_t tbl[12] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
    if (m == 2 && isLeap(y)) return 29;
    return tbl[m - 1];
}

int64_t normalizeYear(int64_t y) {
    // VB6: 两位年份 <50 走 2000s, ≥50 走 1900s
    if (y < 100) return y < 50 ? y + 2000 : y + 1900;
    return y;
}

void trimAscii(std::string& s) {
    size_t b = 0, e = s.size();
    while (b < e && (s[b] == ' ' || s[b] == '\t')) b++;
    while (e > b && (s[e - 1] == ' ' || s[e - 1] == '\t')) e--;
    s = s.substr(b, e - b);
}

} // namespace

bool vb6c3::foldDateLiteralToOADate(const std::string& rawIn, double& out) {
    std::string raw = rawIn;
    if (raw.size() < 3 || raw.front() != '#' || raw.back() != '#') return false;
    raw = raw.substr(1, raw.size() - 2);
    trimAscii(raw);
    if (raw.empty()) return false;

    // 日期段 / 时间段以第一个空格分（VB6 也允许 T 分隔，语料里没有，先不认）
    std::string datePart = raw, timePart;
    size_t sp = raw.find(' ');
    if (sp != std::string::npos) {
        datePart = raw.substr(0, sp);
        timePart = raw.substr(sp + 1);
        trimAscii(timePart);
    }

    int64_t days = 0;
    bool haveDate = false;
    const int64_t oleEpochDays = daysFromCivil(1899, 12, 30);

    if (datePart.find('/') != std::string::npos || datePart.find('-') != std::string::npos) {
        char sep = datePart.find('/') != std::string::npos ? '/' : '-';
        if (datePart.find(sep == '/' ? '-' : '/') != std::string::npos) return false;  // 混用分隔符
        std::vector<std::string> f;
        size_t pos = 0;
        while (true) {
            size_t nx = datePart.find(sep, pos);
            f.push_back(datePart.substr(pos, nx == std::string::npos ? std::string::npos : nx - pos));
            if (nx == std::string::npos) break;
            pos = nx + 1;
        }
        if (f.size() != 3) return false;
        int64_t a, b, c;
        if (!parseInt32(f[0], a) || !parseInt32(f[1], b) || !parseInt32(f[2], c)) return false;
        int64_t year, month, day;
        if (sep == '/') {          // VB6: 斜杠 = M/D/Y
            month = a; day = b; year = c;
        } else {                   // VB6: 连字符 = D-M-Y
            day = a; month = b; year = c;
        }
        year = normalizeYear(year);
        if (year < 100 || year > 9999) return false;
        if (month < 1 || month > 12) return false;
        if (day < 1 || day > daysInMonth(year, month)) return false;
        days = daysFromCivil(year, month, day) - oleEpochDays;
        haveDate = true;
    } else if (!datePart.empty()) {
        // 纯时间（无日期段）：datePart 实际就是时间
        timePart = datePart + (timePart.empty() ? std::string() : std::string(" ") + timePart);
    }

    double frac = 0.0;
    if (!timePart.empty()) {
        std::string tp = timePart;
        // 尾部 AM/PM（也认单独的 A/P）
        int64_t pmAdj = 0;
        std::string tail = tp;
        size_t lastSp = tp.find_last_of(" \t");
        std::string lastTok = (lastSp == std::string::npos) ? tp : tp.substr(lastSp + 1);
        trimAscii(lastTok);
        for (auto& ch : lastTok) ch = static_cast<char>(std::toupper(static_cast<unsigned char>(ch)));
        if (lastTok == "AM" || lastTok == "A" || lastTok == "PM" || lastTok == "P") {
            pmAdj = (lastTok[0] == 'P') ? 1 : 0;
            tp = tp.substr(0, lastSp == std::string::npos ? 0 : lastSp);
            trimAscii(tp);
        }
        std::vector<std::string> hf;
        size_t pos = 0;
        while (true) {
            size_t nx = tp.find(':', pos);
            hf.push_back(tp.substr(pos, nx == std::string::npos ? std::string::npos : nx - pos));
            if (nx == std::string::npos) break;
            pos = nx + 1;
        }
        if (hf.size() < 2 || hf.size() > 3) return false;
        int64_t h, m, s = 0;
        if (!parseInt32(hf[0], h) || !parseInt32(hf[1], m)) return false;
        if (hf.size() == 3 && !parseInt32(hf[2], s)) return false;
        if (m < 0 || m > 59 || s < 0 || s > 59) return false;
        const bool hasAmPm = (lastTok == "AM" || lastTok == "A" || lastTok == "PM" || lastTok == "P");
        if (hasAmPm) {
            if (h < 1 || h > 12) return false;
            if (pmAdj == 1 && h < 12) h += 12;   // PM 1..11 → 13..23（PM 12 就是 12）
            if (pmAdj == 0 && h == 12) h = 0;    // AM 12 → 0
        } else if (h < 0 || h > 23) {
            return false;
        }
        frac = static_cast<double>(h * 3600 + m * 60 + s) / 86400.0;
    }

    if (!haveDate && timePart.empty()) return false;
    out = static_cast<double>(days) + frac;
    return true;
}

} // namespace vb6c3
