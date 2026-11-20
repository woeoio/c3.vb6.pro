// TPF0（Delphi VCL TFiler 二进制流）读取器 —— ai/029 C30-10 P2
// 文件头注释见 tpf0_reader.hpp（含实测字节对照）。

#include "tpf0_reader.hpp"

#include <cstring>
#include <cstdio>
#include <cmath>

namespace vb6c3 {

namespace {

struct Cursor {
    const uint8_t* p;
    size_t len;
    size_t pos = 0;
    bool ok = true;
    std::string err;

    bool need(size_t n) {
        if (pos + n > len) {
            if (ok) {
                char b[128];
                std::snprintf(b, sizeof(b), "越界: pos=%zu 需要 %zu 字节, 总长 %zu", pos, n, len);
                err = b;
            }
            ok = false;
            return false;
        }
        return true;
    }
    uint8_t u8() { if (!need(1)) return 0; return p[pos++]; }
    int8_t  i8()  { return (int8_t)u8(); }
    int16_t i16() { if (!need(2)) return 0; int16_t v = (int16_t)(p[pos] | (p[pos+1] << 8)); pos += 2; return v; }
    int32_t i32() { if (!need(4)) return 0; int32_t v = (int32_t)((uint32_t)p[pos] | ((uint32_t)p[pos+1] << 8)
                     | ((uint32_t)p[pos+2] << 16) | ((uint32_t)p[pos+3] << 24)); pos += 4; return v; }
    int64_t i64() { if (!need(8)) return 0; uint64_t v = 0; for (int k = 7; k >= 0; k--) v = (v << 8) | p[pos + k]; pos += 8; return (int64_t)v; }
    // Delphi ShortString：1 字节长度 + 内容
    std::string shortStr() {
        std::string s;
        uint8_t n = u8();
        if (!need(n)) return s;
        s.assign((const char*)p + pos, n);
        pos += n;
        return s;
    }
    bool tag(const char* t) {
        if (!need(4)) return false;
        if (std::memcmp(p + pos, t, 4) != 0) {
            if (ok) { err = "签名不符（应为 TPF0）"; ok = false; }
            return false;
        }
        pos += 4;
        return true;
    }
};

std::string utf16ToUtf8(const uint8_t* p, size_t bytes) {
    std::string out;
    for (size_t i = 0; i + 1 < bytes; i += 2) {
        uint32_t cp = (uint32_t)(p[i] | (p[i+1] << 8));
        if (cp < 0x80) out.push_back((char)cp);
        else if (cp < 0x800) {
            out.push_back((char)(0xC0 | (cp >> 6)));
            out.push_back((char)(0x80 | (cp & 0x3F)));
        } else {
            out.push_back((char)(0xE0 | (cp >> 12)));
            out.push_back((char)(0x80 | ((cp >> 6) & 0x3F)));
            out.push_back((char)(0x80 | (cp & 0x3F)));
        }
    }
    return out;
}

bool readValue(Cursor& c, Tpf0Value& v, int depth);

bool readProps(Cursor& c, std::vector<Tpf0Prop>& props, int depth) {
    if (depth > 32) { c.err = "嵌套过深"; c.ok = false; return false; }
    for (;;) {
        if (!c.ok) return false;
        size_t save = c.pos;
        std::string name = c.shortStr();
        if (!c.ok) return false;
        if (name.empty()) {           // vaNull 0 就是属性表结束
            return true;
        }
        (void)save;
        Tpf0Prop pr;
        pr.name = name;
        if (!readValue(c, pr.value, depth + 1)) return false;
        props.push_back(std::move(pr));
    }
}

bool readValue(Cursor& c, Tpf0Value& v, int depth) {
    if (depth > 32) { c.err = "嵌套过深"; c.ok = false; return false; }
    uint8_t k = c.u8();
    if (!c.ok) return false;
    v.kind = (Tpf0Kind)k;
    switch (v.kind) {
    case Tpf0Kind::Null:
        return true;
    case Tpf0Kind::List:
        for (;;) {
            uint8_t t = c.u8();
            if (!c.ok) return false;
            if (t == (uint8_t)Tpf0Kind::Null) return true;
            c.pos--;                       // 回退，交给 readValue 正常读
            Tpf0Value item;
            if (!readValue(c, item, depth + 1)) return false;
            v.items.push_back(std::move(item));
        }
    case Tpf0Kind::Int8:  v.num = c.i8();  return c.ok;
    case Tpf0Kind::Int16: v.num = c.i16(); return c.ok;
    case Tpf0Kind::Int32: v.num = c.i32(); return c.ok;
    case Tpf0Kind::Int64: v.num = c.i64(); return c.ok;
    case Tpf0Kind::True:  v.num = 1; return true;
    case Tpf0Kind::False: v.num = 0; return true;
    case Tpf0Kind::String:
    case Tpf0Kind::Ident:
    case Tpf0Kind::Utf8String:
        v.str = c.shortStr();
        return c.ok;
    case Tpf0Kind::WString: {
        int32_t n = c.i32();               // 字节数
        if (!c.ok) return false;
        if (n < 0 || !c.need((size_t)n)) { c.err = "WString 长度非法"; c.ok = false; return false; }
        v.str = utf16ToUtf8(c.p + c.pos, (size_t)n);
        c.pos += (size_t)n;
        return true;
    }
    case Tpf0Kind::Binary: {
        int32_t n = c.i32();
        if (!c.ok) return false;
        if (n < 0 || !c.need((size_t)n)) { c.err = "Binary 长度非法"; c.ok = false; return false; }
        v.bin.assign(c.p + c.pos, c.p + c.pos + n);
        c.pos += (size_t)n;
        return true;
    }
    case Tpf0Kind::Single: {                 // 4 字节 IEEE
        int32_t raw = c.i32();
        if (!c.ok) return false;
        float f; std::memcpy(&f, &raw, 4); v.dbl = (double)f; return true;
    }
    case Tpf0Kind::Double: {
        int64_t raw = c.i64();
        if (!c.ok) return false;
        double d; std::memcpy(&d, &raw, 8); v.dbl = d; return true;
    }
    case Tpf0Kind::Extended: {               // 80 位：低 64 位尾数 + 15 位指数
        if (!c.need(10)) return false;
        uint64_t mant = 0;
        for (int i = 7; i >= 0; i--) mant = (mant << 8) | c.p[c.pos + i];
        uint16_t se = (uint16_t)(c.p[c.pos + 8] | (c.p[c.pos + 9] << 8));
        c.pos += 10;
        int exp = se & 0x7FFF;
        double m = (double)mant;              // 未归一化尾数；TeeChart 用 Extended 存坐标
        double val = (exp == 0) ? 0.0 : m * std::pow(2.0, (double)exp - 16383.0 - 63.0);
        v.dbl = (se & 0x8000) ? -val : val;
        return true;
    }
    case Tpf0Kind::Date: {
        int64_t raw = c.i64();               // Delphi TDateTime 的二进制（double）
        if (!c.ok) return false;
        double d; std::memcpy(&d, &raw, 8); v.dbl = d; return true;
    }
    case (Tpf0Kind)14: {   /* 0x0E：老版 Delphi / TeeChart 把集合写成这个档（见 Collection 那条）*/
        size_t save = c.pos;                 /* kind 已读过，pos 指向负载首字节 */
        Cursor u = c;
        std::vector<std::vector<Tpf0Prop>> items;
        bool sawAny = false;
        for (;;) {
            size_t s2 = u.pos;
            uint8_t tag = u.u8();
            if (!u.ok) break;
            if (tag == (uint8_t)Tpf0Kind::Null) {
                if (sawAny) { v.kind = Tpf0Kind::Collection; v.coll = std::move(items); c = u; return true; }
                break;
            }
            if (tag != (uint8_t)Tpf0Kind::List) { u.pos = s2; break; }
            std::vector<Tpf0Prop> item;
            if (!readProps(u, item, depth + 1)) break;
            items.push_back(std::move(item));
            sawAny = true;
        }
        /* 不是集合就按字面的 Currency 读（8 字节定点） */
        c.pos = save;
        int64_t raw = c.i64();
        if (!c.ok) return false;
        v.kind = Tpf0Kind::Currency;
        v.dbl = (double)raw / 10000.0;
        return true;
    }
    case Tpf0Kind::Set:
        for (;;) {
            std::string id = c.shortStr();
            if (!c.ok) return false;
            if (id.empty()) return true;
            Tpf0Value item; item.kind = Tpf0Kind::Ident; item.str = id;
            v.items.push_back(std::move(item));
        }
    case Tpf0Kind::Collection: {
        /* 两种写法都要接：
         *   ① Delphi 6+ 的 WriteCollection —— 先 Int32 count，再每项 vaList+属性表；
         *   ② TeeChart v5 这份流里的写法 —— **没有 count**，直接一串 vaList 项，
         *      以 vaNull 收尾（实测 AxisOpaqueZoneForm.frx: CustomAxes 后面就是
         *      `01 0A "Horizontal" 08 09 "OtherSide" 08 …`，前面没有计数字）。
         * 两种都在真文件里见过，所以用“先试 A，失败回退 B”的写法，不猜。 */
        const uint8_t* anchor = c.p;
        size_t anchorLen = c.len;
        std::string anchorErr = c.err;
        size_t start = c.pos;
        size_t afterKind = start - 1;   /* 调用方已读过 kind 字节 */
        (void)anchor; (void)anchorLen;
        {
            Cursor t = c;
            int32_t n = t.i32();
            if (t.ok && n >= 0 && n <= 100000) {
                bool good = true;
                std::vector<std::vector<Tpf0Prop>> items;
                for (int32_t i = 0; i < n; i++) {
                    uint8_t tag = t.u8();
                    if (!t.ok || tag != (uint8_t)Tpf0Kind::List) { good = false; break; }
                    std::vector<Tpf0Prop> item;
                    if (!readProps(t, item, depth + 1)) { good = false; break; }
                    items.push_back(std::move(item));
                }
                if (good) { v.coll = std::move(items); c = t; return true; }
            }
            /* 回退：无 count 的写法 */
            Cursor u = c; u.pos = start; u.ok = true; u.err = anchorErr;
            std::vector<std::vector<Tpf0Prop>> items;
            bool sawAny = false;
            for (;;) {
                size_t save = u.pos;
                uint8_t tag = u.u8();
                if (!u.ok) break;
                if (tag == (uint8_t)Tpf0Kind::Null) {   /* 集合结束 */
                    if (sawAny) { v.coll = std::move(items); c = u; return true; }
                    break;
                }
                if (tag != (uint8_t)Tpf0Kind::List) { u.pos = save; break; }
                std::vector<Tpf0Prop> item;
                if (!readProps(u, item, depth + 1)) break;
                items.push_back(std::move(item));
                sawAny = true;
            }
            c.pos = afterKind + 1;
            c.err = anchorErr;
        }
        c.err = "Collection 两种写法都读不下去";
        c.ok = false;
        return false;
    }
    default:
        c.err = "未知的 TValueType";
        c.ok = false;
        return false;
    }
}

bool readComp(Cursor& c, Tpf0Comp& comp, int depth) {
    if (depth > 16) { c.err = "组件嵌套过深"; c.ok = false; return false; }
    comp.cls = c.shortStr();
    if (!c.ok) return false;
    if (comp.cls.empty()) return true;       // 空类名 = 子组件列表结束
    comp.name = c.shortStr();
    if (!c.ok) return false;
    if (!readProps(c, comp.props, depth)) return false;
    for (;;) {
        if (!c.ok) return false;
        size_t save = c.pos;
        uint8_t b = c.u8();
        if (!c.ok) return false;
        if (b == 0) return true;             // 子组件表结束（vaNull）
        c.pos = save;
        Tpf0Comp child;
        if (!readComp(c, child, depth + 1)) {
            /* 尾部截断容错：有几个真样本（DrawlineByCode.frx）在最后一条系列之后
             * 只剩几个字节的残尾，按长度读名字会越界。此时**保留已解析的那部分**
             * （标题/系列都已在），比整块丢掉好 —— 设计期外观本来就是尽力而为。 */
            c.ok = true;
            c.err.clear();
            return true;
        }
        if (child.cls.empty()) return true;
        comp.children.push_back(std::move(child));
    }
}

} // namespace

std::string g_tpf0Error;
size_t g_tpf0Consumed = 0;

const char* tpf0KindName(Tpf0Kind k) {
    switch (k) {
    case Tpf0Kind::Null: return "Null";
    case Tpf0Kind::List: return "List";
    case Tpf0Kind::Int8: return "Int8";
    case Tpf0Kind::Int16: return "Int16";
    case Tpf0Kind::Int32: return "Int32";
    case Tpf0Kind::Extended: return "Extended";
    case Tpf0Kind::String: return "String";
    case Tpf0Kind::Ident: return "Ident";
    case Tpf0Kind::False: return "False";
    case Tpf0Kind::True: return "True";
    case Tpf0Kind::Binary: return "Binary";
    case Tpf0Kind::Set: return "Set";
    case Tpf0Kind::Collection: return "Collection";
    case Tpf0Kind::Single: return "Single";
    case Tpf0Kind::Currency: return "Currency";
    case Tpf0Kind::Date: return "Date";
    case Tpf0Kind::WString: return "WString";
    case Tpf0Kind::Int64: return "Int64";
    case Tpf0Kind::Utf8String: return "Utf8String";
    case Tpf0Kind::Double: return "Double";
    }
    return "?";
}

bool Tpf0Reader::parse(const std::vector<uint8_t>& bytes, Tpf0Comp& out) {
    Cursor c{bytes.data(), bytes.size(), 0, true, {}};
    g_tpf0Error.clear();
    g_tpf0Consumed = 0;
    if (!c.tag("TPF0")) {
        g_tpf0Error = c.err.empty() ? "不是 TPF0 流" : c.err;
        return false;
    }
    out = Tpf0Comp{};
    if (!readComp(c, out, 0)) {
        char b[256];
        std::snprintf(b, sizeof(b), "%s (pos=%zu)", c.err.c_str(), c.pos);
        g_tpf0Error = b;
        return false;
    }
    g_tpf0Consumed = c.pos;
    return true;
}

const std::string& Tpf0Reader::error() { return g_tpf0Error; }
size_t Tpf0Reader::consumed() { return g_tpf0Consumed; }

static bool ieq(const std::string& a, const char* b) {
    size_t i = 0;
    for (; i < a.size() && b[i]; i++) {
        char x = a[i], y = b[i];
        if (x >= 'A' && x <= 'Z') x = (char)(x - 'A' + 'a');
        if (y >= 'A' && y <= 'Z') y = (char)(y - 'A' + 'a');
        if (x != y) return false;
    }
    return i == a.size() && !b[i];
}

const Tpf0Value* Tpf0Reader::findProp(const std::vector<Tpf0Prop>& props, const char* name) {
    for (const auto& p : props) if (ieq(p.name, name)) return &p.value;
    return nullptr;
}

int64_t Tpf0Reader::propInt(const std::vector<Tpf0Prop>& props, const char* name, int64_t def) {
    const Tpf0Value* v = findProp(props, name);
    if (!v) return def;
    switch (v->kind) {
    case Tpf0Kind::Int8: case Tpf0Kind::Int16: case Tpf0Kind::Int32:
    case Tpf0Kind::Int64: case Tpf0Kind::True: case Tpf0Kind::False:
        return v->num;
    case Tpf0Kind::Double: case Tpf0Kind::Single: case Tpf0Kind::Extended:
        return (int64_t)v->dbl;
    default: return def;
    }
}

bool Tpf0Reader::propBool(const std::vector<Tpf0Prop>& props, const char* name, bool def) {
    const Tpf0Value* v = findProp(props, name);
    if (!v) return def;
    if (v->kind == Tpf0Kind::True) return true;
    if (v->kind == Tpf0Kind::False) return false;
    return v->num != 0;
}

std::string Tpf0Reader::propStr(const std::vector<Tpf0Prop>& props, const char* name) {
    const Tpf0Value* v = findProp(props, name);
    if (!v) return std::string();
    if (v->kind == Tpf0Kind::String || v->kind == Tpf0Kind::Ident
        || v->kind == Tpf0Kind::Utf8String || v->kind == Tpf0Kind::WString)
        return v->str;
    return std::string();
}

} // namespace vb6c3
