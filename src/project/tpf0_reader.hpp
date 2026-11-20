#pragma once
// ============================================================
// TPF0（Delphi VCL TFiler 二进制流）读取器 —— ai/029 C30-10 P2
//
// 为什么需要它：TeeChart 是 Delphi 写的 VCL 控件，`tests/TeeChartChartDisplay` 这类
// 厂商示例的设计期外观（Series 条数/类型/颜色/标题、Header、Axis 标题…）全在 `.frx`
// 的 `OleObjectBlob` 里，而那块负载就是 TPF0 流。VB6 建完控件后拿它
// IPersistStreamInit::Load 灌进 OCX；内置控件没有 OCX，只能自己解析。
//
// 实测（tests/TeeChartChartDisplay/DisplayForm.frx，偏移 0 的 LB 块）：
//   54 50 46 30                     "TPF0"
//   0B "TChartChart"               类名（TeeChart 的 TChart 派生类）
//   00                             实例名（根组件为空）
//   04 "Left"  02 00               vaInt8
//   03 "Top"   02 04
//   05 "Width" 03 91 01            vaInt16 = 401   ← 与真 VB6 的客户区 401 宽度一致
//   06 "Height 03 01 01            vaInt16 = 257
//
// 本文件只做**解析**（结构 → 树），映射到图表状态在发码侧（cgen）做：
// 解析器不认识 TeeChart，避免把厂商属性名硬编码进读文件的那一层。
// ============================================================

#include <cstdint>
#include <string>
#include <vector>

namespace vb6c3 {

// Delphi TValueType（TWriter 二进制档位）
enum class Tpf0Kind : uint8_t {
    Null        = 0,
    List        = 1,
    Int8        = 2,
    Int16       = 3,
    Int32       = 4,
    Extended    = 5,
    String      = 6,
    Ident       = 7,
    False       = 8,
    True        = 9,
    Binary      = 10,
    Set         = 11,
    Collection  = 12,
    Single      = 13,
    Currency    = 14,
    Date        = 15,
    WString     = 16,
    Int64       = 17,
    Utf8String  = 18,
    Double      = 19,
};

struct Tpf0Prop;

struct Tpf0Value {
    Tpf0Kind kind = Tpf0Kind::Null;
    int64_t  num = 0;             // Int8/Int16/Int32/Int64（含 True/False）
    double   dbl = 0.0;           // Extended/Single/Currency/Date/Double
    std::string str;              // String/Ident/Utf8String/WString(转 UTF-8)
    std::vector<uint8_t> bin;     // Binary
    std::vector<Tpf0Value> items; // List / Set(每项是 Ident)
    std::vector<std::vector<Tpf0Prop>> coll;  // Collection：每项 = 属性表
};

struct Tpf0Prop {
    std::string name;
    Tpf0Value   value;
};

struct Tpf0Comp {
    std::string cls;              // 类名（TChartChart / TLineSeries / TPieSeries…）
    std::string name;             // 实例名（根组件为空）
    std::vector<Tpf0Prop> props;
    std::vector<Tpf0Comp> children;
};

class Tpf0Reader {
public:
    // 解析一整条流；成功时 out 为根组件。
    // 失败返回 false，error() 给出字节偏移与原因（定位靠偏移，不靠猜）。
    static bool parse(const std::vector<uint8_t>& bytes, Tpf0Comp& out);
    static const std::string& error();
    static size_t consumed();     // 解析根组件用掉的字节数（调试用）

    // 在属性表里按名查（大小写不敏感，Delphi 属性名大小写固定，但别名写法常见）
    static const Tpf0Value* findProp(const std::vector<Tpf0Prop>& props, const char* name);
    // 便捷取值
    static int64_t  propInt(const std::vector<Tpf0Prop>& props, const char* name, int64_t def);
    static bool     propBool(const std::vector<Tpf0Prop>& props, const char* name, bool def);
    static std::string propStr(const std::vector<Tpf0Prop>& props, const char* name);
};

const char* tpf0KindName(Tpf0Kind k);

} // namespace vb6c3
