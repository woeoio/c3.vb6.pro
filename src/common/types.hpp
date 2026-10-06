#pragma once
// vb6c - Visual Basic 6.0 Compiler
// Copyright (c) 2026 vb6.pro project

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>
#include <memory>
#include <unordered_map>
#include <optional>

namespace vb6c3 {

// VB6数据类型枚举
// 使用uint16_t以容纳Array(8192)和ByRef(16384)位标志
enum class Vb6Type : uint16_t {
    Empty = 0,
    Null = 1,
    Integer = 2,     // 16-bit
    Long = 3,        // 32-bit
    Single = 4,      // 32-bit float
    Double = 5,      // 64-bit float
    Currency = 6,    // 64-bit scaled integer
    Date = 7,        // 64-bit double (OLE Date)
    String = 8,      // BSTR
    Object = 9,      // IDispatch*
    Error = 10,      // SCODE
    Boolean = 11,    // 16-bit (0xFFFF = True, 0x0000 = False)
    Variant = 12,    // VARIANT
    DataObject = 13,
    Decimal = 14,    // 96-bit unsigned integer + scaling
    Byte = 17,       // 8-bit unsigned
    ULong = 19,      // unsigned Long (VB7+)
    LongPtr = 20,     // Fix 081e: LongPtr - architecture-width integer (intptr_t)
    // Fix 084m: LongLong 与 LongPtr **分开**。历史实现把 LongLong 也映射到 LongPtr,
    // 于是 x86 下 (intptr_t = 4 字节) 它退化成 32 位 —— 与「有符号 64 位,
    // -2^63 .. 2^63-1」的语义不符, 实测让 Asm 过程的 int64 返回值高位全丢。
    // 语义: 恒为 64 位有符号 (int64_t), 与架构无关 (对比 LongPtr 是架构宽度)。
    LongLong = 21,
    UserDefinedType = 36,
    Array = 8192,    // bit flag
    ByRef = 16384,   // bit flag

    // 编译器内部类型
    Void = 255,
    Unknown = 254,
};

// VB6调用约定
enum class CallConv : uint8_t {
    VBDefault,   // ByRef默认
    StdCall,     // Declare语句Windows API
    CDecl,       // Declare语句C语言
    FastCall,    // 保留
};

// VB6访问级别
enum class AccessLevel : uint8_t {
    Public = 0,
    Private = 1,
    Friend = 2,     // VB6无此关键字，保留
    Protected = 3,  // tB 扩展 (B08a): 只在类家族内可见。追加在末尾 —— Default=Public 是别名,
                    // 插在中间会牵动任何按数值比较的代码
    Default = Public,
};

// 虚方法修饰位 (tB 扩展, ai/022 B08b): 只作用于 Sub/Function/Property 声明。
// VB6 原生三件套互斥 → 一个字段三种取值, 不写成三个 bool (那会出现 6 种非法组合)。
enum class ProcVirt : uint8_t {
    None = 0,        // 未写修饰符 (= VB6 默认: 不可覆盖)
    Overridable,     // 声明可被派生类覆盖
    Overrides,       // 本声明覆盖祖先的 Overridable 成员
    NotOverridable,  // 显式声明不可覆盖 (= None; 只用来与基类意图对照)
};


// VB6过程类型
enum class ProcKind : uint8_t {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
};

// VB6 的文档类别: driver 按源文件扩展名定一次 (唯一写入点在 driver_frontend)，语义层与
// 发码层都读它。`UserControl.x` / `PropertyPage.x` / `Extender.x` / `Ambient.x` 这类
// "文档隐式对象"只在对应类别的文档里存在，判据不该散在字符串猜测里 (账 #217 第二刀)。
enum class DocumentKind : uint8_t {
    Standard = 0,      // .bas / .cls
    Form = 1,          // .frm
    UserControl = 2,   // .ctl
    PropertyPage = 3,  // .pag
};

// VB6类Instancing属性
enum class VBInstancing : uint8_t {
    Private = 1,              // 仅本工程内可见（VB6默认）
    PublicNotCreatable = 2,   // 外部可用但不能New
    SingleUse = 3,            // 外部可New，每个客户独立实例
    GlobalSingleUse = 4,      // 同SingleUse，无需显式创建
    MultiUse = 5,             // 外部可New，多客户共享进程
    GlobalMultiUse = 6,       // 同MultiUse，无需显式创建
};

} // namespace vb6c3
