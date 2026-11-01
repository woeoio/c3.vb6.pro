// res_inventory.cpp - 走一遍 rc.exe 的 .res（RESFMT 资源目录表），回答"里面有没有清单"。
//
// 只读表、不解析资源内容 ⇒ 不依赖 Windows 私有头。
//
// RESFMT 的一条记录 = 32 字节表头 +（可选）紧跟的宽字符串 + 数据（4 字节对齐）：
//   +0  DataSize        0xFFFFFFFF = 目录项（group icon / group cursor），没有数据
//   +4  HeaderSize      恒 32
//   +8  Type            LOWORD==0xFFFF ⇒ HIWORD 是序号；0xFFFFFFFF ⇒ 类型串紧跟表头
//   +12 Name            同上
//   +16 DataVersion  +20 MemoryFlags  +22 LanguageId  +24 Version  +28 Characteristics
// 这里关心的序号：RT_MANIFEST = 24；名字 1 = CREATEPROCESS_MANIFEST_RESOURCE_ID（进程
// 激活上下文的入口），2 = ISOLATED_APPLICATION（**不参与**进程激活上下文 ⇒ 不能当"自带"）。

#include "driver/res_inventory.hpp"

#include <cstddef>
#include <fstream>
#include <vector>

namespace vb6c3 {
namespace {

constexpr std::uint32_t kDirOrStringMarker = 0xFFFFFFFFu;  // DataSize / Type / Name 都用它表"另一种"
constexpr std::uint32_t kOrdinalPrefix     = 0x0000FFFFu;  // LOWORD==0xFFFF ⇒ HIWORD 是序号
constexpr std::size_t   kMinHeaderSize     = 32;

constexpr std::uint16_t kRtManifest            = 24;
constexpr std::uint16_t kRtGroupIcon           = 14;
constexpr std::uint16_t kManifestAsApp         = 1;
constexpr std::uint16_t kManifestIsolated      = 2;

std::uint32_t le32(const std::uint8_t* p) {
    return static_cast<std::uint32_t>(p[0])
         | (static_cast<std::uint32_t>(p[1]) << 8)
         | (static_cast<std::uint32_t>(p[2]) << 16)
         | (static_cast<std::uint32_t>(p[3]) << 24);
}

// 紧跟表头的宽字符串占几字节（含结束符）；找不到结束符给 0，让调用方判"解析失败"。
std::size_t wideStringLen(const std::uint8_t* base, std::size_t remain) {
    for (std::size_t i = 0; i + 1 < remain; i += 2) {
        if (base[i] == 0 && base[i + 1] == 0) return i + 2;
    }
    return 0;
}

bool ordinalOf(std::uint32_t field, std::uint16_t& out) {
    if ((field & 0xFFFFu) != kOrdinalPrefix) return false;
    out = static_cast<std::uint16_t>(field >> 16);
    return true;
}

} // namespace

bool ResFileProbe::scan(const std::uint8_t* data, std::size_t size) {
    std::size_t pos = 0;
    while (pos + kMinHeaderSize <= size) {
        const std::uint8_t* h = data + pos;
        const std::uint32_t dataSize   = le32(h + 0);
        const std::uint32_t headerSize = le32(h + 4);
        const std::uint32_t typeField  = le32(h + 8);
        const std::uint32_t nameField  = le32(h + 12);

        if (headerSize < kMinHeaderSize || headerSize > size - pos) {
            why = "第 " + std::to_string(entries) + " 条的 HeaderSize 不合法";
            return false;
        }
        std::size_t cur = pos + headerSize;

        std::uint16_t typeOrd = 0, nameOrd = 0;
        const bool typeIsOrd = ordinalOf(typeField, typeOrd);
        const bool nameIsOrd = ordinalOf(nameField, nameOrd);

        // 串在前、名字在后，都要跳掉才能定位数据
        if (!typeIsOrd && typeField == kDirOrStringMarker) {
            std::size_t n = wideStringLen(data + cur, size - cur);
            if (!n) { why = "类型串没有结束符"; return false; }
            cur += n;
        }
        if (!nameIsOrd && nameField == kDirOrStringMarker) {
            std::size_t n = wideStringLen(data + cur, size - cur);
            if (!n) { why = "名字串没有结束符"; return false; }
            cur += n;
        }

        ++entries;
        if (typeIsOrd && typeOrd == kRtManifest && nameIsOrd) {
            if      (nameOrd == kManifestAsApp)    hasAppManifest = true;
            else if (nameOrd == kManifestIsolated) hasIsolatedManifest = true;
        }
        if (typeIsOrd && typeOrd == kRtGroupIcon) hasIconGroup = true;

        if (dataSize == kDirOrStringMarker) { pos = cur; continue; }   // 目录项无数据
        if (dataSize > size - cur) { why = "数据长度越过文件尾"; return false; }
        const std::size_t padded =
            (static_cast<std::size_t>(dataSize) + 3u) & ~static_cast<std::size_t>(3u);
        if (padded > size - cur) { why = "对齐后的数据越过文件尾"; return false; }
        pos = cur + padded;
    }
    if (entries == 0) { why = "一条资源都没有"; return false; }
    why = std::to_string(entries) + " 条资源，"
        + (hasAppManifest ? "含 #1 清单"
                          : (hasIsolatedManifest ? "只有 #2(ISOLATED) 清单" : "无 #1 清单"));
    return true;
}

bool ResFileProbe::loadFromPath(const std::string& resPath) {
    std::ifstream f(resPath, std::ios::binary);
    if (!f) { why = "打不开"; return false; }
    f.seekg(0, std::ios::end);
    const std::streamoff n = f.tellg();
    f.seekg(0, std::ios::beg);
    if (n <= 0)                       { why = "尺寸为 0"; return false; }
    if (static_cast<std::size_t>(n) > 64u * 1024u * 1024u)
                                      { why = "大得离谱(>64MB)"; return false; }
    std::vector<std::uint8_t> buf(static_cast<std::size_t>(n));
    if (!f.read(reinterpret_cast<char*>(buf.data()), n)) { why = "读不完整"; return false; }
    parsed = scan(buf.data(), buf.size());
    return parsed;
}

} // namespace vb6c3
