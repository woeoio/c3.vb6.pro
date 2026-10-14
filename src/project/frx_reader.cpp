// ============================================================
// .frx 二进制资源文件读取器实现
// ============================================================
#include "project/frx_reader.hpp"
#include <fstream>
#include <cstring>
#include <cstdlib>
#include <algorithm>

// Windows MultiByteToWideChar for GBK->UTF-8 conversion
#ifdef _WIN32
#include <windows.h>
#endif

namespace vb6c3 {

// 静态成员初始化
std::vector<uint8_t> FrxReader::fileData_;
std::string FrxReader::lastError_;

// ============================================================
// 加载 .frx 文件
// ============================================================
bool FrxReader::load(const std::filesystem::path& frxPath) {
    std::ifstream f(frxPath, std::ios::binary);
    if (!f.is_open()) {
        lastError_ = "Cannot open .frx file: " + frxPath.string();
        return false;
    }
    f.seekg(0, std::ios::end);
    size_t size = (size_t)f.tellg();
    f.seekg(0, std::ios::beg);
    fileData_.resize(size);
    f.read((char*)fileData_.data(), size);
    f.close();
    lastError_.clear();
    return true;
}

// ============================================================
// 读取 little-endian 值
// ============================================================
uint32_t FrxReader::readLE32(const uint8_t* p) {
    return (uint32_t)p[0] | ((uint32_t)p[1] << 8) |
           ((uint32_t)p[2] << 16) | ((uint32_t)p[3] << 24);
}

uint16_t FrxReader::readLE16(const uint8_t* p) {
    return (uint16_t)p[0] | ((uint16_t)p[1] << 8);
}

// ============================================================
// 检测图片格式
// ============================================================
FrxImageFormat FrxReader::detectFormat(const uint8_t* data, size_t size) {
    if (size < 4) return FrxImageFormat::Unknown;
    // JPEG: FF D8 FF
    if (data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF)
        return FrxImageFormat::JPEG;
    // ICO: 00 00 01 00
    if (data[0] == 0x00 && data[1] == 0x00 && data[2] == 0x01 && data[3] == 0x00)
        return FrxImageFormat::ICO;
    // BMP: "BM"
    if (data[0] == 'B' && data[1] == 'M')
        return FrxImageFormat::BMP;
    // PNG: 89 50 4E 47
    if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
        return FrxImageFormat::PNG;
    // GIF: "GIF8"
    if (data[0] == 'G' && data[1] == 'I' && data[2] == 'F' && data[3] == '8')
        return FrxImageFormat::GIF;
    // WMF: D7 CD C6 9A
    if (data[0] == 0xD7 && data[1] == 0xCD && data[2] == 0xC6 && data[3] == 0x9A)
        return FrxImageFormat::WMF;
    // EMF: 01 00 00 00 (first DWORD of EMR_HEADER)
    if (data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x00 && data[3] == 0x00 && size >= 6)
        return FrxImageFormat::EMF;
    return FrxImageFormat::Unknown;
}

// ============================================================
// 读取图片数据
// ============================================================
// .frx Picture 属性格式:
//   普通控件: [4B totalSize] [4B magic "lt\0\0"] [4B imgSize] [imgSize bytes imageData]
//   ImageList: [4B totalSize] [16B GUID] [4B magic "lt\0\0"] [4B imgSize] [imgSize bytes imageData]
FrxPictureData FrxReader::readPicture(size_t offset, bool isImageList) {
    FrxPictureData result;
    if (fileData_.empty()) {
        lastError_ = ".frx file not loaded";
        return result;
    }
    if (offset + 12 > fileData_.size()) {
        lastError_ = "Offset out of range";
        return result;
    }

    const uint8_t* base = fileData_.data() + offset;
    size_t remaining = fileData_.size() - offset;

    // 跳过头部
    size_t headerSize;
    if (isImageList) {
        // ImageList: 4B totalSize + 16B GUID + 4B magic + 4B imgSize = 28 bytes header
        headerSize = 28;
    } else {
        // Normal: 4B totalSize + 4B magic + 4B imgSize = 12 bytes header
        headerSize = 12;
    }

    if (offset + headerSize > fileData_.size()) {
        lastError_ = "Picture header exceeds file bounds";
        return result;
    }

    // 读取图片数据大小
    uint32_t imgSize;
    if (isImageList) {
        // [4B totalSize] [16B GUID] [4B magic] [4B imgSize] [data]
        imgSize = readLE32(base + 24);
    } else {
        // [4B totalSize] [4B magic] [4B imgSize] [data]
        imgSize = readLE32(base + 8);
    }

    if (imgSize == 0 || imgSize > remaining - headerSize) {
        lastError_ = "Invalid image size in .frx";
        return result;
    }

    // 提取图片数据
    const uint8_t* imgData = base + headerSize;
    result.data.assign(imgData, imgData + imgSize);
    result.format = detectFormat(imgData, imgSize);

    return result;
}

// ============================================================
// 读取多行文本 (TextBox.Text)
// ============================================================
// 格式: [1B length] [length bytes GBK text]
// 注意: 长度可能超过255, 实际格式可能是 [4B length] [data]
// 经分析 frxParse 的 Text1.Text: offset 0x27C4
//   27 27 00 00 "Text1\r\n大师的\r\n佛挡杀佛\r\n"
//   -> 4字节长度(0x27=39), 后跟39字节GBK文本
FrxTextData FrxReader::readText(size_t offset) {
    FrxTextData result;
    if (fileData_.empty()) {
        lastError_ = ".frx file not loaded";
        return result;
    }
    if (offset + 4 > fileData_.size()) {
        lastError_ = "Text offset out of range";
        return result;
    }

    const uint8_t* base = fileData_.data() + offset;
    // 尝试4字节长度前缀 (VB6 .frx 中 Text 属性使用 DWORD 长度)
    uint32_t textLen = readLE32(base);
    if (textLen == 0 || offset + 4 + textLen > fileData_.size()) {
        // 回退: 尝试1字节长度前缀
        textLen = base[0];
        if (textLen == 0 || offset + 1 + textLen > fileData_.size()) {
            lastError_ = "Invalid text length in .frx";
            return result;
        }
        result.text = gbkToUtf8(std::string((const char*)(base + 1), textLen));
    } else {
        result.text = gbkToUtf8(std::string((const char*)(base + 4), textLen));
    }
    return result;
}

// ============================================================
// 读取字符串列表 (ListBox.List)
// ============================================================
// 格式: [2B count] [per item: 2B length + length bytes GBK text]
FrxListData FrxReader::readStringList(size_t offset) {
    FrxListData result;
    if (fileData_.empty()) {
        lastError_ = ".frx file not loaded";
        return result;
    }
    if (offset + 2 > fileData_.size()) {
        lastError_ = "String list offset out of range";
        return result;
    }

    const uint8_t* base = fileData_.data() + offset;
    // VB6 .frx List format: [2B count] [2B prefix] [per item: 2B len + GBK text]
    // The prefix value equals count; skip both
    uint16_t count = readLE16(base);
    size_t pos = 4;  // skip count(2B) + prefix(2B)

    for (uint16_t i = 0; i < count; i++) {
        if (offset + pos + 2 > fileData_.size()) break;
        uint16_t strLen = readLE16(base + pos);
        pos += 2;
        if (offset + pos + strLen > fileData_.size()) break;
        std::string gbkStr((const char*)(base + pos), strLen);
        result.items.push_back(gbkToUtf8(gbkStr));
        pos += strLen;
    }
    return result;
}

// ============================================================
// 读取整数列表 (ListBox.ItemData)
// ============================================================
// Fix 195: 实测 (10 个真实工程的 .frx) 这个 blob 与 List **同构** ——
//   [2B count] [2B 前导值] [每项: 2B len + len 字节十进制文本]
// 原实现按"每项 2B 整数"读, 读到的是结构的字节本身 —— 任何工程都只解出
// 恒定假值 1/304/12288/1/304… (实测 frxParse、Charts 2020、VBFlexGridDemo
// 三个工程全是同一串), 也就是说设计期 ItemData 编进 exe 一直是垃圾。
// 段长可自证: frxParse/Form1.frx 的 ItemData@0x9A70 段长 13 字节
// = 4 + 3×(2+1), 与下一个资源偏移 0x9A7D 精确相接; 1 项工程为 7 字节。
FrxIntListData FrxReader::readIntList(size_t offset) {
    FrxIntListData result;
    if (fileData_.empty()) {
        lastError_ = ".frx file not loaded";
        return result;
    }
    if (offset + 4 > fileData_.size()) {
        lastError_ = "Int list offset out of range";
        return result;
    }

    const uint8_t* base = fileData_.data() + offset;
    uint16_t count = readLE16(base);
    size_t pos = 4;  // skip count(2B) + 前导值(2B)

    if (count > 10000) {
        lastError_ = "Invalid int list count in .frx";
        return result;
    }

    for (uint16_t i = 0; i < count; i++) {
        if (offset + pos + 2 > fileData_.size()) break;
        uint16_t strLen = readLE16(base + pos);
        pos += 2;
        if (offset + pos + strLen > fileData_.size()) break;
        std::string s((const char*)(base + pos), strLen);
        pos += strLen;
        // 十进制文本 -> int (容忍负号); 解析不出按 VB6 默认值 0 处理
        result.items.push_back(static_cast<int>(std::strtol(s.c_str(), nullptr, 10)));
    }
    return result;
}

// ============================================================
// Fix 114: 读取"属性包"持久化字符串
// ============================================================
// VB6 .frm/.ctl 中形如 `Caption = "Form2.frx":006E` 的字符串属性以属性包
// (DocProperty) 形式存入 .frx。实测布局 (Charts 2020/Form2.frx):
//   [GUID 16B] [word 0x0011] [word 0x0001] [dword cbBytes] [3B 填充] [UTF-16LE]
// 例: 偏移 0x00 -> "USD $532.00", 0x36 -> "Venta diaria", 0x6E -> "LabelPlus1"。
// 由于不同 VB6 版本填充字节数可能不同, 这里在 offset 起 64 字节范围内扫描
// 首个"长度 dword + 全部可打印 UTF-16"的合法组合, 兼顾偏移前导结构差异。
FrxTextData FrxReader::readDocString(size_t offset) {
    FrxTextData result;
    if (fileData_.empty()) {
        lastError_ = ".frx file not loaded";
        return result;
    }
    if (offset >= fileData_.size()) {
        lastError_ = "DocString offset out of range";
        return result;
    }

    const size_t kScanWindow = 64;
    const size_t kMaxLen = 0x4000;   // 16KB 上限, 防御性
    size_t scanEnd = std::min(fileData_.size(), offset + kScanWindow);

    // 以"数据起点"为主循环: 长度 dword 位于数据起点前 (4+pad) 字节 (pad=0..8),
    // 兼容 VB6 不同版本的填充差异 (实测 pad=3)。
    //
    // 账(乱码): 这里以前有一句 `if ((ds & 1) != 0) continue;` —— **假定 UTF-16 数据必在
    // 偶偏移**。实测 (ComCtlsDemo MainForm.frx) 该串数据落在**奇偏移** 0xA2B; 被跳掉后
    // 扫描退一格取 0xA2A, 每对字节读成 (0x00, 0x49) ⇒ wchar 0x4900 ⇒ "ImageCombo1"
    // 整条变成 CJK 假字 (\u4900\u6D00…), ComboBoxEx 显示的就是这串 (屏上「整条 CJK 乱码」)。
    // 现在**不做奇偶限制**、改**逐候选打分取最优**: 设计期串绝大多数是 ASCII, 而"错位
    // 一个字节"的解码会把低字节恒置 0 (wchar = 0xXX00) ⇒ ASCII 计数为 0; 正确对齐的
    // ASCII 计数 = 码元数。取 ASCII 计数最大者即可稳定选回正确对齐 (平局保留先到者)。
    FrxTextData best;
    int bestAscii = -1;
    for (size_t ds = offset; ds + 2 <= scanEnd; ++ds) {
        for (size_t pad = 0; pad <= 8; ++pad) {
            if (ds < offset + 4 + pad) break;
            size_t lp = ds - (4 + pad);
            uint32_t nBytes = readLE32(fileData_.data() + lp);
            if (nBytes < 2 || nBytes > kMaxLen || (nBytes & 1) != 0) continue;
            if (ds + nBytes > fileData_.size()) continue;
            const uint8_t* s = fileData_.data() + ds;
            // 校验: 每个 UTF-16 码元都必须是可打印字符 (>= 0x20) 或常见空白
            bool ok = true;
            int ascii = 0;
            for (uint32_t i = 0; i < nBytes; i += 2) {
                uint32_t cp = (uint32_t)s[i] | ((uint32_t)s[i + 1] << 8);
                if (cp < 0x0080) ascii++;
                if (cp == 0x0009 || cp == 0x000A || cp == 0x000D) continue;
                if (cp < 0x0020) { ok = false; break; }
            }
            if (!ok) continue;
            if (ascii <= bestAscii) continue;   // 已有更"像文本"的候选
            int wlen = (int)(nBytes / 2);
#ifdef _WIN32
            std::wstring wstr((const wchar_t*)s, wlen);
            int ulen = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), wlen, NULL, 0, NULL, NULL);
            if (ulen <= 0) continue;
            std::string utf8(ulen, 0);
            WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), wlen, &utf8[0], ulen, NULL, NULL);
            best.text = utf8;
#else
            best.text.assign((const char*)s, nBytes);  // 非 Windows 下退化为原始字节
#endif
            bestAscii = ascii;
        }
    }
    if (!best.text.empty()) return best;

    lastError_ = "No valid DocString at offset";
    return result;
}

// ============================================================
// Fix: 属性包块 (UserControl 持久化 PropertyBag) 识别 + 取原始字节
//   块 = [DWORD cbSize][cbSize 字节负载]
//   负载 = [28B 前缀: 00 | GUID(16) | 00010000 | 00000006 | 000000 ]
//          [VB6 PropertyBag 序列化数据: WORD 0xB293 ; WORD catOff ; DWORD cbContent ; items...]
//   判据: 前 64B 内存在 0xB293 且 该处偏移 + cbContent == cbSize (实测恒 28 + cbContent)。
//   这样既命中所有 Init* 属性包, 又不会误判 DocString / 图片块。
// ============================================================
FrxPropBagBlob FrxReader::readPropBagBlob(size_t offset) {
    FrxPropBagBlob out;
    if (offset + 4 > fileData_.size()) return out;
    uint32_t cbSize = readLE32(fileData_.data() + offset);
    if (cbSize < 8 + 12 || offset + 4 + cbSize > fileData_.size()) return out;
    const uint8_t* blob = fileData_.data() + offset + 4;

    size_t scanMax = (cbSize < 64) ? cbSize : 64;
    for (size_t h = 0; h + 8 <= cbSize && h <= scanMax; ++h) {
        if (readLE16(blob + h) != 0xB293) continue;
        uint32_t cbContent = readLE32(blob + h + 4);
        if (cbContent < 8 || cbContent > cbSize) continue;
        if (h + cbContent != cbSize) continue;          // 必须正好补满块
        out.found = true;
        out.bytes.assign(blob, blob + cbSize);
        return out;
    }
    return out;
}

// ============================================================
// GBK -> UTF-8 转换
// ============================================================
std::string FrxReader::gbkToUtf8(const std::string& gbk) {
#ifdef _WIN32
    if (gbk.empty()) return "";
    // GBK -> UTF-16
    int wlen = MultiByteToWideChar(936, 0, gbk.c_str(), (int)gbk.size(), NULL, 0);
    if (wlen <= 0) return gbk;  // fallback
    std::wstring wstr(wlen, 0);
    MultiByteToWideChar(936, 0, gbk.c_str(), (int)gbk.size(), &wstr[0], wlen);
    // UTF-16 -> UTF-8
    int ulen = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), wlen, NULL, 0, NULL, NULL);
    if (ulen <= 0) return gbk;
    std::string utf8(ulen, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), wlen, &utf8[0], ulen, NULL, NULL);
    return utf8;
#else
    return gbk;  // non-Windows: no conversion
#endif
}

} // namespace vb6c3
