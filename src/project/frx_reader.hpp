#pragma once
// ============================================================
// .frx 二进制资源文件读取器
// 在编译时读取 .frx 文件中指定偏移处的二进制数据
// ============================================================

#include <string>
#include <vector>
#include <cstdint>
#include <filesystem>

namespace vb6c3 {

// .frx 数据类型
enum class FrxDataType {
    Picture,        // 图片属性 (Form/PictureBox/Image/CommandButton/OptionButton/CheckBox)
    PictureImageList, // ImageList 的 ListImage 图片
    Text,           // TextBox.Text 多行文本
    ListString,     // ListBox.List 字符串数组
    ListInt,        // ListBox.ItemData 整数数组
    Unknown,
};

// 图片格式
enum class FrxImageFormat {
    JPEG,   // FF D8 FF
    ICO,    // 00 00 01 00
    BMP,    // "BM"
    PNG,    // 89 50 4E 47
    WMF,    // D7 CD C6 9A
    EMF,    // 01 00 00 00
    GIF,    // 47 49 46 38
    Unknown,
};

// 读取结果
struct FrxPictureData {
    std::vector<uint8_t> data;      // 原始图片数据 (已去除 .frx 头部)
    FrxImageFormat format = FrxImageFormat::Unknown;
};

struct FrxTextData {
    std::string text;               // GBK 编码的文本 (已转换为 UTF-8)
};

struct FrxListData {
    std::vector<std::string> items; // 字符串列表 (已转换为 UTF-8)
};

struct FrxIntListData {
    std::vector<int> items;         // 整数列表
};

// Fix: .frx 里的 "属性包块" (UserControl 的持久化 PropertyBag)
//   .frm 记 `InitPanels = "MainForm.frx":0000`, 该偏移处是一整块
//   [DWORD cbSize][28B 前缀][VB6 PropertyBag 序列化数据]。
//   运行期 .ctl 的 ReadProperties 会 `PropBag.Contents = <该 blob>` 再按名读字段,
//   所以编译器只需把**原始字节**交出去 (解析在 RTL 侧做), 见 uc_propbag.c。
struct FrxPropBagBlob {
    bool found = false;                 // 该偏移处是否确为属性包块
    std::vector<uint8_t> bytes;         // 块负载 (cbSize 字节, 含 28B 前缀)
};

// .frx 里的设计期 StdFont 序列化块 (`.frm` 里的 `font = "Form3.frx":0000`)
//   布局实测 (Form3/Form5.frx):
//     [DWORD cbSize][GUID StdFont 16B][DWORD 版本][WORD 字重]
//     [DWORD 字号×10000][BYTE 标志][ASCII 字体名 + NUL][填充到 cbSize]
//   `.ctl` 的 ReadProperties 用 `Set m_Font = .ReadProperty("Font", Ambient.Font)`,
//   属性包里必须是 VT_DISPATCH(IFont), 字符串槽会被忽略 → 退回 ambient 8.25pt。
struct FrxFontData {
    bool found = false;
    std::string name;                   // 字体名 (ASCII, 与 VB6 一致)
    double size = 8.25;                 // 字号 (pt)
    int weight = 400;
    int charset = 0;
};

// .frx 里的 "DocProperty 字节块" (UserControl 设计期二进制属性, 如 `image_1`)
//   布局: [DWORD cbSize][1B 00][GUID 16B][WORD 00][WORD 01][DWORD nBytes]
//         [4B 0][nBytes 负载]  ——  payload 起点 = offset+32, cbSize = 28+nBytes
//   Caption 这类字符串属性同一布局, 只是 payload 是 UTF-16 文本 (readDocString)。
//   图片属性的 payload 是 PNG/JPEG/BMP…, 必须按**原始字节**交给属性包
//   (VB6 里 `ReadProperty("image_N", 0)` 交回 Byte() ⇒ .ctl 的 ItemImage 才拿到图)。
struct FrxBlobBytes {
    bool found = false;
    std::vector<uint8_t> bytes;
};

// .frx 里的 `OleObjectBlob` (设计器 OCX 的持久化数据)
//   布局实测: [4B "LB" + WORD 版本][DWORD 负载长度][DWORD Left][DWORD Top]
//             [DWORD Width][DWORD Height][负载 = 控件 IPersistStream 数据]
//   TeeChart 的负载是 Delphi VCL 的 TPF0 流; VB6 建完控件后用
//   IPersistStreamInit::Load 把**负载**灌回去 (24B 包装头不进流)。
struct FrxOleBlob {
    bool found = false;                 // 该偏移处是否为 LB 包装的 OleObjectBlob
    std::vector<uint8_t> bytes;         // 包装头之后的持久化负载 (交给 OCX)
};

// ============================================================
// FrxReader: 读取 .frx 文件
// ============================================================
class FrxReader {
public:
    // 加载 .frx 文件到内存 (返回是否成功)
    static bool load(const std::filesystem::path& frxPath);

    // 读取图片数据 (Picture 属性)
    // isImageList: 是否为 ImageList 的 ListImage (有额外16字节GUID头部)
    static FrxPictureData readPicture(size_t offset, bool isImageList = false);

    // 读取多行文本 (TextBox.Text)
    static FrxTextData readText(size_t offset);

    // Fix 114: 读取"属性包"持久化字符串 (.frm/.ctl 中 `Xxx = "Form2.frx":006E`)
    //   VB6 属性包 DocProperty 编码: [GUID 16B][word 0x0011][word 0x0001]
    //                               [dword cbBytes][3B 填充][UTF-16LE 文本]
    //   实测于 Charts 2020/Form2.frx: LabelPlus1/2/3 的 Caption 即此格式。
    //   为稳健起见在 offset 起 64 字节内扫描首个合法的"长度+UTF-16 文本"组合。
    static FrxTextData readDocString(size_t offset);

    // 读取字符串列表 (ListBox.List)
    static FrxListData readStringList(size_t offset);

    // Fix: 识别/取出 "属性包块" (见上面 FrxPropBagBlob 注释)
    //   判据: 块大小 cbSize 与前缀里的 PropertyBag 头部自洽 ——
    //   签名 WORD 0xB293 (扫描前 64B) 且 头部偏移 + cbContent == cbSize。
    //   该判据在实测的 InitPanels/InitBands/InitButtons/InitListImages/InitTabs
    //   上逐一命中, 且不会误伤 DocString / 图片块。
    static FrxPropBagBlob readPropBagBlob(size_t offset);

    // 读取设计期 StdFont 序列化块 (见上面 FrxFontData 注释)
    static FrxFontData readFont(size_t offset);

    // 读取设计器 OCX 的 `OleObjectBlob` (见上面 FrxOleBlob 注释)
    static FrxOleBlob readOleBlob(size_t offset);

    // 读取 DocProperty 字节块负载 (见 FrxBlobBytes 注释) —— 只做结构自洽校验,
    // 是否真是图片由调用方 detectFormat 判断。
    static FrxBlobBytes readDocBlobBytes(size_t offset);

    // 读取整数列表 (ListBox.ItemData)
    static FrxIntListData readIntList(size_t offset);

    // 获取最后错误信息
    static const std::string& lastError() { return lastError_; }

    // 检测图片格式
    static FrxImageFormat detectFormat(const uint8_t* data, size_t size);

    // GBK 转 UTF-8
    static std::string gbkToUtf8(const std::string& gbk);

private:
    static std::vector<uint8_t> fileData_;  // 整个 .frx 文件内容
    static std::string lastError_;

    // 从文件数据中读取 little-endian DWORD
    static uint32_t readLE32(const uint8_t* p);
    // 从文件数据中读取 little-endian WORD
    static uint16_t readLE16(const uint8_t* p);
};

} // namespace vb6c3
