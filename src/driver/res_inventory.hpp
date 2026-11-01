// res_inventory.hpp - 读一份 .res（rc.exe 的输出）里**到底有什么资源**，不链接、不加载。
//
// 存在的理由（ai/029 C29-M）：产物清单的让位判据从前只看"工程给了 ResFile= 没有"
// （`driver_link.cpp` 里 `userResFile_.empty()`），于是"带一份只有图标/版本信息的 .res"
// 这种很常见的 VB6 工程会**连内置的 comctl v6 清单一起让掉** ⇒ 产物静默退回 v5.82。
// 正确的判据是问那份 .res 本身：**有没有那条作为进程激活上下文入口的清单资源**。
//
// 这里只解析 rc.exe 的 RESFMT（资源目录表），不依赖任何 Windows 私有头。
#ifndef C3_DRIVER_RES_INVENTORY_H
#define C3_DRIVER_RES_INVENTORY_H

#include <cstddef>
#include <cstdint>
#include <string>

namespace vb6c3 {

class ResFileProbe {
public:
    bool   parsed = false;              // 文件读到且目录走通（false ⇒ 结论不可用）
    bool   hasAppManifest = false;      // RT_MANIFEST(24) + 名字 1 = CREATEPROCESS
    bool   hasIsolatedManifest = false; // RT_MANIFEST(24) + 名字 2 = ISOLATED，**不参与进程激活上下文**
    bool   hasIconGroup = false;        // RT_GROUP_ICON(14)：工程自带图标 ⇒ 不再塞默认窗体图标（账 #305 第303刀）
    int    entries = 0;                 // 走到的资源条数（诊断用）
    std::string why;                    // 一行结论，给 stderr 的 C3: 行用

    // 读 + 解析。返回 false == 读不动/认不动（调用方按"保守让位"处置）。
    bool loadFromPath(const std::string& resPath);

private:
    bool scan(const std::uint8_t* data, std::size_t size);
};

} // namespace vb6c3

#endif // C3_DRIVER_RES_INVENTORY_H
