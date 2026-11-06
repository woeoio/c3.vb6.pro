#pragma once

#include <string>
#include <vector>

namespace vb6c3 {

// 账 #305 第328刀: VB6 把**工程自己的**窗体图标嵌进 EXE —— 来处 = `.vbp` 的 `IconForm=`
// (没写就落 `Startup=` 那枚窗体), 去处 = 那张窗体的 `.frm` 里 `Icon = "<名>.frx":<hex>` 那一格,
// `.frx` 里那条 `"lt"` 记录的载荷就是**一整枚 .ICO** (ICONDIR + 目录 + 各帧)。
// 实测口径 (语料里能按 `ExeName32`/`Path32` 找回自己那台 VB6 产物的 151 份 EXE 工程):
//   · `IconForm` 指向的窗体带 Icon 记录 ⇒ 产物里 RT_ICON 的**帧数字节集**与该记录逐帧相同 = 137/138；
//   · 没写 IconForm 而 Startup 窗体带 Icon ⇒ 同样照抄 = 12/14 (两台的产物比源码旧，是它自己的账)；
//   · 那张窗体**没有** Icon 那一格 ⇒ VB6 塞默认窗体图 (3 帧 304/744/296) = 81/81，正是第303刀那一份。
// 所以这一族只欠一件事：把记录里的帧**逐字节**搬进资源，其余口径沿用第303刀 (组号、语言档、闸门)。
// 帧数可以很多 (语料实测 1..16 帧) ⇒ 默认那一份只有 3 帧, 不能拿来当上限。

struct FormIconFrame {
    unsigned char dir[12];            // ICONDIRENTRY 的前 12 字节 (宽/高/色/平面/位深/帧长/帧偏移)
    std::vector<unsigned char> image;  // 该帧载荷, 逐字节照抄 (BM 或 PNG 都不重排)
};

struct FormIconData {
    std::vector<FormIconFrame> frames;
    std::string recordPath;           // 命中的 .frx (诊断与判据用)
};

// 读一枚 `.frm`：找 `Icon = "<file>.frx":<4hex>` 那一行，按偏移打开同目录的 `.frx`，
// 校验记录头 (4B 记录长 / 4B 型标 "lt" / 4B 内层长) 后取出那枚 .ICO。
// 任何一环不认识 ⇒ false，调用方落回默认窗体图标那一档 (第303刀) —— 宁可不画，不猜。
bool loadFormIconFromFrm(const std::string& frmPathUtf8, FormIconData& out, std::string& why);

// 把取到的 .ICO 拍成两份注入材料：每帧一枚 RT_ICON 载荷 + 一枚 RT_GROUP_ICON 的组字节
// (GRPICONDIR + N × GRPICONDIRENTRY，目录项末字段从 .ICO 的“帧偏移”换成资源 id)。
// 组号与帧 id 起始值都复用默认图标那两位权威 (见 default_form_icon.hpp)，不开第二问。
bool buildFormIconResources(const FormIconData& icon,
                            std::vector<std::vector<unsigned char>>& framesOut,
                            std::vector<unsigned int>& idsOut,
                            std::vector<unsigned char>& groupOut);

}  // namespace vb6c3
