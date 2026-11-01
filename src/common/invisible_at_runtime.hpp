// 账 #305 第二百五十八刀: `UserControl.InvisibleAtRuntime` —— 这一位在 VB6 里的意思是
// **根本不建窗口**（不是「建了再藏」：参考那台里 vis=0 的 UC 窗口照样在窗口枚举里，
// 而声明了 InvisibleAtRuntime 的 `.ctl` 实例在参考里一枚窗口都没有）。
//
// 它住在这枚 UC 自己的 `.ctl` 头（`Begin VB.UserControl <名>` 那一块）里，与 CanGetFocus /
// HasDC 同一张属性表 —— 也就是「控件源自己声明」，不是布局藏起来的。C3 从前完全不认这一位
// （全仓 0 命中），于是 ComCtlsDemo 的三枚 ImageList 各多一枚 83x25 宿主 + 内部 VB.Image 的
// 40x40 Static。
//
// 唯一消费者 = 4 处宿主创建站（cgen_form_create_controls.inc 顶层/容器内、
// cgen_form_frame_menu.inc 同两形），它们把答案折成 vb6_UC_SetWindowless(1) 的置/清。
#ifndef C3_COMMON_INVISIBLE_AT_RUNTIME_HPP
#define C3_COMMON_INVISIBLE_AT_RUNTIME_HPP

#include "project/frm_parser.hpp"

namespace vb6c3 {

// spec = 被实例化那一枚 UC 的 .ctl 解析结果 (可以为空 = 不是工程内 UC)
inline bool ucInvisibleAtRuntime(const FrmFile* spec) {
    if (!spec) return false;
    const auto& props = spec->form.formControl.properties;
    auto it = props.find("InvisibleAtRuntime");
    if (it == props.end()) return false;
    return it->second.intValue != 0;
}

}  // namespace vb6c3

#endif  // C3_COMMON_INVISIBLE_AT_RUNTIME_HPP
