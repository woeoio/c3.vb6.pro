#pragma once
// vb6forms.h - VB6 Win32 窗体运行时 (P7) 公共头（伞头）
// 为 C 代码生成器提供窗体/控件运行时 API 的 C 声明；编译器生成的 C 代码
// （cgen_base）与 RTL 各 vb6forms*.c 都只 include 本文件，不要直接用子头。
// 2026-09-17 按功能域拆为 9 个子头，本文件只负责按依赖顺序 include。
//
// 依赖顺序（勿调整，与拆分前的单文件声明顺序一致）：
//   window -> prop -> prop_ctrl -> prop_pic -> prop_form
//   -> ctrlarr -> mdi -> webview -> controls

#ifdef _WIN32
#include <windows.h>
#endif

#include "vb6forms_window.h"
#include "vb6forms_prop.h"
#include "vb6forms_prop_ctrl.h"
#include "vb6forms_prop_pic.h"
#include "vb6forms_prop_form.h"
#include "vb6forms_ctrlarr.h"
#include "vb6forms_mdi.h"
#include "vb6forms_webview.h"
#include "vb6forms_controls.h"
// Form/Printer 绘图方法家族 (PSet/Line/Circle/Point/Cls)。放在最后: 它 include
// vb6forms.h 自身, 靠这里的末尾位置保证前序声明都可见。
#include "vb6forms_draw.h"
