# OptionButton 控件

**OptionButton** 控件显示一个可以打开或者关闭的选项。

**语法**

**OptionButton**

**说明**

在选项组中用 **OptionButton** 显示选项，用户只能选择其中的一项。在 **Frame** 控件、**PictureBox** 控件或者窗体这样的容器中绘制 **OptionButton** 控件，就可以把这些控件分组。为了在 **Frame** 或者**PictureBox** 中将 **OptionButton** 控件分组，首先绘制 **Frame** 或 **PictureBox**，然后在内部绘制 **OptionButton** 控件。同一容器中的 **OptionButton** 控件为一个组。

**OptionButton** 控件和 **CheckBox** 控件功能相似，但是二者间也存在着重要差别。在选择一个 **OptionButton** 时，同组中的其它 **OptionButton** 控件自动无效。相反，可以选择任意数量的 **CheckBox** 控件。

---

**本项目的实现现状**（实测，不是文档转述）

- 原生窗口是 `BUTTON` + `BS_AUTORADIOBUTTON`。`GotFocus` / `LostFocus` 同样要 **`BS_NOTIFY`**
  才有 `BN_SETFOCUS=6` / `BN_KILLFOCUS=7`（账 #158，2026-09-30 起两条创建路都挂）。
  夹具 `tests/btnfocus` 的 `BF-opt=1/1` 钉这一格。
- 组内方向键（`VK_UP` / `VK_DOWN`）由本项目自己派发：同一容器里所有 `BS_AUTORADIOBUTTON`
  按 **`TabIndex`** 走、到尾回绕，换选走 `BM_CLICK`（= 自动取消同组别人 + 发 `BN_CLICKED`，
  所以 `_Click` 会跟着来）。这一步之前交给你的是对话框管理器：它按 z-order 走，
  而且走到链尾会**跳出容器**（实测 `down=cmdTop1`）⇒ 那是账 #168 的由来。
  **一声方向键 = 恰好一条 `_Click`**（账 #171 起；夹具 `tests/tabwalk` 的 `clicks=3` 钉死条数）。
- 账 #171 修掉的那一条值得记，因为它是**系统替你发的**：单选钮在**焦点进入**时也会向父窗发一条
  `BN_CLICKED`，而那一刻它自己**并没有被勾上**（两头证人：`Value` 与 OS 的 `BM_GETCHECK` 都答「没勾」）。
  光按 `code == 0` 筛不掉它（号就是 0，账 #158 那道筛管的是 `code=6/7`）⇒ 现在的口径是
  **「来这条通知时它自己勾上了没有」**，唯一落点在 RTL `vb6_RadioClickCounts`（`vb6forms_prop.h`），
  发码侧只对 `OptionButton` 的 arm 加这一道筛，别的型（`CommandButton` / `CheckBox`）一字不加 ——
  复选框取消勾选是一次正经 `Click`，不能跟着筛。
  证人两头各钉一处：`tests/btnfocus` 的 `BF-noclick` 现在把单选钮也算进去（修前读 `N`），
  `tests/tabwalk` 的 `clicks` 从「只钉字段存在」翻成钉 `3`（修前读 `6`）。
  排除过的四条（别再往这些方向找）：两条 `SetFocus` 路（VB 方法 vs 直接 user32）读数同形；
  `C3_OCX_NO_DLGMSG=1` 下读数一字不变；`--emit-c` 里这两枚单选钮一次子类化都没挂 ⇒
  **与账 #161 那族（子类重发）不同因**；焦点移到本来就勾着的那枚 = 0 条。
- 焦点进出容器里的单选钮另有一处和 VB6 不同：系统会把 `WS_TABSTOP` 自己挪到**勾选那枚**身上，
  所以「谁是站」读的是创建时存的那份 `VB6_TabStop`，不是实时样式位（账 #163/#168）。
- 同一容器里多枚 OptionButton 的互斥由原生 BUTTON 类给；`Value` 的读写与 `Caption` 那几格本批没动。
