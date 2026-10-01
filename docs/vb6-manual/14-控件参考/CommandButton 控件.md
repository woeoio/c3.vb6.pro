# CommandButton 控件

**CommandButton** 控件可以开始、中断或者结束一个进程。选取这个控件后，**CommandButton** 显示按下的形状，所以有时也称之为下压按钮。

**语法**

**CommandButton**

**说明**

为了在 **CommandButton** 控件上显示文本，需要设置其 **Caption** 属性。可以通过单击 **CommandButton** 选中这个按钮。为了能够在按 ENTER 键时也选中命令按钮，需要将其 **Default** 属性设置为 **True**。为了能够按 ESC 键时也选中 **CommandButton**，则需要将 **CommandButton** 的 **Cancel** 属性设置成 **True**。

---

**本项目的实现现状**（实测，不是文档转述）

- `GotFocus` / `LostFocus` 走父窗 `WM_COMMAND` 的 `BN_SETFOCUS=6` / `BN_KILLFOCUS=7`（账 #158）。
  这两位要求创建时挂 **`BS_NOTIFY`** —— 没挂的 BUTTON 类**一条焦点通知都不发**（裸码实测
  `.build/bnnotify/bnnotify.c`：同窗两枚按钮只差这一位，挂了的程序化 `SetFocus` 与 Tab 键都发，
  没挂的一条不发）。本项目自 2026-09-30 起两条创建路都挂这一位，夹具 `tests/btnfocus` 钉着。
- `Click` 走同一条 `WM_COMMAND` 的 `BN_CLICKED=0`。这一位**不需要** `BS_NOTIFY`，所以按钮的
  `Click` 一直能用；但也因为同一个 id 现在会携 `code=6/7` 进来，派发性只认 `code == 0`。
- 放在 Frame / PictureBox 里的按钮同样有效 —— 派发那两趟（顶层与容器子控件）现在共用同一张码表。
- `Style = 1`（Graphical）走 `BS_PUSHLIKE`，本批没动那一格；`Default` / `Cancel` 的键盘口径也不在本批范围内。
