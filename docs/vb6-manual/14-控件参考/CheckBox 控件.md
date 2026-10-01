# CheckBox 控件

# CheckBox 控件

               

选择 **CheckBox** 控件后，该控件将显示 X，而清除 **CheckBox** 控件后， X 消失。该控件可用来提供 True/False 或者 Yes/No 选项。组中可以使用 **CheckBox** 控件显示多项选择，从而可选择其中的一项或多项。也可以通过对 Value 属性编程设置 **CheckBox** 的值。

**语法**

**CheckBox**

**说明**

**CheckBox** 和 **OptionButton** 控件功能相似，但二者之间也存在着重要差别：在一个窗体中可以同时选择任意数量的 **CheckBox** 控件。而反过来，在一个组中，在任何时侯则只能选择一个 **OptionButton** 控件。

为了在 **CheckBox** 后面显示文本，需要设置 **Caption** 属性。**Value** 属性用来确定控件的状态－选择、清除、或不可用。

---

**本项目的实现现状**（实测，不是文档转述）

- 原生窗口是 `BUTTON` + `BS_AUTOCHECKBOX`。`GotFocus` / `LostFocus` 要创建时挂 **`BS_NOTIFY`**
  才收得到 `BN_SETFOCUS=6` / `BN_KILLFOCUS=7`（账 #158，2026-09-30 起两条创建路都挂；
  裸码实测 `.build/bnnotify/bnnotify.c`）。夹具 `tests/btnfocus` 的 `BF-chk=1/1` 钉这一格。
- `Click` 走 `BN_CLICKED=0`，与 `BS_NOTIFY` 无关，一直是通的；派发现在按 `code == 0` 过滤，
  免得一次焦点移动被算成一次点击。
- `Value`（0 Unchecked / 1 Checked / 2 Grayed）与 `Style = 1`（Graphical）那两格本批没动。
