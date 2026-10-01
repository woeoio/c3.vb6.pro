# ComboBox 控件

# ComboBox 控件

               

**ComboBox** 控件将 **TextBox** 控件和 **ListBox** 控件的特性结合在一起－既可以在控件的文本框部分输入信息，也可以在控件的列表框部分选择一项。

**语法**

**ComboBox**

**说明**

为了添加或删除 **ComboBox** 控件中的项目，需要使用 **AddItem** 或 **RemoveItem** 方法。设置 **List**、**ListCount**、和 **ListIndex** 属性，使访问 **ComboBox** 中的项目成为可能。也可以在设计时使用 **List** 属性将项目添加到列表中。

**注意** 只有当 **ComboBox** 的下拉部分的内容被滚动时，Scroll 事件才在 **ComboBox** 中发生，而不是每次 **ComboBox** 的内容改变时。例如，如果 **ComboBox** 的下拉部分包含五行，并且最顶上的项为突出显示，则在您按完向下箭头键六下（或按一次 PgUp 键）之前 Scroll 事件不发生。再往后，每按一次向上箭头键引发一次 Scroll 事件。

---

**本项目的实现现状**（实测，不是文档转述）

- `GotFocus` / `LostFocus` 走父窗 `WM_COMMAND` 的 `CBN_SETFOCUS=3` / `CBN_KILLFOCUS=4`（账 #158）。
  这一对在 2026-09-30 之前**一直不触发** —— 发码那两支撑的是 `code == 1024 / 2048`，那是 `CBEM_*`
  （发给 ComboBoxEx 的消息号），从来不会作为 notification code 出现。夹具 `tests/combofocus` 钉着：
  程序化 `SetFocus` 一进一出各触发一次，且与 ListBox / TextBox 的焦点通知互不串台
  （`LBN_SETFOCUS` 与 `CBN_KILLFOCUS` **同为 4**，全靠 `id` 过滤分开）。
- `Validate` 走的是另一条派发，一直是对的；带 `Validate` 时 `LostFocus` 会先问 `VB6_ValidateCancel`。
- `Style = 2`（Dropdown List）与 `Style = 0`（Dropdown Combo）在上述通知上一致；`Simple` 样式那一档本批没测。
