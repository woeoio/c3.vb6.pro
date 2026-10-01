# TextBox 控件

# TextBox 控件

               

**TextBox** 控件有时也称作编辑字段或者编辑控件，显示设计时输入的用户输入的、或运行时在代码中赋予控件的信息。

**语法**

**TextBox**

**说明**

为了在 **TextBox** 控件中显示多行文本，要将 **MultiLine** 属性设置为 **True**。如果多行 **TextBox** 没有水平滚动条，那么即使 **TextBox** 调整了大小，文本也会自动换行。为了在 **TextBox** 上定制滚动条组合，需要设置 **ScrollBars** 属性。

如果文本框的 **MultiLine** 属性设置为 **True** 而且它的 **ScrollBars** 没有设置为 **None** (0)，则滚动条总出现在文本框上。

如果将 **MultiLine** 属性设置为 **True**，则可以在 **TextBox** 内用 **Alignment** 属性设置文本的对齐。如果 **MultiLine** 属性是 **False**，则 **Alignment** 属性不起作用。

在 DDE 对话中，**TextBox** 控件还可以起接收端链接的作用。

---

## 本项目的实现口径

原生 `Edit`（配合 `RICHEDIT50W` 的那族是 RichTextBox 页）。这一页只记**量出来的**那几条：

| 写法 | 落点 |
|---|---|
| `TB.ScrollBars` | 创建参数 `WS_HSCROLL` / `WS_VSCROLL`；枚举照 VB6：**0 无 / 1 水平 / 2 垂直 / 3 两者**（账 #108 之前 C3 把 1/2 用反了，与 RichTextBox 两套方向） |
| `TB.MultiLine` | `ES_MULTILINE`（+ `ES_AUTOVSCROLL`）；设计期那条与容器子控件那条同一条映射 |
| 读回 `TB.ScrollBars` | 先读运行期写过的窗口属性，没有就从窗口样式推（`WS_HSCROLL` → 1、`WS_VSCROLL` → 2），所以设计期设 2 就读回 2 |

枚举的证人不是文档而是 VB6 自己存出来的 `.frm`：`tests/VBFlexGridDemo/InputForm.frm` 里
`Begin VB.TextBox Text2` 下挂着 `ScrollBars = 2  'Vertical`。**判据**在 `tests/ctrlprop`
（CP6=1 / CP7=2 / CP8=0：设计期设 1、设 2、没写过）。
