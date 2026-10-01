# PictureBox 控件

# PictureBox 控件

               

**PictureBox** 控件可以显示来自位图、图标或者元文件，以及来自增强的元文件、JPEG 或 GIF 文件的图形。如果控件不足以显示整幅图象，则裁剪图象以适应控件的大小。

**语法**

**PictureBox**

**说明**

也可以用 **PictureBox** 控件将 **OptionButton** 控件分组，并用该控件显示图形方法的输出和 **Print** 方法写入的文本。

为了使 **PictureBox** 控件能够自动调整大小以显示整幅图形，将它的 **AutoSize** 属性设置成 **True**。

可在代码中操作图形属性和方法，以创建动画或进行仿真。对运行时的打印操作，例如修改屏幕窗体格式以便打印，Graphics 属性和事件是很有用的。

在 DDE 对话中，**PictureBox** 控件还可以起接收端链接的作用。

**PictureBox** 控件和 **Data** 控件是唯一可以放置在 MDI 窗体内部区域的标准 Visual Basic 控件。可以使用该控件在内部区域的顶部或底部对控件分组，以创建工具栏或状态栏。

**注意**   Unisys Corporation 有一项专利，该专利声称涉及到 GIF-LZW 压缩技术的某些方面，在该技术中使用了 PictureBox 和 Image 控件。Microsoft Corporation 于1996年9月获得了对 Unisys LZW 专利的使用许可。然而，Microsoft 的许可证并不延伸到那些软件开发商或第三方，他们使用任何 Microsoft 工具包、语言开发或操作系统产品来在他们自己的产品中提供 GIF 读/写和/或任何其他 LZW 能力（例如，通过 DLL 和 API）。

如果您的商业应用程序使用了这些控件之一（并且因此使用了 LZW 技术），您可能会希望获得有关专利的独立的法律意见，详细信息请与 http://www.unisys.com/ 的 Unisys USA 联系。

---

**本项目的实现现状**（实测，不是文档转述）

- 原生窗口是 `STATIC` + `SS_BITMAP | SS_CENTERIMAGE`。**它不进 tab 序**：创建时不立
  `WS_TABSTOP`（账 #164，与 Label / Image / Frame 同一张排除表）。裸码实测：这一立上去，
  对话框管理器就会把静态窗口当成一站（`B1 → Static → B2`），摘掉就变回 `B1 → B2`。
  固定套件 `tests/tabwalk` 的 `TW-picstop=Y` 钉这一格（`TW-seq` 里也没了 `pic1` 这一站）。
- 排除表只管默认值：`.frm` 里**真写了** `TabStop = -1` 的情况仍照写的发（账 #83(a) 的口径）。
  剩下的后果是：它会占一站，而 `SetFocus` 落到静态窗口上本来也拿不到焦点。
- VB6 的 PictureBox 没有 `GotFocus` / `LostFocus` 事件（拿不到焦点），本批也不给它补这对事件。
- 它可以当容器（`Frame` / `PictureBox` / `SSTab` 三型之一）：创建时挂 `WS_EX_CONTROLPARENT`
  （= **0x00010000**）且两条创建路都挂，`tests/tabwalk` 的 `EX-pic=65536` 是运行期证人。
  挂上**对数**之后容器里的子控件已经能走进 tab 序（账 #165 = C29-FS-d：那个出口里手抄的常量
  以前是 `WS_EX_APPWINDOW` 0x00040000，所以有过两轮「样式在、行为没修」的读数）。
  走到的次序目前还是 z-order 而不是 `TabIndex`（账 #163）。
- 图形一侧：设计期 `Picture` 的装载与运行期 `Cls` / `Print` 的落点不在本批范围内。
