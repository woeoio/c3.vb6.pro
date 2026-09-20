# vi 的个人 todo

> 想法池（草稿区），约定见 [README.md](README.md)。决定动手的条目转仓库 Issue 后从这里删除。

## 想法

- 搞个关键字列表，能修改的，自定义中英文，原生命令名随便改

## 群友质询 Declare 64位加宽延伸的待办（2026-09-03）
1. VB6 算术/赋值溢出检查（Error 6）未实现：Long/Integer 运算当前为 32 位静默回绕，VB6 运行时对超范围结果抛错误 6，需评估 RTL 溢出检查或编译开关（ai/009 中列为 P3 可延后，考虑提前）
2. Declare 返回值/句柄赋给 Long 变量会截断：评估加编译警告提示改用 LongPtr 接收句柄/指针，并补充老代码迁移指引
3. Declare ByVal Long 加宽 intptr_t 边界测试：覆盖纯数值参数、负数符号扩展、ByRef Long 不受影响、LongPtr 显式声明等用例，固化"仅加宽不改变语义"的行为

## 群友质询延展：问题与建议（2026-09-03 二轮）
1. [问题] Declare As Long 返回值在 x64 的零扩展符号破坏：ABI 写 EAX 清空 RAX 高位，DLL 返回的 32 位负值（如 -1 错误码）被 intptr_t 全宽读成 0xFFFFFFFF（4294967295）。赋给 Long 变量靠 C 隐式截断还原无恙，但直接比较/位运算/作中间值（如 If Foo() = -1、Foo() And mask）时语义错误
   [建议] Declare 调用消费点按 VB6 语义补 (int32_t) 截断（比较/整值运算路径特判）；真正需要 64 位句柄时由源码显式声明 As LongPtr 获取，保持"As Long=32位有符号"忠实语义，不做返回值静默加宽
2. [问题] UDT 成员不随架构加宽导致 Declare 结构体参数错位：Fix 081e 只作用于直接写 As Long 的参数/返回值，UDT 内部 Long 恒为 4 字节；含句柄/指针成员的结构体（OPENFILENAME.hwndOwner 等）x64 下应为 8 字节，老源码 As Long 生成 4 字节 → 传给 API 该成员之后全部错位；且朴素 struct 无 pack，含 Double/Currency 成员的 UDT 在 x64 默认对齐下尺寸与 VB6 布局可能不同
   [建议] 无法自动判断成员语义，需迁移指引+警告：提示将句柄类成员显式改 As LongPtr（VBA7 官方 API 声明同款做法）；对常见 Win32 API 结构体提供修正声明清单；文档说明 UDT 对齐边界
3. [建议] 补 #If Win64 回归测试固化行为（特性已支持，Fix 081h：win64 = --arch 参数，常量查找小写不敏感，支持 #If/#ElseIf/#Else/#End If/#Const）：覆盖 --arch x64/x86 两套目标分支选择、#If Win64 Then LongPtr #Else Long 的 VBA7 双平台迁移代码、-D 覆盖内置常量

## 群友质询延展：问题与建议（2026-09-03 三轮）
1. [问题] 缺少编译器身份标识常量：内置条件编译常量仅 win16/win32/win64/vba6/vba7/mac/win 七个平台与语言版本属性，无 C3 专属标识；且 vba6 与 vba7 同时为 True（C3 双兼容定位），用户无法用 #If VBA7 Then 区分"当前是 C3 还是原生 VB6/VBA7"环境
   [建议] preprocessor.cpp 内置注册 c3=True（小写键，与其他内置一致）；命名用 C3（条件编译符号表与运行时变量名空间独立，无冲突），保留 -d:C3=False 可覆盖以便测试原生分支；文档给出四分支标准写法：#If C3 / #ElseIf VBA7 And Win64 / #ElseIf VBA7 / #Else，配套回归测试

## VB6 原生 GUI 兼容程度结论存档（2026-09-15）

0. [基线口径] 只谈 VB6 **内置**（工具箱 21 类 + Form/MDIForm + Menu），不含第三方 OCX。当前实测：
   - 控件：14 类全链路可用（Form/MDIForm、PictureBox、Label、TextBox、Frame、CommandButton、CheckBox、OptionButton、ComboBox、ListBox、HScrollBar、VScrollBar、Timer、Image、Menu）→ 21 类里约 67%
   - 属性：50+ 个双向读写（Left/Top/Width/Height/hWnd、Font 六件套、ForeColor/BackColor、Alignment、TabIndex/TabStop/CausesValidation、ToolTipText、Tag、MousePointer/MouseIcon、BorderStyle、Visible/Enabled + 各控件专有）
   - 事件：约 25 个（Click/DblClick/Change/Scroll/KeyDown/Press/Up/MouseDown/Up/Move/Enter/Leave/GotFocus/LostFocus/Validate(Cancel) + 窗体 Load/Unload(Cancel)/QueryUnload(Cancel)/Activate/Deactivate/Resize/鼠标键 + Timer + Menu.Click）
   - **方法：近乎空白**（仅 AddItem/RemoveItem/Clear + Timer 相关）→ 这是最短的一块板，欠账已转 [#2](https://gitcode.com/woeoio/c3.vb6.pro/issues/2)
1. [结论] 属性层最厚、事件层够日常、**方法层和绘图层是主要欠账**。常规 CRUD/工具型界面（文本框+按钮+列表+下拉+复选/单选+Frame+图片+滚动条+菜单+定时器+MDI）能编译并原生跑起来；依赖运行时操纵控件（Move/SetFocus/ZOrder/Refresh）、自绘画图（PSet/Line/Circle/Cls/PaintPicture）、拖放、多层容器嵌套的老程序跑不通，需改源码
2. [发现-关键，已转 #1] Shape/Line/DriveListBox/DirListBox/FileListBox 五类并非"没写"，而是**半接线**：RTL 已实现（vb6forms.c 里 Shape/Line 自绘 WndProc、文件系统控件的 Drive/Path/Pattern/FileName/Refresh 全有），属性读写表已登记、.frm 初值赋值也已生成，但 frm_parser.cpp 的 controlTypeToWin32Class 对这五类返回 nullptr → cgen_form.cpp 的创建循环不发 CreateWindow → vb6_hwnd_ 恒为 NULL → 编译零错但运行时看不见。**性价比最高，差一步** → 细节与验收见 [#1](https://gitcode.com/woeoio/c3.vb6.pro/issues/1)（已认领 woeoio）
3. [发现-次要] Toolbar/StatusBar/CommonDialog/ImageList 走 ActiveX CoCreateInstance 生成 IDispatch*，但无属性/方法映射，且这些 OCX 是 32 位 → x64 进程根本 CoCreate 失败
4. [发现-其他，已转 #4/#5] 容器只遍历 Frame 单层子控件（PictureBox 当容器、多层嵌套不支持）；缇↔像素硬编码 1 比 15（96 DPI），无 per-monitor DPI；缺 Form_Click、Paint、DragDrop/DragOver 事件（容器问题 → [#4](https://gitcode.com/woeoio/c3.vb6.pro/issues/4)；DPI 与事件 → [#5](https://gitcode.com/woeoio/c3.vb6.pro/issues/5)）
5. [结论-对外口径] GUI 是**Win32 原生重实现而非复刻 VB6 运行时**，"形似"可达成、"神似"（像素级渲染、字体度量、VB6 怪癖行为）需逐项对齐；这也是 VBMAN 运行期对齐的主战场

> **「控件 67% → 100%」主方向已全部转为正式 Issue，本文件只留指向不再重复记录**：
> [#1 P0 五类半接线（已认领 woeoio）](https://gitcode.com/woeoio/c3.vb6.pro/issues/1) ·
> [#2 P1 控件方法层](https://gitcode.com/woeoio/c3.vb6.pro/issues/2) ·
> [#3 P1 绘图语句与画布](https://gitcode.com/woeoio/c3.vb6.pro/issues/3) ·
> [#4 P2 容器嵌套与 Data/OLE 报错边界](https://gitcode.com/woeoio/c3.vb6.pro/issues/4) ·
> [#5 P3 Form_Click/Paint 事件与 DPI](https://gitcode.com/woeoio/c3.vb6.pro/issues/5)；路线图摘要见 README 第十章。

## 群友建议：字符串连接现代语法（Fan XiaoLei，2026-09-15）

0. [原话] 群友 Fan XiaoLei：vb6老是&连接太麻烦了。回应：这个可以加入现代语言特性，比如反引号包裹，我记录到 todo 去，兼容 &
1. [建议] 反引号字符串（现代语法糖，`&` 保持完全兼容不废弃）
   - 词法层新增 token：反引号 `` ` `` 在 VB6 词法中无既有用途，可安全收编，不与注释/行继续/字符串冲突
   - 最小档（先做）：多行原始字符串——免 `_` 续行、免双写引号（内嵌 `"` 原样保留），主要解决 SQL/JSON/长文本拼接痛点
   - 进阶档（后评估）：插值语法（如 `` `Hello ${name}` `` 自动展开表达式），与 `&` 混用时语义清晰（反引号串是一个普通字符串字面量）
   - 编码注意：反引号串内容按源文件编码解码（GBK/UTF-8 均支持），生成侧统一转 BSTR
2. [测试] 多行内容与换行符（CRLF/LF 归一）、内嵌双引号、内嵌反引号（转义方案：`` `` `` 双写）、与 `&` 混用、StrComp/Len 等函数消费、插值表达式含对象默认属性

## 群友建议：对外暴露接口供 AI/MCP/CLI 调用（漠漠，2026-09-19）

0. [原话] 群友 漠漠：可以留个接口让 AI 调 mcp 或者 cli 调用吗？把模型对象暴露出来，让 mcp 可以接入，这样就可以用很多免费模型了
1. [理解] 诉求拆两层，需与漠漠确认到底要哪一层
   - 层 A（编译器侧，C3 自身）：把编译器的"模型对象"（AST/符号表/类型与诊断信息）以结构化方式暴露，对外提供 MCP server 或稳定的 CLI 子命令（如 `c3 --json` 输出诊断、补全、跳转信息），让任意 AI 客户端（不限厂商/模型）都能驱动 C3
   - 层 B（使用侧）：C3 用户/作者想换更便宜或免费的模型时，只要对方支持 MCP 就能接入，不必绑定某一家
2. [价值] MCP 已是事实标准（Claude/CodeBuddy/Cursor 等均支持），做一层薄适配即可让所有支持 MCP 的客户端复用；CLI 侧补机器可读输出（JSON）成本最低、优先级最高
3. [待办-候选，未定优先级] 先做 CLI 机器可读输出（`--json` 诊断/编译结果，字段稳定化）；再评估 MCP server（tools：compile / analyze / 查询符号与诊断，resources：项目与错误清单）；同时明确暴露粒度与稳定性承诺，避免内部 AST 结构直接外泄成为长期包袱
4. [待确认] 目标模型/客户端是谁；是否需要反向能力（AI 改源码 → 调 C3 编译 → 回读错误自动修）；"模型对象"具体指编译器内部模型还是别的

## 支持 .lib / .obj 产物（2026-09-20）

0. [诉求] 让 C3 能把 VB6 工程/模块编译成 `.lib` 静态库，交给 C/C++（或其他语言）程序静态编译链接，复用既有 VB6 代码资产
1. [现状] 不支持：driver_link.cpp 产物只有 `.exe` / `.dll`（ActiveX DLL），`--emit-c`/`--emit-llvm` 只停在中产物，无"编 .obj + lib.exe 打包"路径
2. [方案-候选]
   - 新增 `--lib` 开关：各模块照常 cgen → 编成 .obj → 调 MSVC `lib.exe`（publish 自带迷你工具链）打包成 `<工程名>.lib`
   - 同级新增 `--emit-obj`：编到 .obj 即停（不链接不打包），作为独立最终产物保留在 --output-dir；是 .lib 的天然前置（.lib = .obj 集合 + lib.exe 打包），两者共用同一条"对象产出"路径，实现成本最低（现状 .obj 只是 driver_link.cpp /Fo 的中间产物，链接后即弃）
   - 导出边界：需要显式导出清单（哪些 Public 函数/Sub 进 .lib），避免把全模块符号裸漏；可用 `Public` + 约定命名（或 .def 式清单文件）圈定
   - 符号与调用约定：VB6 函数映射为 C 符号需定调用约定（stdcall/cdecl）与命名修饰规则，跨 x86/x64 分别打包
3. [难点-关键：RTL 依赖] VB6 代码生成的 C 依赖 vb6rtl（字符串/数组/COM/内存管理），静态库必须把用到的 RTL 目标文件一并链入 .lib 或要求消费方附加 c3rtl.lib——需设计"自包含 .lib"还是".lib + rtl.lib 两个交付物"
4. [难点-边界] 类型映射：BSTR/VARIANT/UDT 参数与返回值需同步生成 C 头文件（`<工程名>.h`），否则消费方无法声明签名；窗体/GUI 代码依赖消息循环与 WinMain，不适合进静态库，应限定"纯标准模块/类模块函数库"场景并给出明确报错或警告
5. [测试] x64/x86 双架构 .lib 由 cl.exe 的 C 程序链接调用回归；含 BSTR 参数往返、UDT 传参、错误（Err.Raise）如何跨库边界暴露（返回码还是 C++ 异常）待定

