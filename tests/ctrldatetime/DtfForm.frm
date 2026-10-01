VERSION 5.00
Begin VB.Form DtfForm 
   Caption         =   "DtfForm"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "DtfForm"
   ScaleHeight     =   3600
   ScaleWidth      =   6000
   Begin MSComCtl2.DTPicker dt1 
      CheckBox        =   -1  'True
      Format          =   1
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2400
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2400
   End
   Begin MSComCtl2.DTPicker dt3 
      Format          =   2
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      UpDown          =   -1  'True
      Width           =   2400
   End
   Begin MSComCtl2.DTPicker dt4 
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   3
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2400
   End
   Begin VB.Timer evtTimer 
      Interval        =   200
      Left            =   240
      Top             =   2160
   End
End
Attribute VB_Name = "DtfForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ai/029 C29-DT-a：VB6 DTPicker 走**原生** SysDateTimePick32（不加载 MSCOMCT2.OCX —— 它是
' 32 位 inproc，x64 进程里 CoCreateInstance 直接失败，见 029 §三 D6）。
'
' 本格只管"窗口 + 样式 + 标量属性面"，Date 型那三格（Value / MinDate / MaxDate）留 DT-b。
'
' 两条量出来的硬事实，判据的形状就是被它们决定的：
'  ① **CheckBox / UpDown 只能在创建时给**（它们是复选框/微调按钮那两枚子窗口的创建参数）。
'     事后 SetWindowLong 写 DTS_SHOWNONE / DTS_UPDOWN，控件立刻把这两位抹回去 —— DT11 就是
'     钉这条边界的针（写完读回 0、宽度不动）。谁做出真运行期切换，这条必须翻红。
'     Format 不一样：切档运行期**有效**（DT9/DT10 用控件自己算的宽度证它真的重算了）。
'     三条仍然一律从 .frm 立进创建样式，因为设计期就该在创建时就对。
'  ② 样式位的 getter 只证明"写进去了"，不证明"控件按这档在画" —— 因为 SDK 常数
'     DTS_TIMEFORMAT = 0x0009 自带 bit0（就是 DTS_UPDOWN 那位），只对自己的掩码读数会
'     自洽地假绿。所以 DT6-DT10 全部配**控件侧**读数：DTM_GETIDEALSIZE（问控件自己算的
'     "装得下当前格式"的宽度），长日期明显比时间宽，宽度对上了才叫格式真选中。
'
' Date 型那三格在 C29-DT-b（下面 25..36）。DTN_* 事件面（Change / DropDown / CloseUp）
' 由 evtTimer 那一截验，见文件末尾。
'
' 事件判据为什么全在 Timer 里、而且一律比**增量**：Form_Load 那半截自己就往控件写了好几回
' 值（DT26/DT30/DT32..DT34），原生要是因此自发 DTN_DATETIMECHANGE，计数在进 Timer 之前就已经
' 不是 0 了 —— 拿绝对值比会把"我们写的"和"Sim 发的"两件事混在一起（E= 那条原始读数就是留给
' 这一眼分辨的）。

Private gChg1 As Long
Private gDrp1 As Long
Private gClu1 As Long
Private gChg2 As Long

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Sub Form_Load()
    Dim bkg As Long, txt As Long, trl As Long, tbk As Long, ttx As Long
    Dim wLong As Long, wTime As Long, w0 As Long, w1 As Long
    Dim dReq As Date
    ' 五个色值刻意取"三字节都不同且都很小"的组合：系统默认色是 0xFFFFFF / 0 / 0x808080
    ' 那一族，任何一个恰好撞上都会让 DT20 那条"控件之间不共享"变成假绿或假红。
    bkg = 66051: txt = 263430: trl = 460809: tbk = 658188: ttx = 855567

    ' --- 1..5 创建样式进了窗口（三枚各一种设计期组合） ---
    Debug.Print "DT1=" & TF(dt1.Format = 1)
    Debug.Print "DT2=" & TF(dt1.CheckBox = -1)
    Debug.Print "DT3=" & TF(dt2.Format = 0 And dt2.CheckBox = 0 And dt2.UpDown = 0)
    Debug.Print "DT4=" & TF(dt3.UpDown = -1)
    Debug.Print "DT5=" & TF(dt3.Format = 2 And dt3.CheckBox = 0)

    ' --- 6..8 控件侧读数：长日期比时间宽，且三枚各自算各自的 ---
    wLong = dt1.IdealWidth
    wTime = dt3.IdealWidth
    Debug.Print "DT6=" & TF(wLong > 0 And wTime > 0)
    Debug.Print "DT7=" & TF(wLong > wTime)
    Debug.Print "DT8=" & TF(dt2.IdealWidth > 0 And dt2.IdealWidth <> wLong And dt2.IdealWidth <> wTime)

    ' --- 9..11 运行期切档 vs 运行期改 CheckBox：一个真生效、一个不落地（都量过） ---
    ' 原始读数（本机，W 那条）：dt1 长日期+复选框 = 143，dt3 时间 = 64，
    ' dt2 创建时短日期 = 95，运行期切成长日期 = 121。
    ' 121 落在"长日期但不带复选框"那一档 ⇒ Format 这一族**运行期切得动**（样式位 + 控件
    ' 自己算的宽度一起动），而下面 DT11 的 CheckBox 写完读回 0、宽度纹丝不动 ——
    ' 那两位是子窗口（复选框/微调按钮）的创建参数，控件只在 WM_CREATE 时读一次。
    w0 = dt2.IdealWidth
    dt2.Format = 1
    w1 = dt2.IdealWidth
    Debug.Print "DT9=" & TF(dt2.Format = 1 And w1 > w0)
    Debug.Print "DT10=" & TF(w1 > wTime And w1 < wLong)
    dt2.CheckBox = True
    Debug.Print "DT11=" & TF(dt2.CheckBox = 0 And dt2.IdealWidth = w1)
    ' 原始宽度打在针之外：长/时间/切档前/切档后，四条数字一起进日志，判据红了好对
    Debug.Print "W=" & wLong & "/" & wTime & "/" & w0 & "/" & w1

    ' --- 12..14 自定义格式（DTM_SETFORMATW 是真运行期消息；串自存，原生没有 Get 对称项） ---
    dt2.CustomFormat = "yyyy-MM-dd"
    Debug.Print "DT12=" & TF(dt2.CustomFormat = "yyyy-MM-dd")
    Debug.Print "DT13=" & TF(dt2.Format = 3)
    dt2.Format = 0
    Debug.Print "DT14=" & TF(dt2.CustomFormat = "" And dt2.Format = 0)

    ' --- 15..19 下拉月历五色：逐格写、逐格读回自己那格 ---
    dt2.CalendarBackColor = bkg
    dt2.CalendarForeColor = txt
    dt2.CalendarTrailingForeColor = trl
    dt2.CalendarTitleBackColor = tbk
    dt2.CalendarTitleForeColor = ttx
    Debug.Print "DT15=" & TF(dt2.CalendarBackColor = bkg)
    Debug.Print "DT16=" & TF(dt2.CalendarForeColor = txt)
    Debug.Print "DT17=" & TF(dt2.CalendarTrailingForeColor = trl)
    Debug.Print "DT18=" & TF(dt2.CalendarTitleBackColor = tbk)
    Debug.Print "DT19=" & TF(dt2.CalendarTitleForeColor = ttx)

    ' --- 20..21 序号与窗口都不串台：改"标题底色"那一格，另外四格原样；dt1 从没涂过色 ---
    dt2.CalendarTitleBackColor = 1234567
    Debug.Print "DT20=" & TF(dt2.CalendarBackColor = bkg And dt2.CalendarForeColor = txt _
                              And dt2.CalendarTrailingForeColor = trl _
                              And dt2.CalendarTitleForeColor = ttx)
    Debug.Print "DT21=" & TF(dt1.CalendarBackColor <> bkg And dt1.CalendarForeColor <> txt)

    ' --- 22..24 设计期 CustomFormat 那条字符串路（.frm 里带引号 + 当 C 字面量发，两头都得对） ---
    Debug.Print "DT22=" & TF(dt4.Format = 3)
    Debug.Print "DT23=" & TF(dt4.CustomFormat = "yyyy-MM-dd HH:mm")
    Debug.Print "DT24=" & TF(dt4.IdealWidth > w0)

    ' --- 25..36 C29-DT-b：Date 值面（Value / MinDate / MaxDate + "无日期"那一态）---
    ' 换算是 oleaut32 那一对现成函数（VariantTimeToSystemTime / SystemTimeToVariantTime）。
    ' VB 侧 Date 就是 double 序列号，所以 43894 这种整数序列直接当日期用（= 2020-03-04）。
    ' DT30 / DT31 两条按**实测真值**写（原先各猜错过一次，见 029 §九 本格）：
    '   · 设一个低于 MinDate 的日期，控件**直接拒绝、值保持原样**，不是钳到 MinDate；
    '   · DTS_SHOWNONE 那枚**创建时是勾上的**（HasDate=-1、Value=今天），不是未勾。
    Debug.Print "DT25=" & TF(Int(dt2.Value) = Int(Now))
    dReq = 43894
    dt2.Value = dReq
    Debug.Print "DT26=" & TF(CLng(dt2.Value) = 43894)
    dt2.MinDate = 43831
    Debug.Print "DT27=" & TF(CLng(dt2.MinDate) = 43831)
    dt2.MaxDate = 44999
    Debug.Print "DT28=" & TF(CLng(dt2.MaxDate) = 44999)
    ' 改一端必须不动另一端（原生是一张 (min,max) 表 + 有效位标志；只发 GDTR_MIN 会清掉 max）
    Debug.Print "DT29=" & TF(CLng(dt2.MinDate) = 43831)
    dt2.Value = 40000
    Debug.Print "DT30=" & TF(CLng(dt2.Value) = 43894)
    Debug.Print "DT31=" & TF(dt1.HasDate = -1 And Int(dt1.Value) = Int(Now))
    dt1.Value = 43894
    Debug.Print "DT32=" & TF(dt1.HasDate = -1 And CLng(dt1.Value) = 43894)
    dt1.HasDate = False
    ' 未勾这一态原生仍回填一个内部日期（本机 36494）⇒ GetValue 必须认返回标志才回 0
    Debug.Print "DT33=" & TF(dt1.HasDate = 0 And dt1.Value = 0)
    dt1.HasDate = True
    Debug.Print "DT34=" & TF(dt1.HasDate = -1 And CLng(dt1.Value) = 43894)
    Debug.Print "DT35=" & TF(dt3.MinDate = 0 And dt3.MaxDate = 0)
    Debug.Print "DT36=" & TF(CLng(dt3.Value) <> CLng(dt2.Value))

    ' DONE 与 Unload 在 evtTimer_Timer —— 事件判据得等窗体载入完再跑 (照 C29-8c 的先例)。
End Sub

' ---------------- C29-DT-c: DTN_* 事件面 ----------------
' Sim* 是**判据专用**助手 (与 C29-8c 的 TreeView.SimNodeClick / C29-4 的 StatusBar.SimClick
' 同一先例)，不对应任何 VB6 语义：无头环境点不了鼠标、也没有"用户拨了一下日期"这回事，
' 而直接调 handler 会绕开整条派发链 —— 只有从真 WM_NOTIFY 进父窗，才验得到
' "case WM_NOTIFY + 按 code 分流 + 按 hwndFrom 认来源"三段都接上了。
Private Sub evtTimer_Timer()
    Static done As Integer
    Dim b1 As Long, b2 As Long, b3 As Long, b4 As Long
    Dim n41 As Long
    If done Then Exit Sub
    done = 1

    b1 = gChg1: b2 = gDrp1: b3 = gClu1: b4 = gChg2
    ' 原始计数打在针之外，判据红了好对账
    Debug.Print "E=" & b1 & "/" & b2 & "/" & b3 & "/" & b4

    ' --- 37. Change：Sim 发一条真通知，handler 走一次 ---
    dt1.SimChange()
    Debug.Print "DT37=" & TF(gChg1 - b1 = 1)

    ' --- 38..39 分流：弹/收各自那一格加，且都不许顺手把 Change 也点一次 ---
    dt1.SimDropDown()
    dt1.SimCloseUp()
    Debug.Print "DT38=" & TF(gDrp1 - b2 = 1 And gClu1 - b3 = 1)
    Debug.Print "DT39=" & TF(gChg1 - b1 = 1)

    ' --- 40. 认来源：通知从 dt2 发出，dt1 的 handler 不该动 ---
    dt2.SimChange()
    Debug.Print "DT40=" & TF(gChg2 - b4 = 1 And gChg1 - b1 = 1)

    ' --- 41..42 运行期"程序化赋值"到底自不自发通知（实测，不是推的）：
    ' DT41 数的是 `dt2.Value = 43900` 之后 gChg2 又涨了几次。本机量到的答案是**一次都不涨**
    ' (n41 停在 SimChange 那一次的 1) —— 原生 SysDateTimePick32 对 DTM_SETSYSTEMTIME 不发
    ' DTN_DATETIMECHANGE，那条通知只由用户交互驱动。这里刻意照原生、不伪造：VB6 那颗 OCX 里
    ' Value 赋值会 raise Change，哪天要做那条齐平，就得在 vb6_DTP_SetValue 里补发一条，
    ' 届时 DT41 必须翻红逼人来改这条断言（它就是这条口径的哨兵）。
    ' DT42 钉另一半：CheckBox 那条写在 DT11 已证明是空转，所以它一下都不许加。
    dt2.Value = 43900
    n41 = gChg2 - b4
    Debug.Print "DT41=" & TF(n41 = 1)
    dt2.CheckBox = True
    Debug.Print "DT42=" & TF(gChg2 - b4 = n41)

    Debug.Print "E2=" & n41
        Debug.Print "DT43=" & CStr(dt1.CheckBox) & "/" & TypeName(dt1.CheckBox) & "/" & VarType(dt1.CheckBox)
Debug.Print "CTRLDATETIME-DONE"
    Unload Me
End Sub

' VB6 这三条处理器都没有参数，所以生成的回调形参表是空的 (与 ListView 那两条 void* 不同)。
Private Sub dt1_Change()
    gChg1 = gChg1 + 1
End Sub

Private Sub dt1_DropDown()
    gDrp1 = gDrp1 + 1
End Sub

Private Sub dt1_CloseUp()
    gClu1 = gClu1 + 1
End Sub

Private Sub dt2_Change()
    gChg2 = gChg2 + 1
End Sub
