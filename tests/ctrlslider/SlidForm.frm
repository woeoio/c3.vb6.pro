VERSION 5.00
Begin VB.Form SlidForm 
   Caption         =   "SlidForm"
   ClientHeight    =   3200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   LinkTopic       =   "SlidForm"
   ScaleHeight     =   3200
   ScaleWidth      =   4200
   Begin MSComctlLib.Slider sld8 
      Height          =   400
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   1600
      _ExtentX        =   2822
      _ExtentY        =   706
   End
   Begin VB.TextBox txtF 
      Height          =   400
      Left            =   2560
      TabIndex        =   8
      Text            =   "F"
      Top             =   2400
      Width           =   1400
   End
   Begin VB.TextBox txtG 
      Height          =   400
      Left            =   2560
      TabIndex        =   9
      Text            =   "G"
      Top             =   2880
      Width           =   1400
   End
   Begin VB.TextBox txtH 
      FontName        =   "Arial"
      FontSize        =   20
      Height          =   300
      Left            =   3600
      TabIndex        =   10
      Text            =   "H"
      Top             =   120
      Width           =   500
   End
   Begin VB.Timer tGo 
      Enabled         =   -1   'True
      Interval        =   60
      Left            =   3720
      Top             =   2040
   End
   Begin MSComctlLib.Slider sld1 
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Tag             =   ""
      TickFrequency   =   10
      Top             =   120
      Width           =   2000
      _ExtentX        =   3528
      _ExtentY        =   706
   End
   Begin MSComctlLib.Slider sld2 
      Height          =   1200
      Left            =   2400
      Orientation     =   1
      TabIndex        =   1
      Top             =   120
      Width           =   400
      _ExtentX        =   706
      _ExtentY        =   2117
   End
   Begin MSComctlLib.Slider sld3 
      Height          =   400
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2000
      _ExtentX        =   3528
      _ExtentY        =   706
   End
   Begin MSComctlLib.Slider sld4 
      Height          =   400
      LargeChange     =   8
      Left            =   120
      Max             =   100
      Min             =   10
      SelStart        =   20
      SelEnd          =   60
      SelectRange     =   -1   'True
      SmallChange     =   2
      TabIndex        =   3
      TickFrequency   =   5
      Top             =   1200
      Value           =   42
      Width           =   2000
      _ExtentX        =   3528
      _ExtentY        =   706
   End
   Begin MSComctlLib.Slider sld5 
      Height          =   400
      Left            =   120
      TabIndex        =   4
      Tag             =   "dtag"
      TickStyle       =   2
      Top             =   1680
      ToolTipText     =   "dtip"
      Width           =   2000
      _ExtentX        =   3528
      _ExtentY        =   706
   End
   Begin MSComctlLib.Slider sld6 
      Height          =   400
      Left            =   2280
      SelLength       =   15
      SelStart        =   25
      SelectRange     =   -1   'True
      SmallChange     =   1
      TabIndex        =   5
      Top             =   1680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   706
   End
   Begin VB.Frame frmTickBox 
      Caption         =   "tickbox"
      Height          =   1000
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2400
      Begin MSComctlLib.Slider sld7 
         Height          =   400
         LargeChange     =   8
         Left            =   120
         Max             =   100
         Min             =   10
         SelLength       =   15
         SelStart        =   20
         SelectRange     =   -1   'True
         SmallChange     =   2
         TabIndex        =   7
         TickFrequency   =   5
         TickStyle       =   1
         Top             =   240
         Value           =   42
         Width           =   2000
         _ExtentX        =   3528
         _ExtentY        =   706
      End
      Begin VB.TextBox txtI 
         FontSize        =   20
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Text            =   "I"
         Top             =   720
         Width           =   900
      End
   End
End
Attribute VB_Name = "SlidForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ai/029 C29-SL-a: Slider 的窗口 + 创建样式 + 标量属性面（原生 msctls_trackbar32）。
' 登记之前这枚控件**连窗口都没建**（探针 .build/slprobe 实测：CreateControls 里压根没有它，
' vb6_hwnd_sld1 恒 NULL，属性读回全空、写进去静默丢，退出码照旧 0）。
'
' 判据纪律（本格两条实测口径，见 029 §九 C29-SL-0）：
'   1) 方向**不能拿 channel 矩形当证人** —— 实测 TBM_GETCHANNELRECT 的 rect 永远把行程长度
'      放在 x 分量，水平杆与竖直杆答同一组数（第一发探针因此误判过"运行期改不动"）。
'      正解是 TravelIsVert：把滑块推到量程两端各读一次 TBM_GETTHUMBRECT，看位移落在哪根轴。
'      所以每条 Orientation 断言都配一条 TravelIsVert（SL2+SL4、SL6、SL9 就是这几对）。
'   2) TickFrequency 原生**问不出**（没有 GETTICFREQ；GETTIC 答不出、GETTICPOS 与频率无关），
'      所以它只能自存读回；"控件当前有没有画刻度"另用 TickPresent 证（GETTICPOS(0) != -1）。
'   3) 值面（Min/Max/Value/Small·LargeChange/Sel*）刻意**不在本批** —— 归 SL-b，
'      所以这件夹具一次都不碰那些名字（它们还没登记，碰了就落到"拿 HWND 当 IDispatch"那条兜底）。

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

' C29-SL-c: 三条计数针（两条 Change + 一条 Scroll）。事件断言一律**只数增量**，
' 不拿绝对值比 —— 原生自发的那几条什么时候到、到几条，不由我们定。
Private gChg1 As Long
Private gScr1 As Long
Private gChg2 As Long

Private Sub tGo_Timer()
    tGo.Enabled = False
    Debug.Print "SL1-vis=" & CStr(sld1.Visible)
    Debug.Print "SL2-ori=" & CStr(sld1.Orientation)
    Debug.Print "SL3-freq=" & CStr(sld1.TickFrequency)
    Debug.Print "SL4-travel=" & CStr(sld1.TravelIsVert)
    Debug.Print "SL5-tick=" & CStr(sld1.TickPresent)
    Debug.Print "SL6-vertori=" & CStr(sld2.Orientation) & CStr(sld2.TravelIsVert)
    Debug.Print "SL7-nofreq=" & CStr(sld3.TickFrequency) & CStr(sld3.TickPresent)
    sld3.TickFrequency = 5
    Debug.Print "SL8-freqset=" & CStr(sld3.TickFrequency)
    ' 运行期翻方向: 实测**有效**（写 TBS_VERT + 换帧之后滑块位移轴跟着换）
    sld1.Orientation = 1
    Debug.Print "SL9-flip=" & CStr(sld1.Orientation) & CStr(sld1.TravelIsVert)
    sld1.Orientation = 0
    Debug.Print "SL10-flipback=" & CStr(sld1.Orientation) & CStr(sld1.TravelIsVert)
    sld1.Enabled = False
    Debug.Print "SL11-en=" & CStr(sld1.Enabled) & CStr(sld3.Enabled)
    ' ---- C29-SL-b 值面：sld4 设计期写满、sld5 什么都没写（默认档照原生答，不猜 VB6 文档）----
    Debug.Print "SB1-range=" & CStr(sld4.Min) & "/" & CStr(sld4.Max)
    Debug.Print "SB2-value=" & CStr(sld4.Value)
    ' SB3 是 **Init 参数顺序的证人**：先立 range(10..100) 再立 page=8。顺序反过来就会被
    ' range 那次重算把 page 顶成 18（探针实测），这条读数就会变成 2/18。
    Debug.Print "SB3-changes=" & CStr(sld4.SmallChange) & "/" & CStr(sld4.LargeChange)
    Debug.Print "SB4-sel=" & CStr(sld4.SelStart) & "/" & CStr(sld4.SelEnd) _
        & "/" & CStr(sld4.SelectRange)
    sld4.Value = 500
    Debug.Print "SB5-clampmax=" & CStr(sld4.Value)
    sld4.Value = 5
    Debug.Print "SB6-clampmin=" & CStr(sld4.Value)
    sld4.Max = 40
    Debug.Print "SB7-shrink=" & CStr(sld4.Min) & "/" & CStr(sld4.Max) & "/" & CStr(sld4.Value)
    sld4.SelectRange = False
    Debug.Print "SB8-selrange-off=" & CStr(sld4.SelectRange)
    Debug.Print "SB9-defaults=" & CStr(sld5.Min) & "/" & CStr(sld5.Max) & "/" _
        & CStr(sld5.SmallChange) & "/" & CStr(sld5.LargeChange) & "/" & CStr(sld5.Value) & "/" _
        & CStr(sld5.TickFrequency)
    ' 起点越过终点时两端互相顶（原生不接受反向区段）
    sld5.SelectRange = True
    sld5.SelStart = 30
    sld5.SelEnd = 70
    sld5.SelStart = 80
    Debug.Print "SB10-push=" & CStr(sld5.SelStart) & "/" & CStr(sld5.SelEnd)
    ' ---- C29-SL-c: 事件面（Change / Scroll 走 WM_HSCROLL / WM_VSCROLL 那一条通道）----
    ' SimNotify 是**判据专用**助手（与 DT-c 的 SimChange / MV-c 的 SimDateClick / RT-d 的
    ' SimNotify 同先例）：无头环境点不了鼠标，而直接调 handler 会绕开整条派发链 ——
    ' 它先把值推到 pos（真手势都是控件先动、再发通知），再按原生那一档发一条**真**消息进父窗。
    ' 码值/通道/wParam 布局全部实测（探针 .build/slprobe/slmeasure7.c 与 8.c）：
    '   横杆发 WM_HSCROLL、竖杆发 WM_VSCROLL；LOWORD = TB_* 码，5/4 那两档的值在 HIWORD；
    '   一次真拖 = 5×N → 4 → 8；方向键 = 0 → 8；程序化 SETPOS/SETRANGE 一条都不发。
    Dim b1 As Long, b2 As Long, b3 As Long
    sld1.Value = 20                 ' 先把基准摆明（程序化赋值不发通知，SC10 钉这条）
    b1 = gChg1: b2 = gScr1: b3 = gChg2
    Debug.Print "SC0=" & b1 & "/" & b2 & "/" & b3
    sld1.SimNotify(5, 40)           ' TB_THUMBTRACK：值 20→40 ⇒ Change 一次；滑块那两档 ⇒ Scroll 一次
    Debug.Print "SC1-track=" & TF(gChg1 - b1 = 1 And gScr1 - b2 = 1)
    sld1.SimNotify(5, 40)           ' 同值再来一条（拖动不足一像素的那种）：Scroll 涨、Change 不涨
    Debug.Print "SC2-same=" & TF(gChg1 - b1 = 1 And gScr1 - b2 = 2)
    sld1.SimNotify(5, 55)           ' 拖拽中连续触发就是这一串 5 ⇒ Change 跟着涨到 2
    Debug.Print "SC3-next=" & TF(gChg1 - b1 = 2 And gScr1 - b2 = 3)
    sld1.SimNotify(4, 55)           ' TB_THUMBPOSITION（落点与当前同值）⇒ 只算 Scroll
    Debug.Print "SC4-posit=" & TF(gChg1 - b1 = 2 And gScr1 - b2 = 4)
    sld1.SimNotify(8, 55)           ' TB_ENDTRACK 是"收尾"那条，不带新值 ⇒ 两条都不许点
    Debug.Print "SC5-endtrk=" & TF(gChg1 - b1 = 2 And gScr1 - b2 = 4)
    sld1.SimNotify(0, 54)           ' TB_LINEUP（方向键那一档）：值动了 ⇒ Change；不是滑块档 ⇒ Scroll 不涨
    Debug.Print "SC6-line=" & TF(gChg1 - b1 = 3 And gScr1 - b2 = 4)
    Debug.Print "SC7-value=" & CStr(sld1.Value)
    sld2.SimNotify(5, 30)           ' 竖杆：同一条 case、发的是 WM_VSCROLL（sld2 只挂 Change）
    Debug.Print "SC8-vert=" & TF(gChg2 - b3 = 1)
    Debug.Print "SC9-isolate=" & TF(gChg1 - b1 = 3 And gScr1 - b2 = 4)
    sld1.Value = 66                 ' 程序化赋值**不发**通知（实测；DT41 同型的哨兵，哪天齐平就得翻红）
    Debug.Print "SC10-setval=" & TF(gChg1 - b1 = 3 And gScr1 - b2 = 4)
    sld1.SimNotify(5, 66)           ' SetValue 已把基准推到 66 ⇒ 这条同值通知不该被当成"变了"
    Debug.Print "SC11-baseline=" & TF(gChg1 - b1 = 3 And gScr1 - b2 = 5)
    ' ---- C29-SL-d: 常规事件面（Click / DblClick / KeyDown / KeyUp）----
    ' SimStdEvent 也是**判据专用**助手（与上面 SimNotify 同先例）：把一条常规事件的原生
    ' 消息**同步**送进控件自己的过程（SendMessage；实测与 PostMessage 的读数逐字相同，
    ' 只是同步那条不必在两次调用之间夹 DoEvents）。探针 .build/slprobe/slmeasure10.c
    ' （真手势）+ 11.c（合成，P* 与 S* 两段）：
    '   · 控件的 WM_LBUTTONUP / WM_LBUTTONDBLCLK 都经过它的子类过程 ⇒ Click / DblClick 有落点；
    '   · 裸的一条 UP / DBLCLK **不**惊动父窗那条通道（值不动 ⇒ Change、Scroll 都不跟）；
    '   · 按键那条**控件真动**（VK_LEFT 走 SmallChange 一档）并发父窗 code=0 ⇒ Change 跟一次；
    '     紧跟的 KEYUP 发 code=8、值不再动 ⇒ KeyUp 不该顺带点着 Change；
    '   · 真按下时控件自己就把焦点抢过去了（10.c：DOWN 之后跟着 SETFOCUS、"focus after click=1"）
    '     ⇒ 键那两条不靠 WS_TABSTOP 也到得了（Tab 导航本身在账 #83）。
    Dim c0 As Long, d0 As Long, k0 As Long, u0 As Long, x0 As Long
    sld6.Value = 50
    c0 = gClk6: d0 = gDbl6: k0 = gKeyD6: u0 = gKeyUp6: x0 = gChg6
    Debug.Print "SD0=" & c0 & "/" & d0 & "/" & k0 & "/" & u0 & "/" & x0
    sld6.SimStdEvent(0, 0)          ' 抬起：只该点 Click
    Debug.Print "SD1-click=" & TF(gClk6 - c0 = 1 And gDbl6 - d0 = 0 And gChg6 - x0 = 0)
    sld6.SimStdEvent(1, 0)          ' 双击：只该点 DblClick，Click 不再涨（两档不互带）
    Debug.Print "SD2-dbl=" & TF(gDbl6 - d0 = 1 And gClk6 - c0 = 1)
    gKeyC6 = -1
    sld6.SimStdEvent(2, 37)         ' VK_LEFT = vbKeyLeft = 37：KeyDown + 控件真动 + Change 跟一次
    Debug.Print "SD3-keydown=" & TF(gKeyD6 - k0 = 1 And gKeyC6 = 37 And gChg6 - x0 = 1)
    Debug.Print "SD4-value=" & CStr(sld6.Value)
    sld6.SimStdEvent(3, 37)         ' 抬键：只点 KeyUp，值不再动
    Debug.Print "SD5-keyup=" & TF(gKeyUp6 - u0 = 1 And gChg6 - x0 = 1 And sld6.Value = 49)
    sld3.SimStdEvent(0, 0)          ' 认来源：没挂 handler 的那枚被"点"，sld6 的计数一条都不许动
    Debug.Print "SD6-isolate=" & TF(gClk6 - c0 = 1 And gKeyUp6 - u0 = 1)
    ' SD7 是**装不装那一趟**的证人：sld2 只有 Click 与 Change 两个处理器，而 Change 走的是
    ' 父窗那条 WM_VSCROLL（与子类化无关），所以在改之前它压根不会被子类化 —— 处理器编得出来、
    ' 没人给它送消息（实测：BASE 的产物里既没有 sld2 的子类过程、也没有 install 那一行）。
    Dim y0 As Long
    y0 = gClk2
    sld2.SimStdEvent(0, 0)
    Debug.Print "SD7-install=" & TF(gClk2 - y0 = 1)
    ' ---- C29-SL-e: 通用字符串属性面（ToolTipText / Tag）----
    ' 这两条以前压根不在类型表里，而 RTL getter 又声明成 `void*` ⇒ 装箱那一步（C11 _Generic
    ' 的表把"其他指针"送去 VariantObject）把字符串当**对象**装。实测（探针 .build/sltt，
    ' 改之前）：CStr 打空、TypeName 答 "Object"、`If Slider1.Tag = "tag1"` 答假，
    ' 而同一枚属性的 Len / InStr 那些**直接拿指针**的面却是对的 —— 两种答案就是没登记的证状。
    ' 登记成 String（+ getter 的 C 返回型改 wchar_t*）之后消费面统一。SE9 是"没写过就是空串"
    ' 那一条默认档，SE10 是 .frm 里设计期那条到没到窗口（VB6 会到；没到就是设计期那一趟的缺口）。
    Dim sTip As String
    sld3.ToolTipText = "abc"
    sld3.Tag = "t9"
    Debug.Print "SE1-tip=" & CStr(sld3.ToolTipText)
    Debug.Print "SE2-tag=" & CStr(sld3.Tag)
    Debug.Print "SE3-tn=" & TypeName(sld3.ToolTipText)
    Debug.Print "SE4-vt=" & CStr(VarType(sld3.Tag))
    Debug.Print "SE5-len=" & CStr(Len(sld3.ToolTipText))
    Debug.Print "SE6-eqtag=" & TF(sld3.Tag = "t9")
    Debug.Print "SE7-eqtip=" & TF(sld3.ToolTipText = "abc")
    sTip = sld3.ToolTipText
    Debug.Print "SE8-assign=" & CStr(Len(sTip)) & "/" & sTip
    Debug.Print "SE9-unset=" & CStr(sld1.ToolTipText) & "/" & CStr(sld1.Tag)
    Debug.Print "SE10-dt=" & CStr(sld5.ToolTipText)
    Debug.Print "SE12-dttag=" & CStr(sld5.Tag)
    sld5.ToolTipText = "rt"
    sld5.Tag = ""
    Debug.Print "SE13-over=" & CStr(sld5.ToolTipText) & "/" & CStr(sld5.Tag)
    sld5.ToolTipText = "dtip"
    sld5.Tag = "dtag"
    Debug.Print "SE14-emptytag=" & CStr(sld1.Tag) & "/end"
    Debug.Print "SE11-cat=" & "v=" & sld3.Tag
    ' ---- C29-SL-g: 焦点事件面（GotFocus / LostFocus）----
    ' SimStdEvent 的 kind=4 **不是发消息**，是真 SetFocus —— 焦点那两条的原生来源就是窗口
    ' 管理器自己发的 WM_SETFOCUS / WM_KILLFOCUS（探针 slmeasure9.c 的 Q5a、10.c 的 R3 都量到
    ' 它们会到被子类化的轨道条上；伪造那两条消息反而验不到真链路）。
    ' SG2/SG4 是"一次移动两边各发一次"与认来源；改之前这三条处理器全是死的（四类白名单不含 Slider）。
    Dim f0 As Long, f1 As Long, f2 As Long
    f0 = gGot6: f1 = gLost6: f2 = gGot3
    Debug.Print "SG0=" & f0 & "/" & f1 & "/" & f2
    sld6.SimStdEvent(4, 0)          ' 焦点给 sld6 ⇒ 只该点它的 GotFocus
    Debug.Print "SG1-got=" & TF(gGot6 - f0 = 1 And gLost6 - f1 = 0 And gGot3 - f2 = 0)
    sld3.SimStdEvent(4, 0)          ' 移走 ⇒ sld6 失焦、sld3 得焦，两条都该发
    Debug.Print "SG2-move=" & TF(gLost6 - f1 = 1 And gGot3 - f2 = 1 And gGot6 - f0 = 1)
    sld6.SimStdEvent(4, 0)          ' 再回来
    Debug.Print "SG3-back=" & TF(gGot6 - f0 = 2 And gGot3 - f2 = 1)
    sld3.SimStdEvent(4, 0)          ' 认来源：sld6 的 LostFocus 只跟着它自己失焦涨
    Debug.Print "SG4-isolate=" & TF(gLost6 - f1 = 2 And gGot3 - f2 = 2 And gGot6 - f0 = 2)
    ' ---- C29-SL-h: TickStyle 四档（数值读自 OCX 自带的类型库，不是猜的）----
    ' 0=sldBottomRight(那三位样式位全清) 1=sldTopLeft(TBS_TOP) 2=sldBoth(TBS_BOTH)
    ' 3=sldNoTicks(TBS_NOTICKS)。判据一律用**相对高低**，不钉绝对像素 —— chan.top 的绝对值
    ' 随主题与控件高度变（探针 slmeasure12.c 在 40px 高：chan.top 10/20/19/10、thumb.top
    ' 2/10/10/2、numTics 11/11/11/0；夹具这一枚是 400 缇 = 26~27px，见 SH0 那行原始读数）。
    ' **两档之间不可分的那一对**：1 与 2 的几何只差一两个像素，而且**谁高谁低本机都不稳**
    ' （探针 40px 高：chan.top 20 / 19；夹具这一枚 400 缇：18 / 19 —— 顺序正好相反，第一版把
    ' "1 > 2 > 0" 写进判据，SH3 就是这么翻红的）⇒ 1/2 两档之间没有稳的几何维度，只能靠样式位
    ' 读回区分；控件侧稳的维度是「与 0 那档不同」+「3 那档 numTics=0」。
    ' ts=3 与 ts=0 在几何上同形 ⇒ 它唯一的证人是 GetNumTicks=0；而 TickPresent 在那一档
    ' **照旧答「有刻度」**（刻度只是不画、那张表还在，实测 GETTICPOS(0) 仍是 14），别拿它当判据。
    Dim ct0 As Long, ct1 As Long, ct2 As Long, ct3 As Long
    ct0 = sld3.ChannelTop
    Debug.Print "SH1-base=" & CStr(sld3.TickStyle) & "/" & TF(sld3.GetNumTicks > 0)
    sld3.TickStyle = 1
    ct1 = sld3.ChannelTop
    Debug.Print "SH2-top=" & CStr(sld3.TickStyle) & "/" & TF(ct1 > ct0)
    sld3.TickStyle = 2
    ct2 = sld3.ChannelTop
    Debug.Print "SH3-both=" & CStr(sld3.TickStyle) & "/" & TF(ct2 > ct0)
    sld3.TickStyle = 3
    ct3 = sld3.ChannelTop
    Debug.Print "SH4-none=" & CStr(sld3.TickStyle) & "/" & CStr(sld3.GetNumTicks) & "/" & TF(ct3 = ct0)
    sld3.TickStyle = 0
    Debug.Print "SH5-back=" & CStr(sld3.TickStyle) & "/" & TF(sld3.ChannelTop = ct0) & "/" & TF(sld3.GetNumTicks > 0)
    ' SH0 是**原始几何证人行**（四档的 chan.top 绝对值），刻意不登记成针 —— 登记等于把本机
    ' 像素钉进判据。它存在的意义是：哪天两根断言同时变 N，看这行就知道是几何没了还是通道换了。
    Debug.Print "SH0-geom=" & ct0 & "/" & ct1 & "/" & ct2 & "/" & ct3
    ' SH6 是**设计期那一条到没到窗口**的证人：sld5 在 .frm 里写的是 TickStyle = 2，而它与 sld3
    ' 同尺寸（2000x400），所以可以直接比 chan.top 高低 —— 只读回一个自存的数不算数。
    Debug.Print "SH6-dt=" & CStr(sld5.TickStyle) & "/" & TF(sld5.ChannelTop > ct0)
    ' SH7 是越界那一档：写侧只认 1/2/3，其余落 0 ⇒ 答出去的数就是窗口真在走的那一档
    ' （与 Orientation 同一口径，不自存、不猜 VB6 会不会报错 —— OCX 跑不起来，那条真值本机拿不到）。
    sld3.TickStyle = 9
    Debug.Print "SH7-oob=" & CStr(sld3.TickStyle)
    Debug.Print "SH8-type=" & TypeName(sld3.TickStyle) & "/" & CStr(VarType(sld3.TickStyle))
    ' SH9 是**第二条创建路**的证人：sld7 挂在 Frame 里，而那条路以前对 Slider 一格样式都不挂
    ' （连默认的 TBS_AUTOTICKS 都没有，设计期写的 TickStyle 更没人下发 —— 账 #83 的一个具体落点）。
    ' 它与 sld3 同尺寸（2000x400），所以直接拿 chan.top 与 ct0（sld3 在 ts=0 那档的基线）比高低：
    ' 读回 1 只证到样式位，`chan.top > ct0` 才证到控件真按那一档重排了。
    Debug.Print "SH9-child=" & CStr(sld7.TickStyle) & "/" & TF(sld7.ChannelTop > ct0)
    ' ---- C29-SL-i: SelLength / ClearSel（类型库读出来的 VB6 那一面）----
    ' 类型库给的是 SelStart(0x0007) + **SelLength**(0x0008) 这一对，另有一条方法 ClearSel(0x000e)。
    ' VB6 那一面**没有 SelEnd 这个名字**（上面 SB 段用的那条是 SL-b 按原生 TBM_SETSELEND 自己加的）。
    ' 三条实测口径（探针 slmeasure13.c，量程 10..100）写进了 RTL 的注释，这里各钉一条：
    '   · **没碰过的控件 GETSELSTART 答的是量程下限、GETSELEND 答 0** ⇒ 终点比起点小，
    '     拿减法会得出负长度 ⇒ SI1 钉「空选区就是 0」；
    '   · 远端超量程由**控件夹住**（SI3 写 900 只到上限 100），我们不再钳第二遍；
    '   · CLEARSEL 之后两端都答 -1 ⇒ 读数折成 0，而**紧接着写 SelLength 不该静默什么都不发生**
    '     （SI4b 就是钉这一条：空态下锚量程下限，与控件自己那一态一致）。
    ' sld4 在上面 SB 段被改过 Max 与 SelectRange，这里先摆回一个明确的起点（判据自带前提）。
    sld4.SelectRange = True
    sld4.Min = 10
    sld4.Max = 100
    sld4.SelStart = 20
    sld4.SelEnd = 60
    Debug.Print "SI1-len=" & CStr(sld4.SelLength)
    sld4.SelLength = 10
    Debug.Print "SI2-set=" & CStr(sld4.SelStart) & "/" & CStr(sld4.SelEnd) & "/" & CStr(sld4.SelLength)
    sld4.SelLength = 900
    Debug.Print "SI3-clamp=" & CStr(sld4.SelStart) & "/" & CStr(sld4.SelEnd) & "/" & CStr(sld4.SelLength)
    sld4.ClearSel()
    Debug.Print "SI4-clear=" & CStr(sld4.SelLength) & "/" & CStr(sld4.SelStart) & "/" & CStr(sld4.SelEnd)
    sld4.SelLength = 5
    Debug.Print "SI4b-reafter=" & CStr(sld4.SelStart) & "/" & CStr(sld4.SelLength)
    ' SI5/SI6 是「不伪造」那一条：sld3 没挂 SelectRange，原生整条 SETSEL 不生效（实测），
    ' 所以写进去读回来还是 0 —— 与 SelStart/SelEnd 在 SL-b 里同一口径。
    Debug.Print "SI5-nobit=" & CStr(sld3.SelLength) & "/" & CStr(sld3.SelectRange)
    sld3.SelLength = 7
    Debug.Print "SI6-nobit2=" & CStr(sld3.SelLength)
    sld4.SelLength = -3
    Debug.Print "SI7-neg=" & CStr(sld4.SelLength)
    Debug.Print "SI8-type=" & TypeName(sld4.SelLength) & "/" & CStr(VarType(sld4.SelLength))
    ' SI9 是**设计期那一条到没到窗口**的证人：sld6 的 .frm 写的是一对 VB6 名字（SelStart 25 +
    ' SelLength 15），而原生只有 (起, 止) 那一条消息 ⇒ cgen 在设计期那一趟折成 (25, 40)。
    ' 改之前 SelLength 整条被丢（创建参数里只有 -999），读回来是空区段 —— 与账 #142 同一形状。
    Debug.Print "SI9-dt=" & CStr(sld6.SelStart) & "/" & CStr(sld6.SelEnd) & "/" & CStr(sld6.SelLength)
    ' SJ1 是**账 #148 的疑点**：ToolTipText 那条一直只验了"存回来的串"，没验过"到底注册进
    ' tooltip 宿主没有"。裸编探针（无 v6 manifest ⇒ 走 v5）里 TTM_ADDTOOLW 直接返回失败，
    ' 但那不能定罪产品 —— 这里改问宿主 TTM_GETTEXT，三枚一起读：运行期设过的 sld3、设计期
    ' 设过的 sld5 都该 Y，从没设过的 sld1 该 N（要是三条齐 Y 就说明这枚证人问不出东西）。
    Debug.Print "SJ1-reg=" & TF(sld3.ToolTipRegistered) & "/" & TF(sld5.ToolTipRegistered) _
        & "/" & TF(sld1.ToolTipRegistered)
    ' ---- C29-SL-k: Text 气泡 + TextPosition（VB6 那一面：ISlider.Text 0x0010 BSTR /
    '      TextPosition 0x0011 = sldAboveLeft 0 / sldBelowRight 1）----
    ' 气泡不走轨道条自带的 TBS_TOOLTIPS（那颗服务不了自定义串，而且只在创建时才建），
    ' 走 RTL 自己持的一枚 TRACK 型 tooltip ⇒ 三条判据问的都是**宿主**而不是我们的窗口属性：
    ' BubbleVisible（IsWindowVisible）、BubbleText（TTM_GETTEXT 读回）、BubbleTop（气泡上边）。
    ' 摆动它的是派发里那条 BubbleNotify：4/5 摆出来、8 收回去（与 Scroll 的分法同一档）。
    Dim bt0 As Long, bt1 As Long
    sld3.Text = "VOL"
    Debug.Print "SK1-text=" & sld3.Text & "/" & TypeName(sld3.Text) & "/" & CStr(VarType(sld3.Text))
    Debug.Print "SK2-off=" & TF(sld3.BubbleVisible) & "/" & TF(sld1.BubbleVisible)
    sld3.SimNotify(5, 40)
    bt0 = sld3.BubbleTop
    Debug.Print "SK3-show=" & TF(sld3.BubbleVisible) & "/" & sld3.BubbleText & "/" & TF(bt0 > 0)
    sld3.SimNotify(8, 40)
    Debug.Print "SK4-hide=" & TF(sld3.BubbleVisible)
    sld3.TextPosition = 1
    sld3.SimNotify(5, 40)
    bt1 = sld3.BubbleTop
    Debug.Print "SK5-pos=" & CStr(sld3.TextPosition) & "/" & TF(bt1 > bt0)
    sld3.SimNotify(8, 40)
    sld3.TextPosition = 0
    ' SK6：Text 空着 ⇒ 气泡里就是当前的值（原生那颗画的就是数字，这条是我们的口径，
    '      VB6 两者怎么共存本机拿不到真值）。sld1 从没写过 Text，拖到 33。
    sld1.SimNotify(5, 33)
    Debug.Print "SK6-num=" & sld1.BubbleText & "/" & CStr(sld1.Value) _
        & "/" & TF(sld1.BubbleVisible)
    sld1.SimNotify(8, 33)
    Debug.Print "SK7-end=" & TF(sld1.BubbleVisible) & "/" & CStr(sld1.TextPosition)
    ' ---- C29-SL-l: 控件的**零实参方法**两形（账 #143）----
    ' 之前 `Slider1.SetFocus` 这类写法（连不带括号的那一形）从没登记过，两形都落进
    ' vb6_ComCall(裸 HWND, L"SetFocus", NULL, 0) —— 原生控件槽里是句柄不是 IDispatch，
    ' 于是编得过、链接得过、跑起来一声不响。现在两条码头共用一张表，判据问的是**焦点自己**：
    ' sld6 那对 GotFocus/LostFocus 计数器（SG 块留下的现成证人）只有 WM_SETFOCUS 真到才涨。
    ' 每次都要先把焦点挪开再问增量 —— SetFocus 落在**已经有焦点**的窗口上不会重发 WM_SETFOCUS。
    ' 落点刻意选 sld5：**启用**、又没挂任何焦点处理器。第一版这里用 sld1，而 sld1 在
    ' SL11 那一步被 Enabled = False 了 —— 禁用窗口拿不到焦点（Win32 语义，原生就是回 NULL），
    ' 于是焦点从没离开过 sld6，SN2 两条增量双双读成 N（判据自伤，不是产品红）。
    Dim gf0 As Long, gf1 As Long, gl0 As Long, gl1 As Long
    sld5.SetFocus                      ' 先把焦点放到**没挂处理器**的那枚上（当基线）
    gf0 = gGot6
    gl0 = gLost6
    sld6.SetFocus                      ' **不带括号**那一形
    Debug.Print "SN1-bare=" & TF(gGot6 > gf0) & "/" & CStr(gLost6 - gl0)
    gf1 = gGot6
    gl1 = gLost6
    sld5.SetFocus()                    ' **带括号**那一形先把焦点拿走 ⇒ sld6 该发 LostFocus
    sld6.SetFocus()                    ' 再带括号回来 ⇒ 该发 GotFocus（两形同一条原生路）
    Debug.Print "SN2-paren=" & TF(gLost6 > gl1) & "/" & TF(gGot6 > gf1)
    ' ClearSel 的**不带括号**那一形：SI4 钉的是带括号的，那条早就通了；这一形以前掉兜底。
    sld4.SelectRange = True
    sld4.Min = 10
    sld4.Max = 100
    sld4.SelStart = 20
    sld4.SelLength = 10
    sld4.ClearSel
    Debug.Print "SN3-clearsel=" & CStr(sld4.SelLength) & "/" & CStr(sld4.SelStart)

    ' ---- C29-SL-m（账 #150）: With 块里的控件**方法** ----
    ' 这一形以前不是"读数不对"，是**整件工程编译不过**：With 的成员访问只认属性，方法名落到
    ' 那条"未知属性"兜底，发成 `_vb6_with_0.SetFocus()` —— HWND 是 struct 指针，`.成员` 非法
    ' （BASE 实测：BUILD-RC=1 + 三条 VB4001 "Unknown control property '.X' in With block"）。
    ' 焦点面还是问 sld6 那对计数器、只问增量；落点用**启用**的控件（sld5/sld4，理由见 SN1 那段）。
    Dim gf2 As Long, gf3 As Long, gl2 As Long
    sld5.SetFocus()                    ' 先把焦点挪开 —— 目标已有焦点时 SetFocus 不重发 WM_SETFOCUS
    gf2 = gGot6
    With sld6
        .SetFocus                      ' With 里**不带括号**那一形
    End With
    Debug.Print "SN4-with-bare=" & TF(gGot6 > gf2)
    gf3 = gGot6
    gl2 = gLost6
    With sld5
        .SetFocus()                    ' With 里**带括号**那一形：焦点走掉才发 LostFocus
    End With
    Debug.Print "SN5-with-paren=" & TF(gLost6 > gl2) & "/" & TF(gGot6 = gf3)
    With sld4
        .SelStart = 20
        .SelLength = 10
        .ClearSel()                    ' 方法与属性赋值混在同一个 With 块里
    End With
    Debug.Print "SN6-with-clearsel=" & CStr(sld4.SelLength) & "/" & CStr(sld4.SelStart)

    ' ---- C29-SL-n（账 #141）: 只挂 `_DblClick` 的那枚也该被装 ----
    ' 驱动走已发货的那条形（SimStdEvent kind=1 = 一条真 WM_LBUTTONDBLCLK 进控件自己的过程），
    ' 本批零新增方法接线 ⇒ 红点只可能在“装不装”那一趟。
    sld8.SimStdEvent(1, 0)
    Debug.Print "SN7-dblone=" & CStr(gDbl8) & "/" & CStr(gDblCancel)

    ' ---- C29-SL-o（账 #83 的 Slider 半边）: 容器子控件的设计期**值面** ----
    ' sld7 挂在 Frame 里 ⇒ 走的是第二条创建路，而那条路以前一条 vb6_Slider_Init 都不发：
    ' .frm 里写的 Min/Max/Value/… 整条丢掉，控件就停在原生默认档（BASE 实测读出默认）。
    ' 判据形状 = 与顶层那条路**同一批 .frm 属性**逐字对上（探针 ChildProbe 顶层那枚的读数），
    ' 不是"看着非零就算过"。LargeChange 读回不等于写进去的那个数是控件自己的口径
    ' （见 SL-0 那批测量），所以这一格刻意把原值与读回值同时钉住 —— 数字变了就红。
    ' SelEnd 那格是**折出来的**（.frm 里只写 SelStart + SelLength，终点 = 起点 + 长度），
    ' 顶层那趟有的折叠，容器这一路必须有同一条 —— 复制一份发码不算接上。
    Debug.Print "SO1-child=" & CStr(sld7.Min) & "/" & CStr(sld7.Max) & "/" & CStr(sld7.Value) & "/" & CStr(sld7.SmallChange)
    Debug.Print "SO2-child-page=" & CStr(sld7.LargeChange) & "/" & CStr(sld7.TickFrequency)
    Debug.Print "SO3-child-sel=" & CStr(sld7.SelStart) & "/" & CStr(sld7.SelLength) & "/" & TF(sld7.SelectRange)

    ' ---- C29-SL-p（说明 §4 那条例子量出来的）: FontSize 存什么读什么 ----
    ' 点号 → 像素是有损的一步（96 DPI 下 1pt = 1.3333px），旧写法从窗口 LOGFONT 反算，
    ' 于是写 8 读回 8.25、写 10 读回 9.75、写 14 读回 14.25（探针 .build/slfont 那张表）。
    ' 现在请求值按窗口自存，未设过的控件仍走反算（行为一字不动）。
    ' SP2 是**第二头**：自存的数读回来当然还是自存的数，那一条问不出窗口 —— 所以要同时问
    ' 一次像素高度（证人 FontPixelHeight，C3 扩展读数），而且只比**相对**高低、不钉绝对像素
    ' （与 ChannelTop / TickStyle 那条同一教训：换 DPI 就换数）。
    ' SP3 钉的是 0 那一档：SetPropW(0) 等于删属性（账 #107），少了那枚 Set 旗标，
    ' 写 0 会静默变回"没设过"、读回来是系统默认字号。
    txtF.FontSize = 8
    txtF.FontSize = 10
    Debug.Print "SP1-round=" & CStr(txtF.FontSize)
    txtF.FontSize = 14
    Debug.Print "SP1b-round14=" & CStr(txtF.FontSize) & "/" & CStr(txtG.FontSize)
    txtF.FontSize = 26
    Debug.Print "SP2-real=" & TF(txtF.FontPixelHeight > 2 * txtG.FontPixelHeight) & "/" & CStr(txtF.FontSize)
    txtF.FontSize = 0
    Debug.Print "SP3-zero=" & CStr(txtF.FontSize)

    ' ---- C29-SL-q（账 #154）: .frm 写的设计期字体**到不到窗口** ----
    ' 探针 .build/slfont 量的：一枚写 FontName=Consolas + FontSize=14 的文本框运行时读回 空串/8.25，
    ' 另一枚写 FontSize=20 的读回 8.25、窗口像素高度还是默认的 11 ⇒ 解析侧留了这两条属性
    ' （frm_parser.cpp 原样存进 properties），缺的只是发码那一步 —— 与 #125 / #142 同一形状。
    ' txtH 在窗体上（第一条创建路）、txtI 在 Frame 里（第二条）—— 这一族栽过几次"只接一头"，
    ' 所以两条路各钉一枚。SQ3 问的是**窗口**（证人 FontPixelHeight，只比相对高低，不钉绝对像素）。
    Debug.Print "SQ1-dt=" & txtH.FontName & "/" & CStr(txtH.FontSize)
    Debug.Print "SQ2-child-dt=" & CStr(txtI.FontSize)
    Debug.Print "SQ3-real=" & TF(txtH.FontPixelHeight > 2 * txtG.FontPixelHeight)
    Debug.Print "SLIDER-DONE"
    Unload Me
End Sub

' VB6 这两条处理器都没有参数（与 DT-c 那三条同形），所以生成的回调形参表是空的。
Private Sub sld1_Change()
    gChg1 = gChg1 + 1
End Sub

Private Sub sld1_Scroll()
    gScr1 = gScr1 + 1
End Sub

Private Sub sld2_Change()
    gChg2 = gChg2 + 1
End Sub

' C29-SL-d: 常规事件那四条的计数器，加一条 Change（按键会让控件真动，所以它是"顺带的证人"）。
Private gClk6 As Long
Private gDbl6 As Long
Private gKeyD6 As Long
Private gKeyUp6 As Long
Private gChg6 As Long
Private gKeyC6 As Long
Private gClk2 As Long

Private Sub sld6_Click()
    gClk6 = gClk6 + 1
End Sub

Private Sub sld6_DblClick()
    gDbl6 = gDbl6 + 1
End Sub

Private Sub sld6_KeyDown(KeyCode As Integer, Shift As Integer)
    gKeyD6 = gKeyD6 + 1
    gKeyC6 = KeyCode
End Sub

Private Sub sld6_KeyUp(KeyCode As Integer, Shift As Integer)
    gKeyUp6 = gKeyUp6 + 1
End Sub

Private Sub sld6_Change()
    gChg6 = gChg6 + 1
End Sub

' C29-SL-g: 焦点那两条的计数器。sld3 只挂 GotFocus —— 它是"只靠焦点处理器也该被装"的证人。
Private gGot6 As Long
Private gLost6 As Long
Private gGot3 As Long

Private Sub sld6_GotFocus()
    gGot6 = gGot6 + 1
End Sub

Private Sub sld6_LostFocus()
    gLost6 = gLost6 + 1
End Sub

Private Sub sld3_GotFocus()
    gGot3 = gGot3 + 1
End Sub

' SD7 的那枚证人：sld2 除 Change（走父窗那条通道）之外只有 sld2_Click 这一条 ⇒ 它是否被
' 子类化，完全由「装不装那一趟」认不认 _Click 决定（计数器在上面的 SL-d 那块里）。
Private gDbl8 As Long
Private gDblCancel As Long

' C29-SL-n（账 #141）: 这枚滑杆**只挂 `_DblClick` 一条**处理器 —— 装的判据以前少这一形，
' 于是子类过程与 WM_LBUTTONDBLCLK 那条 arm 都生成了却没人 install，处理器编得过、永不触发。
' （sld6 也挂 DblClick，但它同时挂着 Click/KeyDown ⇒ 被顺带装上，看不出这一形死了。）
Private Sub sld8_DblClick(Cancel As Integer)
    gDbl8 = gDbl8 + 1
    gDblCancel = Cancel
End Sub

Private Sub sld2_Click()
    gClk2 = gClk2 + 1
End Sub
