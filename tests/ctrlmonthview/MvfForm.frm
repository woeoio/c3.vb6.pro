VERSION 5.00
Begin VB.Form MvfForm 
   Caption         =   "MvfForm"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "MvfForm"
   ScaleHeight     =   4800
   ScaleWidth      =   9000
   Begin MSComCtl2.MonthView mv1 
      Height          =   2400
      Left            =   240
      MaxSelCount     =   7
      MonthColumns    =   2
      MonthRows       =   1
      MultiSelect     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6000
   End
   Begin MSComCtl2.MonthView mv2 
      Height          =   1560
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   2520
   End
   Begin MSComCtl2.MonthView mv3 
      Height          =   1560
      Left            =   3000
      ShowWeekNumbers =   -1  'True
      TabIndex        =   2
      Top             =   2760
      Width           =   2520
   End
   Begin MSComCtl2.MonthView mv4 
      Height          =   1560
      Left            =   5760
      ShowToday       =   0   'False
      TabIndex        =   3
      Top             =   2760
      Width           =   2520
   End
   Begin VB.Timer evtTimer 
      Interval        =   200
      Left            =   240
      Top             =   4440
   End
End
Attribute VB_Name = "MvfForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ai/029 C29-MV-a：VB6 MonthView 走**原生** SysMonthCal32（不加载 MSCOMCT2.OCX —— 它是
' 32 位 inproc，x64 进程里 CoCreateInstance 直接失败，见 029 §三 D6）。
'
' 1..22 格 = "窗口 + 样式 + 标量属性面"；23..33 是 C29-MV-b 的 Date 值面
' （Value / SelStart / SelEnd）。MCN_SELCHANGE(-749) 那条事件（VB6 的 DateClick）留 MV-c。
'
' 三条与 DTPicker 不同的量出来的事实，判据的形状就是被它们决定的：
'  ① **多月平铺不是一个开关，而是"窗口多大"**。原生没有"给几行几列"这条消息 —— 头 6361 行
'     那段注释教的做法是"把 MCM_GETMINREQRECT 的矩形按想要的月数乘大，控件就自己排几个"。
'     所以 MonthRows/MonthColumns 由 RTL 撑矩形落进窗口尺寸，判据一律问
'     MCM_GETCALENDARCOUNT（眼下真画了几个月）—— 这是控件自己的答案，不是我们那张表。
'  ② **ShowToday 与原生 MCS_NOTODAY 是反的**，而且它管的正是 MV-b 要用的那条"today 高亮"。
'     所以这一格不能像 CheckBox 那样按"写下去读回来"了结 —— 读回来对、观感是另一件事，
'     于是每条样式判据都配一条 MCM_GETMINREQRECT 的控件侧尺寸（要不要今天那一行会改高、
'     装不装周号会改宽），尺寸对不上就翻红。
'  ③ MaxSelCount 是**真往返过控件**的一条（MCM_SET/GETMAXSELCOUNT），不像 CustomFormat 那样
'     得自存。而它只在挂了 MCS_MULTISELECT 的控件上写得住：mv2 没那一位，写 9 之后读回还是默认
'     的 1（MV20 钉这条）。于是这格顺手成了"样式位真不真"的第二个控件侧证人 —— 问的是控件有
'     没有按那位办事，不是"我们的掩码读数说我们写进去了"。
'  ④ 与 DTPicker 相反的一条：这三位样式**运行期写是有效的**（MV21 用 MCM_GETMINREQRECT 的高
'     度跟着变来证）。DTS_SHOWNONE 那两位会被控件抹回去，MCS_ 这三位不会 —— 所以 MonthView
'     没有"只能创建时给"这条边界，设计期照旧立进创建参数只是为了让观感从第一帧就对。

Private gClick1 As Long
Private gClick2 As Long
Private gGot1 As Date

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Sub Form_Load()
    Dim k As Long
    Dim d0 As Long
    Dim bkg As Long, txt As Long, trl As Long, tbk As Long, ttx As Long, mbk As Long
    Dim w2 As Long, w3 As Long, h2 As Long, h4 As Long
    Dim n2 As Long, n4 As Long

    ' 六个色值刻意取"三字节都不同且都很小"的组合，撞上任一默认色（0xFFFFFF / 0 / 0x808080
    ' 那一族）就会让 MV14 那条"控件之间不共享"变成假绿或假红。
    bkg = 66051: txt = 263430: trl = 460809: tbk = 658188: ttx = 855567: mbk = 250131

    ' --- 1..4 创建样式进了窗口（四枚各一种设计期组合；mv2 是全默认的对照） ---
    Debug.Print "MV1=" & TF(mv1.MultiSelect = -1)
    Debug.Print "MV2=" & TF(mv2.MultiSelect = 0 And mv2.ShowWeekNumbers = 0 And mv2.ShowToday = -1)
    Debug.Print "MV3=" & TF(mv3.ShowWeekNumbers = -1 And mv3.MultiSelect = 0)
    Debug.Print "MV4=" & TF(mv4.ShowToday = 0)

    ' --- 5..6 多月平铺：控件自己报"眼下画了几个月" ---
    Debug.Print "MV5=" & TF(mv1.MonthCount = 2)
    Debug.Print "MV6=" & TF(mv2.MonthCount = 1)

    ' --- 7..10 控件侧尺寸：周号加宽、今天那一行加高（问的是 MCM_GETMINREQRECT） ---
    w2 = mv2.MinReqWidth: w3 = mv3.MinReqWidth
    h2 = mv2.MinReqHeight: h4 = mv4.MinReqHeight
    Debug.Print "MV7=" & TF(w2 > 0 And w3 > w2)
    Debug.Print "MV8=" & TF(h2 > 0 And h2 > h4)
    Debug.Print "MV9=" & TF(mv1.MinReqWidth = w2 And mv1.MinReqHeight = h2)
    Debug.Print "MV10=" & TF(mv4.MinReqWidth > 0)

    ' --- 11..17 六格配色（VB6 真名，MonthBackColor 在官方那页的配色段里就有）：逐格写、逐格读回自己那格 ---
    mv2.BackColor = bkg
    mv2.ForeColor = txt
    mv2.TitleBackColor = tbk
    mv2.TitleForeColor = ttx
    mv2.TrailingForeColor = trl
    mv2.MonthBackColor = mbk
    Debug.Print "MV11=" & TF(mv2.BackColor = bkg)
    Debug.Print "MV12=" & TF(mv2.ForeColor = txt)
    Debug.Print "MV13=" & TF(mv2.TitleBackColor = tbk)
    Debug.Print "MV14=" & TF(mv2.TitleForeColor = ttx And mv2.TrailingForeColor = trl)
    Debug.Print "MV15=" & TF(mv2.MonthBackColor = mbk)
    ' 改一格必须不动另外几格（MCSC_ 那六个序号撞车的话就是这条红）
    mv2.TitleBackColor = 1234567
    Debug.Print "MV16=" & TF(mv2.BackColor = bkg And mv2.ForeColor = txt _
                              And mv2.TitleForeColor = ttx And mv2.TrailingForeColor = trl _
                              And mv2.MonthBackColor = mbk)
    ' 底色是每枚控件各自的，mv1 从没涂过色 ⇒ 五条都不该是上面那些值
    Debug.Print "MV17=" & TF(mv1.BackColor <> bkg And mv1.TitleForeColor <> ttx)

    ' --- 18..20 设计期 MaxSelCount 真到了控件；没那枚样式位的控件**问得出默认值、写不进去** ---
    Debug.Print "MV18=" & TF(mv1.MaxSelCount = 7)
    mv1.MaxSelCount = 3
    Debug.Print "MV19=" & TF(mv1.MaxSelCount = 3)
    n2 = mv2.MaxSelCount
    mv2.MaxSelCount = 9
    n4 = mv2.MaxSelCount
    ' 量出来的口径：原生默认上限是 1，而 MCM_SETMAXSELCOUNT 在**没挂 MCS_MULTISELECT** 的控件上
    ' 直接被拒（写完读回还是 1）。这就是"样式位真不真"的控件侧证据 —— 比 GWL_STYLE 那一圈
    ' 读数硬，因为它问的是控件按没按那位在办事。mv1 有那一位，所以上面 MV19 写 3 就读得出 3。
    Debug.Print "MV20=" & TF(n2 = 1 And n4 = 1)
    ' 原始读数打在针之外：mv2 默认上限 / 事后写 9 之后的读数 / 四个尺寸
    Debug.Print "R=" & n2 & "/" & n4 & " " & w2 & "/" & w3 & "/" & h2 & "/" & h4

    ' --- 21 运行期改样式位：读回来是一回事，控件的尺寸读数跟没跟是另一回事 ---
    mv4.ShowToday = True
    Debug.Print "MV21=" & TF(mv4.ShowToday = -1 And mv4.MinReqHeight > h4)

    ' --- 22 通用属性面 + 两枚不串台 ---
    mv3.Enabled = False
    Debug.Print "MV22=" & TF(mv3.Enabled = False And mv2.Enabled = -1)

    ' --- 23..33 C29-MV-b：Date 值面（Value = 单选那格；SelStart/SelEnd = 原生那张两端表）---
    ' 换算与 DTPicker 同一对助手（oleaut32 的 VariantTimeToSystemTime / SystemTimeToVariantTime），
    ' VB 侧 Date 就是 double 序列号，所以 43894 这种整数序列直接当日期用（= 2020-03-04）。
    '
    ' 两条拿 C 探针量出来的硬边界，判据的形状是它们决定的（读数见 029 §九 本格）：
    '  ① **两张表互斥**：没挂 MCS_MULTISELECT ⇒ MCM_GET/SETCURSEL 有效、MCM_GET/SELRANGE 一律
    '     失败；挂了 ⇒ 正好反过来。所以 `Value` 与 `SelStart`/`SelEnd` 谁问得出来，
    '     由那枚样式位决定（MV28 钉非多选那侧，MV31 钉多选那侧，两条合起来才是这条互斥）。
    '  ② SETSELRANGE 只在**当前显示的那一个月**里挑日子，跨月会被夹（探针里 2020-03 的两天
    '     夹成同一天）。所以判据一律以"本月 1 号"为基准往外推，不写死日期 —— CI 上跑的日期
    '     未知，写死就会在月初/月末随机红（同一条纪律见 022 账 #79）。
    d0 = Int(Now) - Day(Now) + 1
    Debug.Print "MV23=" & TF(Int(mv2.Value) = Int(Now))
    mv2.Value = 43894
    Debug.Print "MV24=" & TF(CLng(mv2.Value) = 43894)
    mv2.Value = 44562
    Debug.Print "MV25=" & TF(CLng(mv2.Value) = 44562 And CLng(mv3.Value) <> CLng(mv2.Value))
    mv1.MaxSelCount = 30
    mv1.SelStart = d0 + 1
    mv1.SelEnd = d0 + 5
    Debug.Print "MV26=" & TF(CLng(mv1.SelStart) = d0 + 1 And CLng(mv1.SelEnd) = d0 + 5)
    ' 改一端必须不动另一端（原生是一张 (起,止) 表；只发一端会把另一端拆成 0 年那天）
    mv1.SelStart = d0 + 3
    Debug.Print "MV27=" & TF(CLng(mv1.SelStart) = d0 + 3 And CLng(mv1.SelEnd) = d0 + 5)
    ' 非多选的那枚：范围表问不出、也写不进（两格都回 0，写完还是 0）
    mv2.SelStart = d0 + 1
    Debug.Print "MV28=" & TF(mv2.SelStart = 0 And mv2.SelEnd = 0)
    Debug.Print "MV29=" & TF(Int(mv3.Value) = Int(Now))
    ' 原生月历只有"天"这一格 ⇒ 一天的分数部分落不到控件上，读回来是那天零点
    mv2.Value = 44562.75
    Debug.Print "MV30=" & TF(CLng(mv2.Value) = 44562)
    ' 多选那枚反过来：Value 这条问不通（GET/SETCURSEL 在 MCS_MULTISELECT 下直接失败 ⇒ 回 0）
    mv1.Value = 43894
    Debug.Print "MV31=" & TF(mv1.Value = 0)
    ' 上限到底管不管得住范围：收到 3 天再要 7 天 ⇒ 夹完应当正好 3 天（夹的是哪一端由读数说）
    mv1.MaxSelCount = 3
    mv1.SelStart = d0 + 1
    mv1.SelEnd = d0 + 7
    k = CLng(mv1.SelEnd) - CLng(mv1.SelStart) + 1
    Debug.Print "MV32=" & TF(k = 3)
    ' 归一化方向（RTL vb6_MvSetSelEnd 在发 MCM_SETSELRANGE 之前把宽度挤到 MaxSelCount）：
    '   写 SelStart 定住 Start、把 SelEnd 挤到 Start+Max-1；写 SelEnd 定住 End、把 SelStart
    '   挤到 End-Max+1。这条按 [d0+5, d0+7] 落点钉 —— 前一版把这条钉成"顶回去, 两端都不动",
    '   那是把原生沉默拒收当成了语义，跟 VB6 文档里"另一端跟着调"的口径不一致；跨月那天
    '   (10-01) MV26 的初值退化 range 一起把这个内部矛盾暴露了出来 —— 三条判据现在按同
    '   一种"写完读得到"归一化模型钉在一起。
    Debug.Print "MV33=" & TF(CLng(mv1.SelEnd) = d0 + 7 And CLng(mv1.SelStart) = d0 + 5)
    Debug.Print "S=" & k & "/" & CLng(mv1.SelEnd) & "/" & CLng(mv2.SelStart) & "/" & d0

    ' DONE 与 Unload 在 evtTimer_Timer —— 事件判据得等窗体载入完再跑 (照 C29-DT-c 的先例)。
End Sub

' ---------------- C29-MV-c: DateClick（原生 MCN_SELCHANGE）----------------
' SimDateClick 是**判据专用**助手（与 DT-c 的 SimChange 同一先例）：无头环境点不了鼠标，
' 而直接调 handler 会绕开整条派发链 —— 只有从真 WM_NOTIFY 进父窗，才验得到
' "case WM_NOTIFY + code 分流 + hwndFrom 认来源 + 负载折算成 Date"四段都接上了。
Private Sub evtTimer_Timer()
    Static done As Integer
    Dim b1 As Long, b2 As Long, d0 As Long, e1 As Long
    If done Then Exit Sub
    done = 1

    b1 = gClick1: b2 = gClick2
    d0 = Int(Now) - Day(Now) + 1

    ' --- 34. 走完整派发链，且**参数就是负载里那一天**（折算与属性读数同一把尺）---
    mv1.SimDateClick(d0 + 4)
    Debug.Print "MV34=" & TF(gClick1 - b1 = 1 And CLng(gGot1) = d0 + 4)

    ' --- 35. 认来源：通知从 mv2 发出，mv1 的 handler 不许动 ---
    mv2.SimDateClick(d0 + 6)
    Debug.Print "MV35=" & TF(gClick1 - b1 = 1 And gClick2 - b2 = 1)

    ' --- 36. 程序化改选不叫 DateClick（原生那条通知只由用户交互驱动，与 DT-c 的 Change 同型）---
    mv2.Value = d0 + 8
    mv1.SelStart = d0 + 9
    mv1.SelEnd = d0 + 11
    Debug.Print "MV36=" & TF(gClick1 - b1 = 1 And gClick2 - b2 = 1)

    ' --- 37. 另一枚控件的处理器不许被这条通道顺带叫起来（mv3 没写 DateClick）---
    mv1.SimDateClick(d0 + 5)
    e1 = gClick1
    mv3.SimDateClick(d0 + 5)
    Debug.Print "MV37=" & TF(e1 - b1 = 2 And gClick2 - b2 = 1)

    ' --- 38..40 值面必须是**整数日**（针一律按数值比，不取整）---
    '     comctl **6** 会把**当前挂钟时间**混进 MCM_GETCURSEL 回来的 SYSTEMTIME 里（本机实测
    '     发 43894 回 43894.765127，而不挂 manifest 的那份回 43894.000000）⇒ 不抹平则
    '     `CLng(mv.Value)` 下午进位、上午不进位，而 CI 那批跑在上午 ⇒ 门一直拦不住。
    '     第 40 条另钉止端：控件把 hi 存成【当日 23:59:59.9】，旧写法 `-1.0` 读回的是
    '     46270.999988 那种数（BASE 那轮的 `D=` 原始行就是这个），只有靠 `CLng` 才读对；
    '     现在两端都必须是整日。
    mv2.Value = 43894
    Debug.Print "MV38=" & TF(mv2.Value = 43894)
    mv2.Value = 44562.75
    Debug.Print "MV39=" & TF(mv2.Value = 44562)
    mv1.SelStart = d0 + 3
    mv1.SelEnd = d0 + 5
    Debug.Print "MV40=" & TF(mv1.SelStart = d0 + 3 And mv1.SelEnd = d0 + 5)
    Debug.Print "D=" & CDbl(mv2.Value) & "/" & CDbl(mv1.SelStart) & "/" & CDbl(mv1.SelEnd)

    Debug.Print "E=" & (gClick1 - b1) & "/" & (gClick2 - b2) & "/" & CLng(gGot1)
        Debug.Print "MV41=" & CStr(mv1.MultiSelect) & "/" & CStr(mv4.ShowToday) & "/" & TypeName(mv1.MultiSelect)
Debug.Print "CTRLMONTHVIEW-DONE"
    Unload Me
End Sub

' VB6 的签名是 Sub MonthView1_DateClick(ByVal DateSelected As Date) —— 与 DT-c 那三条
' 无参事件的**关键差别**：这条带一个 Date 参数，要按负载折算再传。传法跟着 ByVal 走
' （ByVal ⇒ 原样传值；Form_MouseDown 那批没写 ByVal 的才是 int16_t*/float* 指针）——
' 这条形状写错时 x64 侥幸读得出正确日期、x86 直接错值，所以 MV34 两架构都得跑。
Private Sub mv1_DateClick(ByVal DateSelected As Date)
    gClick1 = gClick1 + 1
    gGot1 = DateSelected
End Sub

Private Sub mv2_DateClick(ByVal DateSelected As Date)
    gClick2 = gClick2 + 1
End Sub
