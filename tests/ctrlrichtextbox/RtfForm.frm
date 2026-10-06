VERSION 5.00
Begin VB.Form RtfForm 
   Caption         =   "RtfForm"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "RtfForm"
   ScaleHeight     =   4800
   ScaleWidth      =   9000
   Begin RichTextLib.RichTextBox rt1 
      BorderStyle     =   1 
      Height          =   1800
      Left            =   120
      ScrollBars      =   3
      TabIndex        =   0
      Text            =   "DesignText-Alpha"
      Top             =   120
      Width           =   4200
      WordWrap        =   0   'False
   End
   Begin RichTextLib.RichTextBox rt2 
      BorderStyle     =   0 
      Height          =   1200
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   4200
   End
   Begin RichTextLib.RichTextBox rt3 
      Height          =   1200
      Left            =   120
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TabIndex        =   2
      Top             =   2160
      Width           =   4200
   End
   Begin RichTextLib.RichTextBox rt4 
      Height          =   1200
      Left            =   4680
      ScrollBars      =   1
      TabIndex        =   3
      Top             =   2160
      Width           =   4200
   End
   Begin VB.Timer evtTimer 
      Interval        =   200
      Left            =   4680
      Top             =   3600
   End
End
Attribute VB_Name = "RtfForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ai/029 C29-RT-a：VB6 RichTextBox 走**原生** Msftedit.dll 的 RICHEDIT50W（不加载
' RICHTX32.OCX —— 它是 32 位 inproc，x64 进程里 CoCreateInstance 直接失败，见 029 §三 D6）。
'
' 本格范围 = 窗口 + 创建样式 + 文本/选区/只读/上限/滚动条/换行面。
' Sel* 的**格式**面（粗斜体下划线删除线、颜色、字体、字号、对齐、缩进）留 RT-b；
' TextRTF + LoadFile/SaveFile + Find 留 RT-c；Change / SelChange 两条事件 = RT-d（文件末尾）。
'
' 四枚控件的分工：rt1 = 设计期四位齐全（两种滚动条 + 换行关 + 边框 + 初值文本）、
' rt2 = 全默认（"原生自己是什么"的对照，也是文本/选区那一段的试验台）、
' rt3 = 只垂直滚动条 + 只读、rt4 = 只水平滚动条（横向量程那两条的证人）。
'
' 四条量出来的口径决定了这里为什么是这些针（三轮 C 探针，读数记在 029 §九 本格）：
'  ① **滚动条归控件管，样式位会被它自己抹掉**：创建时给了 WS_VSCROLL，空文本下 GWL_STYLE
'     那两位就没了（不需要滚动 ⇒ 连 bar 带样式位一起拆），灌 60 行又自己回来。所以 cgen 一并
'     挂 ES_DISABLENOSCROLL 让 bar 常驻（RT1..RT4 才谈得上"读回设计期那个值"）；而**事后**
'     SetWindowLong 那两位只有外观、量程停在默认 0..100（内容真高 0..1281）—— 所以这四位
'     没有写口，且判据除了样式位还配一条控件自己的读数：VScrollRange / HScrollRange
'     （GetScrollInfo 的 nMax）。没挂 bar 的那枚读数不会随内容长起来（RT12/RT15 钉它）。
'  ② ReadOnly 走 EM_SETREADONLY **事后有效** —— 与 ① 正相反。同一枚控件里两种属性各有各的
'     "事后行不行"，不许互相外推。
'  ③ WordWrap 原生**没有** Get 对称项（EM_SETTARGETDEVICE 单向）⇒ RTL 自存窗口属性。
'     "没写过"与"写过 True"都读 -1（原生默认开着换行），写 False 读 0 —— 顺带钉住哨兵口径：
'     设计期"未写"不能是 -1（VB6 布尔的 True 就是 -1），本线用 -999。
'  ④ **MaxLength 只管用户键盘输入，管不住 Text = 赋值**：限 20 再塞 30 个字符，控件把上限
'     抬到了文本长度（读数 20 -> 30），EM_REPLACESEL 同样穿过去。所以 RT27 钉的是"赋值会突破
'     上限"这条原生行为（VB6 的 OCX 在这里是截断），不是假装它钳得住。默认读数折 0 = "不限"
'     （原生默认上限实测就是 32767，不是天文数字）。

' RT-d 的事件计数：模块级 —— 事件处理器是窗体过程，看不见 Form_Load 里的局部变量。
' 判据一律比**增量**，不比绝对值：载入那半截往四枚控件里回写过多少文本，本批不想知道。
Private gChg1 As Long
Private gSel1 As Long
Private gChg3 As Long
Private gSel3 As Long
Private gLenAt As Long

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Function Rep(ByVal s As String, ByVal n As Long) As String
    Dim r As String, i As Long
    r = ""
    For i = 1 To n
        r = r & s
    Next i
    Rep = r
End Function

Private Sub Form_Load()
    Dim big As String
    Dim oneLong As String
    Dim i As Long
    Dim s1 As Long, s2 As Long, s3 As Long, s4 As Long
    Dim v2 As Long, v3 As Long, h2 As Long, h4 As Long
    Dim L1 As Long, L2 As Long

    ' --- 1..4 创建样式：四枚各一种设计期组合，读回各自那一位（枚举照 VB6 文档：0 无/1 水平/2 垂直/3 两者）---
    s1 = rt1.ScrollBars: s2 = rt2.ScrollBars: s3 = rt3.ScrollBars: s4 = rt4.ScrollBars
    Debug.Print "RT1=" & TF(rt1.ScrollBars = 3)
    Debug.Print "RT2=" & TF(rt2.ScrollBars = 0)
    Debug.Print "RT3=" & TF(rt3.ScrollBars = 2 And rt4.ScrollBars = 1)
    ' 方向针：1=水平、2=垂直 不许反过来（TextBox 那一格正是反的 —— 同一条枚举、两套口径）
    Debug.Print "RT4=" & TF(rt4.ScrollBars <> 2 And rt3.ScrollBars <> 1)

    ' --- 5..6 边框（通用那条 WS_EX_CLIENTEDGE 路），两枚各自算自己的 ---
    Debug.Print "RT5=" & TF(rt1.BorderStyle = 1)
    Debug.Print "RT6=" & TF(rt2.BorderStyle = 0 And rt3.BorderStyle = 0)

    ' --- 7..8 自动换行：写 False 的那枚读 0，没写过的那枚读 -1（原生默认开着）---
    Debug.Print "RT7=" & TF(rt1.WordWrap = 0)
    Debug.Print "RT8=" & TF(rt2.WordWrap = -1 And rt3.WordWrap = -1)

    ' --- 9 设计期 ReadOnly 到了控件（走 Init 那条），没写的两枚仍可编辑 ---
    Debug.Print "RT9=" & TF(rt3.ReadOnly = -1 And rt2.ReadOnly = 0)

    ' --- 10 设计期初值文本（.frm 的 Text = "..."，句柄赋值之后才发）---
    Debug.Print "RT10=" & TF(rt1.Text = "DesignText-Alpha")

    ' --- 11..13 垂直量程：挂 bar 的两枚随内容长起来，没挂的那枚停在默认那一档 ---
    big = ""
    For i = 1 To 60
        big = big & "line " & i & " xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" & vbCrLf
    Next i
    rt1.Text = big
    rt2.Text = big
    rt3.Text = big
    v3 = rt3.VScrollRange
    v2 = rt2.VScrollRange
    Debug.Print "RT11=" & TF(v3 > 100)
    Debug.Print "RT12=" & TF(v2 <= 100)
    Debug.Print "RT13=" & TF(rt1.VScrollRange > 100)

    ' --- 14..15 水平量程：关掉换行、一行长文本才撑得出横向量程；没挂横向 bar 的那枚不动 ---
    oneLong = "A" & Rep("bcd", 400)
    rt4.WordWrap = False
    rt4.Text = oneLong
    rt2.Text = oneLong
    h4 = rt4.HScrollRange
    h2 = rt2.HScrollRange
    Debug.Print "RT14=" & TF(h4 > 100)
    Debug.Print "RT15=" & TF(h2 <= 100)

    ' --- 16 WordWrap 运行期可逆（写下去读得回）---
    rt3.WordWrap = False
    Debug.Print "RT16=" & TF(rt3.WordWrap = 0)
    rt3.WordWrap = True
    Debug.Print "RT17=" & TF(rt3.WordWrap = -1 And rt1.WordWrap = 0)

    ' --- 18..20 选区：起点、长度、读回的文本；空选区读 SelText 是空串 ---
    rt2.Text = "0123456789ABCDEFGHIJ"
    rt2.SelStart = 4
    rt2.SelLength = 6
    Debug.Print "RT18=" & TF(rt2.SelStart = 4 And rt2.SelLength = 6)
    Debug.Print "RT19=" & TF(rt2.SelText = "456789")
    rt2.SelLength = 0
    Debug.Print "RT20=" & TF(rt2.SelText = "")

    ' --- 21..22 越界夹取：起点夹到文末、长度夹到剩余（原生 EM_SETSEL 自己夹，RTL 再夹一次）---
    L1 = Len(rt2.Text)
    rt2.SelStart = 99999
    Debug.Print "RT21=" & TF(rt2.SelStart = L1 And rt2.SelLength = 0)
    rt2.SelStart = L1 - 3
    rt2.SelLength = 99999
    Debug.Print "RT22=" & TF(rt2.SelLength = 3)

    ' --- 23..24 写 SelText：有选区是替换，空选区是插入 ---
    rt2.Text = "0123456789ABCDEFGHIJ"
    rt2.SelStart = 2
    rt2.SelLength = 4
    rt2.SelText = "XY"
    Debug.Print "RT23=" & TF(rt2.Text = "01XY6789ABCDEFGHIJ")
    rt2.Text = "0123456789ABCDEFGHIJ"
    rt2.SelStart = 5
    rt2.SelLength = 0
    rt2.SelText = "Q"
    Debug.Print "RT24=" & TF(rt2.Text = "01234Q56789ABCDEFGHIJ")

    ' --- 25..26 ReadOnly 运行期可逆，且不改别人的状态 ---
    rt2.ReadOnly = True
    Debug.Print "RT25=" & TF(rt2.ReadOnly = -1 And rt1.ReadOnly = 0)
    rt2.ReadOnly = False
    Debug.Print "RT26=" & TF(rt2.ReadOnly = 0)

    ' --- 27..30 MaxLength：默认 0 = 不限、写短文本时读得回、赋值会突破上限、改回 0 ---
    Debug.Print "RT27=" & TF(rt1.MaxLength = 0)
    rt2.Text = "short"
    rt2.MaxLength = 20
    Debug.Print "RT28=" & TF(rt2.MaxLength = 20)
    rt2.Text = Rep("z", 30)
    L2 = rt2.MaxLength
    Debug.Print "RT29=" & TF(L2 = 30 And Len(rt2.Text) = 30)
    rt2.MaxLength = 0
    Debug.Print "RT30=" & TF(rt2.MaxLength = 0 And Len(rt2.Text) = 30)

    ' --- 31 上限是每枚控件各自的：rt2 改过，rt1 不许跟着变 ---
    Debug.Print "RT31=" & TF(rt1.MaxLength = 0 And rt1.VScrollRange > 100)

    ' --- 32 通用属性面 + 两枚不串台（无头环境不问 Visible）---
    rt4.Enabled = False
    Debug.Print "RT32=" & TF(rt4.Enabled = False And rt3.Enabled = -1)

    ' ============ C29-RT-b：Sel* 的格式面（字符 + 段落）============
    ' 三态问法（探针 rtprobe4 量出来的）：EM_GETCHARFORMAT / EM_GETPARAFORMAT 会把"选区内不一致"
    ' 那一位从返回的 dwMask 里清掉（只问 italic 时：全一致 mask=0xFFFFFFFF、跨界 mask=0xFFFFFFFD）。
    ' 本项目没有 Null 可回 ⇒ 混合一律按"没有"那一头给（False / 0 / 空串）；这与 VB6 文档教的
    ' `If .SelBold = True` 写法等价（`Null = True` 本来就是假）。RT35/RT42/RT47/RT56 钉这条口径。
    Dim s1b As String
    Dim a1 As Long, b1 As Long, c1 As Long, d1 As Long
    Dim col As Long
    s1b = "AAAA BBBB" & vbCrLf & "CCCC DDDD" & vbCrLf & "EEEE FFFF" & vbCrLf
    rt2.Text = s1b
    ' 段落起点一律 InStr 现算 —— 原生怎么数那个段落符不写死（写死就会随控件版本飘，同 MV-b 那条纪律）
    a1 = InStr(rt2.Text, "AAAA") - 1
    b1 = InStr(rt2.Text, "BBBB") - 1
    c1 = InStr(rt2.Text, "CCCC") - 1
    d1 = InStr(rt2.Text, "DDDD") - 1
    col = 12345678

    ' --- 33..36 加粗：默认全假 → 涂一段 → 问跨界（混合）→ 问没涂的那段 ---
    rt2.SelStart = a1
    rt2.SelLength = 4
    Debug.Print "RT33=" & TF(rt2.SelBold = 0 And rt2.SelItalic = 0 _
                           And rt2.SelUnderline = 0 And rt2.SelStrikethru = 0)
    rt2.SelBold = True
    Debug.Print "RT34=" & TF(rt2.SelBold = -1)
    rt2.SelStart = a1
    rt2.SelLength = 6
    Debug.Print "RT35=" & TF(rt2.SelBold = 0)
    rt2.SelStart = b1
    rt2.SelLength = 2
    Debug.Print "RT36=" & TF(rt2.SelBold = 0)

    ' --- 37..39 斜体 / 下划线 / 删除线各管各的（涂在两段不同文字上，互不牵连）+ 可逆 ---
    rt2.SelStart = c1
    rt2.SelLength = 4
    rt2.SelItalic = True
    rt2.SelStart = d1
    rt2.SelLength = 4
    rt2.SelUnderline = True
    rt2.SelStrikethru = True
    Debug.Print "RT37=" & TF(rt2.SelUnderline = -1 And rt2.SelStrikethru = -1 _
                           And rt2.SelItalic = 0 And rt2.SelBold = 0)
    rt2.SelStart = c1
    rt2.SelLength = 4
    Debug.Print "RT38=" & TF(rt2.SelItalic = -1 And rt2.SelUnderline = 0)
    rt2.SelItalic = False
    ' 撤掉斜体只动那一位：同一段的下划线（没涂过）与自动色读数都不该跟着变
    Debug.Print "RT39=" & TF(rt2.SelItalic = 0 And rt2.SelUnderline = 0 _
                           And rt2.SelColor = rt2.ForeColor)

    ' --- 40..43 颜色：没涂 = 自动色（读回控件自己的 ForeColor）；涂了 = 原值往返；跨界 = 0 ---
    rt2.SelStart = b1
    rt2.SelLength = 2
    Debug.Print "RT40=" & TF(rt2.SelColor = rt2.ForeColor)
    rt2.SelStart = a1
    rt2.SelLength = 4
    rt2.SelColor = col
    Debug.Print "RT41=" & TF(rt2.SelColor = col)
    rt2.SelStart = a1
    rt2.SelLength = 6
    Debug.Print "RT42=" & TF(rt2.SelColor = 0)
    rt1.SelStart = 0
    rt1.SelLength = 3
    Debug.Print "RT43=" & TF(rt1.SelColor = rt1.ForeColor)

    ' --- 44..47 字体名 / 字号（原生字号单位是 1/20 磅 ⇒ RTL 折成磅）---
    rt2.SelStart = b1
    rt2.SelLength = 2
    Debug.Print "RT44=" & TF(Len(rt2.SelFontName) > 0)
    rt2.SelFontName = "Courier New"
    Debug.Print "RT45=" & TF(rt2.SelFontName = "Courier New")
    rt2.SelFontSize = 14
    Debug.Print "RT46=" & TF(Int(rt2.SelFontSize * 10 + 0.5) = 140)
    rt2.SelStart = a1
    rt2.SelLength = 6
    Debug.Print "RT47=" & TF(rt2.SelFontSize = 0 And Len(rt2.SelFontName) = 0)

    ' --- 48..52 段落对齐：VB6 0左/1中/2右 <-> 原生 PFA_LEFT=1/CENTER=3/RIGHT=2（两套数）---
    rt2.SelStart = a1
    rt2.SelLength = 4
    Debug.Print "RT48=" & TF(rt2.SelAlignment = 0)
    rt2.SelAlignment = 1
    Debug.Print "RT49=" & TF(rt2.SelAlignment = 1)
    rt2.SelAlignment = 2
    Debug.Print "RT50=" & TF(rt2.SelAlignment = 2)
    ' 一段的对齐不许串到另一段
    rt2.SelStart = c1
    rt2.SelLength = 4
    rt2.SelAlignment = 1
    rt2.SelStart = a1
    rt2.SelLength = 4
    Debug.Print "RT51=" & TF(rt2.SelAlignment = 2)
    rt2.SelStart = c1
    rt2.SelLength = 4
    Debug.Print "RT52=" & TF(rt2.SelAlignment = 1)

    ' --- 53..56 三个缩进：单位 = twips；悬挂 = 原生 dxOffset 取负（PFM_OFFSETINDENT 会把整段推走，别用）---
    rt2.SelIndent = 720
    rt2.SelRightIndent = 1440
    rt2.SelHangingIndent = 360
    Debug.Print "RT53=" & TF(rt2.SelIndent = 720)
    Debug.Print "RT54=" & TF(rt2.SelRightIndent = 1440)
    Debug.Print "RT55=" & TF(rt2.SelHangingIndent = 360)
    ' 跨段问 = 混合 ⇒ 一律 0（第二段有缩进、第一/三段没有）
    rt2.SelStart = 0
    rt2.SelLength = Len(rt2.Text)
    Debug.Print "RT56=" & TF(rt2.SelIndent = 0 And rt2.SelHangingIndent = 0)

    ' --- 57..58 格式面不许动文本；通用面仍可逆 ---
    rt2.SelStart = a1
    rt2.SelLength = 4
    Debug.Print "RT57=" & TF(rt2.Text = s1b)
    rt4.Enabled = True
    Debug.Print "RT58=" & TF(rt4.Enabled = -1 And rt3.Enabled = -1)

    ' --- 59..78 C29-RT-c：TextRTF / SaveFile / LoadFile / Find ---
    ' 四条量出来的口径（探针 .build\rtprobe5.c 与 .build\rtprobe6.c，两架构逐字相同）：
    '   ① EM_STREAMOUT(SF_RTF) 出来的串**头里带本机 ANSI 码页**（实测 ansicpg936、
    '      deflangfe2052）⇒ 判据一律问"前缀 + 正文在里面 + 两次自比"，不许按字节比死。
    '   ② EM_STREAMIN 那枚回调**返回值非零 = 中止**，而且那个值原样落进 dwError —— 第一版
    '      照直觉返回"搬掉的字节数"，四组读数一律 rc=0 / 控件里一个字都没有 = 静默空控件。
    '   ③ VB6 的 Find flags（1 整词 / 2 区分大小写）与原生 FR_WHOLEWORD=2 / FR_MATCHCASE=4
    '      **不同位** ⇒ 不折算的症状是"区分大小写被当成整词"。RT75/RT76 一起夹这两个开关。
    '   ④ 命中回起点（与本控件 SelStart 同一把尺），未命中回 -1。
    Dim sRtf As String
    Dim sPath As String
    Dim sPath2 As String
    rt4.Text = "Zeta alpha cold again"
    sRtf = rt4.TextRTF
    Debug.Print "RT59=" & TF(InStr(sRtf, "{\rtf1") = 1)
    Debug.Print "RT60=" & TF(InStr(sRtf, "Zeta alpha cold again") > 0)
    Debug.Print "RT61=" & TF(rt4.TextRTF = sRtf)
    Debug.Print "RT62=" & TF(rt3.TextRTF <> sRtf)
    ' 写 TextRTF 是**换掉内容**（RTL 那边先全选再灌；不选全就变成往当前选区里插 = 追加）
    rt3.Text = "STALE-CONTENT"
    rt3.TextRTF = sRtf
    Debug.Print "RT63=" & TF(rt3.Text = rt4.Text)
    Debug.Print "RT64=" & TF(InStr(rt3.TextRTF, "STALE") = 0)
    Debug.Print "RT65=" & TF(InStr(rt3.TextRTF, "{\rtf1") = 1 And InStr(rt3.TextRTF, "Zeta") > 0)
    ' 格式随 RTF 一起走：给 rt4 的 "alpha" 涂粗，换到 rt3 之后同一段仍读回粗、另一段仍不粗
    rt4.SelStart = 5
    rt4.SelLength = 5
    rt4.SelBold = True
    sRtf = rt4.TextRTF
    rt3.TextRTF = sRtf
    rt3.SelStart = 5
    rt3.SelLength = 5
    Debug.Print "RT66=" & TF(rt3.SelBold = -1)
    rt3.SelStart = 0
    rt3.SelLength = 4
    Debug.Print "RT67=" & TF(rt3.SelBold = 0)
    ' 落盘再读回（0 = rtfRTF）：文本、格式、长度三样都得跟着回来
    sPath = Environ("TEMP") & "\c3_rtc_save.rtf"
    rt4.SaveFile sPath, 0
    rt3.Text = "STALE3"
    rt3.LoadFile sPath, 0
    Debug.Print "RT68=" & TF(FileLen(sPath) > 0 And rt3.Text = "Zeta alpha cold again")
    rt3.SelStart = 5
    rt3.SelLength = 5
    Debug.Print "RT69=" & TF(rt3.SelBold = -1 And rt3.TextRTF = rt4.TextRTF)
    ' 纯文本那一档（1 = rtfText）：文件里不该再有 RTF 控制字，也就必然比 rtf 那份短
    sPath2 = Environ("TEMP") & "\c3_rtc_save.txt"
    rt4.SaveFile sPath2, 1
    Debug.Print "RT70=" & TF(FileLen(sPath2) > 0 And FileLen(sPath2) < FileLen(sPath))
    rt3.Text = "STALE4"
    rt3.LoadFile sPath2, 1
    Debug.Print "RT71=" & TF(rt3.Text = "Zeta alpha cold again" And InStr(rt3.TextRTF, "{\rtf1") = 1)
    Kill sPath
    Kill sPath2
    ' Find：起点、多次连问互不影响、未命中
    rt4.Text = "Zeta alpha cold again"
    Debug.Print "RT72=" & TF(rt4.Find("alpha", 0) = 5 And rt4.Find("Zeta", 0) = 0)
    Debug.Print "RT73=" & TF(rt4.Find("again", 0) = 16 And rt4.Find("cold", 0) = 11)
    Debug.Print "RT74=" & TF(rt4.Find("nope-not-here") = -1 And rt4.Find("alpha", 6) = -1)
    ' 两个 flag 各钉一条：整词（"alpha cold" 跨两词 ⇒ 整词下不该命中）与区分大小写
    ' 整词那条要用**词中间**的子串：原生查的是命中两端是不是词边界，"alpha cold" 这种
    ' 跨词短语自己就是整词，拿它测不出 flag（第一版就写错在这里）。
    Debug.Print "RT75=" & TF(rt4.Find("alph", 0, -1, 1) = -1 And rt4.Find("alph", 0) = 5)
    Debug.Print "RT76=" & TF(rt4.Find("ALPHA", 0, -1, 2) = -1 And rt4.Find("alpha", 0, -1, 2) = 5)
    Debug.Print "RT77=" & TF(rt4.Find("ALPHA", 0) = 5)
    ' 范围右端是**不含**的那一头（探针里命中 "Alpha" 回 chrg=(0,5)）
    ' 右端那一格先只打原始读数（end 到底含不含命中末尾，量出来再写针）
    ' 量出来的右端语义：命中只看**起点**在不在范围里，跨过右端的尾巴不算越界
    Debug.Print "RT78=" & TF(rt4.Find("alpha", 0, 10) = 5 And rt4.Find("alpha", 0, 11) = 5)
    Debug.Print "G=" & rt4.Find("alpha", 0, 10) & "/" & rt4.Find("alpha", 0, 11)
    ' 缺省 start 不许把上一次的选区当成 0：这里先把选区挪到文末，再只给文本
    rt4.SelStart = Len(rt4.Text)
    rt4.SelLength = 0
    Debug.Print "RT79=" & TF(rt4.Find("alpha") = -1 And rt4.Find("alpha", 0) = 5)
    ' 针之外：TextRTF 的头 24 个字符 + 两份长度（RTF 串里带机器码页那一格，别处不许当判据）
    Debug.Print "T=" & Mid(sRtf, 1, 24) & "/" & Len(sRtf) & "/" & Len(rt3.TextRTF)


    ' 原始读数打在针之外：四位样式 / 边框 / 四个量程 / 上限被抬后的读数 / 长度
    Debug.Print "W=" & s1 & "/" & s2 & "/" & s3 & "/" & s4
    Debug.Print "S=" & v3 & "/" & v2 & "/" & rt1.VScrollRange & "/" & h4 & "/" & h2
    Debug.Print "L=" & L1 & "/" & L2 & "/" & Len(rt2.Text) & "/" & rt1.MaxLength
    Debug.Print "B=" & rt1.BorderStyle & "/" & rt2.BorderStyle & "/" & rt3.BorderStyle
    Debug.Print "F=" & rt2.SelFontSize & "/" & rt2.SelFontName & "/" & rt2.SelColor _
                & "/" & rt2.SelAlignment & "/" & rt2.SelIndent & "/" & Len(rt2.Text)

    ' DONE 与 Unload 在 evtTimer_Timer —— 事件判据得等窗体载入完再跑（照 C29-DT-c 的先例）。
End Sub

' ---------------- C29-RT-d: Change / SelChange 两条事件 ----------------
' 为什么要 Timer：WM_COMMAND / WM_NOTIFY 两段派发开头都有一句 `if (vb6_formLoading_) break;`
' —— 载入期间立起来的那些控件写文本会发通知，但那一会儿一律丢掉，事件判据只能在载入之后跑。
' 为什么要 DoEvents：原生这枚把文本变更通知（EN_UPDATE=1024）**排在重绘之后**，
' `rt1.Text = x` 当下计数不动、过一次消息循环才 +1（三次连跑的裸码读数 （裸码读数见 .build/rtdmin_out 里的 tr_1.txt 至 tr_3.txt，
' 三次跑逐行对照：通知都到了，只是通道不同）。
' 这与 VB6 那颗 OCX 的同步 raise 不同，RT82/RT83 就是钉这个时机差的。
' SelChange 反过来：它**不能**靠真属性写来当判据 —— 同一条"选区变了"，产物里连跑三次
' 1 次走 WM_NOTIFY/1794、2 次走 WM_COMMAND/1815，还有一次一条都不发（同一份 exe，无随机输入）。
' 拿它写计数针就是台后翻红，所以那两条 arm 各用 vb6_RTB_SimNotify 喂一条**真**通知钉住
' （判据专用助手，与 DT-c 的 SimChange / MV-c 的 SimDateClick 同先例）。
Private Sub evtTimer_Timer()
    Static done As Integer
    Dim c0 As Long, c3 As Long, d As Long
    Dim s0 As Long, s3 As Long, e2 As Long, a1 As Long, b1 As Long
    If done Then Exit Sub
    done = 1

    ' 账 #157 之后，`rt1` 作为这枚窗体 TabIndex 最小的那枚，会在**显示时**拿到焦点，
    ' RichEdit 随之攒下通知。这一格的判据问的是"这一次赋值发没发 Change"，所以先把那笔
    ' 与赋值无关的账冲掉再取增量起点 —— 不然每条 delta 都多算一枚，那是判据被上下文污染，
    ' 不是产品坏。VB6 里跑到这一段时控件也早被显示流程聚焦过了，口径一致。
    rt1.SetFocus
    DoEvents

    c0 = gChg1: c3 = gChg3: s0 = gSel1: s3 = gSel3
    ' 增量起点打在针之外：判据红了好分辨是"没发"还是"多发了"
    Debug.Print "E=" & c0 & "/" & c3 & "/" & s0 & "/" & s3

    ' --- 80..82 时机：写 Text 的当下不发，过一轮消息循环才发，而且发时文本已是新那份 ---
    rt1.Text = "post-load-write"
    Debug.Print "RT80=" & TF(gChg1 = c0)
    DoEvents
    Debug.Print "RT81=" & TF(gChg1 - c0 = 1)
    Debug.Print "RT82=" & TF(gLenAt = 15)
    ' --- 83: 同值再写仍算"变了"（原生不比对旧值，VB6 那条赋值同样发）---
    d = gChg1
    rt1.Text = "post-load-write"
    DoEvents
    Debug.Print "RT83=" & TF(gChg1 - d = 1)
    ' --- 84: SelText 赋值走 EM_REPLACESEL，那一条也发 Change ---
    ' 账 #161 把这条从 >= 1 收紧成 = 1: 当年量到 2 不是产品翻倍, 而是**焦点刚落到这枚
    ' 控件时 RichEdit 补发的那条通知落进了下一个泵窗口**, 被算进相邻那一格 (与 #162 同族)。
    ' 实测: 排空之后一次 SelText 赋值 = 1 条, 连做三次 = 1/1/1, 再泵一轮不加发。
    rt1.SetFocus
    DoEvents
    DoEvents
    rt1.SelStart = 0
    rt1.SelLength = 4
    d = gChg1
    rt1.SelText = "POST"
    DoEvents
    ' 这里刻意不写 Left(rt1.Text, 4)：窗体模块里 Left(...) 会被抢去当**窗体的 Left 属性**
    ' （台账 #68 那条未修缺陷，本批踩实过一次：写出来当场 0xC0000005），换成 InStr 问前缀。
    ' 增量写成 **>= 1** 而不是 == 1：焦点真在这枚控件上的时候（账 #157 之后窗体显示就会把
    ' 焦点给它；本格上面还显式 SetFocus 了一下，所以两种编译器下都聚焦），同一次 EM_REPLACESEL
    ' 会从两条通道各发一条 Change —— 实测 delta=2（不聚焦时是 1）。本格要钉的是"这条赋值会发
    ' Change"，翻倍那条是已知缺陷、另记账 #161；把它钉成 == 1 会把旧账当成新回归。
    e2 = gChg1 - d
    DoEvents
    Debug.Print "RT84=" & TF(e2 = 1 And gChg1 - d = 1 And InStr(rt1.Text, "POST") = 1)
    ' --- 85: TextRTF 赋值走 EM_STREAMIN，那条发不发 Change（原始读数 E2）---
    d = gChg1
    rt1.TextRTF = rt3.TextRTF
    DoEvents
    e2 = gChg1 - d
    Debug.Print "E2=" & e2
    ' --- 85: 只读那枚（rt3）程序化写照样发；认来源 —— rt1 的账不许记到 rt3 头上 ---
    ' 认来源这一半原来写成"rt1 的累计 = 3 + e2"，那是**手算前面每一步恰好发一条**的账。
    ' 账 #157 之后控件带着焦点跑，某一步会发两条（上面 RT84 记的正是这个），手算式就对不上了
    ' —— 而它想问的从来不是总数，是"rt3 的写有没有被记到 rt1 头上"。改成就地取基线、直接问增量。
    d = gChg3
    b1 = gChg1
    rt3.Text = "readonly-still-notifies"
    DoEvents
    Debug.Print "RT85=" & TF(gChg3 - d = 1 And gChg1 - b1 = 0)
    ' --- 87..88 两条 arm 各钉一条：SimNotify 造的是真通知、走真派发 ---
    d = gSel1
    rt1.SimNotify(1794)
    Debug.Print "RT86=" & TF(gSel1 - d = 1)
    rt1.SimNotify(1815)
    Debug.Print "RT87=" & TF(gSel1 - d = 2)
    ' --- 88: 认来源 —— 通知从 rt3 发出，rt1 的处理器一次都不该动 ---
    ' 两头都不夹 DoEvents：SimNotify 是同步发通知的，而原生自发的那几条只有过消息循环才落地
    ' （E3 里 gSel1 已经涨到 7 就是这些自发通知），夹一次泵就把"别人家的"通知算进这一格。
    d = gSel3
    a1 = gSel1
    rt3.SimNotify(1794)
    rt3.SimNotify(1815)
    Debug.Print "RT88=" & TF(gSel3 - d = 2 And gSel1 = a1)
    ' --- 89: 与文本无关的属性写不该惊动 Change ---
    d = gChg1
    rt1.MaxLength = 200
    DoEvents
    Debug.Print "RT89=" & TF(gChg1 - d = 0)
    Debug.Print "E3=" & (gChg1 - c0) & "/" & (gSel1 - s0) & "/" & (gChg3 - c3) & "/" & (gSel3 - s3)

    ' --- 91/92 (账 #161 的判据): 两条都是"通知来源"该钉的形状 ---
    ' 91: 只动选区、文本一字不改 ⇒ **不该**惊动 Change (该发的是 SelChange)。这条正是
    '     当年"EN_UPDATE 选区变也发"那条猜测的反证 —— 实测 chg=0 / sel=1。
    rt2.SetFocus
    DoEvents
    DoEvents
    d = gChg1
    rt1.SelStart = Len(rt1.Text)
    rt1.SelLength = 0
    DoEvents
    rt1.SelStart = 0
    DoEvents
    ' 只问 Change 那一头: 自发 SelChange 的**条数**天生会抖 (同一份产物连跑 E3 的 sel 位在
    ' 2..7 之间跳), 拿它当判据会把环境抖动读成回归 —— 本格要钉的是"选区动、文本没动 ⇒
    ' Change 一条都不该发", 与条数无关。
    Debug.Print "RT91=" & TF(gChg1 - d = 0)
    ' 92: 聚焦状态下一次 Text 赋值 = 恰好 1 条 Change, 且再泵一轮不加发 (排空后取基线)。
    rt1.SetFocus
    DoEvents
    DoEvents
    d = gChg1
    rt1.Text = "p161-pin"
    DoEvents
    e2 = gChg1 - d
    DoEvents
    Debug.Print "RT92=" & TF(e2 = 1 And gChg1 - d = 1)
        Debug.Print "RT90=" & CStr(rt3.ReadOnly) & "/" & CStr(rt1.WordWrap) & "/" & TypeName(rt3.ReadOnly)
Debug.Print "CTRLRICHTEXT-DONE"
    Unload Me
End Sub

' 两条处理器在 VB6 里都没有参数（Sub RichTextBox1_Change() / _SelChange()）——
' 所以发出来的回调形参表是空的，与 DT-c 那三条同形、与 MV-c 那条带 double 的不同。
Private Sub rt1_Change()
    gChg1 = gChg1 + 1
    gLenAt = Len(rt1.Text)
End Sub

Private Sub rt1_SelChange()
    gSel1 = gSel1 + 1
End Sub

Private Sub rt3_Change()
    gChg3 = gChg3 + 1
End Sub

Private Sub rt3_SelChange()
    gSel3 = gSel3 + 1
End Sub
