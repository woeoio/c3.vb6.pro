VERSION 5.00
Begin VB.Form CtrlSSTab
   Caption         =   "CtrlSSTab"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6000
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrCheck
      Interval        =   100
      Enabled         =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   3
      Tab             =   1
      TabHeight       =   520
      TabOrientation  =   0
      TabStyle        =   0
      TabsPerRow      =   3
      WordWrap        =   0
      TabPic16(0)     =   "CtrlSSTab.frx":0000
      BeginProperty Tabs {1EF78043-95F0-11D0-B849-00A0C90DC8A9}
         NumTabs         =   3
         BeginProperty Tab1 {1EF78045-95F0-11D0-B849-00A0C90DC8A9}
            Caption         =   "常规"
            Object.Tag             =   ""
            ImageVariant    =   2
            VersionOriginal =   32
         EndProperty
         BeginProperty Tab2 {1EF78045-95F0-11D0-B849-00A0C90DC8A9}
            Caption         =   "视图"
            Object.Tag             =   ""
            ImageVariant    =   2
            VersionOriginal =   32
         EndProperty
         BeginProperty Tab3 {1EF78045-95F0-11D0-B849-00A0C90DC8A9}
            Caption         =   "高级"
            Object.Tag             =   ""
            ImageVariant    =   2
            VersionOriginal =   32
         EndProperty
      EndProperty
      Begin VB.Label lblPage0
         Caption         =   "第0页"
         Height          =   255
         Left            =   240
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblPage1
         Caption         =   "第1页"
         Height          =   255
         Left            =   -74760
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblPage2
         Caption         =   "第2页"
         Height          =   255
         Left            =   -149760
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdInTab 
         Caption         =   "IT"
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "f1"
      Height          =   1215
      Left            =   3960
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
      Begin VB.CommandButton cmdInFrame 
         Caption         =   "IF"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "CtrlSSTab"
Option Explicit

Private gTabGot As Long
Private gFrameGot As Long

' P20-42: SSTab 复刻 (SysTabControl32, 不加载 TABCTL32.OCX)。
' 探针口径: ①设计期 Tabs/Tab/TabOrientation/TabStyle/TabsPerRow/WordWrap 原样进 RTL
' ②运行期改 Tab / Tabs / 四个外观属性后回读 ③Tab 是 **0 基** (与 VB6 集合的 1 基不同)。

Private Sub Form_Load()
    Debug.Print "TS0-CAP=" & Me.Caption & " CL=" & Me.Controls.Count

    ' 设计期值
    Debug.Print "TS1-TABS=" & SSTab1.Tabs
    Debug.Print "TS2-TAB=" & SSTab1.Tab
    Debug.Print "TS3-ORIENT=" & SSTab1.TabOrientation
    Debug.Print "TS4-STYLE=" & SSTab1.TabStyle
    Debug.Print "TS5-PERROW=" & SSTab1.TabsPerRow
    Debug.Print "TS6-WRAP=" & SSTab1.WordWrap

    ' 运行期切页 (0 基)
    SSTab1.Tab = 0
    Debug.Print "TS7-SET0=" & SSTab1.Tab
    SSTab1.Tab = 2
    Debug.Print "TS8-SET2=" & SSTab1.Tab
    ' 越界不改值
    SSTab1.Tab = 9
    Debug.Print "TS9-OOR=" & SSTab1.Tab

    ' 运行期改外观
    SSTab1.TabOrientation = 1
    Debug.Print "TS10-ORIENT=" & SSTab1.TabOrientation
    SSTab1.TabStyle = 1
    Debug.Print "TS11-STYLE=" & SSTab1.TabStyle
    SSTab1.TabsPerRow = 4
    Debug.Print "TS12-PERROW=" & SSTab1.TabsPerRow
    SSTab1.WordWrap = -1
    Debug.Print "TS13-WRAP=" & SSTab1.WordWrap

    ' Task #44: TabToolTipText(i) 读写回环 (在 Tabs=3 状态下做, TS14/TS16 会改页数)。
    ' 未设过的页返回空串 —— SSTabEx 的属性语义是空, 不是报错。
    SSTab1.TabToolTipText(0) = "tip0"
    Debug.Print "TS31-TIP0=" & SSTab1.TabToolTipText(0)
    SSTab1.TabToolTipText(2) = "tip2"
    Debug.Print "TS32-TIP2=" & SSTab1.TabToolTipText(2)
    Debug.Print "TS33-TIP1=" & SSTab1.TabToolTipText(1)

    ' 改页数: 先扩到 5, 再缩到 2 (缩容时活动页要跟着落回范围内)
    SSTab1.Tabs = 5
    Debug.Print "TS14-TABS5=" & SSTab1.Tabs
    Debug.Print "TS15-TABAFTERGROW=" & SSTab1.Tab
    SSTab1.Tabs = 2
    Debug.Print "TS16-TABS2=" & SSTab1.Tabs
    Debug.Print "TS17-TABAFTERSHRINK=" & SSTab1.Tab

    Debug.Print "TS18-CAP0=" & SSTab1.TabCaption(0)
    SSTab1.TabCaption(0) = "改过"
    Debug.Print "TS19-CAP0B=" & SSTab1.TabCaption(0)
    Debug.Print "TS20-VIS1=" & SSTab1.TabVisible(1)
    SSTab1.TabVisible(1) = 0
    Debug.Print "TS21-VIS1B=" & SSTab1.TabVisible(1)
    ' 容器语义: 页码由设计期 Left 反推 (VB6 把非活动页 Left 减 75000 缇)
    Debug.Print "TS22-P0LEFT=" & lblPage0.Left
    Debug.Print "TS23-P1LEFT=" & lblPage1.Left
    Debug.Print "TS24-P2LEFT=" & lblPage2.Left
    Debug.Print "CTRLSSTAB-DONE"
End Sub

' 切页显隐必须等窗体真正显示之后再断言: vb6_GetControlVisible 走 IsWindowVisible,
' 沿父链传播 —— Form_Load 阶段连 SSTab1.Visible 都是 0 (DoEvents 也不够)。
' Timer 从消息循环里触发, 那时窗体已经显示, Visible 才是切页的真实结果。
Private Sub tmrCheck_Timer()
    tmrCheck.Enabled = False
    Debug.Print "TS25-TABVIS=" & SSTab1.Visible
    ' 前面的断言把状态改过了 (TS16 缩到 2 页 / TS21 藏了第 1 页), 这里先恢复,
    ' 否则切到 1 会被"不能切到隐藏页"挡下, 切到 2 又越界 —— 看着像显隐失效。
    SSTab1.Tabs = 3
    SSTab1.TabVisible(1) = -1
    Debug.Print "TS25B-TABS=" & SSTab1.Tabs & " VIS1=" & SSTab1.TabVisible(1)
    SSTab1.Tab = 0
    Debug.Print "TS26-AT0-P0VIS=" & lblPage0.Visible & " P1VIS=" & lblPage1.Visible & " P2VIS=" & lblPage2.Visible
    SSTab1.Tab = 1
    Debug.Print "TS27-AT1-P0VIS=" & lblPage0.Visible & " P1VIS=" & lblPage1.Visible & " P2VIS=" & lblPage2.Visible
    SSTab1.Tab = 2
    Debug.Print "TS28-AT2-P0VIS=" & lblPage0.Visible & " P1VIS=" & lblPage1.Visible & " P2VIS=" & lblPage2.Visible
    Debug.Print "CTRLSSTAB-VISDONE"

    ' Click(PreviousTab): 不需要合成消息 —— TabCtrl_SetCurSel 自己就会给父窗口发
    ' TCN_SELCHANGE (走 WM_NOTIFY), 与用户点标签时 comctl32 发的是同一条。
    ' 所以 `SSTab1.Tab = …` 就会触发 Click, 参数 = 切换前的页号。
    SSTab1.Tab = 2
    Debug.Print "TS29-SETTAB2=" & SSTab1.Tab
    SSTab1.Tab = 0
    Debug.Print "TS29B-SETTAB0=" & SSTab1.Tab
    Debug.Print "TS34=" & CStr(SSTab1.WordWrap) & "/" & TypeName(SSTab1.WordWrap) & "/" & VarType(SSTab1.WordWrap)
Debug.Print "CTRLSSTAB-CLICKDONE"

    ' ---- 账 #167 的两步证人（活动页=0，cmdInTab 在这一页上）----
    Dim d1t As Long, d1f As Long
    cmdInFrame.SetFocus        ' 先把焦点安在一个已知落点：#157 的初始焦点可能正好就是 cmdInTab
    DoEvents
    d1t = gTabGot: d1f = gFrameGot
    cmdInTab.SetFocus
    DoEvents
    Debug.Print "TS35-ARM-TAB=" & CStr(gTabGot - d1t) & "/" & CStr(gFrameGot - d1f)
    d1t = gTabGot: d1f = gFrameGot
    cmdInFrame.SetFocus
    DoEvents
    Debug.Print "TS36-ARM-FRAME=" & CStr(gTabGot - d1t) & "/" & CStr(gFrameGot - d1f)
    Unload Me
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Debug.Print "TS30-CLICK-PREV=" & PreviousTab & " NOW=" & SSTab1.Tab
End Sub

' 账 #167 的证人：**容器清单只许有一处权威**。收成单一 `controlIsContainerType` 之前，创建侧
' 认 SSTab 而派发侧只认 Frame/PictureBox ⇒ 同一枚子控件在两侧数到不同的 id，
' `cmdInTab` 的 BN_SETFOCUS 带着创建侧那个号，去派发侧找不到自己的 arm。
' 这里 BASE 实测两步都是 `0/0`（号整体错开，谁也没接上）；在容器嵌得少一格的探针工程里同一
' 条缺陷读出来是 `0/1` —— 号正好撞上**别人的** arm，跑的是别人的 `_GotFocus`。两种都算这一族。
' 两步都钉：第二步反向给 `cmdInFrame` 焦点 —— 若两边全指到 cmdInTab 也一样红。
' 读数只报增量：账 #157 那条初始焦点会自己发一次 GotFocus，先手落点不是这条要问的东西。
Private Sub cmdInTab_GotFocus()
    gTabGot = gTabGot + 1
End Sub

Private Sub cmdInFrame_GotFocus()
    gFrameGot = gFrameGot + 1
End Sub
