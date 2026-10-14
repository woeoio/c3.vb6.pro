VERSION 5.00
Begin VB.Form CtrlStatusBar
   Caption         =   "CtrlStatusBar"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6000
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar StatusBar1
      Align          =   2                     ' AlignBottom
      Height         =   255
      Left           =   0
      Top            =   3045
      Width          =   6000
      _ExtentX       =   10583
      _ExtentY       =   450
      _Anchor        =   10
      Style          =   0                     ' sbrNormal
      Panels(1)      =   "Ready"
         .Key        =   "pr"
         .Style      =   0                     ' sbrText
         .AutoSize   =   1                     ' sbrSpring
         .MinWidth   =   40
      End
      Panels(2)      =   "Tip"
         .Key        =   "tp"
         .Style      =   2                     ' sbrNum
         .Width      =   120
         .MinWidth   =   20
         .AutoSize   =   0                     ' sbrFixed
         .ToolTipText=   "NumLock state"
      End
      Panels(3)      =   ""
         .Style      =   5                     ' sbrTime
         .AutoSize   =   2                     ' sbrContents
      End
   End
   Begin MSComctlLib.StatusBar StatusBar2
      Align          =   2                     ' AlignBottom
      Height         =   255
      Left           =   0
      Top            =   2790
      Width          =   6000
      _ExtentX       =   10583
      _ExtentY       =   450
      _Anchor        =   10
      Style          =   1                     ' sbrSimple
      SimpleText     =   "Simple text here"
   End
End
Attribute VB_Name = "CtrlStatusBar"
Option Explicit

' P20-40: StatusBar 复刻 (msctls_status32, 不加载 mscomctl.ocx)。
' 探针口径: ①设计期面板必须原样进 RTL (Panels 集合 + 各自的 Key/Text/Style/AutoSize/
' Width/MinWidth/ToolTipText) ②运行期 Add/Remove/Clear 对下标与 Key 表的搬动
' ③Align/Style/SimpleText 三个自身属性 ④时间/日期面板的文本由 RTL 自己算
' (SDK 10.0.19041.0 的 commctrl.h 里没有 SBT_TIME/SBT_DATE, 只能自己来)。

' prints a boolean as True/False (this fixture printed bare values only until SB36 moved)
Function SBTF(b As Boolean) As String
    If b Then SBTF = "True" Else SBTF = "False"
End Function

Private Sub Form_Load()
    Dim sbW0 As Long, sbW1 As Long, sbSum As Long, sbTol As Long
    Dim sbH0 As String
    ' SB0 asks whether a handle survives the stringify road: VB6 hands hWnd out as a
    ' Long, so CStr of it is digits. An empty string here meant the pointer went into
    ' the object box and came back as "".
    sbH0 = CStr(Me.hwnd)
    Debug.Print "SB0-HWND-NUM=" & SBTF(Len(sbH0) > 0) & " CAP=" & Me.Caption & " CL=" & Me.Controls.Count
    Debug.Print "SB1-COUNT=" & StatusBar1.Panels.Count
    Debug.Print "SB2-KEY1=" & StatusBar1.Panels(1).Key
    Debug.Print "SB3-TEXT1=" & StatusBar1.Panels(1).Text
    Debug.Print "SB4-STYLE1=" & StatusBar1.Panels(1).Style
    Debug.Print "SB5-AUTOSZ1=" & StatusBar1.Panels(1).AutoSize
    ' Design-block MinWidth is himetric like Object.Width; the getter answers twips -> 22
    Debug.Print "SB6-MINW1=" & StatusBar1.Panels(1).MinWidth
    Debug.Print "SB7-KEY2=" & StatusBar1.Panels(2).Key
    Debug.Print "SB8-TEXT2=" & StatusBar1.Panels(2).Text
    Debug.Print "SB9-STYLE2=" & StatusBar1.Panels(2).Style
    Debug.Print "SB10-W2=" & StatusBar1.Panels(2).Width
    Debug.Print "SB11-AUTOSZ2=" & StatusBar1.Panels(2).AutoSize
    Debug.Print "SB12-TIP2=" & StatusBar1.Panels(2).ToolTipText
    Debug.Print "SB13-STYLE3=" & StatusBar1.Panels(3).Style
    Debug.Print "SB14-IDXBYKEY=" & StatusBar1.Panels("tp").Index
    Debug.Print "SB15-TEXTBYKEY=" & StatusBar1.Panels("tp").Text

    ' 自身属性
    Debug.Print "SB16-ALIGN=" & StatusBar1.Align
    Debug.Print "SB17-STYLE=" & StatusBar1.Style
    Debug.Print "SB18-SIMPLE2=" & StatusBar2.SimpleText
    StatusBar2.SimpleText = "Changed"
    Debug.Print "SB19-SIMPLE2B=" & StatusBar2.SimpleText
    StatusBar2.Style = 0
    Debug.Print "SB20-STYLE2B=" & StatusBar2.Style

    ' 运行期改文本
    StatusBar1.Panels(1).Text = "Busy"
    Debug.Print "SB21-SETTEXT=" & StatusBar1.Panels(1).Text

    ' 运行期 Add
    Debug.Print "SB22-ADDIDX=" & StatusBar1.Panels.Add(4, "extra", "Extra")
    Debug.Print "SB23-COUNT2=" & StatusBar1.Panels.Count
    Debug.Print "SB24-NEWKEY=" & StatusBar1.Panels(4).Key
    Debug.Print "SB25-NEWTEXT=" & StatusBar1.Panels(4).Text

    ' 插到中间: 原第 4 项让位, 新项占位
    Debug.Print "SB26-INSERT=" & StatusBar1.Panels.Add(2, "ins", "Inserted")
    Debug.Print "SB27-COUNT3=" & StatusBar1.Panels.Count
    Debug.Print "SB28-IDXP2=" & StatusBar1.Panels(2).Key
    Debug.Print "SB29-IDXP3=" & StatusBar1.Panels(3).Key

    ' 按 Key 删
    StatusBar1.Panels.Remove "ins"
    Debug.Print "SB30-AFTERRM=" & StatusBar1.Panels.Count
    Debug.Print "SB31-P2KEY=" & StatusBar1.Panels(2).Key
    Debug.Print "SB32-COUNT4=" & StatusBar1.Panels.Count

    ' 按下标删
    StatusBar1.Panels.Remove 4
    Debug.Print "SB33-AFTERRM2=" & StatusBar1.Panels.Count
    Debug.Print "SB34-LASTKEY=" & StatusBar1.Panels(3).Key

    ' 设置类属性回读
    StatusBar1.Panels(1).MinWidth = 77
    Debug.Print "SB35-SETMINW=" & StatusBar1.Panels(1).MinWidth
    ' SB36 used to pin 123, i.e. "whatever you wrote is what you read back". From the 22nd
    ' cut (zhang 206, B41 face 2) this member answers the Laid-out geometry in twips, and
    ' panel 1 is a spring panel whose requested width is meaningless - so that number can
    ' only come from the window. Restated per #157: take the baseline in place and ask for
    ' the increment (writing a width into a spring panel must not move what it reports),
    ' never a hand-added total. The RAW line prints all four numbers.
    sbW0 = StatusBar1.Panels(1).Width
    StatusBar1.Panels(1).Width = 123
    sbW1 = StatusBar1.Panels(1).Width
    sbSum = StatusBar1.Panels(1).Width + StatusBar1.Panels(2).Width _
            + StatusBar1.Panels(3).Width
    Debug.Print "SB36-RAW before=" & sbW0 & " after=" & sbW1 & " sum=" & sbSum _
                & " bar=" & StatusBar1.Width
    Debug.Print "SB36-SPRING-FOLLOWS-WINDOW=" & SBTF((sbW1 > 0) And (sbW1 = sbW0))
    ' SB36-TILES-BAR = the existence witness for zhang 303 (B132), the defect the 22nd cut
    ' exposed by fixing the getter: SbLayout put a sbrContents tile's measured width into
    ' want[] but never into fixedTotal, so the spring was handed that width a second time and
    ' the last right edge ran past the client rect - BASE reads sum=6630 against bar=6000
    ' (the 630 twips = 42px clock text, counted twice), NEW must tile the bar.
    ' Tolerance, not a hand-added total: every tile is folded pixels->twips with rounding, so
    ' three tiles may differ from the bar's own width by a couple of twips (45 twips = 3px).
    sbTol = Abs(sbSum - StatusBar1.Width)
    Debug.Print "SB36-TILES-BAR=" & SBTF((sbSum > 0) And (sbTol <= 45))
    StatusBar1.Panels(1).AutoSize = 0
    Debug.Print "SB37-SETAUTOSZ=" & StatusBar1.Panels(1).AutoSize
    StatusBar1.Panels(1).ToolTipText = "hello"
    Debug.Print "SB38-SETTIP=" & StatusBar1.Panels(1).ToolTipText
    StatusBar1.Panels(1).Style = 6
    Debug.Print "SB39-SETSTYLE=" & StatusBar1.Panels(1).Style

    ' Clear
    StatusBar1.Panels.Clear
    Debug.Print "SB40-AFTERCLR=" & StatusBar1.Panels.Count
    Debug.Print "CTRLSTATUSBAR-DONE"
    Unload Me
End Sub
