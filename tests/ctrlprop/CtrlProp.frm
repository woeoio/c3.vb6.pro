VERSION 5.00
Begin VB.Form CtrlProp 
   Caption         =   "CtrlProp"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List2 
      Height          =   600
      Left            =   2400
      TabIndex        =   4
      Top             =   2160
      Width           =   2000
   End
   Begin VB.ListBox List1 
      Height          =   600
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2000
   End
   Begin VB.TextBox txtH 
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      Height          =   600
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   2000
   End
   Begin VB.TextBox txtV 
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      Height          =   600
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   2000
   End
   Begin VB.TextBox txtOne 
      Height          =   300
      ToolTipText     =   "dtxt"
      Left            =   240
      TabIndex        =   2
      Text            =   "T"
      Top             =   1680
      Width           =   2000
   End
   Begin VB.Label lblArr 
      Caption         =   "A1"
      Height          =   300
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   2000
   End
   Begin VB.Label lblArr 
      Caption         =   "A0"
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2000
   End
   Begin VB.Label lblSingle 
      Caption         =   "S"
      Tag             =   "dtagL"
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   2000
   End
End
Attribute VB_Name = "CtrlProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Fix 194 回归: 控件属性的字符串写入值必须转 BSTR, 且 List(j) 参与字符串相等比较
' 必须按 BSTR 处理 (RTL 里 vb6_GetListItem 的 C 返回类型是 void*)。
' 旧行为: ① lblArr(n).Caption = <数值> 生成的 C 把整数当 BSTR 指针 →
'            非 0 时 0xC000041D (用户回调未处理异常), 为 0 时静默不生效;
'         ② B = List2.List(J) 恒为假 (void* 被 _Generic 判成 VariantObject, CStr 得空串)。
Private Sub Form_Load()
    Dim I As Long
    Dim J As Long
    Dim B As String
    Dim hit As Long
    List1.AddItem "aaa"
    List1.AddItem "bbb"
    List2.AddItem "aaa"
    List2.AddItem "ccc"
    lblArr(0).Caption = List1.ListCount
    lblArr(1).Caption = List2.ListCount
    Debug.Print "CP1=" & lblArr(0).Caption
    Debug.Print "CP2=" & lblArr(1).Caption
    hit = 0
    For I = 0 To List1.ListCount - 1
        B = List1.List(I)
        For J = 0 To List2.ListCount - 1
            If B = List2.List(J) Then
                hit = hit + 1
                Exit For
            End If
        Next
    Next
    Debug.Print "CP3=" & CStr(hit)
    lblSingle.Caption = List1.ListCount
    txtOne.Text = List1.ListCount
    Debug.Print "CP4=" & lblSingle.Caption
    Debug.Print "CP5=" & txtOne.Text
    Debug.Print "CP6=" & txtH.ScrollBars
    Debug.Print "CP7=" & txtV.ScrollBars
    Debug.Print "CP8=" & txtOne.ScrollBars
    ' 账 #108/#107 两条一起钉：ScrollBars 的往返 + BorderStyle=None 设得上去
    ' （List1 是 ListBox，走的正是 vb6forms_style.c 那个"存窗口属性"兜底分支）
    List1.BorderStyle = 0
    Debug.Print "CP9=" & List1.BorderStyle
    List1.BorderStyle = 1
    Debug.Print "CP10=" & List1.BorderStyle
    List1.BorderStyle = 0
    Debug.Print "CP11=" & List1.BorderStyle
    ' 账 #142(C29-SL-e): 通用字符串属性 ToolTipText / Tag 的档位在**通用段**登记，
    ' 所以证人不能只有 Slider 一枚 —— 这里各来一枚 Label 与 TextBox（改之前这两条 CStr
    ' 打空、TypeName 答 "Object"，因为 RTL getter 是 `void*` 而装箱表把"其他指针"送对象）。
    lblSingle.ToolTipText = "ltt"
    txtOne.Tag = "xtt"
    Debug.Print "CP12=" & CStr(lblSingle.ToolTipText) & "/" & TypeName(lblSingle.ToolTipText)
    Debug.Print "CP13=" & CStr(txtOne.Tag) & "/" & TypeName(txtOne.Tag)
    If txtOne.Tag = "xtt" Then
        Debug.Print "CP14=yes"
    Else
        Debug.Print "CP14=no"
    End If
    Debug.Print "CP15=" & CStr(List1.ToolTipText) & "/end"
    ' 账 #142: 设计期那两条字符串属性的证人（通用段被**两条创建路**共用，非 Slider 的控件也得证一次）
    Debug.Print "CP16=" & CStr(lblSingle.Tag) & "/" & CStr(lblSingle.ToolTipText)
    Debug.Print "CP17=" & CStr(txtOne.ToolTipText) & "/" & CStr(txtOne.Tag)
    txtOne.ToolTipText = ""
    Debug.Print "CP18=" & CStr(txtOne.ToolTipText) & "/" & CStr(txtOne.Text)

    Debug.Print "CTRLPROP-DONE"
    Unload Me
End Sub
