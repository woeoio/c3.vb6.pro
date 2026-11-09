VERSION 5.00
Begin VB.UserControl UCProbe 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ScaleHeight     =   600
   ScaleWidth      =   1200
   Begin VB.Label Label1 
      Caption         =   "L1"
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "UCProbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' PD-Parent 语义刀: 三层 Parent 里 UC 侧的两层。
' 真 VB6 6.00.9782 oracle (.build/oracle/parent) 实测:
'   Parent.Name/Caption/TypeName → 顶层窗体 ("Form1" / "HostForm" / "Form1")
'   Label1.Parent Is Me          → True
'   Label1.Parent.Name           → 设计器类名 "UCProbe" (不是实例名)
Public Sub Dump(ByVal tag As String)
    Dim cHost As Long
    Dim cInner As Long
    Debug.Print tag & "-PARENT-NAME=" & Parent.Name
    Debug.Print tag & "-PARENT-CAPTION=" & Parent.Caption
    Debug.Print tag & "-PARENT-TYPENAME=" & TypeName(Parent)
    Debug.Print tag & "-INNER-PARENT-IS-ME=" & CStr(Label1.Parent Is Me)
    Debug.Print tag & "-INNER-PARENT-NAME=" & Label1.Parent.Name
    Debug.Print tag & "-INNER-PARENT-TYPENAME=" & TypeName(Label1.Parent)
    cHost = Parent.BackColor
    cInner = Label1.Parent.BackColor
    Debug.Print tag & "-PARENT-BACKCOLOR-IS-INNER=" & CStr(cHost = cInner)
End Sub
