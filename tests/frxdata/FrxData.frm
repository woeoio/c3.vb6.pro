VERSION 5.00
Begin VB.Form FrxData 
   Caption         =   "FrxData"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   1440
      ItemData        =   "FrxData.frx":0000
      Left            =   360
      List            =   "FrxData.frx":0010
      TabIndex        =   1
      Top             =   1680
      Width           =   2400
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrxData.frx":0000
      Left            =   3120
      List            =   "FrxData.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   1200
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrxData.frx":0026
      Top             =   240
      Width           =   2400
   End
End
Attribute VB_Name = "FrxData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clickN As Long
Dim ok8 As Boolean

' Fix 195 回归: .frx 三种 blob (字符串 / List / ItemData) 的真实布局。
' 旧 readIntList 按"每项 2B 整数"读 ItemData, 读到的是结构的字节本身 ——
' 任何工程都解出恒定假值 1/304/12288..., 设计期 ItemData 编进 exe 一直是垃圾。
' 本夹具的 ItemData 故意取 5/300/-7 三个非默认值: 旧实现必错。
' 中国字串用以校验 GBK 解码: 与 .frm 源码里的字面量比对, 输出只有 ASCII。

Private Function TF(b As Boolean) As String
    If b Then TF = "True" Else TF = "False"
End Function

Private Sub Combo1_Click()
    clickN = clickN + 1
End Sub

Private Sub Form_Load()
    Debug.Print "FD1=" & Replace(Text1.Text, vbCrLf, "|")
    Debug.Print "FD2=" & CStr(List1.ListCount)
    Debug.Print "FD3=" & List1.List(0)
    Debug.Print "FD4=" & CStr(List1.ItemData(0))
    Debug.Print "FD5=" & CStr(List1.ItemData(1))
    Debug.Print "FD6=" & CStr(List1.ItemData(2))
    If List1.List(2) = "你好" Then
        Debug.Print "FD7=OK"
    Else
        Debug.Print "FD7=BAD"
    End If
    ok8 = (Combo1.ListCount = 3) And (Combo1.List(0) = "1234") And _
          (Combo1.ItemData(0) = 5) And (Combo1.ItemData(2) = -7)
    Debug.Print "FD8-COMBO=" & TF(ok8)
    Debug.Print "FD9-COMBO-RAW count=" & CStr(Combo1.ListCount) & " item0=" & Combo1.List(0) & _
                " id0=" & CStr(Combo1.ItemData(0)) & " id2=" & CStr(Combo1.ItemData(2))
    clickN = 0
    Combo1.ListIndex = 2
    Debug.Print "FD10-COMBO-LI-RAW li=" & CStr(Combo1.ListIndex) & " text=" & Combo1.Text & _
                " clicks=" & CStr(clickN)
    Debug.Print "FRXDATA-DONE"
    Unload Me
End Sub
