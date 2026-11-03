VERSION 5.00
Begin VB.MDIForm IdMdi 
   BackColor       =   &H8000000C&
   Caption         =   "ID-MDI"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer idTimer 
      Interval        =   250
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "IdMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' M3 第三百六x刀: 接收者身份的最小复现夹具 (不依赖 GUISample 源码)。
'
' 三条身份线:
'   ① `Set p = Me` 之后的 `p.Caption` 读/写 —— 晚绑定接收者必须是宿主对象本身;
'   ② MDIForm 的 `ActiveForm` —— 必须拿到当前活动子窗体 (不是 Empty, 也不许报
'      C3_HOST_MEMBER_MISS: "认得但无值" 与 "不认识这个成员" 是两件事);
'   ③ `ActiveForm.Caption` 的读/写 —— 拿到子窗体之后属性链要真能落到窗口上。
' 判据只看运行期读数 (Debug.Print), 不看日志里"没有 miss"。

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Sub MDIForm_Load()
    Dim c As IdChild
    Set c = New IdChild
    c.Show
    Debug.Print "CH0=" & TF(Not (c Is Nothing))
    Debug.Print "CH1=" & CStr(c.Caption)
    Debug.Print "CH2=" & TF(Not (ActiveForm Is Nothing))
End Sub

Private Sub idTimer_Timer()
    Static done As Integer
    If done Then Exit Sub
    done = 1

    ' --- ID1..ID2: `Set p = Me` 后晚绑定读写 Caption ---
    Dim p As Object
    Set p = Me
    p.Caption = "ID-SETP"
    Debug.Print "ID1=" & TF(Me.Caption = "ID-SETP")
    Debug.Print "ID2=" & TF(CStr(p.Caption) = "ID-SETP")

    ' --- ID3..ID6: ActiveForm 身份 ---
    Debug.Print "ID3=" & TF(Not (ActiveForm Is Nothing))
    If Not (ActiveForm Is Nothing) Then
        Debug.Print "ID4=" & TF(ActiveForm.Caption = "ID-CHILD")
        Debug.Print "ID5a=" & CStr(TypeName(Me))
        Debug.Print "ID5b=" & CStr(TypeName(ActiveForm))
        Debug.Print "ID5=" & TF(TypeName(ActiveForm) = "IdChild")
        ActiveForm.Caption = "ID-AF"
        Debug.Print "ID6=" & TF(ActiveForm.Caption = "ID-AF")
    End If

    Debug.Print "ID-DONE"
    Unload Me
End Sub
