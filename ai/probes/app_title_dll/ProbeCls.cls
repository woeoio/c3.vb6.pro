VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Probe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' 账 #605 的口径探针：VB6 编出来的 ActiveX DLL 里，App.Title 到底答什么？
' 五条读数各答一问，全部用中括号夹住，空串也看得见于("P1-title=[]")。

Public Function P1() As String
    P1 = "P1-title=[" & App.Title & "]"
End Function

Public Function P2() As String
    P2 = "P2-len=" & Len(App.Title) & " eqEmpty=" & (App.Title = "")
End Function

' 对照：App.EXEName 答谁的名字 —— 它答 DLL  ⇒ App 指的是这枚 DLL；答 cscript ⇒ App 指的是宿主
Public Function P3() As String
    P3 = "P3-exename=[" & App.EXEName & "]"
End Function

' 对照：App.Path 答谁的目录
Public Function P4() As String
    P4 = "P4-path=[" & App.Path & "]"
End Function

' 对照：Title 在运行期可不可写（VB6 文档说 App.Title 是只读的；C3 那一路把写当成了对 g_vb6AppTitle 赋值）
Public Function P5() As String
    On Error Resume Next
    App.Title = "written-by-probe"
    P5 = "P5-after-write=[" & App.Title & "] err=" & Err.Number
    On Error GoTo 0
End Function
