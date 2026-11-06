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

' ---- 追问三发（2026-11-02 补，第330刀之后仍在册的那两档 + 新发现的版本档）----
' P6: App.hInstance 答的是本 DLL 的 HMODULE 还是宿主进程的？
'     own= 那一发用 GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS|UNCHANGED_REFCOUNT,
'     AddressOf ProbeHost.Anchor, ...) 独立量一次「这段代码在哪枚模块」—— 值取自
'     libloaderapi.h:281-283（PIN=1 / UNCHANGED_REFCOUNT=2 / FROM_ADDRESS=4 ⇒ flags=6）。
'     same=True ⇒ hInstance 与自身模块同值；same=False 且 own 非 0 ⇒ VB6 的 App.hInstance 是别的什么。
Private Declare Function GetModuleHandleExW Lib "kernel32" (ByVal flags As Long, ByVal addr As Long, ByRef hModule As Long) As Long

Private Function OwnModule() As Long
    Dim h As Long
    OwnModule = 0
    If GetModuleHandleExW(6&, AddressOf ProbeHost.Anchor, h) <> 0 Then OwnModule = h
End Function

Public Function P6() As String
    Dim hi As Long, ow As Long
    hi = App.hInstance
    ow = OwnModule()
    P6 = "P6-hinst=[" & hi & "] own=[" & ow & "] same=" & CStr(hi = ow) & " nonzero=" & CStr(hi <> 0)
End Function

' P7: App.PrevInstance 在一枚 DLL 里答什么（VB6 手册只讲「本应用程序是否已有另一实例在跑」，
'     而 DLL 里的「本应用程序」到底指宿主还是指这枚组件，手册没说）。报错本身就是答案。
Public Function P7() As String
    On Error Resume Next
    Dim v As Variant
    v = App.PrevInstance
    P7 = "P7-prev=[" & CStr(v) & "] err=" & Err.Number
    On Error GoTo 0
End Function

' P8: App.Major / Minor / Revision / Build —— 本工程 .vbp 里写的是 MajorVer=7 / MinorVer=3 /
'     RevisionVer=11（故意不是 0，否则与「编译器根本没折」那一档分不开）。
'     C3 这一侧今天把这三档硬折成 0（见 cgen_expr_member_form_builtin.inc 的 major/minor/revision）。
Public Function P8() As String
    On Error Resume Next
    P8 = "P8-ver=[" & App.Major & "." & App.Minor & "." & App.Revision & "] build=[" & App.Build & "] err=" & Err.Number
    On Error GoTo 0
End Function
