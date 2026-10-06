Option Explicit
' 账 #184: `AddressOf` 交出去的过程 = Win32 回调, 约定必须是 __stdcall —— VB6 的 AddressOf
' 给的就是 stdcall 调用桩。x86 下本体是 __cdecl, 把本体直接交给 OS 等于每次回调少弹 N*4
' 字节 (实测 VBFlexGrid 的 SUBCLASSPROC 有 6 个形参 ⇒ 每次消息派发漏弹 24 字节, 起窗即
' 堆损坏 0xC0000374, 本地 3/4 复现)。现在发码为被取址的过程另发一枚 __stdcall 桩。
'
' 判据三头 (全部与"桌面上有几枚窗口"无关 ⇒ 不抖):
'   AC1-tag   每次回调收到的 lParam 必须等于送出去的那个值 (参数编组没歪)
'   AC2-echo  同一进程里枚举两遍, 计数必须相同 (回调没被少调/多调)
'   AC3-hwnd  回调里 hWnd 永不为 0 (第一个形参也没歪)
'   AC4-stop  回调返回 0 ⇒ 最多只被调一次 (返回值按约定送到了调用方)
'   AC5-ret   EnumWindows 的返回值与"回调是否返回过 0"自洽 (不猜环境)
'   AC6-plain 没被 AddressOf 取址的过程不受影响
' 形状面 (桩存在 / AddressOf 站点取桩 / 本体仍是 cdecl) 由 Test-CodegenNote 在同名
' 夹具上钉 —— 见 run_tests.ps1 的 addrof_cb_stub_shape。
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Public gCalls As Long
Public gTagBad As Long
Public gZeroHwnd As Long
Public gStopCalls As Long

' 被 AddressOf 取址的 Private 过程 ⇒ 桩也发 static (只在本模块被取址)
Private Function CountWin(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
    gCalls = gCalls + 1
    If lParam <> 4242 Then gTagBad = gTagBad + 1
    If hWnd = 0 Then gZeroHwnd = gZeroHwnd + 1
    CountWin = 1                      ' 非 0 = 继续枚举
End Function

' 被 AddressOf 取址的 Public 过程 ⇒ 桩非 static (别的模块也可能取它的址)
Public Function StopWin(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
    gStopCalls = gStopCalls + 1
    StopWin = 0                       ' 0 = 调用方应当停下来
End Function

' 没被取址的过程: 本体约定与发码形状一个字都不该变
Public Function PlainAdd(ByVal n As Long) As Long
    PlainAdd = n + 1
End Function

Public Sub Main()
    Dim r1 As Long, r2 As Long, rStop As Long
    r1 = EnumWindows(AddressOf CountWin, 4242)
    Dim first As Long
    first = gCalls
    r2 = EnumWindows(AddressOf CountWin, 4242)
    rStop = EnumWindows(AddressOf StopWin, 4242)

    Dim sTag As String
    If gTagBad = 0 Then sTag = "Y" Else sTag = "N"
    Dim sEcho As String
    If first = gCalls - first Then sEcho = "Y" Else sEcho = "N"
    Dim sHwnd As String
    If gZeroHwnd = 0 Then sHwnd = "Y" Else sHwnd = "N"
    Dim sStop As String
    If gStopCalls <= 1 Then sStop = "Y" Else sStop = "N"
    Dim sRet As String
    ' 枚举"被回调打断"当且仅当返回 0 的回调真被调过
    If (rStop = 0) = (gStopCalls = 1) Then sRet = "Y" Else sRet = "N"
    Dim sPlain As String
    If PlainAdd(41) = 42 Then sPlain = "Y" Else sPlain = "N"

    Debug.Print "AC1-tag=" & sTag
    Debug.Print "AC2-echo=" & sEcho
    Debug.Print "AC3-hwnd=" & sHwnd
    Debug.Print "AC4-stop=" & sStop
    Debug.Print "AC5-ret=" & sRet
    Debug.Print "AC6-plain=" & sPlain
    Debug.Print "AC-CALLS=" & CStr(gCalls)
    Debug.Print "AC-DONE"
End Sub
