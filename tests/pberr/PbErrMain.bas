Attribute VB_Name = "PbErrMain"
Option Explicit

' <vbeclipse> 账 #212 的判据驱动。六个读数全取真实输出（本地 x86+x64 逐行相同，
' `.build/b216out/pe/x64.out` 与 `x86.out`），真 VB6 给的是同一组数：
'   PE-1d / PE-2d / PE-null —— 错误 9 落回**抛出那枚实例方法自己**的处理器（不是越过它落到调用方）
'   PE-caller               —— 方法内没有处理器时，9 继续往上交回调用方的处理器
'   PE-state                —— 三次抛出之后那枚对象的成员照旧读得出 20-40（实例栈没留在中间态）
' 这一格没有产品改动：#209/#214 只把「裸读越界」换成「抛 9」，而抛 9 这一跳在实例方法里本来就通，
' 本夹子是把这件事钉住，别下一次改错误面时把它改漏。

Public Sub Main()
    Dim o As PbHold
    Dim r As Long

    Set o = New PbHold
    Debug.Print "PE-1d=" & o.OneDimAbove()
    Debug.Print "PE-2d=" & o.TwoDimAbove()
    Debug.Print "PE-null=" & o.NullDyn()

    On Error GoTo Caller
    r = o.NoHandler(9)
    Debug.Print "PE-caller=NOERR " & r
    GoTo Witness
Caller:
    Debug.Print "PE-caller=" & Err.Number
Witness:
    Debug.Print "PE-state=" & o.At(2) & "-" & o.At(4)
    Debug.Print "PE-DONE"
End Sub
