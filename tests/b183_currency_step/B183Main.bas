Option Explicit

' 账 #305 §11.163 (第一百八十刀): Currency 中间量**每一步都圈**到 4 位。
'
' 真 VB6 探针 .temp/k180p (P183.bas) 读数:
'   T1 c*2*1.0001            -> 5 Double 20.2490247   (Double 段不圈)
'   T2 c+0.00004             -> 6 Currency 10.1235   (末尾装箱圈一次)
'   T3 c*1.00005             -> 5 Double 10.124006175 (Double 段不圈)
'   T4 (c+0.00004)*2         -> 6 Currency 20.247     ★ 中间量先圈
'   T5 c*1.0001*1            -> 5 Double 10.12451235
'   T6 c/3*3                 -> 5 Double 10.1235
'   T7 CCur(c*2*1.0001)      -> 6 Currency 20.249
'   T8 c+0.00004+0.00004     -> 6 Currency 10.1235    ★ 每一步都圈
' 改前 C3: T4=20.2471 / T8=10.1236 (只在末尾圈一次); 其余相同。
'
' 每格四列: VarType | TypeName | CStr(x) | CStr(CLng(x*10000))

Public Sub Main()
    Dim c As Currency
    Dim v As Variant
    c = 10.1235

    v = c * 2 * 1.0001:        PrintV "T1", v
    v = c + 0.00004:           PrintV "T2", v
    v = c * 1.00005:           PrintV "T3", v
    v = (c + 0.00004) * 2:     PrintV "T4", v
    v = c * 1.0001 * 1:        PrintV "T5", v
    v = c / 3 * 3:             PrintV "T6", v
    v = CCur(c * 2 * 1.0001):  PrintV "T7", v
    v = c + 0.00004 + 0.00004: PrintV "T8", v
    Debug.Print "B183-DONE"
End Sub

Private Sub PrintV(ByVal tag As String, ByVal v As Variant)
    Debug.Print tag & "=" & VarType(v) & "|" & TypeName(v) & "|" & _
                  CStr(v) & "|" & CStr(CLng(v * 10000))
End Sub
