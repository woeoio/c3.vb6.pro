Attribute VB_Name = "NullExtProbe"
Option Explicit
' ============================================================================
' NullExtProbe - manual VB6 IDE probe for the cells C3 still lacks readings for
'
' HOW TO USE
'   1. Open VB6, create a new Standard EXE project.
'   2. Project > Add Module > Existing... and pick this file.
'   3. Press Ctrl+G (Immediate window), type:   ProbeExt       then Enter.
'   4. Copy the whole output block back (it starts with "== NullExtProbe =="
'      and ends with "== DONE ==").
'
' WHY On Error Resume Next
'   Every cell is a single statement whose Err.Number is the reading. Without it
'   the IDE would break on the first 94/13 and the rest of the probe would never
'   run; with it the run is continuous and each cell reports its own error.
'
' WHAT IS BEING ASKED (open backlog cells)
'   A1..A9  `Null = ""` 这一类比较的**结果本身**是什么 (Null? False?)
'           —— C3 现在发 `vb6_StrCmp(...) == 0` ⇒ True / err=0，而 F7 真机是 94。
'   W4/W5/W6  Null 从"变体数组元素 / 返回 Null 的函数 / Collection.Item"进来时，
'           赋给 String 目标是不是也按 94 且目标不变。
'   C6..C8  Mid/Left/UCase 收 Null 时的错误号 (C6 在台账里是"未取")。
'   D1..D3  Len(<数>) 的返回档到底是 Integer 还是 Long (VB6 的 TypeName 说法)。
'
' FORMAT
'   <id>  <label>  | <readout>  err=<Err.Number>
'   - "err" 是**那一句**的 Err.Number (每格前 Err.Clear)。
'   - 赋给 s 的格子先把 s 置成 "keep"，好看出"报错后目标有没有被清空"。
' ============================================================================

Private nv As Variant          ' 装 Null 的变体 (与字面量 Null 同一件事, 便于逐句复查)
Private b As Boolean
Private d As Double
Private s As String
Private v As Variant
Private n As Long

Private Function FNull() As Variant
    FNull = Null
End Function

Private Function VStr(ByVal x As Variant) As String
    On Error Resume Next
    If IsNull(x) Then
        VStr = "<Null>"
        Exit Function
    End If
    VStr = CStr(x)
    If Err.Number <> 0 Then VStr = "<CStr err " & Err.Number & ">"
End Function

Private Function TN(ByVal x As Variant) As String
    On Error Resume Next
    TN = TypeName(x)
End Function

Private Sub P(ByVal id As String, ByVal label As String, ByVal x As Variant, ByVal e As Long)
    Debug.Print id & "  " & label & "  | isnull=" & IsNull(x) & " val=" & VStr(x) _
                & " tn=" & TN(x) & " err=" & e
End Sub

Public Sub ProbeExt()
    On Error Resume Next

    nv = Null
    b = True: d = 1.5: s = "keep": n = -1

    Debug.Print "== NullExtProbe =="
    Debug.Print "-- A. comparison result itself --"

    Err.Clear: v = Empty
    v = (nv = "")
    P "A1", "v = (Null = """")", v, Err.Number

    Err.Clear: d = 1.5
    d = (nv = "")
    Debug.Print "A2  d = (Null = """")        | d=" & d & " err=" & Err.Number

    Err.Clear: s = "keep"
    s = (nv = "")
    Debug.Print "A3  s = (Null = """")        | s=[" & s & "] err=" & Err.Number

    Err.Clear: b = True
    b = (nv = "")
    Debug.Print "A4  b = (Null = """")        | b=" & b & " err=" & Err.Number

    Err.Clear: v = Empty
    v = (nv <> "")
    P "A5", "v = (Null <> """")", v, Err.Number

    Err.Clear: v = Empty
    v = (nv = nv)
    P "A6", "v = (Null = Null)", v, Err.Number

    Err.Clear: v = Empty
    v = (nv = 0)
    P "A7", "v = (Null = 0)", v, Err.Number

    Err.Clear: v = Empty
    v = (nv > 0)
    P "A8", "v = (Null > 0)", v, Err.Number

    Err.Clear: s = "keep"
    s = (nv <> "")
    Debug.Print "A9  s = (Null <> """")       | s=[" & s & "] err=" & Err.Number

    Debug.Print "-- W. Null arriving through a container / function --"

    Dim a(1) As Variant
    a(0) = Null

    Err.Clear: s = "keep"
    s = a(0)
    Debug.Print "W4  s = a(0)   (Variant array elem) | s=[" & s & "] err=" & Err.Number

    Err.Clear: n = -1
    n = Len(a(0))
    Debug.Print "W4b n = Len(a(0))                  | n=" & n & " err=" & Err.Number

    Err.Clear: s = "keep"
    s = FNull()
    Debug.Print "W5  s = FNull() (func As Variant)  | s=[" & s & "] err=" & Err.Number

    Err.Clear: n = -1
    n = Len(FNull())
    Debug.Print "W5b n = Len(FNull())               | n=" & n & " err=" & Err.Number

    Dim c As New Collection
    c.Add Null

    Err.Clear: s = "keep"
    s = c(1)
    Debug.Print "W6  s = c(1)   (Collection.Item)   | s=[" & s & "] err=" & Err.Number

    Err.Clear: n = -1
    n = Len(c(1))
    Debug.Print "W6b n = Len(c(1))                  | n=" & n & " err=" & Err.Number

    Err.Clear: v = Empty
    v = a(0)
    P "W7", "v = a(0)  (sanity: should be Null)", v, Err.Number

    Debug.Print "-- C. string builtins fed with Null --"

    Err.Clear: s = "keep"
    s = Mid(nv, 1)
    Debug.Print "C6  s = Mid(Null, 1)               | s=[" & s & "] err=" & Err.Number

    Err.Clear: s = "keep"
    s = Left(nv, 1)
    Debug.Print "C7  s = Left(Null, 1)              | s=[" & s & "] err=" & Err.Number

    Err.Clear: s = "keep"
    s = UCase(nv)
    Debug.Print "C8  s = UCase(Null)                | s=[" & s & "] err=" & Err.Number

    Debug.Print "-- D. Len's own return type --"

    ' 两件事: (1) TypeName 只收**变量名**, 喂表达式就报「变量要求」;
    ' (2) VB6 里 12345 是 **Integer** 字面量 (≤32767), 要问台账里那格
    '     「Len(<Long>)」必须真给一个 Long 变量。两条都靠先落变量解决。
    Dim i1 As Integer
    Dim ln1 As Long
    Dim lv As Variant

    Err.Clear: i1 = 12345: lv = Empty
    lv = Len(i1)
    P "D1", "lv = Len(<Integer var>)", lv, Err.Number

    Err.Clear: ln1 = 12345: lv = Empty
    lv = Len(ln1)
    P "D2", "lv = Len(<Long var>)", lv, Err.Number

    Err.Clear: lv = Empty
    lv = Len("abc")
    P "D3", "lv = Len(""abc"")", lv, Err.Number

    Debug.Print "== DONE =="
End Sub
