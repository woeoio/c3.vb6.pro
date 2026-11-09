Option Explicit

' Randomize 语句的三种形态：裸 / 正种 / 负种。
' 负种（`Randomize -1`）曾被通用语句路径当成中缀减法 ⇒ 发 `(vb6_Randomize - 1)`
' 并报一条假 VB3001（PhotoDemon pdPaletteChild.cls 的形态）。这里钉住两件事：
'   1) 三种形态都能编、都能跑；
'   2) 同一个种子两次投放得到的首个 Rnd 相同（种子确实传进去了）。

Public Sub Main()
    Dim a As Long, b As Long
    Dim c As Long, d As Long
    Dim ok As Boolean

    Randomize 5
    a = CLng(Rnd() * 1000)
    Randomize 5
    b = CLng(Rnd() * 1000)

    Randomize -1
    c = CLng(Rnd() * 1000)
    Randomize -1
    d = CLng(Rnd() * 1000)

    Randomize
    ok = (a = b) And (c = d)

    Debug.Print "RZ-OK=" & ok
End Sub
