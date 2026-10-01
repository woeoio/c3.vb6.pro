' <vbeclipse>: vbTextCompare 对表夹具。修复前实测: RTL 里 5 处 (void)compare 把
' compare 静默丢掉 (StrComp/Replace/Split/Filter/InStrRev 一律按二进制比), 而
' InStr 的第 4 参被 codegen 截掉、"字符串优先的三参形" 三个实参错接到
' (start, haystack, needle) 三槽 ⇒ needle 收到整数 1 → 解引用地址 1 → 段错误。
' 判据载荷全 ASCII ⇒ 结果与码页无关 (ai/029 的 ASCII 载荷纪律)。
Option Explicit

Public Sub Main()
    Dim s As String
    s = "Hello World, hello WORLD"
    Debug.Print "TC-instr-b="; InStr(s, "world")
    Debug.Print "TC-instr-t="; InStr(s, "world", vbTextCompare)
    Debug.Print "TC-instr4-t="; InStr(8, s, "world", vbTextCompare)
    Debug.Print "TC-instr4-b="; InStr(8, s, "world", vbBinaryCompare)
    Debug.Print "TC-revb="; InStrRev(s, "hello")
    Debug.Print "TC-revt="; InStrRev(s, "hello", -1, vbTextCompare)

    Debug.Print "TC-rep-b=["; Replace("aAaA", "a", "-", 1, -1, vbBinaryCompare); "]"
    Debug.Print "TC-rep-t=["; Replace("aAaA", "a", "-", 1, -1, vbTextCompare); "]"

    Dim sp As Variant
    sp = Split("aXbXc", "x", -1, vbBinaryCompare)
    Debug.Print "TC-split-b-ub="; UBound(sp)
    sp = Split("aXbXc", "x", -1, vbTextCompare)
    Debug.Print "TC-split-t-ub="; UBound(sp); " ["; Join(sp, "|"); "]"

    Dim src As Variant
    src = Array("abc", "xyz", "ABD")
    Dim fb As Variant
    Dim ft As Variant
    fb = Filter(src, "AB", True, vbBinaryCompare)
    ft = Filter(src, "AB", True, vbTextCompare)
    Debug.Print "TC-filter-b-ub="; UBound(fb)
    Debug.Print "TC-filter-t-ub="; UBound(ft); " ["; Join(ft, "|"); "]"

    Debug.Print "TC-sc-b="; StrComp("a", "B", vbBinaryCompare)
    Debug.Print "TC-sc-t="; StrComp("a", "B", vbTextCompare)
    Debug.Print "TC-sc-eq="; StrComp("AB", "ab", vbTextCompare)
    Debug.Print "TC-DONE"
End Sub
