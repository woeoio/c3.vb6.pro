Attribute VB_Name = "HostDecimal"
Option Explicit

' Acct #305 knife 54 (B165): the host-store leg used to drop Decimal/I8/Date/Currency into
' default => Empty, so a value put into a Collection came back as 0/Empty. The store now asks
' the two ring authorities, so a stored Decimal survives on BOTH arches (x86 keeps DECIMAL as
' a pointer in our slot, which is why borrowing alone read back garbage there).
' The array / ByRef legs are deliberately NOT taken in this knife: they still answer Empty,
' and B165-arr is pinned as the negative witness for that.
Sub Main()
    Dim c As Collection
    Set c = New Collection

    Dim v As Variant
    v = CDec(5)
    Debug.Print "B165-box=" & VarType(v) & "|" & TypeName(v) & "|" & CStr(v)
    c.Add v, "d"
    Debug.Print "B165-coll=" & VarType(c("d")) & "|" & TypeName(c("d")) & "|" & CStr(c("d"))

    Dim w As Variant
    w = CDec(7)
    c.Add w, "e"
    ' two live decimals: the second store must not take the first one's buffer away
    Debug.Print "B165-two=" & CStr(c("d")) & "/" & CStr(c("e"))

    Dim e2 As Variant
    For Each e2 In c
        Debug.Print "B165-fe-1=" & VarType(e2) & "|" & TypeName(e2) & "|" & CStr(e2)
        Exit For
    Next e2
    Dim last As Variant
    For Each e2 In c
        last = e2
    Next e2
    Debug.Print "B165-fe-2=" & VarType(last) & "|" & TypeName(last) & "|" & CStr(last)

    ' witnesses that must not move
    Dim s As Variant
    s = "abc"
    c.Add s, "s"
    Dim l As Variant
    l = 42
    c.Add l, "l"
    Debug.Print "B165-str=" & VarType(c("s")) & "|" & TypeName(c("s")) & "|" & CStr(c("s"))
    Debug.Print "B165-long=" & VarType(c("l")) & "|" & CStr(c("l"))

    ' release side: drop the first item, the survivors must still read their own values
    c.Remove 1
    Debug.Print "B165-rem=" & CStr(c.Count) & "/" & CStr(c("e")) & "/" & CStr(c("l"))

    ' The two other consumers of vb6_VariantCopy: assigning a Variant array element and
    ' copying a whole Variant array. Both used to share the x86 DECIMAL buffer with their
    ' source (Copy did not know that tier at all -> two owners of one 16-byte block).
    Dim av(0 To 1) As Variant
    av(0) = CDec(5)
    av(1) = CDec(7)
    Debug.Print "B165-elem=" & CStr(av(0)) & "/" & CStr(av(1))
    Dim bv() As Variant
    bv = av
    Debug.Print "B165-whole=" & CStr(bv(0)) & "/" & CStr(bv(1)) & "/" & CStr(av(0))

    ' deliberately unchanged leg (B159/B161 owns arrays across this ring)
    Dim a() As Long
    ReDim a(0 To 1)
    a(0) = 11
    Dim arrV As Variant
    arrV = a
    c.Add arrV, "ar"
    Debug.Print "B165-arr=" & VarType(c("ar")) & "|" & TypeName(c("ar"))

    Debug.Print "B165-DONE"
End Sub
