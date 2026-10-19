' Ledger B181 (knife 71) regression fixture. StrPtr / StrConv of a Byte() array reached through
' a ByRef FORMAL parameter used to hand the SafeArray DESCRIPTOR address to a Win32 formal typed
' BSTR, and cl only warned, so nothing was loud. Two independent reasons, both measured on the
' previous build (x64 and x86, this file's own shapes): the byte-array name set had NO
' formal-parameter registration path (module-level and local arrays registered, parameters did
' not), and the "(*x)" unwrap refused anything 4 characters long, so a one-letter formal never
' reached the lookup at all. BASE readings were: formal pointer <> caller pointer, delta non-zero,
' and lstrlenW of the formal pointer read 3 instead of 2 (it was reading the descriptor header).
'
' After the fix the three paths answer the same way. Readings must be identical on x64 and x86.
' Nothing here pins a pointer VALUE - only that the paths agree, that the data behind the pointer
' reads back as UTF-16 "AB" (length 2), and that the pointer is not NULL.
Option Explicit

Private Const BV_APP As String = "C3ProbeB181"

Declare Function lstrlenW Lib "kernel32" (ByVal p As LongPtr) As Long

Private m_Module() As Byte

Private Sub FillArray(ByRef data() As Byte)
    ReDim data(5)
    data(0) = 65
    data(1) = 0
    data(2) = 66
    data(3) = 0
    data(4) = 0
    data(5) = 0
End Sub

' one-letter formal: the shape that used to fail the length gate
Private Function SamePtrShortName(ByRef b() As Byte, ByVal want As LongPtr) As String
    If StrPtr(b) = want Then SamePtrShortName = "Y" Else SamePtrShortName = "N"
End Function

Private Function LenShortName(ByRef b() As Byte) As Long
    LenShortName = lstrlenW(StrPtr(b))
End Function

' long formal: the shape that used to fail because parameters were never registered
Private Function SamePtrLastName(ByRef data() As Byte, ByVal want As LongPtr) As String
    If StrPtr(data) = want Then SamePtrLastName = "Y" Else SamePtrLastName = "N"
End Function

Private Function LenLastName(ByRef data() As Byte) As Long
    LenLastName = lstrlenW(StrPtr(data))
End Function

Private Function LenModuleArray() As Long
    LenModuleArray = lstrlenW(StrPtr(m_Module))
End Function

' Never called. These two hold the shapes the emit needles read: the sibling station (StrConv)
' redirecting a formal, and the negative control that a NON-byte formal keeps the plain call.
' StrConv's own semantics for Byte() with embedded NULs is a different account, so no behaviour
' pin is put on it here.
Private Function StrConvWitness(ByRef data() As Byte) As String
    StrConvWitness = StrConv(data, vbUnicode)
End Function

Private Function PlainTextWitn(ByRef txt() As String) As Long
    PlainTextWitn = lstrlenW(StrPtr(txt(0)))
End Function

Public Sub Main()
    Dim buf() As Byte
    Dim pCaller As LongPtr

    FillArray buf
    FillArray m_Module
    pCaller = StrPtr(buf)

    Debug.Print BV_APP & "-P1-short-formal-same=" & SamePtrShortName(buf, pCaller)
    Debug.Print BV_APP & "-P2-long-formal-same=" & SamePtrLastName(buf, pCaller)
    Debug.Print BV_APP & "-P3-short-formal-len=" & CStr(LenShortName(buf))
    Debug.Print BV_APP & "-P4-long-formal-len=" & CStr(LenLastName(buf))
    Debug.Print BV_APP & "-P5-module-len=" & CStr(LenModuleArray())
    Debug.Print BV_APP & "-P6-caller-nonzero=" & CStr(pCaller <> 0)
    Debug.Print BV_APP & "-DONE"
End Sub
