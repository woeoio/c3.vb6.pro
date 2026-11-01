Option Explicit

' Fixture for the type-authority check: the SAME VB type written two ways must map to
' ONE C type. OLE_COLOR is a name the compiler knows (COM alias), so the library-
' qualified spelling stdole.OLE_COLOR has to fold onto the same int32_t. StdFont is a
' genuinely external type here (no project symbol), so BOTH spellings stay void* -
' that side is the guard against folding too much (a real COM type must not be pulled
' native just because someone wrote the prefix). ASCII comments on purpose: .bas/.frm
' sources are read as ANSI by the driver.

Public Sub BareColor(ByVal c As OLE_COLOR)
    Debug.Print c
End Sub

Public Sub QualColor(ByVal c As stdole.OLE_COLOR)
    Debug.Print c
End Sub

Public Sub BareFont(ByVal f As StdFont)
    Debug.Print f.Name
End Sub

Public Sub QualFont(ByVal f As stdole.StdFont)
    Debug.Print f.Name
End Sub

' Second family (knife 267): a dotted name's prefix is WHERE the type lives, not part
' of the type. When the library itself is spelled with a leading "vb" (VBRUN = the
' ComCtls runtime type library), the "starts with Vb => enum => Long" heuristic in
' TypeSystem::resolveTypeName swallowed the whole string, so the qualified spelling
' folded a pointer-returning getter into int32_t - on x64 that truncates the
' collection address, the RTL stops recognising it by pointer identity, and For Each
' calls IDispatch on the garbage. Bare and qualified spellings must answer the SAME
' way (void*). The two answers that come from whole-string rules must NOT be re-folded
' by the leaf retry: an *Enum suffix stays Long, an MSDATASRC.* member stays String.

Public Function BareCollEnum() As ParentControls
    Set BareCollEnum = Nothing
End Function

Public Function QualCollEnum() As VBRUN.ParentControls
    Set QualCollEnum = Nothing
End Function

Public Sub TakeQualColl(ByVal c As VBRUN.ContainedControls)
    Debug.Print c Is Nothing
End Sub

Public Function QualSheetEnum() As VBRUN.SheetTypeEnum
    QualSheetEnum = 3
End Function

Public Function QualDataMember() As MSDATASRC.DataMember
    QualDataMember = "x"
End Function
