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
