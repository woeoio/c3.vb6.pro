Option Explicit

Public gTerms As Long
Public gColl As Collection

Private Function MakeColl() As Collection
    Set MakeColl = New Collection
End Function

Private Sub A_FnRet()
    With MakeColl()
        .Add New B136Item
        Debug.Print "B136-A1-count=" & .Count
    End With
    Debug.Print "B136-A2-terms=" & gTerms
End Sub

Private Sub B_FnRetExit()
    With MakeColl()
        .Add New B136Item
        Debug.Print "B136-B1-count=" & .Count
        Exit Sub
    End With
    Debug.Print "B136-B1b-unreachable"
End Sub

Private Sub C_FieldBorrow()
    With gColl
        .Add New B136Item
        Debug.Print "B136-C1-count=" & .Count
    End With
    Debug.Print "B136-C2-terms=" & gTerms
End Sub

Private Sub D_LocalBorrow()
    Dim c As Collection
    Set c = gColl
    With c
        .Add New B136Item
        Debug.Print "B136-D1-count=" & .Count
    End With
    Debug.Print "B136-D2-terms=" & gTerms
End Sub

Sub Main()
    Set gColl = New Collection
    gColl.Add New B136Item
    A_FnRet
    B_FnRetExit
    Debug.Print "B136-B2-terms=" & gTerms
    C_FieldBorrow
    D_LocalBorrow
    Debug.Print "B136-Z-terms=" & gTerms
    Debug.Print "B136-DONE"
End Sub