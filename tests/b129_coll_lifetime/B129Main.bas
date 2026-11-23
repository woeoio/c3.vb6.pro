Option Explicit

Public gTerms As Long

' 载体寿命三岔: (甲) Set Nothing / (乙) 重新赋值丢掉旧载体 / (丙) 两枚槽指向同一枚载体。
' B129Item 只在引用真退到 0 时出厂 (Class_Terminate 加一), 所以每一条 terms 都是
' 「载体到 0 时把元素一起放开」这一件事的读数。
Private Sub CaseNothing()
    Dim c As Collection
    Set c = New Collection
    c.Add New B129Item
    c.Add New B129Item
    Debug.Print "B129-A-count=" & c.Count
    Set c = Nothing
End Sub

Private Sub CaseReassign()
    Dim c As Collection
    Set c = New Collection
    c.Add New B129Item
    Set c = New Collection
    c.Add New B129Item
    Set c = Nothing
End Sub

Private Sub CaseAlias()
    Dim c As Collection
    Dim a As Collection
    Set c = New Collection
    c.Add New B129Item
    Set a = c
    Set c = Nothing
    Debug.Print "B129-B-alive=" & a.Count
    Set a = Nothing
End Sub

Private Sub CaseNest()
    Dim inner As Collection
    Dim outer As Collection
    Set inner = New Collection
    Set outer = New Collection
    inner.Add New B129Item
    outer.Add inner
    Set inner = Nothing
    Debug.Print "B129-C-inner=" & outer(1).Count
    Set outer = Nothing
End Sub

Sub Main()
    CaseNothing
    Debug.Print "B129-A1-terms=" & gTerms
    CaseReassign
    Debug.Print "B129-A2-terms=" & gTerms
    CaseAlias
    Debug.Print "B129-B-terms=" & gTerms
    CaseNest
    Debug.Print "B129-DONE"
End Sub