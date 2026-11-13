Attribute VB_Name = "ModMain"
Option Explicit

Sub Main()
    Dim t As New ClsEat
    Dim s As String
    s = "AAA"
    t.Eat s
    ' Before fix: the caller's "AAA" block was already freed by the callee -- the
    ' next same-bucket (12B) allocation reuses it in free order, so the dangling
    ' s reads back "Q1". After fix: copy semantics, s is still "AAA".
    Dim a As String
    a = "Q1"
    Debug.Print "S=" & s
    ' Zero-regression arm: a ByRef string param must NOT get a copy -- ByRef
    ' semantics are aliasing by definition, and copying would break the link
    ' (the callee would no longer write back into the caller's variable).
    Dim r As String
    r = "BBB"
    ByRefEat r
    Debug.Print "R=" & r
End Sub

Public Sub ByRefEat(ByRef s As String)
    s = Trim(s) & "!"
End Sub
