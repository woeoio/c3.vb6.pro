Attribute VB_Name = "mdlK174"
Option Explicit

' <vbeclipse> fixture driver (knife 174 / account 471 station A):
' the loop variable is a legacy interface VIEW (value struct),
' the elements come back from the collection as our own wrappers.
' Before the fix every For Each line here was cl C2440 (Ray.vbp x14).

Public Function ViewValue(ByVal target As IK174) As Long
    ViewValue = target.Value()
End Function

Public Sub Main()
    Dim col As New Collection
    col.Add New K174Alpha
    col.Add New K174Beta
    col.Add New K174Plain
    col.Add 7
    Dim o As IK174
    Dim total As Long
    Dim skipped As Long
    For Each o In col
        If o Is Nothing Then
            skipped = skipped + 1
        Else
            total = total + ViewValue(o)
        End If
    Next
    Debug.Print "K174-sum="; total
    Debug.Print "K174-skipped="; skipped
    Debug.Print "K174-DONE"
End Sub
