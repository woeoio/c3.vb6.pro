Option Explicit

' Account #611: VB6 TypeName(CreateObject("Scripting.Dictionary")) = "Dictionary".
' C3 used to answer "Object" for every external IDispatch, which made
' JSonWrite.Serialize reject it with "Unsupported object type: Object"
' (JSonFast button 7 killed the process on click).
' Only objects created by CreateObject are tagged; other external COM
' objects still answer "Object" so existing corpus stays unchanged.
Sub Main()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Debug.Print "tn=" & TypeName(d)
    d.Add "k", "v"
    Debug.Print "item=" & d("k")
End Sub
