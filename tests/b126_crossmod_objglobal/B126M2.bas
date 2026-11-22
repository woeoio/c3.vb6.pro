Option Explicit

Sub Main()
    PrepareIt
    Debug.Print "C3ProbeB126-A1=" & TypeName(gObj)
    gObj.Seed = 8
    Debug.Print "C3ProbeB126-A2=" & CStr(gObj.Seed)
    Debug.Print "C3ProbeB126-DONE"
End Sub
