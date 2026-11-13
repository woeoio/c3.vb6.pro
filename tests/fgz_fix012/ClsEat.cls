VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ClsEat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ByVal string param that gets assigned in the body -- VB6 semantics: the callee
' works on its own copy.
' Before fix: the C signature BSTR p aliased the caller's string, and
' vb6_BSTR_AssignMove frees the old value first => it freed the caller's string
' (root cause of upstream FIX-012; the upstream project saw "*Model_Space" drawn).
Public Sub Eat(ByVal s As String)
    s = Trim(s)
End Sub
