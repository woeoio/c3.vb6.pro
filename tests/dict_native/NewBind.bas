Attribute VB_Name = "NewBind"
Option Explicit

' C30-10 第三十刀：`As New <外部 ProgID>` 的判据夹具。
' 覆盖四条路：局部/模块级 × VBScript.RegExp / Scripting.Dictionary。
' 期望（VB6 语义，真组件对照见 ai/039 §11.4）：
'   n1 local_re=True typename=RegExp
'   n2 module_re=True typename=RegExp
'   n3 local_dict=1 typename=Dictionary
'   n4 module_dict=1 typename=Dictionary

Public gRe As New VBScript.RegExp
Public gDict As New Scripting.Dictionary

Sub Main()
    Dim re As New VBScript.RegExp
    re.Pattern = "ab"
    Debug.Print "n1 local_re=" & CStr(re.Test("xaby")) & " typename=" & TypeName(re)

    gRe.Pattern = "cd"
    Debug.Print "n2 module_re=" & CStr(gRe.Test("xcdy")) & " typename=" & TypeName(gRe)

    Dim d As New Scripting.Dictionary
    d.Add "k", 1
    Debug.Print "n3 local_dict=" & CStr(d.Count) & " typename=" & TypeName(d)

    gDict.Add "k", 2
    Debug.Print "n4 module_dict=" & CStr(gDict.Count) & " typename=" & TypeName(gDict)
End Sub
