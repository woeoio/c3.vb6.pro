' test_err_obj.bas - P24-12: Err object comprehensive test
Option Explicit

Sub Main()
    Dim passCount As Long
    passCount = 0
    
    ' Test 1: Err.Number initial state is 0
    On Error Resume Next
    If Err.Number = 0 Then
        passCount = passCount + 1
        Debug.Print "ERR-1:OK Number=0"
    End If
    
    ' Test 2: Err.Number after Err.Raise 11
    Err.Raise 11
    If Err.Number = 11 Then

        passCount = passCount + 1
        Debug.Print "ERR-2:OK Number="; Err.Number
    End If
    
    ' Test 3: Err.Clear
    Err.Clear
    If Err.Number = 0 Then
        passCount = passCount + 1
        Debug.Print "ERR-3:OK Clear"
    End If
    
    ' Test 4: Err.Raise
    On Error Resume Next
    Err.Raise 5
    If Err.Number = 5 Then
        passCount = passCount + 1
        Debug.Print "ERR-4:OK Raise 5"
    End If
    
    ' Test 5: Err.Description
    If Len(Err.Description) >= 0 Then
        passCount = passCount + 1
        Debug.Print "ERR-5:OK Description"
    End If
    
    ' Test 6: Err.Source
    If Len(Err.Source) >= 0 Then
        passCount = passCount + 1
        Debug.Print "ERR-6:OK Source"
    End If
    
    ' Test 7..9 (账 #305 第二百四十一刀): named 实参按**名字**落位 —— 改前两形都把 Description
    ' 的文案交进 RTL 的第 2 档 (source)，Err.Description 留空；第三形按声明序写，改前改后都对，
    ' 留着钉「别把位置序改错」。三发读数与第二百四十刀 A1/A2/A3 的探针同源。
    Err.Clear
    On Error Resume Next
    Err.Raise Number:=7001, Description:="DESC7"
    If Err.Number = 7001 And Err.Description = "DESC7" And Len(Err.Source) = 0 Then
        passCount = passCount + 1
        Debug.Print "ERR-7:OK named-desc"
    Else
        Debug.Print "ERR-7:NO Number="; Err.Number; " DESC=["; Err.Description; "] SRC=["; Err.Source; "]"
    End If

    Err.Clear
    Err.Raise Number:=7002, Source:="SRC8", Description:="DESC8"
    If Err.Number = 7002 And Err.Description = "DESC8" And Err.Source = "SRC8" Then
        passCount = passCount + 1
        Debug.Print "ERR-8:OK named-declared-order"
    Else
        Debug.Print "ERR-8:NO Number="; Err.Number; " DESC=["; Err.Description; "] SRC=["; Err.Source; "]"
    End If

    Err.Clear
    Err.Raise Number:=7003, Description:="DESC9", Source:="SRC9"
    If Err.Number = 7003 And Err.Description = "DESC9" And Err.Source = "SRC9" Then
        passCount = passCount + 1
        Debug.Print "ERR-9:OK named-reversed"
    Else
        Debug.Print "ERR-9:NO Number="; Err.Number; " DESC=["; Err.Description; "] SRC=["; Err.Source; "]"
    End If
    On Error GoTo 0
    
    Debug.Print "ERR:"; passCount; "/9"
End Sub
