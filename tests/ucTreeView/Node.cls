VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Selected As Boolean
Public Text As String
Public key As Variant
Public Data As Variant
Public SubText As String
Public Image As Variant
Public SelectedImage As Variant
Public Checked As Boolean
Public PlusButton As Boolean
Public Expanded As Boolean
Public FontBold As Boolean
Public FontUnderline As Boolean
Public TextColor As Long
Public BackColor As Long
Public RadioStyle As Boolean
Public HideCheck As Boolean
Public Disabled As Boolean
Public LinkStyle As Boolean
Public Badge As String
Public BadgeBackColor As String
Public BorderColor As Long
Public BorderWidth As Long
Public Childs As Collection
Public Parent As Node
Public IsChild As Boolean
Public ToolTipText As String
Public Height As Long
Public TextWidth As Long
Public Left As Long
Public Top As Long
Public Right As Long
Public Bottom As Long

Public Sub Clear()
    Set Childs = New Collection
End Sub

Public Sub Remove(index)
    Childs.Remove index
End Sub

Public Function Count() As Long
    Count = Childs.Count
End Function

Public Function index() As Long
    Dim i As Long
    If Parent Is Nothing Then Exit Function
    
    For i = 1 To Parent.Childs.Count
        If Parent.Childs(i) Is Me Then
            index = i
            Exit Function
        End If
    Next
End Function

Public Function Move(index As Integer, NewPosition As Integer)
    Dim tempNode As Node
    
    If index < 1 Or index > Childs.Count Then
        Err.Raise vbObjectError + 513, , "Índice fuera de rango: " & index
        Exit Function
    End If
    
    If NewPosition < 1 Or NewPosition > Childs.Count + 1 Then
        Err.Raise vbObjectError + 514, , "Nueva posición fuera de rango: " & NewPosition
        Exit Function
    End If
    
    Set tempNode = Me.Childs(index)

    Childs.Remove index

    If NewPosition <= Childs.Count Then
        Childs.Add tempNode, tempNode.key, NewPosition
    Else
        Childs.Add tempNode, tempNode.key
    End If
End Function

Public Function Switch(Index1 As Integer, Index2 As Integer)
    Dim TempNode1 As Node
    Dim TempNode2 As Node
    Dim MaxIndex As Integer
    
    ' Verifica que los índices sean válidos
    MaxIndex = Childs.Count
    If Index1 < 1 Or Index1 > MaxIndex Or Index2 < 1 Or Index2 > MaxIndex Then
        Err.Raise vbObjectError + 513, , "Índices fuera de rango."
        Exit Function
    End If
    
    ' Capturar los nodos en las posiciones dadas
    Set TempNode1 = Childs(Index1)
    Set TempNode2 = Childs(Index2)
    
    ' Si los índices son consecutivos, manejar de forma especial
    If Abs(Index1 - Index2) = 1 Then
        ' Eliminar primero el nodo de mayor índice
        If Index1 > Index2 Then
            Childs.Remove Index1
            Childs.Add TempNode1, TempNode1.key, Index2
        Else
            Childs.Remove Index2
            Childs.Add TempNode2, TempNode2.key, Index1
        End If
    Else
        ' Eliminar primero el nodo de mayor índice para no afectar el otro índice
        If Index1 > Index2 Then
            Childs.Remove Index1
            Childs.Remove Index2
        Else
            Childs.Remove Index2
            Childs.Remove Index1
        End If
        
        ' Agregar los nodos en las posiciones intercambiadas
        If Index1 <= MaxIndex - 1 Then
            Childs.Add TempNode2, TempNode2.key, Index1
        Else
            Childs.Add TempNode2, TempNode2.key
        End If
        
        If Index2 <= MaxIndex - 1 Then
            Childs.Add TempNode1, TempNode1.key, Index2
        Else
            Childs.Add TempNode1, TempNode1.key
        End If
    End If
End Function


Public Function PreviousNode() As Node
    Dim i As Long
    For i = 1 To Parent.Childs.Count
        If Parent.Childs(i) Is Me Then
            If i > 1 Then
                Set PreviousNode = Parent.Childs(i - 1)
            Else
                Set PreviousNode = Nothing
            End If
            Exit Function
        End If
    Next
End Function

Public Function NextNode() As Node
    Dim i As Long
    For i = 1 To Parent.Childs.Count
        If Parent.Childs(i) Is Me Then
            If i < Parent.Childs.Count Then
                Set NextNode = Parent.Childs(i + 1)
            Else
                Set NextNode = Nothing
            End If
            Exit Function
        End If
    Next
End Function

Public Function LastSibling() As Node
    Set LastSibling = Parent.Childs(Parent.Childs.Count)
End Function

Public Function FirstSibling() As Node
    Set FirstSibling = Parent.Childs(1)
End Function

Public Function FindNodeByKey(SearchKey As Variant) As Node
    Dim N As Node
    For Each N In Childs
        If N.key = SearchKey Then
            Set FindNodeByKey = N
            Exit Function
        End If
    Next
End Function

Public Function FindNodeByText(searchText As String) As Node
    Dim N As Node
    For Each N In Childs
        If N.Text = searchText Then
            Set FindNodeByText = N
            Exit Function
        End If
    Next
End Function

Public Sub Sort(Optional ByVal sortBy As String = "Text", Optional ByVal ascending As Boolean = True)
    Dim i As Integer, j As Integer
    Dim comparisonResult As Integer

    ' Bubble Sort
    For i = 1 To Childs.Count - 1
        For j = i + 1 To Childs.Count
            comparisonResult = CompareNodes(Childs(i), Childs(j), sortBy)
            If (ascending And comparisonResult > 0) Or (Not ascending And comparisonResult < 0) Then
                Me.Switch i, j
            End If
        Next j
    Next i
End Sub

Private Function CompareNodes(node1 As Node, node2 As Node, ByVal sortBy As String) As Integer
    Select Case sortBy
        Case "Text"
            CompareNodes = StrComp(node1.Text, node2.Text, vbTextCompare)
        Case "key"
            CompareNodes = StrComp(CStr(node1.key), CStr(node2.key), vbTextCompare)
        Case "SubText"
            CompareNodes = StrComp(node1.SubText, node2.SubText, vbTextCompare)
        Case "Data"
            CompareNodes = StrComp(node1.Data, node2.Data, vbTextCompare)
    End Select
End Function

Public Function Add(ByVal sText As String, _
        Optional ByVal key As Variant, _
        Optional ByVal Image As Variant, _
        Optional ByVal SelectedImage As Variant, _
        Optional ByVal bForcePlusButton As Boolean, _
        Optional Before As Variant, _
        Optional After As Variant _
        ) As Node
        
        Dim cNode As Node
        Set cNode = New Node
        cNode.key = key
        cNode.Text = sText
        cNode.Image = Image
        cNode.SelectedImage = SelectedImage
        cNode.PlusButton = bForcePlusButton
        Set cNode.Parent = Me
        Childs.Add cNode, key, Before, After
        Set Add = cNode
End Function

Function GetFullPath(Optional UseKey As Boolean, Optional PathSeparator As String = "\") As String
    Dim Path As String
    
    Dim N As Node
    
    Set N = Me

    ' Agrega el texto del nodo actual
    If UseKey Then
        Path = N.key
    Else
        Path = N.Text
    End If
    
    ' Recorre los padres del nodo actual
    Do While Not N.Parent Is Nothing
        Set N = N.Parent
        If UseKey Then
            Path = N.key & PathSeparator & Path
        Else
            Path = N.Text & PathSeparator & Path
        End If
    Loop
    
    ' Elimina el PathSeparator inicial si existe
    If VBA.Left$(Path, 1) = PathSeparator Then
        Path = Mid(Path, 2)
    End If
    
    GetFullPath = Path
End Function

Private Sub Class_Initialize()
    Set Childs = New Collection
    TextColor = -1
    BackColor = -1
    BadgeBackColor = vbRed
    BorderColor = &HFFCCCCCC
End Sub
