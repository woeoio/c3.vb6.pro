VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Drag & Drop"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3855
   LinkTopic       =   "Form4"
   ScaleHeight     =   5475
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucTreeViewPlus ucTreeViewPlus1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   5395026
      FullRowSelect   =   0   'False
      ShowTreeLines   =   -1  'True
      BeginProperty SubTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PlusMinusColor  =   5395026
      PlusMinusHotColor=   16737123
      NodeSpacing     =   24
      Margin          =   6
      MousePointerHands=   -1  'True
      AllowEdit       =   -1  'True
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mX As Single
Dim mY As Single

Dim mDragNode As Node
Dim DestNode As Node

Private Sub Form_Load()
    Dim i As Long, j As Long
    With ucTreeViewPlus1
        .ImageListInitialize
        .ImageListAddImage App.Path & "\img\folderClose.png"
        .ImageListAddImage App.Path & "\img\folderOpen.png"
        .ImageListAddImage App.Path & "\img\file.png"
        .OLEDropMode = 1
    End With
    
    For i = 1 To 5
        With ucTreeViewPlus1.Nodes.Add("Folder " & i, , 1, 2)
            For j = 1 To 4
                Call .Add("File " & j, , 3)
            Next
        End With
    Next
End Sub

Private Sub ucTreeViewPlus1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mX = x
    mY = Y
End Sub

Private Sub ucTreeViewPlus1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        If Abs(x - mX) > 2 And Abs(Y - mY) > 2 Then
            If Not mDragNode Is Nothing Then Exit Sub
            Set mDragNode = ucTreeViewPlus1.GetHotNode()
            If Not mDragNode Is Nothing Then
                mDragNode.Selected = True
                ucTreeViewPlus1.OLEDrag
            End If
        End If
    End If
End Sub

Private Sub ucTreeViewPlus1_OLECompleteDrag(Effect As Long)
    mDragNode.Selected = False
    ucTreeViewPlus1.SelNode = mDragNode
    Set mDragNode = Nothing
    Set DestNode = Nothing
    ucTreeViewPlus1.Update
End Sub

Private Sub ucTreeViewPlus1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
    Set DestNode = ucTreeViewPlus1.HitTest(x / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
    If Not DestNode Is Nothing Then
        If mDragNode.Parent Is DestNode Then
            Effect = vbDropEffectNone
            Exit Sub
        End If
    
    
        If Not DestNode Is mDragNode Then
            If DestNode.Image = 1 Then
                ucTreeViewPlus1.SelNode = DestNode
            End If
        Else
            Set DestNode = Nothing
        End If
    Else
        Set DestNode = ucTreeViewPlus1.Nodes
    End If
    ucTreeViewPlus1.Refresh
End Sub

Private Sub ucTreeViewPlus1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    Dim index As Long

    If Not DestNode Is Nothing Then
        If Not DestNode Is mDragNode Then
            index = DestNode.index
            If mDragNode Is Nothing Then Exit Sub
            mDragNode.Parent.Childs.Remove mDragNode.index
            If index = 0 Then  'Raiz "\"
               Set mDragNode.Parent = DestNode
               DestNode.Childs.Add mDragNode, mDragNode.key, 1
               DestNode.Sort
            Else
                If DestNode.Image = 1 Then  'Folder
                    Set mDragNode.Parent = DestNode
                    DestNode.Childs.Add mDragNode, mDragNode.key
                    DestNode.Sort
                Else 'File
                    Set mDragNode.Parent = DestNode.Parent
                    DestNode.Parent.Childs.Add mDragNode, mDragNode.key
                    DestNode.Parent.Sort
                End If
            End If
        End If
    End If
End Sub

Private Sub ucTreeViewPlus1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData ObjPtr(mDragNode)
End Sub
