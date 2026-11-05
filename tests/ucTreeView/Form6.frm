VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00525252&
   Caption         =   "Agenda tema oscuro"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form6"
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   719
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4320
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   600
   End
   Begin Proyecto1.ucTreeViewPlus ucTreeViewPlus1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   10398
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ContainerBkColor=   4275513
      BackColor       =   4275513
      BorderRadius    =   0
      BorderColor     =   0
      SelectionColor  =   15629646
      PlusMinusVisible=   0   'False
      BeginProperty SubTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PlusMinusColor  =   16777215
      PlusMinusHotColor=   16737123
      NodeSpacing     =   8
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2280
         Picture         =   "Form6.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Delete contact"
         Top             =   2280
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   
Const Contacts = "Acosta Lucas,Alvarez Juan,Andrade Maria,Arias Pedro,Ayala Carolina,Becerra Julian,Benitez Laura,Blanco Sergio,Bustos Gabriela,Campos Mariana,Cardozo Alejandro,Carrasco Fernanda,Castillo Raul,Cortez Lorena,Cruz Rodrigo,Dominguez Javier,Escobar Lucia,Farias Gustavo,Franco Jose,Garcia Ana,Gonzalez Clara,Guerrero Marcos,Gutierrez Pilar,Herrera Francisco,Iglesias Rafael,Juarez Patricia,Leiva Adrian,Lopez Emiliano,Luna Valeria,Maldonado Cristian,Marquez Andrea,Martinez Sebastian,Medina Laura,Mendoza Javier,Miranda Jorge,Molina Daniela,Morales Ignacio,Muñoz Federico,Navarro Alejandro,Nunez Juliana,Olivares Hernan,Ortega Esteban,Paredes Silvina,Ponce Ignacio,Quiroga Luis,Ramos Fabian,Reyes Florencia,Rios Federico,Rivera Diego,Romero Gabriela,Rosa Pablo,Ruiz Emiliano,Salas Fernanda,Sanchez Marta,Silva Martin,Sosa Guillermo,Suarez Francisco,Tapia Matias,Torres Cecilia,Valdez Marcos,Vargas Luis,Vazquez Ana,Vega Luciana,Velazquez Raul,Vera Santiago,Villalba Julieta,Zamora Ignacio"
Const ColoresPastel = "&HEFD1BA, &HFDD2C2, &HEFD9C1, &HE4D2B7, &HCCE2CB, &HFFE4D0, &HFFECB3, &HF6DBA9, &HD4A5A5, &HF0B27A, &HC5E1A5, &HFDDAC6, &HE1F5FE, &HFFF59D, &HFFCCBC, &HCE93D8, &HF48FB1, &HBCAAA4, &HA5D6A7, &FFAB91"

Private Sub Form_Load()
    Dim ArrContacts() As String
    Dim Colores() As String
    Dim lColor As Long
    Dim i As Long
    Dim Mail As String
    Dim imgIndex As Long
    Dim LastLetter As String
    Dim CurrentLetter As String
    Dim oNode As Node
    
    ArrContacts = Split(Contacts, ",")
    Colores = Split(ColoresPastel, ",")

    ucTreeViewPlus1.ImageListInitialize 40, 40
    ucTreeViewPlus1.ImagesRadius = 38
    
    For i = 0 To UBound(ArrContacts)
        lColor = Val(Colores(RandomBetween(0, UBound(Colores) - 1)))
        ucTreeViewPlus1.ImageListAddImage GetContactImgProfile(ArrContacts(i), lColor)
    Next
    
    For i = 0 To UBound(ArrContacts)
        CurrentLetter = Left$(ArrContacts(i), 1)
        If CurrentLetter <> LastLetter Then
            LastLetter = CurrentLetter
            Set oNode = ucTreeViewPlus1.Nodes.Add(LastLetter)
            oNode.FontBold = True
            oNode.BackColor = &H595652
        End If
        Mail = LCase(Replace$(ArrContacts(i), " ", "")) & "@gmail.com"
        imgIndex = imgIndex + 1
        With oNode.Add(ArrContacts(i), , imgIndex)
            .SubText = Mail
            '.BadgeBackColor = -1
            '.Badge = CStr(i)
        End With
    Next
    
    ucTreeViewPlus1.ExpandAllNodes
    ucTreeViewPlus1.Update
End Sub


Private Function GetContactImgProfile(sContact As String, BackColor As Long) As StdPicture
    Dim sNames() As String
    Dim Letters As String
    Dim TW As Long, TH As Long
    sNames = Split(sContact, " ")
    Letters = UCase(Left$(sNames(0), 1) & Left$(sNames(1), 1))
    
    With Picture1
        TW = .TextWidth(Letters)
        TH = .TextHeight(Letters)
        .BackColor = BackColor
        .CurrentX = Picture1.ScaleWidth / 2 - TW / 2
        .CurrentY = Picture1.ScaleHeight / 2 - TH / 2
        Picture1.Print Letters
        Set GetContactImgProfile = .Image
    End With

End Function

Function RandomBetween(minValue As Integer, maxValue As Integer) As Integer
    RandomBetween = Int((maxValue - minValue + 1) * Rnd + minValue)
End Function


Private Sub Image1_Click()
    ucTreeViewPlus1.GetHotNode.Parent.Remove ucTreeViewPlus1.GetHotNode.index
    Image1.Visible = False
    ucTreeViewPlus1.Update
    
End Sub

Private Sub ucTreeViewPlus1_NodeMouseEnter(Node As Node)
    If Len(Node.SubText) Then 'identificamos que sea un contacto y no un nodo padre
        Image1.Top = (Node.Top + (Node.Bottom - Node.Top) / 2) * Screen.TwipsPerPixelY - Image1.Height / 2
        Image1.Left = ucTreeViewPlus1.Width * Screen.TwipsPerPixelX - Image1.Width - 280
        Image1.Visible = True
    Else
        Image1.Visible = False
    End If
End Sub

Private Sub ucTreeViewPlus1_NodeMouseLeave(Node As Node)
     Image1.Visible = False
End Sub

Private Sub ucTreeViewPlus1_ScrollChange(ByVal Value As Long, ByVal Max As Long)
     Image1.Visible = False
End Sub
