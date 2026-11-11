VERSION 5.00
Begin VB.Form MenuArr 
   Caption         =   "MenuArr"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Begin VB.Menu MnuFileTop 
      Caption         =   "File"
      Begin VB.Menu MnuFile 
         Caption         =   "New..."
         Index           =   0
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Open recent"
         Index           =   1
         Begin VB.Menu MnuRecent 
            Caption         =   "empty"
         End
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Import"
         Index           =   2
         Begin VB.Menu MnuImport 
            Caption         =   "From clipboard"
         End
      End
   End
   Begin VB.Menu MnuEditTop 
      Caption         =   "Edit"
      Begin VB.Menu MnuUndo 
         Caption         =   "Undo"
      End
   End
End
Attribute VB_Name = "MenuArr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
