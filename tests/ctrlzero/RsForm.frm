VERSION 5.00
Begin VB.Form RsForm 
   Caption         =   "RsForm"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4560
   Begin VB.Data DataZ 
      Caption         =   "DataZ"
      Connect         =   ""
      DatabaseName    =   "."
      RecordSource    =   "SELECT id FROM [x]"
      Height          =   375
      Left            =   240
      Top             =   2520
      Width           =   5880
   End
End
Attribute VB_Name = "RsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ledger 278 knife 15 (B124): Data1.Recordset.<member> has exactly ONE codegen shape
' today - the real IDispatch from vb6_Data_RecordsetObj. The sentinel counts these ten
' shapes and pins that no direct vb6_Data_<member>() call is emitted any more.
Public Sub Methods()
    DataZ.Recordset.Refresh
    Call DataZ.Recordset.MoveFirst
    DataZ.Recordset.MoveLast
    DataZ.Recordset.MoveNext
    DataZ.Recordset.MovePrevious
End Sub
Public Sub Reads()
    Dim sZ As String
    sZ = DataZ.Recordset.Fields("id").Value
    sZ = DataZ.Recordset.Fields(1)
    If DataZ.Recordset.BOF Then sZ = "a"
    sZ = DataZ.Recordset.RecordCount
End Sub
Public Sub InWith()
    With DataZ.Recordset
        .Refresh
    End With
End Sub
