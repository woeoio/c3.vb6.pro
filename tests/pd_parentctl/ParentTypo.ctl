VERSION 5.00
Begin VB.UserControl ParentTypo 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1200
End
Attribute VB_Name = "ParentTypo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' 负控 (pd_parentctl): 只放行 `Parent` 这一个内建名。拼成 `Paren` 必须照旧报 VB3001，
' 否则"放行"就变成"任何前缀像 Parent 的名字都行"。

Public Function TypoCount() As Long
    Dim c As Object
    Dim n As Long

    For Each c In Paren.Controls
        n = n + 1
    Next c

    TypoCount = n
End Function
