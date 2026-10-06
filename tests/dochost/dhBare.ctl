VERSION 5.00
Begin VB.UserControl dhBare 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ScaleHeight     =   1200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1200
End
Attribute VB_Name = "dhBare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> 账 #219 的判据夹子 (裸写那一面, UserControl)。VB6 允许在 .ctl/.pag 里把文档
' 自带的成员裸写 (等价于 <对象>.<成员>)。发码侧从账 #159 起就由 kHostPseudoRows 那张表回答,
' 语义层以前不问它 => 每条合法裸写配一句 VB3001 (语料实测 .ctl 里 hDC 占 24 条)。
' 判据两头: 一, 表里带 HPF_BARE 的名字一个都不许报 (本夹子); 二, 打错的名字必须继续报
' (dhTypo.pag) —— 放行的是那张表, 不是一串名字。

Public Function BareVals() As String
    BareVals = CStr(ScaleWidth) & "/" & CStr(hDC) & "/" & CStr(hWnd)
End Function
