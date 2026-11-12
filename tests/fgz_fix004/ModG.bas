Attribute VB_Name = "ModG"
Option Explicit

' 富贵竹 FIX-004 (账 #610 第340刀): 唯一的**真全局** —— 显式 Public 关键字。
' 同名的窗体 Dim 副本不许把它连坐成 static (修前: ModG.h 无 extern,
' ModU 裸名引用 C2065, 本地探针 .local/fgz-probes/fix004 实锤)。
Public gShared As Long

Public Sub SetG(ByVal n As Long)
    gShared = n
End Sub
