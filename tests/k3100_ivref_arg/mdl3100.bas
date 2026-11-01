Attribute VB_Name = "mdl3100"
Option Explicit

Private g_val As Long
Public g_term As Long

Private Sub Take(ByRef x As IK3100)
    g_val = g_val + 1
End Sub

Private Sub TakeVal(ByVal x As IK3100)
    g_val = 1
End Sub

Private Function MakeI() As IK3100
    Set MakeI = New H3100
End Function

Public Sub Main()
    Dim c As H3100
    Set c = New H3100
    Dim v As IK3100
    Set v = c
    Dim n As H3100
    Take MakeI()
    TakeVal c
    Debug.Print "K3100-v=" & g_val
    ' 第二百零五刀形①: 借来的**实例** (裸类变量) 进视图槽 —— 交付侧铸带 NULL 守卫的视图,
    ' copy-in 在这一格先付一枚、调用行之后交回一枚, 净 0 ⇒ c 既不提前死也不永不退。
    Take c
    Debug.Print "K3100-b=" & g_val
    ' 第二百零五刀形②: 借来的**视图** (视图变量) 进视图槽 —— 值原样交, 缺的是先付那一枚;
    ' 修前这里净 -1 (当场不炸, 到过程出口拆槽时 0xC0000374 = 过退的双自由)。
    Take v
    Debug.Print "K3100-w=" & g_val
    ' 第二百零五刀形③: 主人 nullable 那一档 —— n 从没 Set 过, 交出去的是 NULL 而非野地址。
    Take n
    Debug.Print "K3100-n=" & g_val
    Debug.Print "K3100-h=" & c.Hits
    ' 第二百零四刀形: `New <类>` 自己交出那一枚 ⇒ 只交回、不先付。
    Take New H3100
    Debug.Print "K3100-v2=" & g_val
    Debug.Print "K3100-t=" & g_term
    Set v = Nothing
    Set c = Nothing
    Debug.Print "K3100-t2=" & g_term
    Debug.Print "K3100-DONE"
End Sub
