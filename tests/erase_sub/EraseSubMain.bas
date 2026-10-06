Attribute VB_Name = "EraseSubMain"
Option Explicit

' 账 #186 的判据夹具: `Erase m_items(1).bvData` —— 销毁 **UDT 数组那一格里的**动态数组。
'   修前 parser 一律报 VB2001 "Erase 不支持带下标的形式"，真工程里这句合法且到处在用
'   (Charts 2020/ucTreeMaps 的 PropPagFMR.pag:720 `Erase m_tvFiles(lIndex).bvData` 就卡在这)。
'   EraseStmt 只带名字字符串、表达不了下标 ⇒ 发码需要一个表达式树 (与 ReDim 的 Fix 100 同一套机制),
'   出口是 VB6_SA_AT(vb6_type_TElem, m_items, 1).bvData —— 销毁它并置 NULL。
' 两头都要钉 (缺一头就是假绿):
'   ①别的格不受影响 (s0=30 / s2=50) —— 拦"静默退化成销毁整个数组"那一族错法;
'   ②那一格销毁后还能重新分配并用 (s1=20 len=1) —— 拦"没置 NULL"：ReDim 自己会先销毁旧数组,
'     没置 NULL 就是**双释放**, 现场直接崩; 反过来"没真销毁"会读到残留的 40。
' 负控 = 修复前的编译器编同一份夹具: parser 报 VB2001, BUILD rc=1, 一条读数也不会出现。

Private Type TElem
    nTag As Long
    bvData() As Byte
End Type

Private m_items(2) As TElem

Private Sub FillOne(ByVal idx As Long, ByVal n As Long, ByVal unit As Byte)
    Dim k As Long
    m_items(idx).nTag = idx * 100
    ReDim m_items(idx).bvData(n - 1)
    For k = 0 To n - 1
        m_items(idx).bvData(k) = unit
    Next
End Sub

Private Function SumOf(ByVal idx As Long) As Long
    Dim k As Long, s As Long
    For k = 0 To UBound(m_items(idx).bvData)
        s = s + m_items(idx).bvData(k)
    Next
    SumOf = s
End Function

Sub Main()
    FillOne 0, 3, 10          ' 和 30
    FillOne 1, 4, 10          ' 和 40
    FillOne 2, 5, 10          ' 和 50
    Debug.Print "EA00-BEFORE s0=" & SumOf(0) & " s1=" & SumOf(1) & " s2=" & SumOf(2)

    Erase m_items(1).bvData

    Debug.Print "EA01-KEEP s0=" & SumOf(0) & " s2=" & SumOf(2)

    ReDim m_items(1).bvData(1)
    m_items(1).bvData(0) = 10
    m_items(1).bvData(1) = 10
    Debug.Print "EA02-RECYCLE s1=" & SumOf(1) & " len=" & UBound(m_items(1).bvData)

    Debug.Print "EA-CNT keep_ok=" & CStr(SumOf(0) = 30 And SumOf(2) = 50) & _
                " recycle_ok=" & CStr(SumOf(1) = 20 And UBound(m_items(1).bvData) = 1)
    Debug.Print "EA-DONE"
End Sub
