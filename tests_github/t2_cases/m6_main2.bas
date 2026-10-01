' m6_main2.bas - 测试 Module.Method 跨模块调用语法
' 限定符必须写成 m6_math2 —— 那是本工程里那个模块的**名字** (VB6 与 C3 都取 .bas 的
' 文件名)。本模块没写 Option Explicit, 所以旧稿里的 `MathUtils.Add(...)` 不会报错,
' 只会被按 VB6 的隐式声明规矩当成一个空的 Variant 变量走晚绑定, 恒返 0 (实测)。
Sub Main()
    Dim result As Long
    result = m6_math2.Add(10, 20)
    Debug.Print result
    result = m6_math2.Multiply(3, 7)
    Debug.Print result
End Sub
