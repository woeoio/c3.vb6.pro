' C30-10 内置 Dictionary 的行为基准（跑在**真 scrrun.dll** 上，cscript 当 oracle）
'
' 为什么先有这份：Dictionary 的坑不在 Add/Item 这种主干，而在
'   * CompareMode 0/1/2 下"键相等"到底怎么算（1 与 "1"、Empty、Nothing）
'   * 删掉中间一个键之后 Keys 的顺序（插入序还是别的）
'   * 各种缺失/重复/空键的错误号（32811 / 457 / 94 / 5）
' 这些不能靠猜，先把真组件的答案钉成文件，实现侧逐行对照。
'
' 输出格式：id|值   或  id|err|错误号   （两边必须逐字一致）

Option Explicit

Dim fso, gOut, gDir
Set fso = CreateObject("Scripting.FileSystemObject")
gDir = fso.GetParentFolderName(WScript.ScriptFullName)
Set gOut = fso.OpenTextFile(fso.BuildPath(gDir, "dict_oracle.txt"), 2, True)

Sub P(ByVal s)
    gOut.WriteLine s
End Sub

Dim d

' ---- 基本 ----
Set d = CreateObject("Scripting.Dictionary")
P "d00|typename=" & TypeName(d)
P "d00b|isobject=" & CStr(IsObject(d))
d.Add "a", 1
P "d01|item=" & CStr(d.Item("a"))
P "d02|default=" & CStr(d("a"))
P "d03|exists_missing=" & CStr(d.Exists("zz"))
P "d04|count=" & CStr(d.Count)
P "d05|cmpmode=" & CStr(d.CompareMode)

Set d = Nothing

' ---- 错误号 ----
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
On Error Resume Next
Err.Clear
Dim v
v = d.Item("nope")
P "d10|item_missing_err=" & CStr(Err.Number)
Err.Clear
d.Add "a", 2
P "d11|add_dup_err=" & CStr(Err.Number)
Err.Clear
d.Remove "nope"
P "d12|remove_missing_err=" & CStr(Err.Number)
On Error GoTo 0
Set d = Nothing

' ---- 插入序 / 删除后的顺序 ----
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "b", 2
d.Add "c", 3
P "d20|keys_abc=" & Join(d.Keys, ",")
P "d21|items_abc=" & Join(d.Items, ",")
d.Remove "b"
P "d22|keys_after_remove_b=" & Join(d.Keys, ",")
d.Add "b", 9
P "d23|keys_readd_b=" & Join(d.Keys, ",")
P "d24|items_readd_b=" & Join(d.Items, ",")
d.RemoveAll
P "d25|count_after_removeall=" & CStr(d.Count)
Set d = Nothing

' ---- 键的类型：1 vs "1" vs 1.0（Binary 模式）----
Set d = CreateObject("Scripting.Dictionary")
d.Add 1, "int1"
P "d30|exists_str_1=" & CStr(d.Exists("1"))
P "d31|exists_int_1=" & CStr(d.Exists(1))
P "d32|exists_dbl_1=" & CStr(d.Exists(CDbl(1)))
Set d = Nothing

' ---- CompareMode=1 (Text) ----
Set d = CreateObject("Scripting.Dictionary")
d.CompareMode = 1
d.Add "Key", "v1"
P "d40|text_exists_lower=" & CStr(d.Exists("key"))
P "d41|text_item_lower=" & CStr(d.Item("key"))
On Error Resume Next
Err.Clear
d.Add "KEY", "v2"
P "d42|text_add_case_dup_err=" & CStr(Err.Number)
On Error GoTo 0
P "d43|text_count=" & CStr(d.Count)
Set d = Nothing

' ---- CompareMode=1 下的数字键 ----
Set d = CreateObject("Scripting.Dictionary")
d.CompareMode = 1
d.Add 1, "one"
P "d50|text_exists_str1=" & CStr(d.Exists("1"))
Set d = Nothing

' ---- CompareMode=2 (Database) 与 1 的差别（尾随空格）----
Set d = CreateObject("Scripting.Dictionary")
d.CompareMode = 2
d.Add "ab", 1
P "d55|db_exists_trailing_space=" & CStr(d.Exists("ab  "))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.CompareMode = 1
d.Add "ab", 1
P "d56|text_exists_trailing_space=" & CStr(d.Exists("ab  "))
Set d = Nothing

' ---- 空键 / Nothing / Null ----
Set d = CreateObject("Scripting.Dictionary")
On Error Resume Next
Err.Clear
d.Add Empty, "e"
P "d60|add_empty_err=" & CStr(Err.Number)
P "d61|exists_empty=" & CStr(d.Exists(Empty))
Err.Clear
d.Add Nothing, "n"
P "d62|add_nothing_err=" & CStr(Err.Number)
Err.Clear
d.Add Null, "z"
P "d63|add_null_err=" & CStr(Err.Number)
On Error GoTo 0
Set d = Nothing

' ---- Item 赋值即建键 ----
Set d = CreateObject("Scripting.Dictionary")
d("newkey") = 5
P "d70|exists_after_item_set=" & CStr(d.Exists("newkey"))
P "d71|count_after_item_set=" & CStr(d.Count)
P "d72|read_back=" & CStr(d("newkey"))
Set d = Nothing

' ---- Keys/Items 是数组吗 ----
Set d = CreateObject("Scripting.Dictionary")
d.Add "x", 10
P "d80|keys_isarray=" & CStr(IsArray(d.Keys))
P "d81|keys_lbound=" & CStr(LBound(d.Keys))
P "d82|keys_ubound=" & CStr(UBound(d.Keys))
P "d83|items_isarray=" & CStr(IsArray(d.Items))
Set d = Nothing

' ---- For Each 走键 ----
Set d = CreateObject("Scripting.Dictionary")
d.Add "k1", 1
d.Add "k2", 2
Dim k, acc
acc = ""
For Each k In d
    acc = acc & CStr(k) & ";"
Next
P "d90|foreach_keys=" & acc
Set d = Nothing

' ---- 有元素时改 CompareMode ----
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
On Error Resume Next
Err.Clear
d.CompareMode = 1
P "d95|cmpmode_change_with_items_err=" & CStr(Err.Number)
On Error GoTo 0
Set d = Nothing

' ---- 值类型：对象值 / 数组值 ----
Set d = CreateObject("Scripting.Dictionary")
Dim arr(2)
arr(0) = "p"
arr(1) = "q"
d.Add "arr", arr
Dim back
back = d("arr")
P "d96|array_value_elem=" & CStr(back(1))
Set d = Nothing

gOut.Close
