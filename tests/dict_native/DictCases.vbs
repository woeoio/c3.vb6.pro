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

' ---- 错误号（每例用新字典，避免上一条的错误/状态污染读数）----
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
On Error Resume Next
Err.Clear
Dim v
v = d.Item("nope")
P "d10|item_missing_err=" & CStr(Err.Number)
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
On Error Resume Next
Err.Clear
d.Add "a", 2
P "d11|add_dup_err=" & CStr(Err.Number)
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
On Error Resume Next
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
Dim its
its = d.Items
P "d21|items_abc=" & CStr(its(0)) & "," & CStr(its(1)) & "," & CStr(its(2))
d.Remove "b"
P "d22|keys_after_remove_b=" & Join(d.Keys, ",")
d.Add "b", 9
P "d23|keys_readd_b=" & Join(d.Keys, ",")
its = d.Items
P "d24|items_readd_b=" & CStr(its(0)) & "," & CStr(its(1)) & "," & CStr(its(2))
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
d.Item("newkey") = 5
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

' ---- S1: key type truth table (all readings from real scrrun) ----
Dim dv, dk, acc2, kk

Set d = CreateObject("Scripting.Dictionary")
d.Add True, "t"
P "d100|bool_key: true=" & CStr(d.Exists(True)) & " neg1=" & CStr(d.Exists(-1)) & " one=" & CStr(d.Exists(1)) & " strTrue=" & CStr(d.Exists("True"))
Set d = Nothing

dv = CDate("2026-11-16 10:30:00")
Set d = CreateObject("Scripting.Dictionary")
d.Add dv, "dt"
P "d101|date_key: same=" & CStr(d.Exists(dv)) & " asnum=" & CStr(d.Exists(CDbl(dv))) & " asstr=" & CStr(d.Exists(CStr(dv)))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add CCur(1.5), "cur"
P "d102|cur_key: cur=" & CStr(d.Exists(CCur(1.5))) & " dbl=" & CStr(d.Exists(CDbl(1.5))) & " int1=" & CStr(d.Exists(1))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.CompareMode = 2
d.Add "Abc", 1
P "d103|db_mode: lower=" & CStr(d.Exists("abc")) & " pad=" & CStr(d.Exists("Abc ")) & " leadpad=" & CStr(d.Exists(" Abc"))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add Nothing, "n"
P "d104|nothing_key: zero=" & CStr(d.Exists(0)) & " nothing=" & CStr(d.Exists(Nothing)) & " empStr=" & CStr(d.Exists(""))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add Empty, "e"
P "d105|empty_key: zero=" & CStr(d.Exists(0)) & " empty=" & CStr(d.Exists(Empty)) & " empStr=" & CStr(d.Exists("")) & " null=" & CStr(d.Exists(Null))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add Null, "nu"
P "d106|null_key: null=" & CStr(d.Exists(Null)) & " empty=" & CStr(d.Exists(Empty))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
dk = CInt(7)
d.Add dk, "i"
P "d107|int7_key: long7=" & CStr(d.Exists(CLng(7))) & " dbl7=" & CStr(d.Exists(CDbl(7))) & " str7=" & CStr(d.Exists("7"))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "b", 2
d.Add "c", 3
On Error Resume Next
Err.Clear
acc2 = ""
For Each kk In d
    acc2 = acc2 & kk & ";"
    If kk = "a" Then d.Remove "a"
Next
P "d108|foreach_remove: seq=" & acc2 & " err=" & CStr(Err.Number) & " count=" & CStr(d.Count)
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "b", 2
d.Item("a") = 99
P "d109|item_set_existing: keys=" & Join(d.Keys, ",") & " val=" & CStr(d("a"))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
On Error Resume Next
Err.Clear
d.CompareMode = 3
P "d110|cmpmode3: err=" & CStr(Err.Number) & " readback=" & CStr(d.CompareMode)
Err.Clear
d.CompareMode = -1
P "d111|cmpmode_neg: err=" & CStr(Err.Number)
On Error GoTo 0
Set d = Nothing

Dim ks
Set d = CreateObject("Scripting.Dictionary")
d.Add "x", 1
ks = d.Keys
P "d112|vartypes: count=" & CStr(VarType(d.Count)) & " keys=" & CStr(VarType(ks)) & " keyelem=" & CStr(VarType(ks(0)))
P "d121|vartype_ccur=" & CStr(VarType(CCur(1.5))) & " dbl=" & CStr(VarType(CDbl(1.5))) & " int=" & CStr(VarType(1))
P "d122|ccur_value=" & CStr(CCur(1.5))
Set d = CreateObject("Scripting.Dictionary")
d.Add CCur(1.5), "cur"
Dim sk
sk = d.Keys
P "d123|stored_key: vartype=" & CStr(VarType(sk(0))) & " val=" & CStr(sk(0))
Set d = Nothing

Dim inner
Set inner = CreateObject("Scripting.Dictionary")
inner.Add "i", 1
Set d = CreateObject("Scripting.Dictionary")
d.Add "o", inner
P "d113|object_value: isobject=" & CStr(IsObject(d("o"))) & " typename=" & TypeName(d("o"))
Set d = Nothing
Set inner = Nothing

' CompareMode=3 was accepted silently (readback=3). What does it compare like?
Set d = CreateObject("Scripting.Dictionary")
On Error Resume Next
d.CompareMode = 3
Err.Clear
d.Add "A", 1
d.Add "a", 2
P "d114|cm3_dedupe: add_lower_err=" & CStr(Err.Number) & " count=" & CStr(d.Count) & " readback=" & CStr(d.CompareMode)
On Error GoTo 0
Set d = Nothing

' Empty key under Text mode: still equal to 0 and ""?
Set d = CreateObject("Scripting.Dictionary")
d.CompareMode = 1
d.Add Empty, "e"
P "d115|empty_key_text: zero=" & CStr(d.Exists(0)) & " empty=" & CStr(d.Exists(Empty)) & " empStr=" & CStr(d.Exists(""))
Set d = Nothing

' ---- Key（属性写，改名）与 HashVal（未文档化成员）；第五十八刀 ----
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
On Error Resume Next
d.Key("a") = "b"
P "d116|key_rename: err=" & CStr(Err.Number) & " hasB=" & CStr(d.Exists("b")) & " hasA=" & CStr(d.Exists("a")) & " val=" & CStr(d("b"))
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
On Error Resume Next
d.Key("zz") = "yy"
P "d117|key_missing: err=" & CStr(Err.Number) & " count=" & CStr(d.Count)
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "b", 2
On Error Resume Next
d.Key("a") = "b"
P "d118|key_collision: err=" & CStr(Err.Number) & " count=" & CStr(d.Count) & " a=" & CStr(d("a")) & " b=" & CStr(d("b"))
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "b", 2
d.Add "c", 3
On Error Resume Next
d.Key("a") = "z"
P "d119|key_order: keys=" & Join(d.Keys, ",")
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
On Error Resume Next
d.Key("a") = "a"
P "d120|key_same_name: err=" & CStr(Err.Number) & " count=" & CStr(d.Count)
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
On Error Resume Next
d.Key("A") = "b"
P "d121|key_case_binary: err=" & CStr(Err.Number)
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.CompareMode = 1
d.Add "a", 1
On Error Resume Next
d.Key("A") = "b"
P "d122|key_case_text: err=" & CStr(Err.Number) & " hasB=" & CStr(d.Exists("b"))
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add 1, "x"
On Error Resume Next
d.Key(1) = 2
P "d123|key_numkey: err=" & CStr(Err.Number) & " has2=" & CStr(d.Exists(2)) & " val=" & CStr(d(2))
On Error GoTo 0
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "abc", 2
P "d124|hashval_str: a=" & CStr(d.HashVal("a")) & " abc=" & CStr(d.HashVal("abc")) & " hello=" & CStr(d.HashVal("hello")) & " empty=" & CStr(d.HashVal(""))
P "d125|hashval_missing: val=" & CStr(d.HashVal("nope")) & " type=" & TypeName(d.HashVal("nope"))
Set d = Nothing

Set d = CreateObject("Scripting.Dictionary")
On Error Resume Next
d("k") = "v"
P "d126|default_put: err=" & CStr(Err.Number) & " count=" & CStr(d.Count) & " val=" & CStr(d("k"))
On Error GoTo 0
Set d = Nothing

gOut.Close
