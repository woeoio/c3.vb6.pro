# Bool / ByRef 口径探针（给真 VB6 环境用）

这两问在本机没有 VB6 服务端可问，所以用一组最小 VB6 件取真读数。**只需回贴三段文本**，不用分析。

## 要回答的两问

| 编号 | 问题 | 我们现在的实现 |
|---|---|---|
| Q1 | VB6 **服务端**的 `ByVal s As String` 形参收到自动化客户端递来的 `VT_BOOL` 时，体里看到的是 `True/False` 还是 `-1/0`？ | 走 OLE `VariantChangeType` ⇒ `-1`（`vb6_ComArgNormalize`） |
| Q2 | VB6 **客户端**（晚绑定 `As Object`）调 `Bump(ByRef v As Long)` 之后，调用方那个变量会变成 11 吗？ | 我们的客户端按值打包 ⇒ 变量不变（写回丢失） |

---

## 步骤

### 1. 编译服务端

用 VB6 打开 `ai/probes/bool_string/BoolProbe.vbp`：

- 工程里**没有写任何 Reference 行**（不需要 stdole）；若 IDE 自己弹引用对话框，直接确定/取消都行。
- `文件 > 生成 BoolProbe.dll`，存到本目录（`ai/probes/bool_string/`）。

### 2. 注册（二选一）

- **有管理员**：在 cmd 里 `regsvr32 /s "完整路径\BoolProbe.dll"`；用完 `regsvr32 /u /s` 注销。
- **没有管理员**：回到 VB6 IDE，直接 **F5 运行 BoolProbe 工程**（提示启动部件时选"等待组件被创建"），
  IDE 在跑的时候 `CreateObject("BoolProbe.ProbeCls")` 一样能建出来。跑完 Esc 停掉即可。

### 3. 跑自动化客户端（VBScript）

```
cscript //nologo "完整路径\probe.vbs"
```

### 4. 跑 VB6 客户端（晚绑定）

用 VB6 打开 `ai/probes/bool_string/vbclient/BoolProbeClient.vbp`，**F5**。
它会弹一个 MsgBox，并把同样内容写到 `%TEMP%\boolprobe_client.txt`。

---

## 回贴什么

1. `cscript` 打印的 8 行（A1–A8）。
2. VB6 客户端弹框里的 4 行（B1–B4），或 `%TEMP%\boolprobe_client.txt` 的内容。
3. 一句话：`生成 BoolProbe.dll` 有没有报错。

## 每行的意义（我这边据此定口径）

- **A1/A2**：`EchoStr(True)` / `EchoStr(False)` —— **这就是 Q1 的答案**。
  `[True]`/`[False]` ⇒ VB6 走语言口径；`[-1]`/`[0]` ⇒ VB6 也走 OLE 转换，我们不用改。
- A3/A4：非布尔对照，应分别为 `[12.75]`、`[40]`（证明夹具本身是通的）。
- A5：`EchoLng(True)`，预期 `-1`（布尔进数值形参）。
- A6/A7：`EchoVar(True)` / `EchoVar(12.75)`，看 TypeName 与字面；预期 `Boolean=True`、`Double=12.75`。
- A8：VBScript 客户端 + `Bump`。VBScript 的变量传递与 VB6 不同，这行只作参考。
- **B4**：VB6 客户端的 `Bump` 之后 `x` 的值 —— **这是 Q2 的答案**。
  `x=11` ⇒ VB6 客户端确实按引用发（我们的客户端应当照做）；`x=10` ⇒ 两边一致，Q2 可以销。

## 服务端源码（`ProbeCls.cls`，供你核对，不用改）

```vb
Public Function EchoStr(ByVal s As String) As String
    EchoStr = "[" & s & "]"
End Function
Public Function EchoLng(ByVal n As Long) As Long
    EchoLng = n
End Function
Public Function EchoVar(ByVal v As Variant) As String
    EchoVar = TypeName(v) & "=" & CStr(v)
End Function
Public Function Bump(ByRef v As Long) As Long
    v = v + 1
    Bump = v
End Function
```


## 读数（2026-10-26，真 VB6 6.0，已收到）

```
A1 EchoStr(True)  = [True]
A2 EchoStr(False) = [False]
A3 EchoStr(12.75) = [12.75]
A4 EchoStr(40)    = [40]
A5 EchoLng(True)  = -1
A6 EchoVar(True)  = Boolean=True
A7 EchoVar(12.75) = Double=12.75
A8  ------------  = 类型不匹配（VBScript 调 o.Bump(x)）
B1 EchoStr(True)  = [True]
B2 EchoStr(False) = [False]
B3 EchoLng(True)  = -1
B4 Bump: x=11 r=11
```

结论：

- **Q1 已定**：服务端 `As String` 形参收 `VT_BOOL` 走**语言**口径（`True`/`False`），不是 OLE 的 `-1`/`0`。
  落点 = `vb6_ComChangeType`（唯一出口），第 238 刀。
- **Q2 已定**：`B4` 那个客户端是 `Dim o As Object` + `CreateObject`，**晚绑定**，写回仍然发生 ⇒ 客户端该递 `VT_BYREF|T`。
- A8 的类型不匹配不是矛盾：VBScript 的变量是 `Variant`（递 `VT_BYREF|VT_VARIANT`），真 VB6 服务端拒；
  VB6 的 `x As Long` 递的是 `VT_BYREF|VT_I4`，服务端接受并写回。差别在**实参带不带型**。
- 包装坑（已修）：`vbclient/Module1.bas` 里 `"\boolprobe_client.txt"` 的 `\b` 曾被写成退格符 0x08，
  真 VB6 报「错误的文件名或号码」(错误 52)。往 .bas/.vbs/.ps1 里写反斜杠必须显式拼 `chr(92)`。
