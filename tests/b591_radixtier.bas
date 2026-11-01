Attribute VB_Name = "ModB591"
Option Explicit

' 账 #305 第313刀 (§B239，任务 #591) = 无后缀进制字面量的**位模式**按写下的那一档补码读。
'
' VB6 的档位 = 装得下这个位模式的最小有符号类型，于是：
'   &HBF32 = -16590   &HFFFF = -1   &H8000 = -32768   &O177777 = -1   &B16个1 = -1
'   &HFFFFFFFF = -1  (32 位档, 老规矩逐字保留)
' 分界量的是**数值位宽**、不是位数 —— 同一发探针里上下各实测一发：
'   &H10000 = +65536 (不动)   &O200000 = +65536 (不动)   &O100000 = -32768 (要动)
' 读数由 cscript 的 VBScript 引擎给出 (.local/k313 两轮探针: tier.vbs + octbin + mask)。
'
' 病不是洁癖。第312刀符号化打在**一条语句**上: ComCtlsDemo 点 Command2 →
' Common.PictureFromHandle 里 `CLSID.Data2 = &HBF32` 发成 vb6_ChkInt(48946) → 运行期错误 6
' 打死进程 (改前本夹具同样一开场就死, rc=6 且 stdout 零行 —— 这就是 BASE 负控)。
' 位运算那一族更广: `255 Or &HFF00` 在 VB 里是 -1 (掩码自己按 16 位补码带符号扩展),
' 按 +65280 算得的是另一个数; 真语料 D:\vb_yqt4qPac 里落在这一档的字面量有 1285 处。
'
' 证人七条 = 本刀**不许动**的形: 16 位档以下 (255)、以上 (&H10000 / &O200000)、
' 32 位档 (&HFFFFFFFF = -1 / &H7FFFFFFF)、以及 % 与 & 两枚显式后缀 (&HFFFF& = 65535)。
' &B 那一族 VB6 压根没有 (C3 的扩展, VBScript 对它报语法错), 按同一位主人走。

Sub Main()
    Dim i As Integer
    i = &HBF32
    Debug.Print "RT-store=" & CStr(i)
    Debug.Print "RT-h8000=" & CStr(&H8000)
    Debug.Print "RT-hffff=" & CStr(&HFFFF)
    Debug.Print "RT-h8def=" & CStr(&H8DEF)
    Debug.Print "RT-hf082=" & CStr(&HF082)
    Debug.Print "RT-o177777=" & CStr(&O177777)
    Debug.Print "RT-o100000=" & CStr(&O100000)
    Debug.Print "RT-b16ones=" & CStr(&B1111111111111111)
    Debug.Print "RT-mask=" & CStr(255 Or &HFF00)
    Debug.Print "RT-w-low=" & CStr(&HFF)
    Debug.Print "RT-w-above16=" & CStr(&H10000)
    Debug.Print "RT-w-o200000=" & CStr(&O200000)
    Debug.Print "RT-w-h32=" & CStr(&HFFFFFFFF)
    Debug.Print "RT-w-max=" & CStr(&H7FFFFFFF)
    Debug.Print "RT-w-amp=" & CStr(&HFFFF&)
    Debug.Print "RT-w-pct=" & CStr(&HFF%)
    Debug.Print "RT-DONE"
End Sub
