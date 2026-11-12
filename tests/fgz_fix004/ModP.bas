Attribute VB_Name = "ModP"
Option Explicit

' Fix 093b 护栏: 两份**真 Public** 同名 → 各自强制 static (LNK2005 原意图不变)。
Public gDuo As Long
