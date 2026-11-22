Option Explicit

' 出厂计数: Class_Terminate 每响一次加一 (夹具靠它判"这一次真的拆了")。
Private Sub Class_Terminate()
    gTerms = gTerms + 1
End Sub