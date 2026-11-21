Attribute VB_Name = "Mod60_オートフィルタ"
Option Explicit

Public Sub オートフィルタ設定()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.UsedRange.AutoFilter
End Sub

Public Sub オートフィルタ解除()
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
End Sub
