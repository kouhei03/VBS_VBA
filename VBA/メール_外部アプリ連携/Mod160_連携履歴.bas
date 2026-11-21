Attribute VB_Name = "Mod160_連携履歴"
Option Explicit

Public Sub 連携処理実行履歴成否ステータス記録()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "連携履歴")
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = InputBox("処理名を入力", "連携履歴")
    ws.Cells(r, 3).Value = InputBox("ステータスを入力（成功/警告/失敗）", "連携履歴")
    ws.Cells(r, 4).Value = InputBox("詳細を入力", "連携履歴")
    ws.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Columns("A:D").AutoFit
End Sub
