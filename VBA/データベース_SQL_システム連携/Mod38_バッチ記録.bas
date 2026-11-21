Attribute VB_Name = "Mod38_バッチ記録"
Option Explicit

Public Sub 定期バッチ開始終了件数DBログ記録()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Format$(Now, "yyyy/mm/dd hh:mm:ss")
    ws.Cells(r, 2).Value = InputBox("処理名", "バッチ記録")
    ws.Cells(r, 3).Value = InputBox("ステータス（開始/終了/失敗）", "バッチ記録")
    ws.Cells(r, 4).Value = InputBox("処理件数", "バッチ記録")
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("バッチログ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "バッチログ"
    End If
End Function
