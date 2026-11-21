Attribute VB_Name = "Mod99_差戻理由"
Option Explicit

Public Sub 差戻理由修正コメントログ集計()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim log As Worksheet: Set log = GetOrCreateSheet
    Dim r As Long: r = log.Cells(log.Rows.Count, 1).End(xlUp).Row + 1
    log.Cells(r, 1).Value = Now
    log.Cells(r, 2).Value = CStr(ws.Cells(Selection.Row, 1).Value)
    log.Cells(r, 3).Value = InputBox("差し戻し理由・修正コメント", "差戻ログ")
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("差戻ログ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "差戻ログ"
    End If
End Function
