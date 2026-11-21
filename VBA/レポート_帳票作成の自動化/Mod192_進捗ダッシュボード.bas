Attribute VB_Name = "Mod192_進捗ダッシュボード"
Option Explicit

Public Sub 進捗ダッシュボードKPIカード達成バー更新()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("ダッシュボード")
    ws.Cells.Clear
    ws.Range("A1:C1").Merge
    ws.Cells(1, 1).Value = "進捗ダッシュボード"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(3, 1).Value = "KPI"
    ws.Cells(3, 2).Value = "達成率"
    ws.Cells(3, 3).Value = "進捗バー"
    ws.Cells(4, 1).Value = "全体進捗"
    ws.Cells(4, 2).Value = 0.75
    ws.Cells(4, 2).NumberFormat = "0%"
    ws.Cells(4, 3).Formula = "=REPT(""■"",ROUND(B4*20,0))"
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = name
    End If
End Function
