Attribute VB_Name = "Mod89_進捗集約"
Option Explicit

Public Sub 複数担当者進捗シート集約全体更新()
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    Dim outRow As Long: outRow = 1
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "全体進捗" Then
            Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 2 Then
                ws.Range("A2:F" & lastRow).Copy output.Cells(outRow, 1)
                outRow = output.Cells(output.Rows.Count, 1).End(xlUp).Row + 1
            End If
        End If
    Next ws
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("全体進捗")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "全体進捗"
    End If
End Function
