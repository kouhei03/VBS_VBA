Attribute VB_Name = "Mod77_複数シート統合"
Option Explicit

Public Sub 複数シート明細を統合()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim out As Worksheet: Set out = GetOrCreateSheet(wb, "統合明細")
    out.Cells.Clear
    Dim outputRow As Long: outputRow = 1
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name <> out.Name Then
            Dim lastRow As Long: lastRow = GetLastRow(ws)
            Dim lastCol As Long: lastCol = GetLastCol(ws)
            If lastRow > 0 And lastCol > 0 Then
                If outputRow = 1 Then
                    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy out.Cells(outputRow, 1)
                    outputRow = lastRow + 1
                Else
                    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy out.Cells(outputRow, 1)
                    outputRow = outputRow + lastRow - 1
                End If
            End If
        End If
    Next ws
End Sub
