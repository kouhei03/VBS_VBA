Attribute VB_Name = "Mod58_差分比較"
Option Explicit

Public Sub 新旧データの差分をハイライト()
    Dim oldName As String: oldName = InputBox("旧データシート名", "差分比較")
    Dim newName As String: newName = InputBox("新データシート名", "差分比較")
    If Not SheetExists(ThisWorkbook, oldName) Or Not SheetExists(ThisWorkbook, newName) Then Exit Sub
    Dim oldWs As Worksheet: Set oldWs = ThisWorkbook.Worksheets(oldName)
    Dim newWs As Worksheet: Set newWs = ThisWorkbook.Worksheets(newName)
    Dim lastRow As Long: lastRow = Application.WorksheetFunction.Max(GetLastRow(oldWs), GetLastRow(newWs))
    Dim lastCol As Long: lastCol = Application.WorksheetFunction.Max(GetLastCol(oldWs), GetLastCol(newWs))
    Dim r As Long, c As Long
    For r = 1 To lastRow
        For c = 1 To lastCol
            If CStr(oldWs.Cells(r, c).Value) <> CStr(newWs.Cells(r, c).Value) Then
                newWs.Cells(r, c).Interior.Color = RGB(255, 230, 153)
            Else
                newWs.Cells(r, c).Interior.Pattern = xlNone
            End If
        Next c
    Next r
End Sub
