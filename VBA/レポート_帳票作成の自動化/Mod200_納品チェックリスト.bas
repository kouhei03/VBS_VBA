Attribute VB_Name = "Mod200_納品チェックリスト"
Option Explicit

Public Sub 納品提出用パッケージ最終チェックリスト作成()
    Dim checklist As Worksheet: Set checklist = GetOrCreateSheet("納品チェックリスト")
    checklist.Cells.Clear
    checklist.Range("A1:D1").Value = Array("シート名", "最終行", "最終列", "確認")
    Dim r As Long: r = 2
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "納品チェックリスト" Then
            checklist.Cells(r, 1).Value = ws.Name
            Dim lastCellRow As Range
            Dim lastCellColumn As Range
            Set lastCellRow = ws.Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious)
            Set lastCellColumn = ws.Cells.Find("*", , xlFormulas, , xlByColumns, xlPrevious)
            Dim lastRow As Long: lastRow = 0
            Dim lastColumn As Long: lastColumn = 0
            If Not lastCellRow Is Nothing Then lastRow = lastCellRow.Row
            If Not lastCellColumn Is Nothing Then lastColumn = lastCellColumn.Column
            checklist.Cells(r, 2).Value = lastRow
            checklist.Cells(r, 3).Value = lastColumn
            checklist.Cells(r, 4).Value = "未確認"
            r = r + 1
        End If
    Next ws
    checklist.Columns("A:D").AutoFit
    checklist.Activate
    MsgBox "納品・提出用チェックリストを作成しました。", vbInformation
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
