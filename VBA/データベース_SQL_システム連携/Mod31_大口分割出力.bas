Attribute VB_Name = "Mod31_大口分割出力"
Option Explicit

Public Sub DB抽出大口データシートブック分割()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim pageSize As Variant
    pageSize = Application.InputBox("1シートあたりの行数", "分割出力", 50000, , , , , 1)
    If pageSize = False Or Not IsNumeric(pageSize) Then Exit Sub
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim startRow As Long: startRow = 2
    Dim part As Long: part = 1
    Do While startRow <= lastRow
        Dim output As Worksheet: Set output = ThisWorkbook.Worksheets.Add
        output.Name = "分割_" & Format$(part, "000")
        ws.Rows(1).Copy output.Rows(1)
        ws.Range(startRow & ":" & WorksheetFunction.Min(startRow + CLng(pageSize) - 1, lastRow)).Copy output.Rows(2)
        startRow = startRow + CLng(pageSize)
        part = part + 1
    Loop
End Sub
