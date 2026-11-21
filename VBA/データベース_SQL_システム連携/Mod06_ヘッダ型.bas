Attribute VB_Name = "Mod06_ヘッダ型"
Option Explicit

Public Sub 抽出ヘッダ自動出力データ型整形()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastColumn As Long: lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastColumn
        ws.Cells(1, c).Font.Bold = True
        ws.Cells(1, c).Interior.Color = RGB(217, 234, 247)
        ws.Columns(c).AutoFit
        If IsDateStringColumn(ws, c) Then ws.Columns(c).NumberFormat = "yyyy/mm/dd"
        If IsNumericColumn(ws, c) Then ws.Columns(c).NumberFormat = "#,##0"
    Next c
End Sub

Private Function IsDateStringColumn(ByVal ws As Worksheet, ByVal c As Long) As Boolean
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, c).Value)) <> "" And Not IsDate(ws.Cells(r, c).Value) Then Exit Function
    Next r
    IsDateStringColumn = True
End Function

Private Function IsNumericColumn(ByVal ws As Worksheet, ByVal c As Long) As Boolean
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, c).Value)) <> "" And Not IsNumeric(ws.Cells(r, c).Value) Then Exit Function
    Next r
    IsNumericColumn = True
End Function
