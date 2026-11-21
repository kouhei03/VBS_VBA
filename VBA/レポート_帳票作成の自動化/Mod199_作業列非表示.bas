Attribute VB_Name = "Mod199_作業列非表示"
Option Explicit

Public Sub 印刷不要作業列メモ非表示()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyword As String: keyword = InputBox("非表示にする列見出しキーワード（例: メモ）", "作業列非表示")
    If Trim$(keyword) = "" Then Exit Sub
    Dim lastColumn As Long: lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastColumn
        If InStr(1, CStr(ws.Cells(1, c).Value), keyword, vbTextCompare) > 0 Then
            ws.Columns(c).Hidden = True
        End If
    Next c
End Sub
