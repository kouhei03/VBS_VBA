Attribute VB_Name = "Mod16_固定長"
Option Explicit

Public Sub 固定長バイト指定パディング生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim length As Variant
    length = Application.InputBox("各フィールドの固定文字数", "固定長", 20, , , , , 1)
    If length = False Or Not IsNumeric(length) Then Exit Sub
    Dim output As Worksheet: Set output = GetOrCreateSheet
    Dim r As Long
    For r = 1 To lastRow
        output.Cells(r, 1).Value = Left$(CStr(ws.Cells(r, 1).Value) & String$(CLng(length), " "), CLng(length))
    Next r
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("固定長出力")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "固定長出力"
    End If
End Function
