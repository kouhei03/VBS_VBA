Attribute VB_Name = "Mod24_棚卸突合"
Option Explicit

Public Sub 棚卸システム出力実地棚卸表突合()
    Dim systemWs As Worksheet: Set systemWs = ActiveSheet
    Dim actualWs As Worksheet
    Dim actualName As String: actualName = InputBox("実地棚卸シート名", "棚卸突合")
    On Error Resume Next
    Set actualWs = ThisWorkbook.Worksheets(actualName)
    On Error GoTo 0
    If actualWs Is Nothing Then Exit Sub
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("商品コード", "システム在庫", "実地在庫")
    Dim r As Long
    For r = 2 To systemWs.Cells(systemWs.Rows.Count, 1).End(xlUp).Row
        output.Cells(r, 1).Value = systemWs.Cells(r, 1).Value
        output.Cells(r, 2).Value = systemWs.Cells(r, 2).Value
        output.Cells(r, 3).Value = actualWs.Cells(r, 2).Value
        If output.Cells(r, 2).Value <> output.Cells(r, 3).Value Then
            output.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("棚卸突合")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "棚卸突合"
    End If
End Function
