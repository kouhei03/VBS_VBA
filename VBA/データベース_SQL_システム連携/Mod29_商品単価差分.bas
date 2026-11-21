Attribute VB_Name = "Mod29_商品単価差分"
Option Explicit

Public Sub 商品マスタ単価改定前後差分抽出()
    Dim oldWs As Worksheet: Set oldWs = ActiveSheet
    Dim newName As String: newName = InputBox("改定後シート名", "単価差分")
    Dim newWs As Worksheet
    On Error Resume Next
    Set newWs = ThisWorkbook.Worksheets(newName)
    On Error GoTo 0
    If newWs Is Nothing Then Exit Sub
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:D1").Value = Array("商品コード", "改定前", "改定後", "差額")
    Dim r As Long, outRow As Long: outRow = 2
    For r = 2 To oldWs.Cells(oldWs.Rows.Count, 1).End(xlUp).Row
        Dim key As String: key = CStr(oldWs.Cells(r, 1).Value)
        Dim found As Range
        Set found = newWs.Columns(1).Find(key, , xlValues, xlWhole)
        If Not found Is Nothing Then
            Dim oldPrice As Double: oldPrice = CDbl(oldWs.Cells(r, 2).Value)
            Dim newPrice As Double: newPrice = CDbl(found.Offset(0, 1).Value)
            If oldPrice <> newPrice Then
                output.Cells(outRow, 1).Value = key
                output.Cells(outRow, 2).Value = oldPrice
                output.Cells(outRow, 3).Value = newPrice
                output.Cells(outRow, 4).Value = newPrice - oldPrice
                outRow = outRow + 1
            End If
        End If
    Next r
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("単価差分")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "単価差分"
    End If
End Function
