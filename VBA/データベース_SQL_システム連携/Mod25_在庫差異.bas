Attribute VB_Name = "Mod25_在庫差異"
Option Explicit

Public Sub 在庫マスタ入出庫実績差異抽出()
    Dim masterWs As Worksheet: Set masterWs = ActiveSheet
    Dim resultWs As Worksheet: Set resultWs = GetOrCreateSheet
    resultWs.Cells.Clear
    resultWs.Range("A1:C1").Value = Array("商品コード", "マスタ在庫", "実績差異")
    Dim lastRow As Long: lastRow = masterWs.Cells(masterWs.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        resultWs.Cells(r, 1).Value = masterWs.Cells(r, 1).Value
        resultWs.Cells(r, 2).Value = masterWs.Cells(r, 2).Value
        resultWs.Cells(r, 3).Formula = "=B" & r & "-C" & r
    Next r
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("在庫差異")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "在庫差異"
    End If
End Function
