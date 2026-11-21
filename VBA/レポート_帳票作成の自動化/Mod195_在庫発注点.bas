Attribute VB_Name = "Mod195_在庫発注点"
Option Explicit

Public Sub 在庫一覧発注点アラート帳票作成()
    Dim source As Worksheet: Set source = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet("在庫一覧")
    output.Cells.Clear
    source.UsedRange.Copy output.Range("A1")
    output.Cells(1, 4).Value = "発注点判定"
    Dim lastRow As Long: lastRow = output.Cells(output.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(output.Cells(r, 2).Value) And IsNumeric(output.Cells(r, 3).Value) Then
            If output.Cells(r, 2).Value <= output.Cells(r, 3).Value Then
                output.Cells(r, 4).Value = "発注"
                output.Rows(r).Interior.Color = RGB(255, 199, 206)
            Else
                output.Cells(r, 4).Value = "十分"
            End If
        End If
    Next r
    output.Columns.AutoFit
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
