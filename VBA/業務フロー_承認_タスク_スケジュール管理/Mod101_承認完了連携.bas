Attribute VB_Name = "Mod101_承認完了連携"
Option Explicit

Public Sub 承認完了承認後工程連携データ作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:F1").Value = Array("案件番号", "次工程", "担当者", "期限", "金額", "ステータス")
    Dim r As Long, outRow As Long: outRow = 2
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 4).Value) = "承認" Then
            ws.Rows(r).Copy output.Rows(outRow)
            output.Cells(outRow, 6).Value = "連携対象"
            outRow = outRow + 1
        End If
    Next r
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("承認後連携")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "承認後連携"
    End If
End Function
