Attribute VB_Name = "Mod19_取込チェック"
Option Explicit

Public Sub 取込ファイル必須文字数桁数チェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim report As Worksheet: Set report = GetOrCreateSheet
    report.Cells.Clear
    report.Range("A1:C1").Value = Array("セル", "項目", "エラー")
    Dim r As Long: r = 2
    Dim cell As Range
    For Each cell In ws.UsedRange
        If cell.Column = 1 And Trim$(CStr(cell.Value)) = "" And cell.Row > 1 Then
            report.Cells(r, 1).Value = cell.Address
            report.Cells(r, 2).Value = "必須項目"
            report.Cells(r, 3).Value = "空欄"
            r = r + 1
        End If
    Next cell
    report.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("取込チェック")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "取込チェック"
    End If
End Function
