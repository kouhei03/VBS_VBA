Attribute VB_Name = "Mod34_クレンジング検証"
Option Explicit

Public Sub 連携前クレンジング検証レポート()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim report As Worksheet: Set report = GetOrCreateSheet
    report.Cells.Clear
    report.Range("A1:C1").Value = Array("セル", "問題", "件数")
    Dim r As Long: r = 2
    Dim cell As Range
    For Each cell In ws.UsedRange
        If VarType(cell.Value) = vbString And Len(CStr(cell.Value)) <> Len(Trim$(CStr(cell.Value))) Then
            report.Cells(r, 1).Value = cell.Address
            report.Cells(r, 2).Value = "前後空白"
            r = r + 1
        End If
    Next cell
    report.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("事前検証")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "事前検証"
    End If
End Function
