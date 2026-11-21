Attribute VB_Name = "Mod161_請求書"
Option Explicit

Public Sub 請求書自動作成明細転記()
    Dim source As Worksheet: Set source = ActiveSheet
    Dim template As Worksheet: Set template = ChooseSheet("請求書テンプレート")
    If template Is Nothing Then
        MsgBox "請求書テンプレートシートを選択できません。", vbExclamation
        Exit Sub
    End If
    Dim output As Worksheet: Set output = GetOrCreateSheet("請求書")
    template.UsedRange.Copy output.Range("A1")
    Dim lastRow As Long: lastRow = source.Cells(source.Rows.Count, 1).End(xlUp).Row
    source.Range(source.Cells(2, 1), source.Cells(lastRow, 5)).Copy output.Range("A10")
    output.Cells(5, 1).Value = source.Cells(1, 1).Value
    output.Columns("A:H").AutoFit
    MsgBox "請求書を作成しました。", vbInformation
End Sub

Private Function ChooseSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set ChooseSheet = ThisWorkbook.Worksheets(name)
End Function

Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = name
    End If
End Function
