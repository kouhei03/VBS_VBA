Attribute VB_Name = "Mod196_顧客案件ビュー"
Option Explicit

Public Sub 顧客台帳案件管理定型ビュー作成()
    Dim source As Worksheet: Set source = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet("顧客案件ビュー")
    output.Cells.Clear
    source.UsedRange.Copy output.Range("A1")
    output.Rows(1).Font.Bold = True
    output.Rows(1).Interior.Color = RGB(47, 85, 151)
    output.Rows(1).Font.Color = RGB(255, 255, 255)
    output.Columns.AutoFit
    With output.Sort
        .SortFields.Clear
        .SortFields.Add Key:=output.Range("A2"), Order:=xlAscending
        .SetRange output.UsedRange
        .Header = xlYes
        .Apply
    End With
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
