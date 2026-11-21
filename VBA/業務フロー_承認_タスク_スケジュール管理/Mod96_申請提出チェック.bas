Attribute VB_Name = "Mod96_申請提出チェック"
Option Explicit

Public Sub 申請書起案書提出前必須項目チェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("セル", "項目", "状態")
    Dim r As Long: r = 2
    Dim c As Range
    For Each c In ws.UsedRange
        If c.Row > 1 And Trim$(CStr(c.Value)) = "" Then
            output.Cells(r, 1).Value = c.Address
            output.Cells(r, 2).Value = "必須項目"
            output.Cells(r, 3).Value = "未入力"
            r = r + 1
        End If
    Next c
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("申請提出チェック")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "申請提出チェック"
    End If
End Function
