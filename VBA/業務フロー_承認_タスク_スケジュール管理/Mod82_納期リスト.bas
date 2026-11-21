Attribute VB_Name = "Mod82_納期リスト"
Option Explicit

Public Sub 納期締切N日以内リストアップ()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim days As Variant: days = Application.InputBox("何日以内か", "納期リスト", 7, , , , , 1)
    If days = False Or Not IsNumeric(days) Then Exit Sub
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    Dim r As Long, outRow As Long: outRow = 2
    output.Range("A1:D1").Value = Array("案件", "納期", "残日数", "担当者")
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(r, 2).Value) Then
            Dim remaining As Long: remaining = DateDiff("d", Date, CDate(ws.Cells(r, 2).Value))
            If remaining >= 0 And remaining <= CLng(days) Then
                ws.Rows(r).Copy output.Rows(outRow)
                output.Cells(outRow, 3).Value = remaining
                outRow = outRow + 1
            End If
        End If
    Next r
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("納期リスト")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "納期リスト"
    End If
End Function
