Attribute VB_Name = "Mod81_期限切れタスク"
Option Explicit

Public Sub 期限切れタスク自動抽出強調()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If IsDate(ws.Cells(r, 2).Value) And CDate(ws.Cells(r, 2).Value) < Date _
           And UCase$(CStr(ws.Cells(r, 3).Value)) <> "完了" Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
            ws.Cells(r, 4).Value = "期限切れ"
        End If
    Next r
End Sub
