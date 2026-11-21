Attribute VB_Name = "Mod93_シフトルール"
Option Explicit

Public Sub シフト配置ルール連勤資格人数チェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim maxDays As Variant
    maxDays = Application.InputBox("連勤上限日数", "シフトルール", 5, , , , , 1)
    If maxDays = False Or Not IsNumeric(maxDays) Then Exit Sub
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 3).Value) = "連勤" And Val(CStr(ws.Cells(r, 4).Value)) > CLng(maxDays) Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub
