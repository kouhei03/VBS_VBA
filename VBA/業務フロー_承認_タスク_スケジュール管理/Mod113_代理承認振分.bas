Attribute VB_Name = "Mod113_代理承認振分"
Option Explicit

Public Sub 代理承認不在期間タスク振分補助()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim absent As String: absent = InputBox("不在担当者名", "代理振分")
    Dim substitute As String: substitute = InputBox("代理担当者名", "代理振分")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 2).Value) = absent Then
            ws.Cells(r, 2).Value = substitute
            ws.Cells(r, 5).Value = "代理振分"
        End If
    Next r
End Sub
