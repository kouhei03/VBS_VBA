Attribute VB_Name = "Mod85_リマインド下書き"
Option Explicit

Public Sub 担当者リマインド督促メール下書き生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim outlook As Object: Set outlook = CreateObject("Outlook.Application")
    Dim r As Long, count As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 3).Value) <> "完了" And Trim$(CStr(ws.Cells(r, 4).Value)) <> "" Then
            Dim mail As Object: Set mail = outlook.CreateItem(0)
            mail.To = CStr(ws.Cells(r, 4).Value)
            mail.Subject = "【リマインド】" & CStr(ws.Cells(r, 1).Value)
            mail.Body = CStr(ws.Cells(r, 2).Value) & "様" & vbCrLf & _
                        "タスク「" & CStr(ws.Cells(r, 1).Value) & "」の進捗をご確認ください。"
            mail.Save
            count = count + 1
        End If
    Next r
    MsgBox count & " 件の下書きを作成しました。", vbInformation
End Sub
