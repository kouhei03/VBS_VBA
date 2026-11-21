Attribute VB_Name = "Mod123_宛先一括メール"
Option Explicit

Public Sub セル指定宛先一括メール作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim mail As Object: Set mail = outlook.CreateItem(0)
        mail.To = CStr(ws.Cells(r, 1).Value)
        mail.CC = CStr(ws.Cells(r, 2).Value)
        mail.BCC = CStr(ws.Cells(r, 3).Value)
        mail.Subject = CStr(ws.Cells(r, 4).Value)
        mail.Body = CStr(ws.Cells(r, 5).Value)
        mail.Save
    Next r
    MsgBox lastRow - 1 & " 件の下書きを作成しました。", vbInformation
End Sub
