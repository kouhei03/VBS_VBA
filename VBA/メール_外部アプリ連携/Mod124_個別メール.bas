Attribute VB_Name = "Mod124_個別メール"
Option Explicit

Public Sub 担当者取引先別個別メール生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> "" Then
            Dim mail As Object: Set mail = outlook.CreateItem(0)
            mail.To = CStr(ws.Cells(r, 2).Value)
            mail.Subject = CStr(ws.Cells(r, 3).Value)
            mail.Body = CStr(ws.Cells(r, 1).Value) & "様（" & CStr(ws.Cells(r, 4).Value) & "）" & vbCrLf & CStr(ws.Cells(r, 5).Value)
            mail.Save
        End If
    Next r
    MsgBox "個別メールの下書きを生成しました。", vbInformation
End Sub
