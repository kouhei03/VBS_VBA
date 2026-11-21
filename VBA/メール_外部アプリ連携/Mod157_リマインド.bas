Attribute VB_Name = "Mod157_リマインド"
Option Explicit

Public Sub リマインドメール期限判定自動抽出送信()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim count As Long
    Dim r As Long
    For r = 2 To lastRow
        If IsDate(ws.Cells(r, 3).Value) And CDate(ws.Cells(r, 3).Value) <= Date Then
            Dim mail As Object: Set mail = outlook.CreateItem(0)
            mail.To = CStr(ws.Cells(r, 1).Value)
            mail.Subject = "期限確認: " & CStr(ws.Cells(r, 2).Value)
            mail.Body = "期限日 " & CStr(ws.Cells(r, 3).Text) & " を過ぎています。ご確認ください。"
            If MsgBox("リマインドメールを送信しますか？" & vbCrLf & mail.To, _
                      vbYesNo + vbQuestion) = vbYes Then
                mail.Send
                count = count + 1
            End If
        End If
    Next r
    MsgBox count & " 件送信しました。", vbInformation
End Sub
