Attribute VB_Name = "Mod128_安全送信"
Option Explicit

Public Sub 送信前確認ダイアログ安全送信()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("宛先を入力してください", "安全送信")
    mail.CC = InputBox("CC（不要なら空欄）", "安全送信")
    mail.Subject = InputBox("件名を入力してください", "安全送信")
    mail.Body = InputBox("本文を入力してください", "安全送信")
    Dim confirm As String
    confirm = "宛先: " & mail.To & vbCrLf & "件名: " & mail.Subject & vbCrLf & "本文を送信しますか？"
    If MsgBox(confirm, vbYesNo + vbQuestion) = vbYes Then mail.Send
End Sub
