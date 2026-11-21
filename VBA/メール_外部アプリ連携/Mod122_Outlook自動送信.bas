Attribute VB_Name = "Mod122_Outlook自動送信"
Option Explicit

Public Sub Outlookメール自動送信()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("宛先を入力してください", "メール送信")
    mail.Subject = InputBox("件名を入力してください", "メール送信")
    mail.Body = InputBox("本文を入力してください", "メール送信")
    If MsgBox("メールを送信しますか？", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    mail.Send
    MsgBox "送信を開始しました。", vbInformation
End Sub
