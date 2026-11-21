Attribute VB_Name = "Mod121_Outlook下書き"
Option Explicit

Public Sub Outlookメール下書き作成()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("宛先を入力してください", "下書き作成")
    mail.Subject = InputBox("件名を入力してください", "下書き作成")
    mail.Body = InputBox("本文を入力してください", "下書き作成")
    mail.Save
    MsgBox "Outlookの下書きを作成しました。", vbInformation
End Sub
