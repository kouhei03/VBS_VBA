Attribute VB_Name = "Mod158_集計通知連携"
Option Explicit

Public Sub 集計完了通知チャットメール連携()
    Dim summary As String: summary = InputBox("集計完了内容", "完了通知")
    If Trim$(summary) = "" Then Exit Sub
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("通知先メールアドレス", "完了通知")
    mail.Subject = "集計処理完了通知"
    mail.Body = summary
    mail.Save
    Dim data As Object: Set data = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    data.SetText "[集計完了] " & summary
    data.PutInClipboard
    MsgBox "メール下書きとチャット貼付用テキストを作成しました。", vbInformation
End Sub
