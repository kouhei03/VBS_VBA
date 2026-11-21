Attribute VB_Name = "Mod158_集計通知連携"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 集計完了通知チャットメール連携を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「集計完了通知チャットメール連携」を選択して実行する。

' ── 集計完了通知チャットメール連携 ──
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




