Attribute VB_Name = "Mod122_Outlook自動送信"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: Outlookメール自動送信を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「Outlookメール自動送信」を選択して実行する。

' ── Outlookメール自動送信 ──
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




