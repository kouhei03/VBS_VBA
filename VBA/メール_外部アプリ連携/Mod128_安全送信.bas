Attribute VB_Name = "Mod128_安全送信"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 送信前確認ダイアログ安全送信を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「送信前確認ダイアログ安全送信」を選択して実行する。

' ── 送信前確認ダイアログ安全送信 ──
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




