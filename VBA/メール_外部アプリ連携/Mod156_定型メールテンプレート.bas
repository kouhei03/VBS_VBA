Attribute VB_Name = "Mod156_定型メールテンプレート"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 定型案内メールテンプレート切替送信を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「定型案内メールテンプレート切替送信」を選択して実行する。

' ── 定型案内メールテンプレート切替送信 ──
Public Sub 定型案内メールテンプレート切替送信()
    Dim template As String
    template = InputBox("テンプレートを入力: 1=案内、2=催促、3=完了通知", "テンプレートメール", "1")
    Dim subject As String, body As String
    Select Case template
        Case "1": subject = "ご案内"; body = "以下の内容をご案内します。"
        Case "2": subject = "ご確認のお願い"; body = "お手数ですがご確認をお願いいたします。"
        Case "3": subject = "処理完了のお知らせ"; body = "処理が完了しました。"
        Case Else: Exit Sub
    End Select
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("宛先", "テンプレートメール")
    mail.Subject = subject
    mail.Body = body
    If MsgBox("下書き保存または送信しますか？" & vbCrLf & "いいえ=下書き、はい=送信", _
              vbYesNoCancel + vbQuestion) = vbYes Then mail.Send Else mail.Save
End Sub




