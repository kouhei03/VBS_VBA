Attribute VB_Name = "Mod121_Outlook下書き"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: Outlookメール下書き作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「Outlookメール下書き作成」を選択して実行する。

' ── Outlookメール下書き作成 ──
Public Sub Outlookメール下書き作成()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("宛先を入力してください", "下書き作成")
    mail.Subject = InputBox("件名を入力してください", "下書き作成")
    mail.Body = InputBox("本文を入力してください", "下書き作成")
    mail.Save
    MsgBox "Outlookの下書きを作成しました。", vbInformation
End Sub




