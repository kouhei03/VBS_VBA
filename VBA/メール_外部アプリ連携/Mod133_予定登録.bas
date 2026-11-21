Attribute VB_Name = "Mod133_予定登録"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: Outlook予定表予定登録を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「Outlook予定表予定登録」を選択して実行する。

' ── Outlook予定表予定登録 ──
Public Sub Outlook予定表予定登録()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim appointment As Object: Set appointment = outlook.CreateItem(1)
    appointment.Subject = InputBox("予定の件名", "予定登録")
    appointment.Start = CDate(InputBox("開始日時（例: 2026/12/01 10:00）", "予定登録"))
    appointment.End = CDate(InputBox("終了日時（例: 2026/12/01 11:00）", "予定登録"))
    appointment.Location = InputBox("場所", "予定登録")
    appointment.Body = InputBox("内容", "予定登録")
    appointment.ReminderSet = True
    appointment.Save
    MsgBox "予定を登録しました。", vbInformation
End Sub




