Attribute VB_Name = "Mod133_予定登録"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod133_予定登録
' - ファイル名: Mod133_予定登録.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub Outlook予定表予定登録() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub Outlook予定表予定登録()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
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

