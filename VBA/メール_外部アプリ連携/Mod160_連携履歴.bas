Attribute VB_Name = "Mod160_連携履歴"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod160_連携履歴
' - ファイル名: Mod160_連携履歴.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 連携処理実行履歴成否ステータス記録() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 連携処理実行履歴成否ステータス記録()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "連携履歴")
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = InputBox("処理名を入力", "連携履歴")
    ws.Cells(r, 3).Value = InputBox("ステータスを入力（成功/警告/失敗）", "連携履歴")
    ws.Cells(r, 4).Value = InputBox("詳細を入力", "連携履歴")
    ws.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Columns("A:D").AutoFit
End Sub

