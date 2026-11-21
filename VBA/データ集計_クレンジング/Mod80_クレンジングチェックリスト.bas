Attribute VB_Name = "Mod80_クレンジングチェックリスト"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod80_クレンジングチェックリスト
' - ファイル名: Mod80_クレンジングチェックリスト.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub クレンジング前後チェックリスト作成() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub クレンジング前後チェックリスト作成()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "クレンジングチェック")
    out.Cells.Clear
    out.Cells(1, 1).Value = "項目"
    out.Cells(1, 2).Value = "件数"
    out.Cells(1, 3).Value = "合計"
    out.Cells(2, 1).Value = "全行"
    out.Cells(2, 2).Value = GetLastRow(ws) - 1
    out.Cells(3, 1).Value = "空白セル"
    out.Cells(3, 2).Value = Application.CountBlank(ws.UsedRange)
    out.Cells(4, 1).Value = "数値合計"
    out.Cells(4, 3).Value = Application.WorksheetFunction.Sum(ws.UsedRange)
End Sub

