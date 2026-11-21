Attribute VB_Name = "Mod21_トレーラ"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod21_トレーラ
' - ファイル名: Mod21_トレーラ.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub レコード件数合計金額トレーラ付与() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub レコード件数合計金額トレーラ付与()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Cells(lastRow + 1, 1).Value = "TRAILER"
    ws.Cells(lastRow + 1, 2).Value = lastRow - 1
    ws.Cells(lastRow + 1, 3).Value = Application.WorksheetFunction.Sum(ws.Columns(3))
End Sub

