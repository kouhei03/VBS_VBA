Attribute VB_Name = "Mod24_日付演算"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod24_日付演算
' - ファイル名: Mod24_日付演算.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 日付加算(ByVal baseDate As Date, ByVal days As Long) As Date の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 日付加算(ByVal baseDate As Date, ByVal days As Long) As Date
    日付加算 = DateAdd("d", days, baseDate)
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 日付減算(ByVal baseDate As Date, ByVal days As Long) As Date の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 日付減算(ByVal baseDate As Date, ByVal days As Long) As Date
    日付減算 = DateAdd("d", -days, baseDate)
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 日数差(ByVal laterDate As Date, ByVal earlierDate As Date) As Long の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 日数差(ByVal laterDate As Date, ByVal earlierDate As Date) As Long
    日数差 = DateDiff("d", earlierDate, laterDate)
End Function

