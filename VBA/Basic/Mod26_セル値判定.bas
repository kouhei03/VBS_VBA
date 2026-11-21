Attribute VB_Name = "Mod26_セル値判定"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod26_セル値判定
' - ファイル名: Mod26_セル値判定.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function セルが空か判定(ByVal target As Range) As Boolean の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function セルが空か判定(ByVal target As Range) As Boolean
    セルが空か判定 = IsEmpty(target.Value) Or Trim$(CStr(target.Value)) = ""
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function セルがNullか判定(ByVal target As Range) As Boolean の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function セルがNullか判定(ByVal target As Range) As Boolean
    セルがNullか判定 = IsNull(target.Value)
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function セルがエラーか判定(ByVal target As Range) As Boolean の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function セルがエラーか判定(ByVal target As Range) As Boolean
    セルがエラーか判定 = IsError(target.Value)
End Function

