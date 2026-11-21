Attribute VB_Name = "Mod37_カスタム関数"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod37_カスタム関数
' - ファイル名: Mod37_カスタム関数.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' 数式例: =税込金額(A1,0.1)
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 税込金額(ByVal amount As Double, _ の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 税込金額(ByVal amount As Double, _
                         Optional ByVal taxRate As Double = 0.1) As Double
    税込金額 = amount * (1 + taxRate)
End Function

' 数式例: =文字列結合UDF(A1,B1," ")
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 文字列結合UDF(ByVal firstValue As Variant, _ の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 文字列結合UDF(ByVal firstValue As Variant, _
                             ByVal secondValue As Variant, _
                             Optional ByVal separator As String = "") As String
    文字列結合UDF = CStr(firstValue) & separator & CStr(secondValue)
End Function

