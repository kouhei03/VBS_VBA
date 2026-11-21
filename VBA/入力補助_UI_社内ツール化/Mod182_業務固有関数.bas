Attribute VB_Name = "Mod182_業務固有関数"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod182_業務固有関数
' - ファイル名: Mod182_業務固有関数.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 業務コード整形(ByVal value As Variant) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 業務コード整形(ByVal value As Variant) As String
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If VarType(value) = vbString Then
        業務コード整形 = UCase$(Replace(Trim$(CStr(value)), "-", ""))
    End If
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 営業日判定(ByVal value As Variant) As Boolean の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 営業日判定(ByVal value As Variant) As Boolean
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If IsDate(value) Then
        営業日判定 = (Weekday(CDate(value), vbMonday) <= 5)
    End If
End Function

