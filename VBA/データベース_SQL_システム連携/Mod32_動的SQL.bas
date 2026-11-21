Attribute VB_Name = "Mod32_動的SQL"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod32_動的SQL
' - ファイル名: Mod32_動的SQL.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub セルから動的SQL生成() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub セルから動的SQL生成()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim startDate As String: startDate = CStr(ActiveSheet.Range("B2").Value)
    Dim endDate As String: endDate = CStr(ActiveSheet.Range("B3").Value)
    Dim branch As String: branch = CStr(ActiveSheet.Range("B4").Value)
    Dim code As String: code = CStr(ActiveSheet.Range("B5").Value)
    Dim sql As String
    sql = "SELECT * FROM 対象テーブル WHERE 1=1"
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If startDate <> "" Then sql = sql & " AND 日付 >= '" & startDate & "'"
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If endDate <> "" Then sql = sql & " AND 日付 <= '" & endDate & "'"
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If branch <> "" Then sql = sql & " AND 拠点='" & Replace(branch, "'", "''") & "'"
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If code <> "" Then sql = sql & " AND コード='" & Replace(code, "'", "''") & "'"
    ActiveSheet.Range("B7").Value = sql
End Sub

