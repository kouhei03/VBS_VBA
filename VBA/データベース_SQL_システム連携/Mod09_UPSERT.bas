Attribute VB_Name = "Mod09_UPSERT"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod09_UPSERT
' - ファイル名: Mod09_UPSERT.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub UPSERT更新または新規登録() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub UPSERT更新または新規登録()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim connectionString As String: connectionString = InputBox("接続文字列", "UPSERT")
    Dim sqlTemplate As String
    sqlTemplate = InputBox("キープレースホルダ@KEY、値プレースホルダ@VALUEのUPDATE文またはMERGE文", "UPSERT")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If connectionString = "" Or sqlTemplate = "" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To lastRow
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim sql As String: sql = Replace(sqlTemplate, "@KEY", Replace(CStr(ws.Cells(r, 1).Value), "'", "''"))
        sql = Replace(sql, "@VALUE", Replace(CStr(ws.Cells(r, 2).Value), "'", "''"))
        connection.Execute sql
    Next r
    connection.Close
    MsgBox "UPSERT処理を完了しました。", vbInformation
End Sub

