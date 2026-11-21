Attribute VB_Name = "Mod07_一括INSERT"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod07_一括INSERT
' - ファイル名: Mod07_一括INSERT.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub シートからDBテーブルへ一括INSERT() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub シートからDBテーブルへ一括INSERT()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim connectionString As String: connectionString = InputBox("接続文字列", "INSERT")
    Dim tableName As String: tableName = InputBox("登録先テーブル名", "INSERT")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If connectionString = "" Or tableName = "" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastColumn As Long: lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    connection.BeginTrans
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Failed
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To lastRow
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim columns As String: columns = ""
        Dim values As String: values = ""
        Dim c As Long
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For c = 1 To lastColumn
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If c > 1 Then
                columns = columns & ","
                values = values & ","
            End If
            columns = columns & "[" & CStr(ws.Cells(1, c).Value) & "]"
            values = values & "'" & Replace(CStr(ws.Cells(r, c).Value), "'", "''") & "'"
        Next c
        connection.Execute "INSERT INTO [" & tableName & "] (" & columns & ") VALUES (" & values & ")"
    Next r
    connection.CommitTrans
    connection.Close
    MsgBox lastRow - 1 & " 件登録しました。", vbInformation
    Exit Sub
Failed:
    connection.RollbackTrans
    connection.Close
    MsgBox "INSERTでエラー: " & Err.Description, vbCritical
End Sub

