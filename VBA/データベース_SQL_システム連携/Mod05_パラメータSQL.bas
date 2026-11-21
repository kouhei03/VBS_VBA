Attribute VB_Name = "Mod05_パラメータSQL"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod05_パラメータSQL
' - ファイル名: Mod05_パラメータSQL.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub パラメータ付きSQL実行() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub パラメータ付きSQL実行()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim connectionString As String: connectionString = ReadConnectionString()
    Dim sql As String: sql = InputBox("条件列に?を使ったSELECT文を入力", "パラメータSQL")
    Dim parameterValue As String: parameterValue = InputBox("?へ渡す値", "パラメータSQL")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If connectionString = "" Or sql = "" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim command As Object: Set command = CreateObject("ADODB.Command")
    Dim parameter As Object
    connection.Open connectionString
    command.ActiveConnection = connection
    command.CommandText = sql
    command.CommandType = 1
    Set parameter = command.CreateParameter("@p1", 200, 1, 255, parameterValue)
    command.Parameters.Append parameter
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim recordset As Object: Set recordset = command.Execute
    Dim i As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
    Next i
    ActiveSheet.Range("A2").CopyFromRecordset recordset
    recordset.Close
    connection.Close
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function ReadConnectionString() As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function ReadConnectionString() As String
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ReadConnectionString = CStr(ThisWorkbook.Worksheets("接続設定").Range("B2").Value)
End Function

