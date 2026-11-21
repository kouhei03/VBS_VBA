Attribute VB_Name = "Mod12_ストアドプロシージャ"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod12_ストアドプロシージャ
' - ファイル名: Mod12_ストアドプロシージャ.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub ストアドプロシージャ呼出し戻り値取得() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub ストアドプロシージャ呼出し戻り値取得()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim connectionString As String: connectionString = InputBox("接続文字列", "ストアドプロシージャ")
    Dim procedureName As String: procedureName = InputBox("プロシージャ名", "ストアドプロシージャ")
    Dim parameterValue As String: parameterValue = InputBox("引数（空欄可）", "ストアドプロシージャ")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If connectionString = "" Or procedureName = "" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim command As Object: Set command = CreateObject("ADODB.Command")
    connection.Open connectionString
    command.ActiveConnection = connection
    command.CommandText = procedureName
    command.CommandType = 4
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If parameterValue <> "" Then
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim parameter As Object
        Set parameter = command.CreateParameter("P1", 200, 1, 255, parameterValue)
        command.Parameters.Append parameter
    End If
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim recordset As Object: Set recordset = command.Execute
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not recordset Is Nothing Then
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If recordset.State <> 0 Then
            ActiveSheet.Range("A1").CopyFromRecordset recordset
            recordset.Close
        End If
    End If
    connection.Close
End Sub

