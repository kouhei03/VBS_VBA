Attribute VB_Name = "Mod01_AccessADO"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod01_AccessADO
' - ファイル名: Mod01_AccessADO.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub AccessデータベースADO接続データ抽出() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub AccessデータベースADO接続データ抽出()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim databasePath As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Access", "*.accdb;*.mdb"
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show <> -1 Then Exit Sub
        databasePath = .SelectedItems(1)
    End With
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim sql As String: sql = InputBox("SELECT文を入力してください", "Access抽出")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If UCase$(Trim$(Left$(sql, 6))) <> "SELECT" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    connection.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & databasePath
    recordset.Open sql, connection
    WriteRecordsetActiveSheet recordset
    recordset.Close
    connection.Close
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub WriteRecordsetActiveSheet(ByVal recordset As Object) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub WriteRecordsetActiveSheet(ByVal recordset As Object)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim i As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
    Next i
    ActiveSheet.Range("A2").CopyFromRecordset recordset
End Sub

