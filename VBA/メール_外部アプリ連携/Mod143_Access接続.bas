Attribute VB_Name = "Mod143_Access接続"
Option Explicit
' ファイル処理概要: ADOでDBへ接続を実行する。
' ファイルないコード処理詳細解説: Accessデータベース接続クエリSQL実行を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「Accessデータベース接続クエリSQL実行」を選択して実行する。

' ── Accessデータベース接続クエリSQL実行 ──
Public Sub Accessデータベース接続クエリSQL実行()
    Dim databasePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Access", "*.accdb;*.mdb"
        If .Show <> -1 Then Exit Sub
        databasePath = .SelectedItems(1)
    End With
    Dim sql As String: sql = InputBox("実行するSQLを入力してください", "Access SQL")
    If Trim$(sql) = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath
    connection.Execute sql
    connection.Close
    MsgBox "SQLを実行しました。", vbInformation
End Sub




