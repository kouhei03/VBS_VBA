Attribute VB_Name = "Mod144_Access取込出力"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、レコードセットをシートへ転記を実行する。
' ファイルないコード処理詳細解説: AccessテーブルExcel取込書出しを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「AccessテーブルExcel取込書出し」を選択して実行する。

' ── AccessテーブルExcel取込書出し ──
Public Sub AccessテーブルExcel取込書出し()
    Dim databasePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Access", "*.accdb;*.mdb"
        If .Show <> -1 Then Exit Sub
        databasePath = .SelectedItems(1)
    End With
    Dim tableName As String: tableName = InputBox("テーブル名", "Access取込")
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath
    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT * FROM [" & tableName & "]", connection
    ActiveSheet.Range("A2").CopyFromRecordset recordset
    Dim i As Long
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
        ActiveSheet.Cells(1, i + 1).Font.Bold = True
    Next i
    recordset.Close
    connection.Close
End Sub




