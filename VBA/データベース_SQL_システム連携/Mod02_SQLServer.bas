Attribute VB_Name = "Mod02_SQLServer"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、レコードセットをシートへ転記を実行する。
' ファイルないコード処理詳細解説: SQLServer接続SELECTクエリ実行を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「SQLServer接続SELECTクエリ実行」を選択して実行する。

' ── SQLServer接続SELECTクエリ実行 ──
Public Sub SQLServer接続SELECTクエリ実行()
    Dim connectionString As String
    connectionString = InputBox("接続文字列を入力", "SQL Server")
    Dim sql As String: sql = InputBox("SELECT文を入力", "SQL Server")
    If connectionString = "" Or UCase$(Left$(Trim$(sql), 6)) <> "SELECT" Then Exit Sub
    ExecuteSelectToSheet connectionString, sql
End Sub

' ── ExecuteSelectToSheet ──
Private Sub ExecuteSelectToSheet(ByVal connectionString As String, ByVal sql As String)
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    connection.Open connectionString
    recordset.Open sql, connection
    Dim i As Long
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
    Next i
    ActiveSheet.Range("A2").CopyFromRecordset recordset
    recordset.Close
    connection.Close
End Sub




