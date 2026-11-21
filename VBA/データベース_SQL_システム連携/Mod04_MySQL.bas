Attribute VB_Name = "Mod04_MySQL"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、レコードセットをシートへ転記を実行する。
' ファイルないコード処理詳細解説: MySQL接続集計データ取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「MySQL接続集計データ取得」を選択して実行する。

' ── MySQL接続集計データ取得 ──
Public Sub MySQL接続集計データ取得()
    Dim connectionString As String
    connectionString = InputBox("MySQL接続文字列（Provider=MySQL Provider等）", "MySQL")
    Dim sql As String
    sql = InputBox("集計SELECT文を入力（例: SELECT 部門, COUNT(*) 件数 GROUP BY 部門）", "MySQL")
    If connectionString = "" Or UCase$(Left$(Trim$(sql), 6)) <> "SELECT" Then Exit Sub
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




