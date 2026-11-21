Attribute VB_Name = "Mod03_Oracle"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、レコードセットをシートへ転記を実行する。
' ファイルないコード処理詳細解説: Oracle接続ビュー参照を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「Oracle接続ビュー参照」を選択して実行する。

' ── Oracle接続ビュー参照 ──
Public Sub Oracle接続ビュー参照()
    Dim connectionString As String
    connectionString = InputBox("Oracle接続文字列（例: Provider=OraOLEDB.Oracle;Data Source=...）", "Oracle")
    Dim viewName As String: viewName = InputBox("参照するビュー名", "Oracle")
    If connectionString = "" Or viewName = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    connection.Open connectionString
    recordset.Open "SELECT * FROM " & viewName, connection
    Dim i As Long
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
    Next i
    ActiveSheet.Range("A2").CopyFromRecordset recordset
    recordset.Close
    connection.Close
End Sub




