Attribute VB_Name = "Mod22_JOIN抽出"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、レコードセットをシートへ転記を実行する。
' ファイルないコード処理詳細解説: JOIN明細マスタ統合抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「JOIN明細マスタ統合抽出」を選択して実行する。

' ── JOIN明細マスタ統合抽出 ──
Public Sub JOIN明細マスタ統合抽出()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "JOIN抽出")
    Dim sql As String: sql = InputBox("JOINを含むSELECT文を入力", "JOIN抽出")
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




