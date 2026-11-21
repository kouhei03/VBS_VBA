Attribute VB_Name = "Mod23_集計抽出"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、レコードセットをシートへ転記を実行する。
' ファイルないコード処理詳細解説: 集計クエリ日次月次抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「集計クエリ日次月次抽出」を選択して実行する。

' ── 集計クエリ日次月次抽出 ──
Public Sub 集計クエリ日次月次抽出()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "集計抽出")
    Dim sql As String: sql = InputBox("GROUP BY / SUM / COUNTを含むSELECT文", "集計抽出")
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




