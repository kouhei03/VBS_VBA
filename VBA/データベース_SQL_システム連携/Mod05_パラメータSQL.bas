Attribute VB_Name = "Mod05_パラメータSQL"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、レコードセットをシートへ転記、ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: パラメータ付きSQL実行を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「パラメータ付きSQL実行」を選択して実行する。

' ── パラメータ付きSQL実行 ──
Public Sub パラメータ付きSQL実行()
    Dim connectionString As String: connectionString = ReadConnectionString()
    Dim sql As String: sql = InputBox("条件列に?を使ったSELECT文を入力", "パラメータSQL")
    Dim parameterValue As String: parameterValue = InputBox("?へ渡す値", "パラメータSQL")
    If connectionString = "" Or sql = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim command As Object: Set command = CreateObject("ADODB.Command")
    Dim parameter As Object
    connection.Open connectionString
    command.ActiveConnection = connection
    command.CommandText = sql
    command.CommandType = 1
    Set parameter = command.CreateParameter("@p1", 200, 1, 255, parameterValue)
    command.Parameters.Append parameter
    Dim recordset As Object: Set recordset = command.Execute
    Dim i As Long
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
    Next i
    ActiveSheet.Range("A2").CopyFromRecordset recordset
    recordset.Close
    connection.Close
End Sub

' ── ReadConnectionString ──
Private Function ReadConnectionString() As String
    On Error Resume Next
    ReadConnectionString = CStr(ThisWorkbook.Worksheets("接続設定").Range("B2").Value)
End Function




