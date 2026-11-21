Attribute VB_Name = "Mod12_ストアドプロシージャ"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、レコードセットをシートへ転記を実行する。
' ファイルないコード処理詳細解説: ストアドプロシージャ呼出し戻り値取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「ストアドプロシージャ呼出し戻り値取得」を選択して実行する。

' ── ストアドプロシージャ呼出し戻り値取得 ──
Public Sub ストアドプロシージャ呼出し戻り値取得()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "ストアドプロシージャ")
    Dim procedureName As String: procedureName = InputBox("プロシージャ名", "ストアドプロシージャ")
    Dim parameterValue As String: parameterValue = InputBox("引数（空欄可）", "ストアドプロシージャ")
    If connectionString = "" Or procedureName = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim command As Object: Set command = CreateObject("ADODB.Command")
    connection.Open connectionString
    command.ActiveConnection = connection
    command.CommandText = procedureName
    command.CommandType = 4
    If parameterValue <> "" Then
        Dim parameter As Object
        Set parameter = command.CreateParameter("P1", 200, 1, 255, parameterValue)
        command.Parameters.Append parameter
    End If
    Dim recordset As Object: Set recordset = command.Execute
    If Not recordset Is Nothing Then
        If recordset.State <> 0 Then
            ActiveSheet.Range("A1").CopyFromRecordset recordset
            recordset.Close
        End If
    End If
    connection.Close
End Sub




