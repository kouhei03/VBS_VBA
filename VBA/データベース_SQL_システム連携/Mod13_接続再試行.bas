Attribute VB_Name = "Mod13_接続再試行"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: DB接続エラータイムアウト再試行ログを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「DB接続エラータイムアウト再試行ログ」を選択して実行する。

' ── DB接続エラータイムアウト再試行ログ ──
Public Sub DB接続エラータイムアウト再試行ログ()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "再試行")
    If connectionString = "" Then Exit Sub
    Dim retries As Long: retries = 3
    Dim i As Long
    For i = 1 To retries
        Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
        connection.ConnectionTimeout = 10
        On Error Resume Next
        connection.Open connectionString
        If Err.Number = 0 And connection.State = 1 Then
            connection.Close
            LogResult "接続成功", "試行 " & i
            MsgBox "接続成功しました。", vbInformation
            Exit Sub
        End If
        LogResult "接続失敗", Err.Description
        Err.Clear
        Application.Wait Now + TimeValue("00:00:02")
    Next i
    MsgBox "再試行後も接続できませんでした。", vbCritical
End Sub

' ── LogResult ──
Private Sub LogResult(ByVal status As String, ByVal detail As String)
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = status
    ws.Cells(r, 3).Value = detail
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("接続ログ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "接続ログ"
    End If
End Function




