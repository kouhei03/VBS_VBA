Attribute VB_Name = "Mod11_トランザクション"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: トランザクション実行を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「トランザクション実行」を選択して実行する。

' ── トランザクション実行 ──
Public Sub トランザクション実行()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "トランザクション")
    If connectionString = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    connection.BeginTrans
    On Error GoTo Failed
    Dim sql As String
    sql = InputBox("実行するSQLを1件入力", "トランザクション")
    connection.Execute sql
    If MsgBox("Commitしますか？いいえでRollbackします。", vbYesNo + vbQuestion) = vbYes Then
        connection.CommitTrans
    Else
        connection.RollbackTrans
    End If
    connection.Close
    Exit Sub
Failed:
    connection.RollbackTrans
    connection.Close
    MsgBox "Rollbackしました: " & Err.Description, vbCritical
End Sub




