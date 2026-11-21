Attribute VB_Name = "Mod10_DELETE"
Option Explicit
' ファイル処理概要: ADOでDBへ接続を実行する。
' ファイルないコード処理詳細解説: 条件指定DELETE件数確認を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「条件指定DELETE件数確認」を選択して実行する。

' ── 条件指定DELETE件数確認 ──
Public Sub 条件指定DELETE件数確認()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "DELETE")
    Dim sql As String: sql = InputBox("DELETE文を入力してください", "DELETE")
    If connectionString = "" Or UCase$(Trim$(Left$(sql, 6))) <> "DELETE" Then Exit Sub
    If MsgBox("レコードを削除しますか？" & vbCrLf & sql, vbYesNo + vbCritical) <> vbYes Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    Dim count As Long
    connection.Execute sql, count
    connection.Close
    MsgBox count & " 件削除しました。", vbInformation
End Sub




