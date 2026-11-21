Attribute VB_Name = "Mod37_ステータス更新"
Option Explicit
' ファイル処理概要: ADOでDBへ接続を実行する。
' ファイルないコード処理詳細解説: 帳票出力連動DBステータス更新を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「帳票出力連動DBステータス更新」を選択して実行する。

' ── 帳票出力連動DBステータス更新 ──
Public Sub 帳票出力連動DBステータス更新()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "ステータス更新")
    Dim sql As String: sql = InputBox("ステータス更新SQL", "ステータス更新")
    If connectionString = "" Or sql = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    connection.Execute sql
    connection.Close
    MsgBox "ステータスを更新しました。", vbInformation
End Sub




