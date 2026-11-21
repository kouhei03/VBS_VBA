Attribute VB_Name = "Mod33_段階集計"
Option Explicit
' ファイル処理概要: ADOでDBへ接続を実行する。
' ファイルないコード処理詳細解説: ビュー一時テーブル段階集計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「ビュー一時テーブル段階集計」を選択して実行する。

' ── ビュー一時テーブル段階集計 ──
Public Sub ビュー一時テーブル段階集計()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "段階集計")
    Dim sql As String: sql = InputBox("一時テーブルまたはビューを作成するSQL", "段階集計")
    If connectionString = "" Or sql = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    connection.Execute sql
    connection.Close
    MsgBox "段階集計処理を実行しました。", vbInformation
End Sub




