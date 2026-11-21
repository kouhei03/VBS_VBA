Attribute VB_Name = "Mod09_UPSERT"
Option Explicit
' ファイル処理概要: ADOでDBへ接続を実行する。
' ファイルないコード処理詳細解説: UPSERT更新または新規登録を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「UPSERT更新または新規登録」を選択して実行する。

' ── UPSERT更新または新規登録 ──
Public Sub UPSERT更新または新規登録()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "UPSERT")
    Dim sqlTemplate As String
    sqlTemplate = InputBox("キープレースホルダ@KEY、値プレースホルダ@VALUEのUPDATE文またはMERGE文", "UPSERT")
    If connectionString = "" Or sqlTemplate = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    Dim r As Long
    For r = 2 To lastRow
        Dim sql As String: sql = Replace(sqlTemplate, "@KEY", Replace(CStr(ws.Cells(r, 1).Value), "'", "''"))
        sql = Replace(sql, "@VALUE", Replace(CStr(ws.Cells(r, 2).Value), "'", "''"))
        connection.Execute sql
    Next r
    connection.Close
    MsgBox "UPSERT処理を完了しました。", vbInformation
End Sub




