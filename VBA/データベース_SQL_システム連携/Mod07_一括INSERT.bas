Attribute VB_Name = "Mod07_一括INSERT"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: シートからDBテーブルへ一括INSERTを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「シートからDBテーブルへ一括INSERT」を選択して実行する。

' ── シートからDBテーブルへ一括INSERT ──
Public Sub シートからDBテーブルへ一括INSERT()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "INSERT")
    Dim tableName As String: tableName = InputBox("登録先テーブル名", "INSERT")
    If connectionString = "" Or tableName = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastColumn As Long: lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    connection.BeginTrans
    Dim r As Long
    On Error GoTo Failed
    For r = 2 To lastRow
        Dim columns As String: columns = ""
        Dim values As String: values = ""
        Dim c As Long
        For c = 1 To lastColumn
            If c > 1 Then
                columns = columns & ","
                values = values & ","
            End If
            columns = columns & "[" & CStr(ws.Cells(1, c).Value) & "]"
            values = values & "'" & Replace(CStr(ws.Cells(r, c).Value), "'", "''") & "'"
        Next c
        connection.Execute "INSERT INTO [" & tableName & "] (" & columns & ") VALUES (" & values & ")"
    Next r
    connection.CommitTrans
    connection.Close
    MsgBox lastRow - 1 & " 件登録しました。", vbInformation
    Exit Sub
Failed:
    connection.RollbackTrans
    connection.Close
    MsgBox "INSERTでエラー: " & Err.Description, vbCritical
End Sub




