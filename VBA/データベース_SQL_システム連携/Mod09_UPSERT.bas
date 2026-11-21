Attribute VB_Name = "Mod09_UPSERT"
Option Explicit

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
