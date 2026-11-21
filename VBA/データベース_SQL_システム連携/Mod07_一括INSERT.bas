Attribute VB_Name = "Mod07_一括INSERT"
Option Explicit

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
