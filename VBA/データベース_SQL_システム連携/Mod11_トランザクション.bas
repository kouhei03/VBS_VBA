Attribute VB_Name = "Mod11_トランザクション"
Option Explicit

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
