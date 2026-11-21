Attribute VB_Name = "Mod37_ステータス更新"
Option Explicit

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
