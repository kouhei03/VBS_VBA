Attribute VB_Name = "Mod10_DELETE"
Option Explicit

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
