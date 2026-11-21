Attribute VB_Name = "Mod143_Access接続"
Option Explicit

Public Sub Accessデータベース接続クエリSQL実行()
    Dim databasePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Access", "*.accdb;*.mdb"
        If .Show <> -1 Then Exit Sub
        databasePath = .SelectedItems(1)
    End With
    Dim sql As String: sql = InputBox("実行するSQLを入力してください", "Access SQL")
    If Trim$(sql) = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath
    connection.Execute sql
    connection.Close
    MsgBox "SQLを実行しました。", vbInformation
End Sub
