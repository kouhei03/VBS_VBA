Attribute VB_Name = "Mod01_AccessADO"
Option Explicit

Public Sub AccessデータベースADO接続データ抽出()
    Dim databasePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Access", "*.accdb;*.mdb"
        If .Show <> -1 Then Exit Sub
        databasePath = .SelectedItems(1)
    End With
    Dim sql As String: sql = InputBox("SELECT文を入力してください", "Access抽出")
    If UCase$(Trim$(Left$(sql, 6))) <> "SELECT" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    connection.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & databasePath
    recordset.Open sql, connection
    WriteRecordsetActiveSheet recordset
    recordset.Close
    connection.Close
End Sub

Private Sub WriteRecordsetActiveSheet(ByVal recordset As Object)
    Dim i As Long
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
    Next i
    ActiveSheet.Range("A2").CopyFromRecordset recordset
End Sub
