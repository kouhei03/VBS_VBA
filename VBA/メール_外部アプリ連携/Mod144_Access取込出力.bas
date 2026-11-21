Attribute VB_Name = "Mod144_Access取込出力"
Option Explicit

Public Sub AccessテーブルExcel取込書出し()
    Dim databasePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Access", "*.accdb;*.mdb"
        If .Show <> -1 Then Exit Sub
        databasePath = .SelectedItems(1)
    End With
    Dim tableName As String: tableName = InputBox("テーブル名", "Access取込")
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath
    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT * FROM [" & tableName & "]", connection
    ActiveSheet.Range("A2").CopyFromRecordset recordset
    Dim i As Long
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
        ActiveSheet.Cells(1, i + 1).Font.Bold = True
    Next i
    recordset.Close
    connection.Close
End Sub
