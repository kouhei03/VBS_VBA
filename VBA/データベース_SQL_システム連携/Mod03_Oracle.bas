Attribute VB_Name = "Mod03_Oracle"
Option Explicit

Public Sub Oracle接続ビュー参照()
    Dim connectionString As String
    connectionString = InputBox("Oracle接続文字列（例: Provider=OraOLEDB.Oracle;Data Source=...）", "Oracle")
    Dim viewName As String: viewName = InputBox("参照するビュー名", "Oracle")
    If connectionString = "" Or viewName = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    connection.Open connectionString
    recordset.Open "SELECT * FROM " & viewName, connection
    Dim i As Long
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
    Next i
    ActiveSheet.Range("A2").CopyFromRecordset recordset
    recordset.Close
    connection.Close
End Sub
