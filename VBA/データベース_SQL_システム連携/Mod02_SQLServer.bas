Attribute VB_Name = "Mod02_SQLServer"
Option Explicit

Public Sub SQLServer接続SELECTクエリ実行()
    Dim connectionString As String
    connectionString = InputBox("接続文字列を入力", "SQL Server")
    Dim sql As String: sql = InputBox("SELECT文を入力", "SQL Server")
    If connectionString = "" Or UCase$(Left$(Trim$(sql), 6)) <> "SELECT" Then Exit Sub
    ExecuteSelectToSheet connectionString, sql
End Sub

Private Sub ExecuteSelectToSheet(ByVal connectionString As String, ByVal sql As String)
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    Dim recordset As Object: Set recordset = CreateObject("ADODB.Recordset")
    connection.Open connectionString
    recordset.Open sql, connection
    Dim i As Long
    For i = 0 To recordset.Fields.Count - 1
        ActiveSheet.Cells(1, i + 1).Value = recordset.Fields(i).Name
    Next i
    ActiveSheet.Range("A2").CopyFromRecordset recordset
    recordset.Close
    connection.Close
End Sub
