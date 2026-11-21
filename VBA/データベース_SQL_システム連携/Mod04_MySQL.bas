Attribute VB_Name = "Mod04_MySQL"
Option Explicit

Public Sub MySQL接続集計データ取得()
    Dim connectionString As String
    connectionString = InputBox("MySQL接続文字列（Provider=MySQL Provider等）", "MySQL")
    Dim sql As String
    sql = InputBox("集計SELECT文を入力（例: SELECT 部門, COUNT(*) 件数 GROUP BY 部門）", "MySQL")
    If connectionString = "" Or UCase$(Left$(Trim$(sql), 6)) <> "SELECT" Then Exit Sub
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
