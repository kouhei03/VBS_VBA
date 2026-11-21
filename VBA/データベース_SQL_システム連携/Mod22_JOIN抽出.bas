Attribute VB_Name = "Mod22_JOIN抽出"
Option Explicit

Public Sub JOIN明細マスタ統合抽出()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "JOIN抽出")
    Dim sql As String: sql = InputBox("JOINを含むSELECT文を入力", "JOIN抽出")
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
