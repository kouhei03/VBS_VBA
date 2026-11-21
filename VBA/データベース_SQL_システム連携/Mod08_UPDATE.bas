Attribute VB_Name = "Mod08_UPDATE"
Option Explicit

Public Sub 既存レコードUPDATE更新()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "UPDATE")
    Dim tableName As String: tableName = InputBox("更新テーブル名", "UPDATE")
    Dim keyColumn As String: keyColumn = InputBox("キー列名", "UPDATE")
    Dim updateColumn As String: updateColumn = InputBox("更新列名", "UPDATE")
    If connectionString = "" Or tableName = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim keyIndex As Long: keyIndex = FindHeader(ws, keyColumn)
    Dim updateIndex As Long: updateIndex = FindHeader(ws, updateColumn)
    If keyIndex = 0 Or updateIndex = 0 Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    Dim r As Long, count As Long
    For r = 2 To lastRow
        connection.Execute "UPDATE [" & tableName & "] SET [" & updateColumn & "]='" & _
            Replace(CStr(ws.Cells(r, updateIndex).Value), "'", "''") & "' WHERE [" & keyColumn & "]='" & _
            Replace(CStr(ws.Cells(r, keyIndex).Value), "'", "''") & "'", count
    Next r
    connection.Close
    MsgBox "更新処理を完了しました。", vbInformation
End Sub

Private Function FindHeader(ByVal ws As Worksheet, ByVal name As String) As Long
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If CStr(ws.Cells(1, c).Value) = name Then FindHeader = c: Exit Function
    Next c
End Function
