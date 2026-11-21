Attribute VB_Name = "Mod32_動的SQL"
Option Explicit

Public Sub セルから動的SQL生成()
    Dim startDate As String: startDate = CStr(ActiveSheet.Range("B2").Value)
    Dim endDate As String: endDate = CStr(ActiveSheet.Range("B3").Value)
    Dim branch As String: branch = CStr(ActiveSheet.Range("B4").Value)
    Dim code As String: code = CStr(ActiveSheet.Range("B5").Value)
    Dim sql As String
    sql = "SELECT * FROM 対象テーブル WHERE 1=1"
    If startDate <> "" Then sql = sql & " AND 日付 >= '" & startDate & "'"
    If endDate <> "" Then sql = sql & " AND 日付 <= '" & endDate & "'"
    If branch <> "" Then sql = sql & " AND 拠点='" & Replace(branch, "'", "''") & "'"
    If code <> "" Then sql = sql & " AND コード='" & Replace(code, "'", "''") & "'"
    ActiveSheet.Range("B7").Value = sql
End Sub
