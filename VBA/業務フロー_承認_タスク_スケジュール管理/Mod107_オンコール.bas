Attribute VB_Name = "Mod107_オンコール"
Option Explicit

Public Sub オンコール当番自動割当確認表()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("日付", "担当者", "確認")
    Dim startDate As Variant: startDate = Application.InputBox("開始日", "オンコール", Date, , , , , 1)
    If startDate = False Then Exit Sub
    Dim people As String: people = InputBox("担当者をカンマ区切りで入力", "オンコール")
    If Trim$(people) = "" Then Exit Sub
    Dim personList As Variant: personList = Split(people, ",")
    Dim i As Long
    For i = 0 To 13
        ws.Cells(i + 2, 1).Value = CDate(startDate) + i
        ws.Cells(i + 2, 2).Value = Trim$(CStr(personList(i Mod (UBound(personList) + 1))))
    Next i
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("オンコール確認表")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "オンコール確認表"
    End If
End Function
