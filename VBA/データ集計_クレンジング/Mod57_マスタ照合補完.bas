Attribute VB_Name = "Mod57_マスタ照合補完"
Option Explicit

Public Sub マスタデータを照合して補完()
    Dim targetName As String: targetName = InputBox("対象シート名", "マスタ照合")
    Dim masterName As String: masterName = InputBox("マスタシート名", "マスタ照合")
    If Not SheetExists(ThisWorkbook, targetName) Or Not SheetExists(ThisWorkbook, masterName) Then Exit Sub
    Dim targetWs As Worksheet: Set targetWs = ThisWorkbook.Worksheets(targetName)
    Dim masterWs As Worksheet: Set masterWs = ThisWorkbook.Worksheets(masterName)
    Dim masterMap As Object: Set masterMap = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To GetLastRow(masterWs)
        masterMap(CStr(masterWs.Cells(r, 1).Value)) = masterWs.Cells(r, 2).Value
    Next r
    For r = 2 To GetLastRow(targetWs)
        Dim key As String: key = CStr(targetWs.Cells(r, 1).Value)
        If Len(CStr(targetWs.Cells(r, 2).Value)) = 0 And masterMap.Exists(key) Then
            targetWs.Cells(r, 2).Value = masterMap(key)
        End If
    Next r
End Sub
