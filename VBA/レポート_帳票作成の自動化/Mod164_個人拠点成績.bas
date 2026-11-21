Attribute VB_Name = "Mod164_個人拠点成績"
Option Explicit

Public Sub 個人別拠点別成績評価シート作成()
    Dim source As Worksheet: Set source = ActiveSheet
    Dim mode As String: mode = InputBox("区分列番号を入力（個人または拠点の列）", "成績表")
    If Not IsNumeric(mode) Then Exit Sub
    Dim keyCol As Long: keyCol = CLng(mode)
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long: lastRow = source.Cells(source.Rows.Count, keyCol).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If CStr(source.Cells(r, keyCol).Value) <> "" Then
            groups(CStr(source.Cells(r, keyCol).Value)) = True
        End If
    Next r
    Dim key As Variant
    For Each key In groups.Keys
        Dim ws As Worksheet: Set ws = GetOrCreateSheet(SanitizeName(CStr(key)))
        ws.Cells.Clear
        ws.Cells(1, 1).Value = CStr(key)
        source.Rows(1).Copy ws.Rows(2)
        Dim outRow As Long: outRow = 3
        For r = 2 To lastRow
            If CStr(source.Cells(r, keyCol).Value) = CStr(key) Then
                source.Rows(r).Copy ws.Rows(outRow)
                outRow = outRow + 1
            End If
        Next r
        ws.Columns.AutoFit
    Next key
End Sub

Private Function SanitizeName(ByVal name As String) As String
    Dim chars As Variant: chars = Array(":", "\", "/", "?", "*", "[", "]")
    Dim i As Long
    For i = LBound(chars) To UBound(chars): name = Replace(name, chars(i), "_"): Next i
    If Len(name) > 31 Then name = Left$(name, 31)
    If name = "" Then name = "成績"
    SanitizeName = name
End Function

Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = name
    End If
End Function
