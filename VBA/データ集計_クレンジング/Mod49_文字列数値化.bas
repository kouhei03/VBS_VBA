Attribute VB_Name = "Mod49_文字列数値化"
Option Explicit

Public Sub 文字列数値を数値化()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(Trim$(CStr(c.Value))) And Len(Trim$(CStr(c.Value))) > 0 Then
            c.Value = CDbl(Replace(Trim$(CStr(c.Value)), ",", ""))
        End If
    Next c
End Sub
