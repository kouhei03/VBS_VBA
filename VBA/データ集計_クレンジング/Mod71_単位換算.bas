Attribute VB_Name = "Mod71_単位換算"
Option Explicit

Public Sub 円から千円へ換算()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then c.Value = CDbl(c.Value) / 1000
    Next c
End Sub

Public Sub kgからgへ換算()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then c.Value = CDbl(c.Value) * 1000
    Next c
End Sub
