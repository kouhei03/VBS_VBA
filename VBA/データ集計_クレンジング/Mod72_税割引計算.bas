Attribute VB_Name = "Mod72_税割引計算"
Option Explicit

Public Sub 消費税込み金額を計算()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rate As Double: rate = CDbl(InputBox("税率（例: 10%=0.10）", "消費税", "0.10"))
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then c.Value = CDbl(c.Value) * (1 + rate)
    Next c
End Sub

Public Sub 割引価格を計算()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rate As Double: rate = CDbl(InputBox("割引率（例: 20%=0.20）", "割引計算", "0.20"))
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then c.Value = CDbl(c.Value) * (1 - rate)
    Next c
End Sub
