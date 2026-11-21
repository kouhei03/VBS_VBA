Attribute VB_Name = "Mod69_テキスト列分割"
Option Explicit

Public Sub テキスト列を区切り文字で分割()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim separator As String: separator = InputBox("区切り文字", "テキスト分割", " ")
    If separator = "" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        Dim parts As Variant: parts = Split(CStr(c.Value), separator)
        c.Resize(1, UBound(parts) + 1).Value = parts
    Next c
End Sub
