Attribute VB_Name = "Mod50_先頭ゼロ維持"
Option Explicit

Public Sub 数値コードを文字列化してゼロ維持()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then
            c.NumberFormat = "@"
            c.Value = Format$(CDec(c.Value), "0")
        End If
    Next c
End Sub
