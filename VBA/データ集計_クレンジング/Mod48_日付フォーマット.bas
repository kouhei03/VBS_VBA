Attribute VB_Name = "Mod48_日付フォーマット"
Option Explicit

Public Sub 日付フォーマットを統一()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range: Set rng = Selection
    Dim v As Variant
    For Each v In rng
        If IsDate(v.Value) Then
            v.Value = DateValue(v.Value)
            v.NumberFormat = "yyyy/mm/dd"
        End If
    Next v
End Sub
