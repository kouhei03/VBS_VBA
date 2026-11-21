Attribute VB_Name = "Mod12_一次元配列"
Option Explicit

Public Sub 範囲を一次元配列へ読み込み()
    Dim rng As Range
    Set rng = Application.Selection
    If rng Is Nothing Then Exit Sub

    Dim values As Variant
    Dim singleValues() As Variant
    values = rng.Value
    If Not IsArray(values) Then
        ReDim singleValues(0 To 0)
        singleValues(0) = values
    Else
        Dim count As Long
        count = rng.Cells.Count
        ReDim singleValues(0 To count - 1)

        Dim i As Long
        For i = 1 To count
            singleValues(i - 1) = rng.Cells(i).Value
        Next i
    End If
End Sub
