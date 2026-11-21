Attribute VB_Name = "Mod27_選択範囲情報"
Option Explicit

Public Sub 選択範囲情報を表示()
    Dim rng As Range
    Set rng = Application.Selection
    MsgBox "セル数: " & rng.Cells.Count & vbCrLf & _
           "行数: " & rng.Rows.Count & vbCrLf & _
           "列数: " & rng.Columns.Count, vbInformation
End Sub
