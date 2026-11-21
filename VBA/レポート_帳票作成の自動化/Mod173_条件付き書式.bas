Attribute VB_Name = "Mod173_条件付き書式"
Option Explicit

Public Sub 条件付き書式達成期限超過上下位設定()
    If Selection Is Nothing Then Exit Sub
    With Selection.FormatConditions
        .Delete
        .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .Item(.Count).Interior.Color = RGB(255, 199, 206)
        .Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="100"
        .Item(.Count).Interior.Color = RGB(198, 239, 206)
        .AddTop10
        .Item(.Count).TopBottom.TopBottom = xlTop10Top
        .Item(.Count).Interior.Color = RGB(255, 235, 156)
    End With
    MsgBox "条件付き書式を設定しました。", vbInformation
End Sub
