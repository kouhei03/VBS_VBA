Attribute VB_Name = "Mod63_ピボット"
Option Explicit

Public Sub ピボットテーブル更新()
    Dim pt As PivotTable
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    MsgBox "ピボットテーブルを更新しました。", vbInformation
End Sub
