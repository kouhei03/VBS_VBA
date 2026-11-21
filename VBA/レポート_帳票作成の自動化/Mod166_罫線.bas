Attribute VB_Name = "Mod166_罫線"
Option Explicit

Public Sub 動的範囲罫線自動描画()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim used As Range: Set used = ws.UsedRange
    With used.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    MsgBox "使用範囲に罫線を描画しました。", vbInformation
End Sub
