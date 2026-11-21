Attribute VB_Name = "Mod07_処理時間計測"
Option Explicit

Public Sub 処理時間を計測してログ出力()
    Dim started As Double
    started = Timer

    ' ここに処理を記述

    Dim elapsed As Double
    elapsed = Timer - started
    Debug.Print "処理時間: " & Format$(elapsed, "0.000") & " 秒"
    MsgBox "処理時間: " & Format$(elapsed, "0.000") & " 秒", vbInformation
End Sub
