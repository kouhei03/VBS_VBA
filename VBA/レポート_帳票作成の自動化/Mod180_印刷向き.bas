Attribute VB_Name = "Mod180_印刷向き"
Option Explicit

Public Sub 印刷の向き自動設定()
    Dim mode As String: mode = InputBox("向き: 1=縦、2=横", "印刷設定", "2")
    With ActiveSheet.PageSetup
        If mode = "1" Then .Orientation = xlPortrait Else .Orientation = xlLandscape
    End With
End Sub
