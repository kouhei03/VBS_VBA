Attribute VB_Name = "Mod181_印刷縮尺"
Option Explicit

Public Sub 印刷拡大縮小横1縦自動設定()
    Dim mode As String: mode = InputBox("設定: 1=横1、2=縦自動、3=縮尺指定", "印刷縮尺")
    With ActiveSheet.PageSetup
        If mode = "1" Then
            .FitToPagesWide = 1
            .FitToPagesTall = 0
        ElseIf mode = "2" Then
            .FitToPagesWide = 0
            .FitToPagesTall = 1
        ElseIf mode = "3" Then
            .Zoom = Application.InputBox("縮尺率（%）", "印刷縮尺", 100, , , , , 1)
        End If
    End With
End Sub
