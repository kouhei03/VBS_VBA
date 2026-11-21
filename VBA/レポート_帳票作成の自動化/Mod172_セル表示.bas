Attribute VB_Name = "Mod172_セル表示"
Option Explicit

Public Sub セル折り返し縮小表示設定()
    Dim mode As String
    mode = InputBox("設定を入力: 1=折り返し、2=縮小表示", "セル表示")
    If mode = "1" Then
        Selection.WrapText = True
    ElseIf mode = "2" Then
        Selection.WrapText = False
        Selection.ShrinkToFit = True
    End If
End Sub
