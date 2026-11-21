Attribute VB_Name = "Mod14_動的配列"
Option Explicit

Public Sub 動的配列の再確保()
    Dim values() As String
    ReDim values(0 To 2)
    values(0) = "A"
    values(1) = "B"
    values(2) = "C"

    ReDim Preserve values(0 To 4)
    values(3) = "D"
    values(4) = "E"
End Sub
