Attribute VB_Name = "Mod19_文字大小統一"
Option Explicit

Public Function 英字を大文字に(ByVal text As String) As String
    英字を大文字に = UCase$(text)
End Function

Public Function 英字を小文字に(ByVal text As String) As String
    英字を小文字に = LCase$(text)
End Function
