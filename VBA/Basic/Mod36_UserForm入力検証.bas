Attribute VB_Name = "Mod36_UserForm入力検証"
Option Explicit

Public Function 入力値を検証(ByVal text As String, ByVal allowEmpty As Boolean, _
                              ByVal numericOnly As Boolean) As Boolean
    If text = "" Then
        入力値を検証 = allowEmpty
        Exit Function
    End If
    If numericOnly And Not IsNumeric(text) Then
        入力値を検証 = False
        Exit Function
    End If
    入力値を検証 = True
End Function
