Attribute VB_Name = "Mod17_文字列空白除去"
Option Explicit

Public Function 前後空白を除去(ByVal text As Variant) As String
    If IsError(text) Or IsNull(text) Or IsEmpty(text) Then
        前後空白を除去 = ""
    Else
        前後空白を除去 = Trim$(CStr(text))
    End If
End Function
