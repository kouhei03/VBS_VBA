Attribute VB_Name = "Mod18_全角半角変換"
Option Explicit

Public Function 全角英数を半角に(ByVal text As String) As String
    全角英数を半角に = StrConv(text, vbNarrow)
End Function

Public Function 半角英数を全角に(ByVal text As String) As String
    半角英数を全角に = StrConv(text, vbWide)
End Function

Public Function 半角カナを全角に(ByVal text As String) As String
    半角カナを全角に = StrConv(text, vbWide + vbKatakana)
End Function
