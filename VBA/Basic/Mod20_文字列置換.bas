Attribute VB_Name = "Mod20_文字列置換"
Option Explicit

Public Function 特定文字を置換(ByVal text As String, ByVal searchText As String, _
                               ByVal replaceText As String) As String
    特定文字を置換 = Replace(text, searchText, replaceText)
End Function
