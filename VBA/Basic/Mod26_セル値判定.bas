Attribute VB_Name = "Mod26_セル値判定"
Option Explicit

Public Function セルが空か判定(ByVal target As Range) As Boolean
    セルが空か判定 = IsEmpty(target.Value) Or Trim$(CStr(target.Value)) = ""
End Function

Public Function セルがNullか判定(ByVal target As Range) As Boolean
    セルがNullか判定 = IsNull(target.Value)
End Function

Public Function セルがエラーか判定(ByVal target As Range) As Boolean
    セルがエラーか判定 = IsError(target.Value)
End Function
