Attribute VB_Name = "Mod25_曜日営業日"
Option Explicit

Public Function 曜日判定(ByVal targetDate As Date) As String
    曜日判定 = Format$(targetDate, "aaaa")
End Function

Public Function 営業日判定(ByVal targetDate As Date, Optional ByVal holidays As Variant) As Boolean
    Dim result As Boolean
    result = (Weekday(targetDate, vbMonday) <= 5)
    If result And Not IsMissing(holidays) Then
        If TypeName(holidays) = "Range" Then
            result = (Application.CountIf(holidays, targetDate) = 0)
        End If
    End If
    営業日判定 = result
End Function

Public Function 営業日加算(ByVal baseDate As Date, ByVal days As Long) As Date
    Dim count As Long
    Dim result As Date
    result = baseDate
    Do While count < days
        result = DateAdd("d", 1, result)
        If 営業日判定(result) Then count = count + 1
    Loop
    営業日加算 = result
End Function
