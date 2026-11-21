Attribute VB_Name = "Mod24_日付演算"
Option Explicit

Public Function 日付加算(ByVal baseDate As Date, ByVal days As Long) As Date
    日付加算 = DateAdd("d", days, baseDate)
End Function

Public Function 日付減算(ByVal baseDate As Date, ByVal days As Long) As Date
    日付減算 = DateAdd("d", -days, baseDate)
End Function

Public Function 日数差(ByVal laterDate As Date, ByVal earlierDate As Date) As Long
    日数差 = DateDiff("d", earlierDate, laterDate)
End Function
