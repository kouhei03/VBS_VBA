Attribute VB_Name = "Mod23_数値丸め"
Option Explicit

Public Function 数値を丸める(ByVal value As Double, ByVal digits As Long) As Double
    数値を丸める = Round(value, digits)
End Function

Public Function 数値を切り上げる(ByVal value As Double, ByVal digits As Long) As Double
    Dim factor As Double
    factor = 10 ^ digits
    数値を切り上げる = Application.Ceiling(value * factor, 1) / factor
End Function

Public Function 数値を切り捨てる(ByVal value As Double, ByVal digits As Long) As Double
    Dim factor As Double
    factor = 10 ^ digits
    数値を切り捨てる = Fix(value * factor) / factor
End Function
