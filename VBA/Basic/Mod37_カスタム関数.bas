Attribute VB_Name = "Mod37_カスタム関数"
Option Explicit

' 数式例: =税込金額(A1,0.1)
Public Function 税込金額(ByVal amount As Double, _
                         Optional ByVal taxRate As Double = 0.1) As Double
    税込金額 = amount * (1 + taxRate)
End Function

' 数式例: =文字列結合UDF(A1,B1," ")
Public Function 文字列結合UDF(ByVal firstValue As Variant, _
                             ByVal secondValue As Variant, _
                             Optional ByVal separator As String = "") As String
    文字列結合UDF = CStr(firstValue) & separator & CStr(secondValue)
End Function
