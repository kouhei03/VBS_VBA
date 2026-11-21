Attribute VB_Name = "Mod13_二次元配列"
Option Explicit

Public Sub 配列を範囲へ一括書き出し()
    Dim values(1 To 3, 1 To 2) As Variant
    values(1, 1) = "商品"
    values(1, 2) = "数量"
    values(2, 1) = "りんご"
    values(2, 2) = 10
    values(3, 1) = "みかん"
    values(3, 2) = 20
    ActiveSheet.Range("A1").Resize(UBound(values, 1), UBound(values, 2)).Value = values
End Sub
