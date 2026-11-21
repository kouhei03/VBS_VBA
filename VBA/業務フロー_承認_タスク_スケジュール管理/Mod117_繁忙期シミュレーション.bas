Attribute VB_Name = "Mod117_繁忙期シミュレーション"
Option Explicit

Public Sub 繁忙期閑散期要員計画シミュレーション補助()
    Dim workload As Variant: workload = Application.InputBox("予測業務量", "要員シミュレーション", 100, , , , , 1)
    Dim productivity As Variant: productivity = Application.InputBox("1人あたり処理能力", "要員シミュレーション", 10, , , , , 1)
    If workload = False Or productivity = False Then Exit Sub
    Dim required As Double
    required = CDbl(workload) / CDbl(productivity)
    ActiveSheet.Range("A1").Value = "必要要員数"
    ActiveSheet.Range("B1").Value = Application.WorksheetFunction.RoundUp(required, 0)
End Sub
