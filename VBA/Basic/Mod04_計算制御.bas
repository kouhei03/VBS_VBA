Attribute VB_Name = "Mod04_計算制御"
Option Explicit

Public Sub 計算を手動にする()
    Application.Calculation = xlCalculationManual
End Sub

Public Sub 計算を自動に戻す(Optional ByVal recalculateNow As Boolean = True)
    Application.Calculation = xlCalculationAutomatic
    If recalculateNow Then Application.CalculateFullRebuild
End Sub
