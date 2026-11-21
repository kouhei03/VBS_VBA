Attribute VB_Name = "Mod05_アラート制御"
Option Explicit

Public Sub アラートを一時停止(ByVal state As Boolean)
    Application.DisplayAlerts = Not state
End Sub

Public Sub アラート制御終了()
    Application.DisplayAlerts = True
End Sub
