Attribute VB_Name = "Mod03_画面更新制御"
Option Explicit

Public Sub 高速処理を開始()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "処理中..."
End Sub

Public Sub 高速処理を終了()
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
