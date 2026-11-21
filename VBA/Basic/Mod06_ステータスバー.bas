Attribute VB_Name = "Mod06_ステータスバー"
Option Explicit

Public Sub 進捗表示(ByVal currentValue As Long, ByVal totalValue As Long)
    If totalValue <= 0 Then
        Application.StatusBar = "処理中..."
    Else
        Application.StatusBar = "処理中: " & currentValue & " / " & totalValue & _
                                " (" & Format$(currentValue / totalValue, "0%") & ")"
    End If
End Sub

Public Sub 進捗表示終了()
    Application.StatusBar = False
End Sub
