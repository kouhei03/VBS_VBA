Attribute VB_Name = "Mod39_引数受け渡し"
Option Explicit

Public Sub 親プロシージャ()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    子プロシージャ ws, "引き渡した値", 100
End Sub

Public Sub 子プロシージャ(ByVal ws As Worksheet, ByVal text As String, ByVal number As Long)
    ws.Range("A1").Value = text
    ws.Range("B1").Value = number
End Sub
