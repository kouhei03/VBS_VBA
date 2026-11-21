Attribute VB_Name = "Mod31_ブック保護"
Option Explicit

Public Sub ブック保護(ByVal wb As Workbook, ByVal password As String)
    wb.Protect Password:=password, Structure:=True, Windows:=False
End Sub

Public Sub ブック保護解除(ByVal wb As Workbook, ByVal password As String)
    wb.Unprotect Password:=password
End Sub
