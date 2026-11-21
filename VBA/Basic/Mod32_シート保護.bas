Attribute VB_Name = "Mod32_シート保護"
Option Explicit

Public Sub シート保護(ByVal ws As Worksheet, ByVal password As String)
    ws.Protect Password:=password, DrawingObjects:=True, Contents:=True, _
               Scenarios:=True, AllowFiltering:=True
End Sub

Public Sub シート保護解除(ByVal ws As Worksheet, ByVal password As String)
    ws.Unprotect Password:=password
End Sub
