Attribute VB_Name = "Mod34_UserForm表示"
Option Explicit

Public Sub UserFormを表示()
    frmBasicSample.Show
End Sub

Public Sub UserFormを非表示()
    On Error Resume Next
    If frmBasicSample.Visible Then frmBasicSample.Hide
    On Error GoTo 0
End Sub
