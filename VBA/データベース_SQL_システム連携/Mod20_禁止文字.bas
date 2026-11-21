Attribute VB_Name = "Mod20_禁止文字"
Option Explicit

Public Sub 禁止文字制御文字除去エスケープ()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim c As Range
    For Each c In ws.UsedRange
        If VarType(c.Value) = vbString Then
            c.Value = EscapeText(CStr(c.Value))
        End If
    Next c
End Sub

Private Function EscapeText(ByVal text As String) As String
    text = Replace(text, Chr$(0), "")
    text = Replace(text, Chr$(1), "")
    text = Replace(text, Chr$(13), "")
    text = Replace(text, Chr$(10), "")
    text = Replace(text, Chr$(9), "")
    text = Replace(text, "'", "''")
    text = Replace(text, Chr$(34), """""")
    EscapeText = text
End Function
