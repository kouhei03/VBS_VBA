Attribute VB_Name = "Mod150_クリップボード"
Option Explicit

Public Sub クリップボードテキストコピー取得()
    Dim mode As String: mode = InputBox("操作を入力: 1=コピー、2=取得", "クリップボード")
    Dim data As Object: Set data = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If mode = "1" Then
        data.SetText InputBox("コピーするテキスト", "クリップボード")
        data.PutInClipboard
    ElseIf mode = "2" Then
        data.GetFromClipboard
        MsgBox data.GetText, vbInformation
    End If
End Sub
