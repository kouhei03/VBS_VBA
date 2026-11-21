Attribute VB_Name = "Mod107_ネットワーク共有"
Option Explicit

Public Sub ネットワークドライブ共有フォルダ接続確認()
    Dim sharePath As String
    sharePath = InputBox("確認する共有パスまたはドライブを入力（例: Z: または \\server\share）", "接続確認")
    If Trim$(sharePath) = "" Then Exit Sub
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(sharePath) Then
        MsgBox "接続済みです。" & vbCrLf & sharePath, vbInformation
    Else
        MsgBox "接続またはアクセスができません。" & vbCrLf & sharePath, vbCritical
    End If
End Sub
