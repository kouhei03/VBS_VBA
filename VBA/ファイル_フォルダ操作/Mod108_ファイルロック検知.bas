Attribute VB_Name = "Mod108_ファイルロック検知"
Option Explicit

Public Sub ファイルロック使用中状態検知()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileNumber As Integer
    fileNumber = FreeFile
    On Error Resume Next
    Open filePath For Binary Lock Read Write As #fileNumber
    If Err.Number <> 0 Then
        MsgBox "使用中またはロックされています。" & vbCrLf & filePath, vbCritical
    Else
        Close #fileNumber
        MsgBox "ロックされていません。", vbInformation
    End If
    On Error GoTo 0
End Sub
