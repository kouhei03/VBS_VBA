Attribute VB_Name = "Mod102_別名保存"
Option Explicit

Public Sub ブック別名保存日付バージョン付与()
    Dim mode As String
    mode = InputBox("付与形式を入力: 1=日付、2=バージョン、3=両方", "別名保存", "1")
    Dim suffix As String
    If mode = "1" Then suffix = "_" & Format$(Now, "yyyymmdd")
    If mode = "2" Then suffix = "_v" & Format$(1, "00")
    If mode = "3" Then suffix = "_v1_" & Format$(Now, "yyyymmdd")
    If suffix = "" Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim baseName As String: baseName = fso.GetBaseName(ThisWorkbook.Name)
    Dim extension As String: extension = fso.GetExtensionName(ThisWorkbook.Name)
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim destination As String
    destination = fso.BuildPath(destinationFolder, baseName & suffix & "." & extension)
    ThisWorkbook.SaveCopyAs destination
    MsgBox "別名保存しました。" & vbCrLf & destination, vbInformation
End Sub
