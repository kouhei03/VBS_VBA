Attribute VB_Name = "Mod90_ファイル移動"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: ファイルを移動を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「ファイルを移動」を選択して実行する。

' ── ファイルを移動 ──
Public Sub ファイルを移動()
    Dim sourcePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        sourcePath = .SelectedItems(1)
    End With
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "移動先フォルダを選択"
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(destinationFolder) Then fso.CreateFolder destinationFolder
    Dim destinationPath As String
    destinationPath = fso.BuildPath(destinationFolder, fso.GetFileName(sourcePath))
    If fso.FileExists(destinationPath) Then
        If MsgBox("同名ファイルがあります。上書きしますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    fso.MoveFile sourcePath, destinationPath
    MsgBox "移動しました。", vbInformation
End Sub




