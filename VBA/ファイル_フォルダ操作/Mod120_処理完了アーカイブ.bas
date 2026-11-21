Attribute VB_Name = "Mod120_処理完了アーカイブ"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: 処理完了後元データをアーカイブ移動を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「処理完了後元データをアーカイブ移動」を選択して実行する。

' ── 処理完了後元データをアーカイブ移動 ──
Public Sub 処理完了後元データをアーカイブ移動()
    Dim sourceFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "元データフォルダを選択"
        If .Show <> -1 Then Exit Sub
        sourceFolder = .SelectedItems(1)
    End With
    Dim archiveRoot As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "アーカイブ先フォルダを選択"
        If .Show <> -1 Then Exit Sub
        archiveRoot = .SelectedItems(1)
    End With

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim archiveFolder As String
    archiveFolder = fso.BuildPath(archiveRoot, "Archive_" & Format$(Now, "yyyymmdd_hhnnss"))
    fso.CreateFolder archiveFolder
    If MsgBox("処理完了後の元データをアーカイブへ移動しますか？", _
              vbYesNo + vbQuestion) <> vbYes Then
        fso.DeleteFolder archiveFolder
        Exit Sub
    End If
    Dim folder As Object: Set folder = fso.GetFolder(sourceFolder)
    Dim file As Object
    For Each file In folder.Files
        fso.MoveFile file.Path, fso.BuildPath(archiveFolder, file.Name)
    Next file
    MsgBox "元データをアーカイブへ移動しました。" & vbCrLf & archiveFolder, vbInformation
End Sub




