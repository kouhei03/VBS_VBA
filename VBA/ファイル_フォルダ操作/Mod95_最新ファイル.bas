Attribute VB_Name = "Mod95_最新ファイル"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: 最新更新ファイル特定とオープンを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「最新更新ファイル特定とオープン」を選択して実行する。

' ── 最新更新ファイル特定とオープン ──
Public Sub 最新更新ファイル特定とオープン()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim latestFile As Object
    Dim file As Object
    For Each file In fso.GetFolder(folderPath).Files
        If latestFile Is Nothing Or file.DateLastModified > latestFile.DateLastModified Then
            Set latestFile = file
        End If
    Next file
    If latestFile Is Nothing Then
        MsgBox "ファイルがありません。", vbInformation
    Else
        MsgBox "最新ファイル: " & latestFile.Name, vbInformation
        Shell "explorer.exe """ & latestFile.Path & """", vbNormalFocus
    End If
End Sub




