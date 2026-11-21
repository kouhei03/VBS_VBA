Attribute VB_Name = "Mod83_フォルダ存在確認"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: フォルダ存在確認と作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「フォルダ存在確認と作成」を選択して実行する。

' ── フォルダ存在確認と作成 ──
Public Sub フォルダ存在確認と作成()
    Dim folderPath As String
    folderPath = InputBox("確認・作成するフォルダのフルパスを入力してください", "フォルダ作成")
    If Trim$(folderPath) = "" Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        MsgBox "フォルダは既に存在します。" & vbCrLf & folderPath, vbInformation
    Else
        If MsgBox("フォルダを作成しますか？" & vbCrLf & folderPath, vbYesNo + vbQuestion) = vbYes Then
            CreateFolderTree fso, folderPath
            MsgBox "フォルダを作成しました。", vbInformation
        End If
    End If
End Sub

' ── CreateFolderTree ──
Private Sub CreateFolderTree(ByVal fso As Object, ByVal folderPath As String)
    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    If parentPath <> "" And Not fso.FolderExists(parentPath) Then CreateFolderTree fso, parentPath
    fso.CreateFolder folderPath
End Sub




