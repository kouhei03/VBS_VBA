Attribute VB_Name = "Mod107_ネットワーク共有"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: ネットワークドライブ共有フォルダ接続確認を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「ネットワークドライブ共有フォルダ接続確認」を選択して実行する。

' ── ネットワークドライブ共有フォルダ接続確認 ──
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




