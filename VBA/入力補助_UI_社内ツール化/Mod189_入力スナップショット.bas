Attribute VB_Name = "Mod189_入力スナップショット"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: 入力前スナップショット自動バックアップを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「入力前スナップショット自動バックアップ」を選択して実行する。

' ── 入力前スナップショット自動バックアップ ──
Public Sub 入力前スナップショット自動バックアップ()
    If Not ThisWorkbook.Saved Then
        If ThisWorkbook.Path = "" Then
            MsgBox "保存済みブックのみバックアップできます。", vbExclamation
            Exit Sub
        End If
    End If
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(ThisWorkbook.FullName) Then
        Dim backup As String
        backup = fso.GetParentFolderName(ThisWorkbook.FullName) & "\snapshot_" & _
                 Format$(Now, "yyyymmdd_hhnnss") & "_" & fso.GetFileName(ThisWorkbook.FullName)
        fso.CopyFile ThisWorkbook.FullName, backup, True
        MsgBox "スナップショットを保存しました。" & vbCrLf & backup, vbInformation
    End If
End Sub




