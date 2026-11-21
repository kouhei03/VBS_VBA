Attribute VB_Name = "Mod108_ファイルロック検知"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: ファイルロック使用中状態検知を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「ファイルロック使用中状態検知」を選択して実行する。

' ── ファイルロック使用中状態検知 ──
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




