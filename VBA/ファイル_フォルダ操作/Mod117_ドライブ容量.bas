Attribute VB_Name = "Mod117_ドライブ容量"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: ドライブ空き容量チェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「ドライブ空き容量チェック」を選択して実行する。

' ── ドライブ空き容量チェック ──
Public Sub ドライブ空き容量チェック()
    Dim driveLetter As String
    driveLetter = InputBox("確認するドライブ名を入力（例: C:）", "ドライブ容量", "C:")
    If Trim$(driveLetter) = "" Then Exit Sub
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.DriveExists(driveLetter) Then
        MsgBox "ドライブが見つかりません。", vbCritical
        Exit Sub
    End If
    Dim drive As Object: Set drive = fso.GetDrive(driveLetter)
    Dim thresholdGB As Variant
    thresholdGB = Application.InputBox("警告する空き容量（GB）", "ドライブ容量", 10, , , , , 1)
    If thresholdGB = False Or Not IsNumeric(thresholdGB) Then Exit Sub
    Dim freeGB As Double: freeGB = CDbl(drive.FreeSpace) / 1024 / 1024 / 1024
    If freeGB < CDbl(thresholdGB) Then
        MsgBox "空き容量が少ない可能性があります。" & vbCrLf & Format$(freeGB, "0.00") & " GB", vbCritical
    Else
        MsgBox "空き容量: " & Format$(freeGB, "0.00") & " GB", vbInformation
    End If
End Sub




