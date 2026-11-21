Attribute VB_Name = "Mod154_外部アプリ待機"
Option Explicit
' ファイル処理概要: Windowsシェルを操作を実行する。
' ファイルないコード処理詳細解説: 外部アプリ起動時プロセス待機制御を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「外部アプリ起動時プロセス待機制御」を選択して実行する。

' ── 外部アプリ起動時プロセス待機制御 ──
Public Sub 外部アプリ起動時プロセス待機制御()
    Dim exePath As String: exePath = InputBox("起動するEXEのフルパス", "プロセス待機")
    If Trim$(exePath) = "" Then Exit Sub
    CreateObject("WScript.Shell").Run """" & exePath & """", 1, True
    MsgBox "プロセスが終了しました。", vbInformation
End Sub




