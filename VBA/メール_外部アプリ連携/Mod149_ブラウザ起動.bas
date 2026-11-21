Attribute VB_Name = "Mod149_ブラウザ起動"
Option Explicit
' ファイル処理概要: Windowsシェルを操作を実行する。
' ファイルないコード処理詳細解説: 指定URLブラウザ起動を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「指定URLブラウザ起動」を選択して実行する。

' ── 指定URLブラウザ起動 ──
Public Sub 指定URLブラウザ起動()
    Dim url As String: url = InputBox("開くURL", "ブラウザ起動")
    If Trim$(url) = "" Then Exit Sub
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    shell.Run "explorer.exe """ & url & """", 1, False
End Sub




