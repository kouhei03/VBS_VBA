Attribute VB_Name = "Mod145_外部EXE実行"
Option Explicit
' ファイル処理概要: Windowsシェルを操作を実行する。
' ファイルないコード処理詳細解説: 外部EXEバッチスクリプト実行を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「外部EXEバッチスクリプト実行」を選択して実行する。

' ── 外部EXEバッチスクリプト実行 ──
Public Sub 外部EXEバッチスクリプト実行()
    Dim command As String
    command = InputBox("実行するコマンドまたはフルパスを入力", "外部プログラム実行")
    If Trim$(command) = "" Then Exit Sub
    CreateObject("WScript.Shell").Run command, 1, False
End Sub




