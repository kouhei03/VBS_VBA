Attribute VB_Name = "Mod146_コマンド結果"
Option Explicit
' ファイル処理概要: Windowsシェルを操作を実行する。
' ファイルないコード処理詳細解説: コマンドプロンプト実行結果取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「コマンドプロンプト実行結果取得」を選択して実行する。

' ── コマンドプロンプト実行結果取得 ──
Public Sub コマンドプロンプト実行結果取得()
    Dim command As String: command = InputBox("実行するコマンドを入力", "コマンド実行")
    If Trim$(command) = "" Then Exit Sub
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim output As String
    output = shell.Exec("%COMSPEC% /c " & command).StdOut.ReadAll
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "コマンド結果")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = output
End Sub




