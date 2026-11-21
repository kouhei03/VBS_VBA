Attribute VB_Name = "Mod07_処理時間計測"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 処理時間を計測してログ出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「処理時間を計測してログ出力」を選択して実行する。

' ── 処理時間を計測してログ出力 ──
Public Sub 処理時間を計測してログ出力()
    Dim started As Double
    started = Timer

    ' ここに処理を記述

    Dim elapsed As Double
    elapsed = Timer - started
    Debug.Print "処理時間: " & Format$(elapsed, "0.000") & " 秒"
    MsgBox "処理時間: " & Format$(elapsed, "0.000") & " 秒", vbInformation
End Sub




