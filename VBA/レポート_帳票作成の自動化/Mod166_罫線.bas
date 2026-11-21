Attribute VB_Name = "Mod166_罫線"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 動的範囲罫線自動描画を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「動的範囲罫線自動描画」を選択して実行する。

' ── 動的範囲罫線自動描画 ──
Public Sub 動的範囲罫線自動描画()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim used As Range: Set used = ws.UsedRange
    With used.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    MsgBox "使用範囲に罫線を描画しました。", vbInformation
End Sub




