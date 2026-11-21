Attribute VB_Name = "Mod14_動的配列"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 動的配列の再確保を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「動的配列の再確保」を選択して実行する。

' ── 動的配列の再確保 ──
Public Sub 動的配列の再確保()
    Dim values() As String
    ReDim values(0 To 2)
    values(0) = "A"
    values(1) = "B"
    values(2) = "C"

    ReDim Preserve values(0 To 4)
    values(3) = "D"
    values(4) = "E"
End Sub




