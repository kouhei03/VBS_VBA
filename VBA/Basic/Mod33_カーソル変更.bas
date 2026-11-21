Attribute VB_Name = "Mod33_カーソル変更"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: カーソルを砂時計にするを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「カーソルを砂時計にする」を選択して実行する。

' ── カーソルを砂時計にする ──
Public Sub カーソルを砂時計にする()
    Application.Cursor = xlWait
End Sub

' ── カーソルを標準に戻す ──
Public Sub カーソルを標準に戻す()
    Application.Cursor = xlDefault
End Sub




