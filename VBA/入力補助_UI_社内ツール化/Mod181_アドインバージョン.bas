Attribute VB_Name = "Mod181_アドインバージョン"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: アドインバージョン更新通知最新版案内を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「アドインバージョン更新通知最新版案内」を選択して実行する。

' ── アドインバージョン更新通知最新版案内 ──
Public Sub アドインバージョン更新通知最新版案内()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.Range("A1").Value = "現在バージョン"
    ws.Range("B1").Value = "1.0.0"
    ws.Range("A2").Value = "更新日"
    ws.Range("B2").Value = Date
    MsgBox "バージョン情報をシートへ記録しました。", vbInformation
End Sub




