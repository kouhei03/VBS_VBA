Attribute VB_Name = "Mod232_帳票表示チェック"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 帳票改ページ印刷範囲表示崩れチェックリストを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「帳票改ページ印刷範囲表示崩れチェックリスト」を選択して実行する。

' ── 帳票改ページ印刷範囲表示崩れチェックリスト ──
Public Sub 帳票改ページ印刷範囲表示崩れチェックリスト()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.Cells(1, 1).Value = "チェック項目"
    ws.Cells(1, 2).Value = "結果"
    ws.Cells(2, 1).Value = "改ページ"
    ws.Cells(3, 1).Value = "印刷範囲"
    ws.Cells(4, 1).Value = "表示崩れ"
End Sub




