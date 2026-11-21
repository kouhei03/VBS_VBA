Attribute VB_Name = "Mod62_小計行"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 小計行を挿入を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「小計行を挿入」を選択して実行する。

' ── 小計行を挿入 ──
Public Sub 小計行を挿入()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.UsedRange.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(3), _
                          Replace:=True, PageBreaks:=False, SummaryBelowData:=True
End Sub




