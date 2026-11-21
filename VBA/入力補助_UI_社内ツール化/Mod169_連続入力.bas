Attribute VB_Name = "Mod169_連続入力"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 連続入力モード登録後次入力へを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「連続入力モード登録後次入力へ」を選択して実行する。

' ── 連続入力モード登録後次入力へ ──
Public Sub 連続入力モード登録後次入力へ()
    ActiveSheet.Cells(ActiveSheet.Cells.Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    MsgBox "次の入力行を選択しました。連続入力モードです。", vbInformation
End Sub




