Attribute VB_Name = "Mod24_日付演算"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 日付加算を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「日付加算」を選択して実行する。

' ── 日付加算 ──
Public Function 日付加算(ByVal baseDate As Date, ByVal days As Long) As Date
    日付加算 = DateAdd("d", days, baseDate)
End Function

' ── 日付減算 ──
Public Function 日付減算(ByVal baseDate As Date, ByVal days As Long) As Date
    日付減算 = DateAdd("d", -days, baseDate)
End Function

' ── 日数差 ──
Public Function 日数差(ByVal laterDate As Date, ByVal earlierDate As Date) As Long
    日数差 = DateDiff("d", earlierDate, laterDate)
End Function




