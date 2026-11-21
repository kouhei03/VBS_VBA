Attribute VB_Name = "Mod23_数値丸め"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 数値を丸めるを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「数値を丸める」を選択して実行する。

' ── 数値を丸める ──
Public Function 数値を丸める(ByVal value As Double, ByVal digits As Long) As Double
    数値を丸める = Round(value, digits)
End Function

' ── 数値を切り上げる ──
Public Function 数値を切り上げる(ByVal value As Double, ByVal digits As Long) As Double
    Dim factor As Double
    factor = 10 ^ digits
    数値を切り上げる = Application.Ceiling(value * factor, 1) / factor
End Function

' ── 数値を切り捨てる ──
Public Function 数値を切り捨てる(ByVal value As Double, ByVal digits As Long) As Double
    Dim factor As Double
    factor = 10 ^ digits
    数値を切り捨てる = Fix(value * factor) / factor
End Function




