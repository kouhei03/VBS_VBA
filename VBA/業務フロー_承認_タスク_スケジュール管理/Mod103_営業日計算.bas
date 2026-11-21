Attribute VB_Name = "Mod103_営業日計算"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 営業日カレンダー期日納期締切計算を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「営業日カレンダー期日納期締切計算」を選択して実行する。

' ── 営業日カレンダー期日納期締切計算 ──
Public Sub 営業日カレンダー期日納期締切計算()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim baseDate As Variant: baseDate = Application.InputBox("基準日", "営業日計算", Date, , , , , 1)
    Dim addDays As Variant: addDays = Application.InputBox("加算営業日数", "営業日計算", 5, , , , , 1)
    If baseDate = False Or addDays = False Then Exit Sub
    Dim dateValue As Date: dateValue = CDate(baseDate)
    Dim count As Long
    Do While count < CLng(addDays)
        dateValue = dateValue + 1
        If Weekday(dateValue, vbMonday) <= 5 Then count = count + 1
    Loop
    ws.Range("A1").Value = "計算期日"
    ws.Range("B1").Value = dateValue
    ws.Range("B1").NumberFormat = "yyyy/mm/dd"
End Sub




