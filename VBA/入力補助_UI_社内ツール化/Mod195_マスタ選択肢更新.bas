Attribute VB_Name = "Mod195_マスタ選択肢更新"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 入力ルール変更マスタ選択肢自動反映を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「入力ルール変更マスタ選択肢自動反映」を選択して実行する。

' ── 入力ルール変更マスタ選択肢自動反映 ──
Public Sub 入力ルール変更マスタ選択肢自動反映()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim masterName As String: masterName = InputBox("選択肢マスタシート名", "選択肢更新")
    Dim master As Worksheet
    On Error Resume Next
    Set master = ThisWorkbook.Worksheets(masterName)
    On Error GoTo 0
    If master Is Nothing Then Exit Sub
    Dim lastRow As Long: lastRow = master.Cells(master.Rows.Count, 1).End(xlUp).Row
    ws.Range("A2:A" & lastRow).Name = "DynamicChoices"
    MsgBox "マスタ選択肢をDynamicChoicesへ反映しました。", vbInformation
End Sub




