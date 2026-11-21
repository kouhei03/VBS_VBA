Attribute VB_Name = "Mod81_期限切れタスク"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 期限切れタスク自動抽出強調を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「期限切れタスク自動抽出強調」を選択して実行する。

' ── 期限切れタスク自動抽出強調 ──
Public Sub 期限切れタスク自動抽出強調()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If IsDate(ws.Cells(r, 2).Value) And CDate(ws.Cells(r, 2).Value) < Date _
           And UCase$(CStr(ws.Cells(r, 3).Value)) <> "完了" Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
            ws.Cells(r, 4).Value = "期限切れ"
        End If
    Next r
End Sub




