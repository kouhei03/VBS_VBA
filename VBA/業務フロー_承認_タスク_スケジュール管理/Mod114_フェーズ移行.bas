Attribute VB_Name = "Mod114_フェーズ移行"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 案件フェーズ移行チェック項目達成確認を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「案件フェーズ移行チェック項目達成確認」を選択して実行する。

' ── 案件フェーズ移行チェック項目達成確認 ──
Public Sub 案件フェーズ移行チェック項目達成確認()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim complete As Boolean: complete = True
    Dim r As Long
    For r = 2 To lastRow
        If UCase$(CStr(ws.Cells(r, 2).Value)) <> "完了" Then complete = False
    Next r
    ws.Range("E1").Value = "フェーズ移行判定"
    ws.Range("E2").Value = IIf(complete, "移行可", "移行不可")
End Sub




