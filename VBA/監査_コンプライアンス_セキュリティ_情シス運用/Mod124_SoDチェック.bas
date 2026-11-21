Attribute VB_Name = "Mod124_SoDチェック"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: SoD違反過剰権限組合せチェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「SoD違反過剰権限組合せチェック」を選択して実行する。

' ── SoD違反過剰権限組合せチェック ──
Public Sub SoD違反過剰権限組合せチェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim rights As String: rights = "|" & CStr(ws.Cells(r, 2).Value) & "|"
        If InStr(rights, "|承認|") > 0 And InStr(rights, "|実行|") > 0 Then
            ws.Cells(r, 3).Value = "SoD違反候補"
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub




