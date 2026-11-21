Attribute VB_Name = "Mod142_情報持出証跡"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 秘密保持情報持出外部共有証跡確認を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「秘密保持情報持出外部共有証跡確認」を選択して実行する。

' ── 秘密保持情報持出外部共有証跡確認 ──
Public Sub 秘密保持情報持出外部共有証跡確認()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 3).Value) = "外部共有" Or CStr(ws.Cells(r, 3).Value) = "持ち出し" Then
            ws.Rows(r).Interior.Color = RGB(255, 235, 156)
        End If
    Next r
End Sub




