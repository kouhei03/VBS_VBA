Attribute VB_Name = "Mod216_数式パターン検証"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 期待計算式パターン行ずれ検証を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「期待計算式パターン行ずれ検証」を選択して実行する。

' ── 期待計算式パターン行ずれ検証 ──
Public Sub 期待計算式パターン行ずれ検証()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 2).HasFormula Then
            If InStr(ws.Cells(r, 2).Formula, CStr(r)) = 0 Then
                ws.Cells(r, 2).Interior.Color = RGB(255, 199, 206)
            End If
        End If
    Next r
End Sub




