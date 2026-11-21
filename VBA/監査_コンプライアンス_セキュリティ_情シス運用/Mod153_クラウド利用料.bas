Attribute VB_Name = "Mod153_クラウド利用料"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: クラウド利用料リソースタグ漏れ異常増減チェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「クラウド利用料リソースタグ漏れ異常増減チェック」を選択して実行する。

' ── クラウド利用料リソースタグ漏れ異常増減チェック ──
Public Sub クラウド利用料リソースタグ漏れ異常増減チェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ws.Cells(r, 3).Value)) = "" Then
            ws.Rows(r).Interior.Color = RGB(255, 235, 156)
        End If
        If IsNumeric(ws.Cells(r, 4).Value) And CDbl(ws.Cells(r, 4).Value) > 100000 Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub




