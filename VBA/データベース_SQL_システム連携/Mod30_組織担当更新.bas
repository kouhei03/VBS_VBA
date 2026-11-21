Attribute VB_Name = "Mod30_組織担当更新"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 組織部署マスタ変更履歴担当割当更新を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「組織部署マスタ変更履歴担当割当更新」を選択して実行する。

' ── 組織部署マスタ変更履歴担当割当更新 ──
Public Sub 組織部署マスタ変更履歴担当割当更新()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 4).Value) = "変更" Then
            ws.Cells(r, 5).Value = "担当割当更新対象"
        Else
            ws.Cells(r, 5).Value = ""
        End If
    Next r
End Sub




