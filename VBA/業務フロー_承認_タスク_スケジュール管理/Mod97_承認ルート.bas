Attribute VB_Name = "Mod97_承認ルート"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 金額権限規定承認ルート承認者判定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「金額権限規定承認ルート承認者判定」を選択して実行する。

' ── 金額権限規定承認ルート承認者判定 ──
Public Sub 金額権限規定承認ルート承認者判定()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsNumeric(ws.Cells(r, 2).Value) Then
            If CDbl(ws.Cells(r, 2).Value) >= 1000000 Then
                ws.Cells(r, 3).Value = "役員承認"
            ElseIf CDbl(ws.Cells(r, 2).Value) >= 100000 Then
                ws.Cells(r, 3).Value = "部門長承認"
            Else
                ws.Cells(r, 3).Value = "課長承認"
            End If
        End If
    Next r
End Sub




