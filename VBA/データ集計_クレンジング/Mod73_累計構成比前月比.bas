Attribute VB_Name = "Mod73_累計構成比前月比"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 累計構成比前月比を算出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「累計構成比前月比を算出」を選択して実行する。

' ── 累計構成比前月比を算出 ──
Public Sub 累計構成比前月比を算出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim valueCol As Long: valueCol = CLng(InputBox("数値列番号", "派生指標", "3"))
    Dim total As Double
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        If IsNumeric(ws.Cells(r, valueCol).Value) Then total = total + CDbl(ws.Cells(r, valueCol).Value)
    Next r
    Dim running As Double
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    ws.Cells(1, lastCol + 1).Value = "累計"
    ws.Cells(1, lastCol + 2).Value = "構成比"
    ws.Cells(1, lastCol + 3).Value = "前月比"
    For r = 2 To GetLastRow(ws)
        If IsNumeric(ws.Cells(r, valueCol).Value) Then
            Dim value As Double: value = CDbl(ws.Cells(r, valueCol).Value)
            running = running + value
            ws.Cells(r, lastCol + 1).Value = running
            ws.Cells(r, lastCol + 2).Value = IIf(total <> 0, value / total, 0)
            ws.Cells(r, lastCol + 2).NumberFormat = "0.0%"
            If r > 2 And IsNumeric(ws.Cells(r - 1, valueCol).Value) And CDbl(ws.Cells(r - 1, valueCol).Value) <> 0 Then
                ws.Cells(r, lastCol + 3).Value = (value / CDbl(ws.Cells(r - 1, valueCol).Value)) - 1
                ws.Cells(r, lastCol + 3).NumberFormat = "0.0%"
            End If
        End If
    Next r
End Sub




