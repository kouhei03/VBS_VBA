Attribute VB_Name = "Mod151_障害影響範囲"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 障害影響範囲対象ユーザー拠点即時抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「障害影響範囲対象ユーザー拠点即時抽出」を選択して実行する。

' ── 障害影響範囲対象ユーザー拠点即時抽出 ──
Public Sub 障害影響範囲対象ユーザー拠点即時抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyword As String: keyword = InputBox("影響システム・障害名", "影響範囲")
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    Dim r As Long, outRow As Long: outRow = 1
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If InStr(1, CStr(ws.Cells(r, 2).Value), keyword, vbTextCompare) > 0 _
           Or InStr(1, CStr(ws.Cells(r, 3).Value), keyword, vbTextCompare) > 0 Then
            ws.Rows(r).Copy output.Rows(outRow)
            outRow = outRow + 1
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("障害影響範囲")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "障害影響範囲"
    End If
End Function




