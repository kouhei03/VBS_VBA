Attribute VB_Name = "Mod91_勤怠抽出"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 打刻漏れ残業休日出勤抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「打刻漏れ残業休日出勤抽出」を選択して実行する。

' ── 打刻漏れ残業休日出勤抽出 ──
Public Sub 打刻漏れ残業休日出勤抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:D1").Value = Array("氏名", "日付", "区分", "備考")
    Dim r As Long, outRow As Long: outRow = 2
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ws.Cells(r, 2).Value)) = "" Or _
           Trim$(CStr(ws.Cells(r, 3).Value)) = "" Or _
           CDbl(Val(CStr(ws.Cells(r, 4).Value))) > 8 Or _
           CStr(ws.Cells(r, 5).Value) = "休日出勤" Then
            ws.Rows(r).Copy output.Rows(outRow)
            outRow = outRow + 1
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("勤怠アラート")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "勤怠アラート"
    End If
End Function




