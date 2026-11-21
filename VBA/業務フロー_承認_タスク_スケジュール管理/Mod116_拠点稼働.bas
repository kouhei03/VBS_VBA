Attribute VB_Name = "Mod116_拠点稼働"
Option Explicit
' ファイル処理概要: ブック・シートを処理、辞書でキー照合、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 複数拠点チーム稼働状況一覧を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数拠点チーム稼働状況一覧」を選択して実行する。

' ── 複数拠点チーム稼働状況一覧 ──
Public Sub 複数拠点チーム稼働状況一覧()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim key As String: key = CStr(ws.Cells(r, 1).Value) & "/" & CStr(ws.Cells(r, 2).Value)
        If dic.Exists(key) Then dic(key) = CLng(dic(key)) + 1 Else dic(key) = 1
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    Dim outRow As Long: outRow = 1
    Dim item As Variant
    For Each item In dic.Keys
        output.Cells(outRow, 1).Value = item
        output.Cells(outRow, 2).Value = dic(item)
        outRow = outRow + 1
    Next item
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("拠点稼働")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "拠点稼働"
    End If
End Function




