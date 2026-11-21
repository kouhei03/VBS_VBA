Attribute VB_Name = "Mod42_完全一致行抽出削除"
Option Explicit
' ファイル処理概要: 辞書でキー照合を実行する。
' ファイルないコード処理詳細解説: 完全一致行を抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「完全一致行を抽出」を選択して実行する。

' ── 完全一致行を抽出 ──
Public Sub 完全一致行を抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    If lastRow < 3 Then Exit Sub

    Dim out As Worksheet
    Set out = GetOrCreateSheet(ws.Parent, "完全一致行")
    out.Cells.Clear
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim outputRow As Long: outputRow = 2
    out.Rows(1).Value = ws.Rows(1).Value

    Dim r As Long
    For r = 2 To lastRow
        Dim key As String: key = ""
        Dim c As Long
        For c = 1 To lastCol
            key = key & CStr(ws.Cells(r, c).Value) & vbNullChar
        Next c
        If dic.Exists(key) Then
            out.Rows(outputRow).Value = ws.Rows(r).Value
            outputRow = outputRow + 1
        Else
            dic.Add key, r
        End If
    Next r
    MsgBox "完全一致行を抽出しました。", vbInformation
End Sub




