Attribute VB_Name = "Mod118_申請番号採番"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 申請管理番号自動採番重複防止を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「申請管理番号自動採番重複防止」を選択して実行する。

' ── 申請管理番号自動採番重複防止 ──
Public Sub 申請管理番号自動採番重複防止()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim prefix As String: prefix = InputBox("番号の接頭辞", "番号採番", "REQ")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then
            Dim candidate As String
            candidate = prefix & "-" & Format$(Now, "yyyymmdd") & "-" & Format$(r - 1, "000")
            ws.Cells(r, 1).Value = candidate
            Exit Sub
        End If
    Next r
End Sub




