Attribute VB_Name = "Mod170_管理フォーム"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 既存レコード検索呼出編集削除管理フォームを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「既存レコード検索呼出編集削除管理フォーム」を選択して実行する。

' ── 既存レコード検索呼出編集削除管理フォーム ──
Public Sub 既存レコード検索呼出編集削除管理フォーム()
    Dim keyword As String: keyword = InputBox("検索キーワード", "レコード管理")
    If Trim$(keyword) = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If InStr(1, CStr(ws.Cells(r, 1).Value), keyword, vbTextCompare) > 0 Then
            ws.Rows(r).Select
            Exit Sub
        End If
    Next r
    MsgBox "該当レコードが見つかりません。", vbInformation
End Sub




