Attribute VB_Name = "Mod141_規程チェックリスト"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 社内規程セキュリティポリシーチェックリスト集計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「社内規程セキュリティポリシーチェックリスト集計」を選択して実行する。

' ── 社内規程セキュリティポリシーチェックリスト集計 ──
Public Sub 社内規程セキュリティポリシーチェックリスト集計()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim countComplete As Long, countIncomplete As Long
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 3).Value) = "完了" Then countComplete = countComplete + 1 _
        Else countIncomplete = countIncomplete + 1
    Next r
    MsgBox "完了: " & countComplete & vbCrLf & "未完了: " & countIncomplete, vbInformation
End Sub




