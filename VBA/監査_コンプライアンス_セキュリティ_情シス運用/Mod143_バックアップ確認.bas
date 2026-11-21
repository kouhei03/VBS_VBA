Attribute VB_Name = "Mod143_バックアップ確認"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: バックアップジョブ成否最終実行容量確認を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「バックアップジョブ成否最終実行容量確認」を選択して実行する。

' ── バックアップジョブ成否最終実行容量確認 ──
Public Sub バックアップジョブ成否最終実行容量確認()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 2).Value) <> "成功" Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
        If IsDate(ws.Cells(r, 3).Value) And CDate(ws.Cells(r, 3).Value) < DateAdd("d", -2, Now) Then
            ws.Cells(r, 4).Value = "最終実行が古い"
        End If
    Next r
End Sub




