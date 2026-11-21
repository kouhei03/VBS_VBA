Attribute VB_Name = "Mod123_休眠特権"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 一定期間未利用休眠特権アカウント抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「一定期間未利用休眠特権アカウント抽出」を選択して実行する。

' ── 一定期間未利用休眠特権アカウント抽出 ──
Public Sub 一定期間未利用休眠特権アカウント抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim months As Variant: months = Application.InputBox("未利用月数", "休眠アカウント", 6, , , , , 1)
    If months = False Or Not IsNumeric(months) Then Exit Sub
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(r, 3).Value) Then
            If CDate(ws.Cells(r, 3).Value) < DateAdd("m", -CLng(months), Date) _
               Or CStr(ws.Cells(r, 4).Value) = "特権" Then
                ws.Rows(r).Interior.Color = RGB(255, 199, 206)
            End If
        End If
    Next r
End Sub




