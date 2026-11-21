Attribute VB_Name = "Mod93_シフトルール"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: シフト配置ルール連勤資格人数チェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「シフト配置ルール連勤資格人数チェック」を選択して実行する。

' ── シフト配置ルール連勤資格人数チェック ──
Public Sub シフト配置ルール連勤資格人数チェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim maxDays As Variant
    maxDays = Application.InputBox("連勤上限日数", "シフトルール", 5, , , , , 1)
    If maxDays = False Or Not IsNumeric(maxDays) Then Exit Sub
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 3).Value) = "連勤" And Val(CStr(ws.Cells(r, 4).Value)) > CLng(maxDays) Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub




