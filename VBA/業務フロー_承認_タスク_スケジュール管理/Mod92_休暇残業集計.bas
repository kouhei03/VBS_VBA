Attribute VB_Name = "Mod92_休暇残業集計"
Option Explicit
' ファイル処理概要: 辞書でキー照合を実行する。
' ファイルないコード処理詳細解説: 休暇申請残業申請集計上限チェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「休暇申請残業申請集計上限チェック」を選択して実行する。

' ── 休暇申請残業申請集計上限チェック ──
Public Sub 休暇申請残業申請集計上限チェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim threshold As Variant
    threshold = Application.InputBox("残業上限時間", "申請上限", 45, , , , , 1)
    If threshold = False Or Not IsNumeric(threshold) Then Exit Sub
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim owner As String: owner = CStr(ws.Cells(r, 1).Value)
        If IsNumeric(ws.Cells(r, 3).Value) Then
            If dic.Exists(owner) Then dic(owner) = CDbl(dic(owner)) + CDbl(ws.Cells(r, 3).Value) Else dic(owner) = CDbl(ws.Cells(r, 3).Value)
        End If
    Next r
    Dim key As Variant
    For Each key In dic.Keys
        For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If CStr(ws.Cells(r, 1).Value) = CStr(key) And CDbl(dic(key)) > CDbl(threshold) Then
                ws.Rows(r).Interior.Color = RGB(255, 199, 206)
            End If
        Next r
    Next key
End Sub




