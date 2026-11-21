Attribute VB_Name = "Mod134_環境設定差異"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 開発検証本番設定パラメータ差異確認を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「開発検証本番設定パラメータ差異確認」を選択して実行する。

' ── 開発検証本番設定パラメータ差異確認 ──
Public Sub 開発検証本番設定パラメータ差異確認()
    Dim env1 As String: env1 = InputBox("比較元シート名", "環境差異")
    Dim env2 As String: env2 = InputBox("比較先シート名", "環境差異")
    Dim ws1 As Worksheet, ws2 As Worksheet
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets(env1)
    Set ws2 = ThisWorkbook.Worksheets(env2)
    On Error GoTo 0
    If ws1 Is Nothing Or ws2 Is Nothing Then Exit Sub
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:D1").Value = Array("項目", env1, env2, "判定")
    Dim r As Long
    For r = 2 To ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
        Dim found As Range
        Set found = ws2.Columns(1).Find(CStr(ws1.Cells(r, 1).Value), , xlValues, xlWhole)
        output.Cells(r, 1).Value = ws1.Cells(r, 1).Value
        output.Cells(r, 2).Value = ws1.Cells(r, 2).Value
        If found Is Nothing Then
            output.Cells(r, 4).Value = "比較先なし"
        Else
            output.Cells(r, 3).Value = found.Offset(0, 1).Value
            output.Cells(r, 4).Value = IIf(CStr(ws1.Cells(r, 2).Value) = CStr(found.Offset(0, 1).Value), "一致", "差異")
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("環境設定差異")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "環境設定差異"
    End If
End Function




