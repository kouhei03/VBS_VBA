Attribute VB_Name = "Mod125_権限申請整合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、辞書でキー照合、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 権限申請承認履歴実付与権限整合確認を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「権限申請承認履歴実付与権限整合確認」を選択して実行する。

' ── 権限申請承認履歴実付与権限整合確認 ──
Public Sub 権限申請承認履歴実付与権限整合確認()
    Dim approved As Worksheet: Set approved = ActiveSheet
    Dim actualName As String: actualName = InputBox("実際の権限シート名", "権限整合")
    Dim actual As Worksheet
    On Error Resume Next
    Set actual = ThisWorkbook.Worksheets(actualName)
    On Error GoTo 0
    If actual Is Nothing Then Exit Sub
    Dim actualRights As Object: Set actualRights = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To actual.Cells(actual.Rows.Count, 1).End(xlUp).Row
        actualRights(CStr(actual.Cells(r, 1).Value) & "|" & CStr(actual.Cells(r, 2).Value)) = True
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("申請者", "権限", "整合")
    Dim outRow As Long: outRow = 2
    For r = 2 To approved.Cells(approved.Rows.Count, 1).End(xlUp).Row
        Dim key As String: key = CStr(approved.Cells(r, 1).Value) & "|" & CStr(approved.Cells(r, 2).Value)
        output.Cells(outRow, 1).Value = approved.Cells(r, 1).Value
        output.Cells(outRow, 2).Value = approved.Cells(r, 2).Value
        output.Cells(outRow, 3).Value = IIf(actualRights.Exists(key), "一致", "不一致")
        outRow = outRow + 1
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("権限整合")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "権限整合"
    End If
End Function




