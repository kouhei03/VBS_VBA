Attribute VB_Name = "Mod230_権限テストマトリクス"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 権限ロールアクセス可否マトリクス集計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「権限ロールアクセス可否マトリクス集計」を選択して実行する。

' ── 権限ロールアクセス可否マトリクス集計 ──
Public Sub 権限ロールアクセス可否マトリクス集計()
    Dim roles As Variant: roles = Array("管理者", "一般", "閲覧者")
    Dim actions As Variant: actions = Array("登録", "更新", "削除", "閲覧")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    Dim c As Long, r As Long
    For c = LBound(actions) To UBound(actions)
        ws.Cells(1, c + 2).Value = actions(c)
    Next c
    For r = LBound(roles) To UBound(roles)
        ws.Cells(r + 2, 1).Value = roles(r)
        Dim i As Long
        For i = LBound(actions) To UBound(actions)
            ws.Cells(r + 2, i + 2).Value = IIf(roles(r) = "管理者" Or actions(i) = "閲覧", "可", "不可")
        Next i
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("権限テスト")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "権限テスト"
    End If
End Function




