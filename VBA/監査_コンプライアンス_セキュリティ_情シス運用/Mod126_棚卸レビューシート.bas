Attribute VB_Name = "Mod126_棚卸レビューシート"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 部門別管理者別アクセス棚卸レビュー作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「部門別管理者別アクセス棚卸レビュー作成」を選択して実行する。

' ── 部門別管理者別アクセス棚卸レビュー作成 ──
Public Sub 部門別管理者別アクセス棚卸レビュー作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:E1").Value = Array("部門", "管理者", "ユーザー", "権限", "判定")
    ws.UsedRange.Copy output.Range("A2")
    output.Columns.AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("棚卸レビュー")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "棚卸レビュー"
    End If
End Function




