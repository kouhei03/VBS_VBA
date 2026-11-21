Attribute VB_Name = "Mod207_テストマスタセットアップ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: テスト用マスタ初期状態セットアップシート生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テスト用マスタ初期状態セットアップシート生成」を選択して実行する。

' ── テスト用マスタ初期状態セットアップシート生成 ──
Public Sub テスト用マスタ初期状態セットアップシート生成()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("マスタ名", "初期値", "状態")
    ws.Cells(2, 1).Value = "共通マスタ"
    ws.Cells(2, 2).Value = "初期値"
    ws.Cells(2, 3).Value = "未実行"
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("テスト初期状態")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "テスト初期状態"
    End If
End Function




