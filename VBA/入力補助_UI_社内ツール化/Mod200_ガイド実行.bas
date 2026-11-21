Attribute VB_Name = "Mod200_ガイド実行"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 業務マニュアル操作手順連動ガイド実行を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「業務マニュアル操作手順連動ガイド実行」を選択して実行する。

' ── 業務マニュアル操作手順連動ガイド実行 ──
Public Sub 業務マニュアル操作手順連動ガイド実行()
    Dim guide As String
    guide = InputBox("手順名を入力（例: 月次締め、申請登録、出力）", "ガイド付き実行")
    If Trim$(guide) = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Range("A1").Value = "実行ガイド"
    ws.Range("A2").Value = guide
    ws.Range("A3").Value = "1. 入力データを確認"
    ws.Range("A4").Value = "2. 入力チェックを実行"
    ws.Range("A5").Value = "3. 結果を確認して保存"
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("実行ガイド")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "実行ガイド"
    End If
End Function




