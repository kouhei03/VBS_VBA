Attribute VB_Name = "Mod202_正常系ダミーデータ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 正常系氏名住所日付金額ダミーデータ量産を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「正常系氏名住所日付金額ダミーデータ量産」を選択して実行する。

' ── 正常系氏名住所日付金額ダミーデータ量産 ──
Public Sub 正常系氏名住所日付金額ダミーデータ量産()
    Dim count As Variant: count = Application.InputBox("生成件数", "正常系データ", 100, , , , , 1)
    If count = False Or Not IsNumeric(count) Then Exit Sub
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("氏名", "住所", "日付", "金額")
    Dim i As Long
    For i = 1 To CLng(count)
        ws.Cells(i + 1, 1).Value = "テストユーザー" & i
        ws.Cells(i + 1, 2).Value = "東京都テスト区テスト町" & i & "番地"
        ws.Cells(i + 1, 3).Value = DateAdd("d", i, Date)
        ws.Cells(i + 1, 4).Value = i * 100
    Next i
    ws.Columns("A:D").AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("正常系データ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "正常系データ"
    End If
End Function




