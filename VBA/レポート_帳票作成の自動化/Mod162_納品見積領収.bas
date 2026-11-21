Attribute VB_Name = "Mod162_納品見積領収"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 納品書見積書領収書自動生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「納品書見積書領収書自動生成」を選択して実行する。

' ── 納品書見積書領収書自動生成 ──
Public Sub 納品書見積書領収書自動生成()
    Dim typeName As String
    typeName = InputBox("帳票種類を入力: 納品書 / 見積書 / 領収書", "帳票生成")
    If typeName <> "納品書" And typeName <> "見積書" And typeName <> "領収書" Then Exit Sub
    Dim output As Worksheet: Set output = GetOrCreateSheet(typeName)
    output.Cells.Clear
    output.Cells(1, 1).Value = typeName
    output.Range("A1:F1").Merge
    output.Range("A1:F1").HorizontalAlignment = xlCenter
    output.Cells(3, 1).Value = "日付"
    output.Cells(3, 2).Value = Date
    output.Cells(3, 2).NumberFormat = "yyyy/mm/dd"
    output.Cells(5, 1).Value = "品名"
    output.Cells(5, 2).Value = "数量"
    output.Cells(5, 3).Value = "単価"
    output.Cells(5, 4).Value = "金額"
    output.Range("A5:D5").Interior.Color = RGB(217, 234, 247)
    output.Columns("A:D").AutoFit
    MsgBox typeName & "を作成しました。", vbInformation
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = name
    End If
End Function




