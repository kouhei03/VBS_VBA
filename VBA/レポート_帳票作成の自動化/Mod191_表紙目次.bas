Attribute VB_Name = "Mod191_表紙目次"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 表紙目次シート自動生成ハイパーリンクを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「表紙目次シート自動生成ハイパーリンク」を選択して実行する。

' ── 表紙目次シート自動生成ハイパーリンク ──
Public Sub 表紙目次シート自動生成ハイパーリンク()
    Dim indexSheet As Worksheet: Set indexSheet = GetOrCreateSheet("目次")
    indexSheet.Cells.Clear
    indexSheet.Cells(1, 1).Value = "目次"
    indexSheet.Cells(1, 1).Font.Size = 18
    indexSheet.Cells(1, 1).Font.Bold = True
    Dim r As Long: r = 3
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "目次" Then
            indexSheet.Cells(r, 1).Value = ws.Name
            indexSheet.Hyperlinks.Add Anchor:=indexSheet.Cells(r, 1), Address:="", SubAddress:="'" & ws.Name & "'!A1"
            r = r + 1
        End If
    Next ws
    indexSheet.Columns("A:A").AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        GetOrCreateSheet.Name = name
    End If
End Function




