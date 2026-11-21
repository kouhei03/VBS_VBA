Attribute VB_Name = "Mod12_一次元配列"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 範囲を一次元配列へ読み込みを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「範囲を一次元配列へ読み込み」を選択して実行する。

' ── 範囲を一次元配列へ読み込み ──
Public Sub 範囲を一次元配列へ読み込み()
    Dim rng As Range
    Set rng = Application.Selection
    If rng Is Nothing Then Exit Sub

    Dim values As Variant
    Dim singleValues() As Variant
    values = rng.Value
    If Not IsArray(values) Then
        ReDim singleValues(0 To 0)
        singleValues(0) = values
    Else
        Dim count As Long
        count = rng.Cells.Count
        ReDim singleValues(0 To count - 1)

        Dim i As Long
        For i = 1 To count
            singleValues(i - 1) = rng.Cells(i).Value
        Next i
    End If
End Sub




