Attribute VB_Name = "Mod02_最終列取得"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 最終列取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「最終列取得」を選択して実行する。

' ── 最終列取得 ──
Public Function 最終列取得(ByVal ws As Worksheet, Optional ByVal rowNumber As Long = 0) As Long
    Dim found As Range
    If rowNumber > 0 Then
        Set found = ws.Rows(rowNumber).Find(What:="*", LookIn:=xlFormulas, _
                                            SearchDirection:=xlPrevious)
        If Not found Is Nothing Then 最終列取得 = found.Column
    Else
        Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        If Not found Is Nothing Then 最終列取得 = found.Column
    End If
End Function




