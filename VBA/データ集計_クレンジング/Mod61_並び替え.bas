Attribute VB_Name = "Mod61_並び替え"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 複数キーで並び替えを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数キーで並び替え」を選択して実行する。

' ── 複数キーで並び替え ──
Public Sub 複数キーで並び替え()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyText As String: keyText = InputBox("並び替え列番号（例: 1,3）", "並び替え")
    If keyText = "" Then Exit Sub
    Dim parts As Variant: parts = Split(keyText, ",")
    If UBound(parts) = 0 Then
        ws.UsedRange.Sort Key1:=ws.Cells(1, CLng(parts(0))), Header:=xlYes
    Else
        ws.UsedRange.Sort Key1:=ws.Cells(1, CLng(parts(0))), _
                          Key2:=ws.Cells(1, CLng(parts(1))), Header:=xlYes
    End If
End Sub




