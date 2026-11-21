Attribute VB_Name = "Mod16_Collection"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: Collectionで可変リスト管理を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「Collectionで可変リスト管理」を選択して実行する。

' ── Collectionで可変リスト管理 ──
Public Sub Collectionで可変リスト管理()
    Dim items As New Collection
    items.Add "りんご"
    items.Add "みかん"
    items.Add "ぶどう"

    Dim i As Long
    For i = 1 To items.Count
        Debug.Print items(i)
    Next i

    items.Remove 2
End Sub




