Attribute VB_Name = "Mod163_連動リスト"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 連動コンボボックスリストボックス動的更新を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「連動コンボボックスリストボックス動的更新」を選択して実行する。

' ── 連動コンボボックスリストボックス動的更新 ──
Public Sub 連動コンボボックスリストボックス動的更新()
    Dim category As String: category = InputBox("カテゴリを入力", "連動リスト")
    If Trim$(category) = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.Range("A1").Value = "選択カテゴリ"
    ws.Range("B1").Value = category
    ws.Range("A2").Value = "連動候補"
    ws.Range("B2").Value = "候補A"
    ws.Range("B3").Value = "候補B"
End Sub




