Attribute VB_Name = "Mod39_引数受け渡し"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 親プロシージャを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「親プロシージャ」を選択して実行する。

' ── 親プロシージャ ──
Public Sub 親プロシージャ()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    子プロシージャ ws, "引き渡した値", 100
End Sub

' ── 子プロシージャ ──
Public Sub 子プロシージャ(ByVal ws As Worksheet, ByVal text As String, ByVal number As Long)
    ws.Range("A1").Value = text
    ws.Range("B1").Value = number
End Sub




