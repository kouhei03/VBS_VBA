Attribute VB_Name = "Mod18_全角半角変換"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 全角英数を半角にを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「全角英数を半角に」を選択して実行する。

' ── 全角英数を半角に ──
Public Function 全角英数を半角に(ByVal text As String) As String
    全角英数を半角に = StrConv(text, vbNarrow)
End Function

' ── 半角英数を全角に ──
Public Function 半角英数を全角に(ByVal text As String) As String
    半角英数を全角に = StrConv(text, vbWide)
End Function

' ── 半角カナを全角に ──
Public Function 半角カナを全角に(ByVal text As String) As String
    半角カナを全角に = StrConv(text, vbWide + vbKatakana)
End Function




