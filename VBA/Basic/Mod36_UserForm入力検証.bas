Attribute VB_Name = "Mod36_UserForm入力検証"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 入力値を検証を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「入力値を検証」を選択して実行する。

' ── 入力値を検証 ──
Public Function 入力値を検証(ByVal text As String, ByVal allowEmpty As Boolean, _
                              ByVal numericOnly As Boolean) As Boolean
    If text = "" Then
        入力値を検証 = allowEmpty
        Exit Function
    End If
    If numericOnly And Not IsNumeric(text) Then
        入力値を検証 = False
        Exit Function
    End If
    入力値を検証 = True
End Function




