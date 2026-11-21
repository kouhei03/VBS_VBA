Attribute VB_Name = "Mod166_入力形式検証"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 入力形式半角英数コードメール形式チェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「入力形式半角英数コードメール形式チェック」を選択して実行する。

' ── 入力形式半角英数コードメール形式チェック ──
Public Sub 入力形式半角英数コードメール形式チェック()
    If Selection Is Nothing Then Exit Sub
    Dim text As String: text = CStr(Selection.Value)
    Dim mode As String: mode = InputBox("形式: 1=半角英数、2=コード、3=メール", "形式検証")
    Dim valid As Boolean
    If mode = "1" Then valid = (Len(text) = Len(StrConv(text, vbFromUnicode)))
    If mode = "2" Then valid = (text Like "[A-Z0-9]*")
    If mode = "3" Then valid = (InStr(text, "@") > 0 And InStr(text, ".") > InStr(text, "@"))
    If valid Then MsgBox "形式は正しいです。", vbInformation _
    Else MsgBox "入力形式を確認してください。", vbExclamation
End Sub




