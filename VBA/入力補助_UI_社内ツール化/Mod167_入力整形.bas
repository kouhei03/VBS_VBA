Attribute VB_Name = "Mod167_入力整形"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 入力値自動整形を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「入力値自動整形」を選択して実行する。

' ── 入力値自動整形 ──
Public Sub 入力値自動整形()
    If Selection Is Nothing Then Exit Sub
    Dim c As Range
    For Each c In Selection
        If VarType(c.Value) = vbString Then
            c.Value = UCase$(Trim$(StrConv(CStr(c.Value), vbNarrow)))
        End If
    Next c
End Sub




