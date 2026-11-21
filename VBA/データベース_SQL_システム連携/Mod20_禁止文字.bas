Attribute VB_Name = "Mod20_禁止文字"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 禁止文字制御文字除去エスケープを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「禁止文字制御文字除去エスケープ」を選択して実行する。

' ── 禁止文字制御文字除去エスケープ ──
Public Sub 禁止文字制御文字除去エスケープ()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim c As Range
    For Each c In ws.UsedRange
        If VarType(c.Value) = vbString Then
            c.Value = EscapeText(CStr(c.Value))
        End If
    Next c
End Sub

' ── EscapeText ──
Private Function EscapeText(ByVal text As String) As String
    text = Replace(text, Chr$(0), "")
    text = Replace(text, Chr$(1), "")
    text = Replace(text, Chr$(13), "")
    text = Replace(text, Chr$(10), "")
    text = Replace(text, Chr$(9), "")
    text = Replace(text, "'", "''")
    text = Replace(text, Chr$(34), """""")
    EscapeText = text
End Function




