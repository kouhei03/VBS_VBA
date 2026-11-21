Attribute VB_Name = "Mod36_二重登録チェック"
Option Explicit
' ファイル処理概要: 辞書でキー照合を実行する。
' ファイルないコード処理詳細解説: 二重登録防止キー重複チェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「二重登録防止キー重複チェック」を選択して実行する。

' ── 二重登録防止キー重複チェック ──
Public Sub 二重登録防止キー重複チェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyColumn As Variant
    keyColumn = Application.InputBox("キー列番号", "重複チェック", 1, , , , , 1)
    If keyColumn = False Then Exit Sub
    Dim keys As Object: Set keys = CreateObject("Scripting.Dictionary")
    keys.CompareMode = vbTextCompare
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, CLng(keyColumn)).End(xlUp).Row
        Dim key As String: key = CStr(ws.Cells(r, CLng(keyColumn)).Value)
        If key <> "" Then
            If keys.Exists(key) Then
                ws.Rows(r).Interior.Color = RGB(255, 199, 206)
            Else
                keys.Add key, True
            End If
        End If
    Next r
    MsgBox "重複行を赤色で表示しました。", vbInformation
End Sub




