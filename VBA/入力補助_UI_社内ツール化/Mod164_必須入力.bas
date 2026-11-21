Attribute VB_Name = "Mod164_必須入力"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 必須入力未入力チェックフォーカス誘導を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「必須入力未入力チェックフォーカス誘導」を選択して実行する。

' ── 必須入力未入力チェックフォーカス誘導 ──
Public Sub 必須入力未入力チェックフォーカス誘導()
    Dim r As Long
    For r = 2 To ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ActiveSheet.Cells(r, 1).Value)) = "" Then
            ActiveSheet.Cells(r, 1).Select
            MsgBox r & "行目の必須項目が未入力です。", vbExclamation
            Exit Sub
        End If
    Next r
    MsgBox "必須項目に未入力はありません。", vbInformation
End Sub




