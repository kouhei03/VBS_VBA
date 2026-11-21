Attribute VB_Name = "Mod165_テンプレート複製"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: テンプレートシート複製対象名リネームを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テンプレートシート複製対象名リネーム」を選択して実行する。

' ── テンプレートシート複製対象名リネーム ──
Public Sub テンプレートシート複製対象名リネーム()
    Dim source As Worksheet: Set source = ActiveSheet
    Dim targetName As String: targetName = InputBox("複製後のシート名", "テンプレート複製")
    If Trim$(targetName) = "" Then Exit Sub
    Dim copy As Worksheet
    source.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set copy = ActiveSheet
    copy.Name = targetName
    copy.Cells.Replace What:="{対象名}", Replacement:=targetName, LookAt:=xlPart
    MsgBox "テンプレートを複製しました。", vbInformation
End Sub




