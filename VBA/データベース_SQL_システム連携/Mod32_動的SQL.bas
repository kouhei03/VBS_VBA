Attribute VB_Name = "Mod32_動的SQL"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: セルから動的SQL生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「セルから動的SQL生成」を選択して実行する。

' ── セルから動的SQL生成 ──
Public Sub セルから動的SQL生成()
    Dim startDate As String: startDate = CStr(ActiveSheet.Range("B2").Value)
    Dim endDate As String: endDate = CStr(ActiveSheet.Range("B3").Value)
    Dim branch As String: branch = CStr(ActiveSheet.Range("B4").Value)
    Dim code As String: code = CStr(ActiveSheet.Range("B5").Value)
    Dim sql As String
    sql = "SELECT * FROM 対象テーブル WHERE 1=1"
    If startDate <> "" Then sql = sql & " AND 日付 >= '" & startDate & "'"
    If endDate <> "" Then sql = sql & " AND 日付 <= '" & endDate & "'"
    If branch <> "" Then sql = sql & " AND 拠点='" & Replace(branch, "'", "''") & "'"
    If code <> "" Then sql = sql & " AND コード='" & Replace(code, "'", "''") & "'"
    ActiveSheet.Range("B7").Value = sql
End Sub




