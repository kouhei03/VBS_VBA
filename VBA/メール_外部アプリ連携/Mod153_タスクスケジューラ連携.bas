Attribute VB_Name = "Mod153_タスクスケジューラ連携"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: タスクスケジューラ連携起動設計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「タスクスケジューラ連携起動設計」を選択して実行する。

' ── タスクスケジューラ連携起動設計 ──
Public Sub タスクスケジューラ連携起動設計()
    Dim scriptPath As String: scriptPath = InputBox("定期実行するVBSまたはBATのパス", "タスクスケジューラ")
    Dim taskName As String: taskName = InputBox("タスク名", "タスクスケジューラ")
    If Trim$(scriptPath) = "" Or Trim$(taskName) = "" Then Exit Sub
    Dim command As String
    command = "schtasks.exe /Create /TN """ & taskName & """ /TR """"" & scriptPath & """" & """ /SC DAILY /ST 09:00"
    MsgBox "以下のschtasksコマンドを管理者権限で実行してください。" & vbCrLf & command, vbInformation
End Sub




