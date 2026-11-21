Attribute VB_Name = "Mod185_誤操作防止"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 誤操作防止確認ダイアログセル保護を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「誤操作防止確認ダイアログセル保護」を選択して実行する。

' ── 誤操作防止確認ダイアログセル保護 ──
Public Sub 誤操作防止確認ダイアログセル保護()
    If MsgBox("選択範囲を保護しますか？", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    Selection.Locked = True
    ActiveSheet.Protect Password:="locked", UserInterfaceOnly:=True
    MsgBox "選択範囲を保護しました。", vbInformation
End Sub




