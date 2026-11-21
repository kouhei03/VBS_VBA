Attribute VB_Name = "Mod184_ヘッダーフッター"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: ヘッダーフッターページ番号日時社名設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ヘッダーフッターページ番号日時社名設定」を選択して実行する。

' ── ヘッダーフッターページ番号日時社名設定 ──
Public Sub ヘッダーフッターページ番号日時社名設定()
    Dim companyName As String: companyName = InputBox("社名", "ヘッダー/フッター")
    With ActiveSheet.PageSetup
        .LeftHeader = companyName
        .CenterHeader = "&F"
        .RightHeader = "&D &T"
        .LeftFooter = "社名: " & companyName
        .CenterFooter = "Page &P / &N"
    End With
End Sub




