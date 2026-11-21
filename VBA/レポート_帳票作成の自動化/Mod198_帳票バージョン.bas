Attribute VB_Name = "Mod198_帳票バージョン"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 帳票バージョン作成日時自動印字を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「帳票バージョン作成日時自動印字」を選択して実行する。

' ── 帳票バージョン作成日時自動印字 ──
Public Sub 帳票バージョン作成日時自動印字()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim version As String: version = InputBox("バージョン（例: v1.0）", "帳票バージョン", "v1.0")
    With ws.PageSetup
        .RightHeader = "Ver. " & version
        .RightFooter = "作成日時: " & Format$(Now, "yyyy/mm/dd hh:mm")
    End With
    ws.Cells(1, ws.UsedRange.Columns.Count + 1).Value = "Ver." & version
    ws.Cells(2, ws.UsedRange.Columns.Count).Value = "作成: " & Format$(Now, "yyyy/mm/dd hh:mm")
End Sub




