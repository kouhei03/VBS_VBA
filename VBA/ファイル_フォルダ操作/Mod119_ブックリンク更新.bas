Attribute VB_Name = "Mod119_ブックリンク更新"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 複数ブックリンク更新制御を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数ブックリンク更新制御」を選択して実行する。

' ── 複数ブックリンク更新制御 ──
Public Sub 複数ブックリンク更新制御()
    Dim mode As String
    mode = InputBox("リンク更新を制御: 1=更新する、2=更新しない、3=状態確認", "リンク更新", "2")
    If mode = "1" Then
        Application.AskToUpdateLinks = True
        ThisWorkbook.UpdateLinks = xlUpdateLinksAlways
    ElseIf mode = "2" Then
        Application.AskToUpdateLinks = False
        ThisWorkbook.UpdateLinks = xlUpdateLinksNever
    ElseIf mode = "3" Then
        Dim links As Variant
        links = ThisWorkbook.LinkSources(xlExcelLinks)
        If IsEmpty(links) Then
            MsgBox "外部リンクはありません。", vbInformation
        Else
            MsgBox "ブックの外部リンク数: " & UBound(links) - LBound(links) + 1, vbInformation
        End If
    End If
End Sub




