Attribute VB_Name = "Mod67_桁数検出"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 桁数不足超過を検出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「桁数不足超過を検出」を選択して実行する。

' ── 桁数不足超過を検出 ──
Public Sub 桁数不足超過を検出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim col As Long: col = CLng(InputBox("対象列番号", "桁数検出", "1"))
    Dim minLen As Long: minLen = CLng(InputBox("最小桁数", "桁数検出", "1"))
    Dim maxLen As Long: maxLen = CLng(InputBox("最大桁数", "桁数検出", "10"))
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        Dim length As Long: length = Len(CStr(ws.Cells(r, col).Value))
        If length < minLen Or length > maxLen Then
            ws.Cells(r, col).Interior.Color = RGB(255, 192, 0)
        End If
    Next r
End Sub




