Attribute VB_Name = "Mod51_複数条件合計"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 複数条件で合計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数条件で合計」を選択して実行する。

' ── 複数条件で合計 ──
Public Sub 複数条件で合計()
    Dim rng As Range: Set rng = ActiveSheet.UsedRange
    Dim sumCol As Long: sumCol = CLng(InputBox("合計列番号", "複数条件合計", "3"))
    Dim criteriaCol As Long: criteriaCol = CLng(InputBox("条件列番号", "複数条件合計", "1"))
    Dim criterion As String: criterion = InputBox("条件値", "複数条件合計")
    Dim total As Double
    Dim r As Long
    For r = 2 To GetLastRow(ActiveSheet)
        If CStr(ActiveSheet.Cells(r, criteriaCol).Value) = criterion Then
            If IsNumeric(ActiveSheet.Cells(r, sumCol).Value) Then
                total = total + CDbl(ActiveSheet.Cells(r, sumCol).Value)
            End If
        End If
    Next r
    MsgBox "合計: " & Format$(total, "#,##0"), vbInformation
End Sub




