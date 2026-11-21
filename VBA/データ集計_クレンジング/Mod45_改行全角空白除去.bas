Attribute VB_Name = "Mod45_改行全角空白除去"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 改行コードと全角空白を除去を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「改行コードと全角空白を除去」を選択して実行する。

' ── 改行コードと全角空白を除去 ──
Public Sub 改行コードと全角空白を除去()
    Dim rng As Range
    Set rng = IIf(TypeName(Selection) = "Range", Selection, ActiveSheet.UsedRange)
    Dim arr As Variant: arr = rng.Value
    Dim r As Long, c As Long
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If VarType(arr(r, c)) = vbString Then
                arr(r, c) = Replace(Replace(Replace(arr(r, c), Chr$(13), ""), Chr$(10), ""), "　", "")
            End If
        Next c
    Next r
    rng.Value = arr
End Sub




