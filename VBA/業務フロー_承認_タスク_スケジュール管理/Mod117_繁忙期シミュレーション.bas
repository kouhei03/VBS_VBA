Attribute VB_Name = "Mod117_繁忙期シミュレーション"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 繁忙期閑散期要員計画シミュレーション補助を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「繁忙期閑散期要員計画シミュレーション補助」を選択して実行する。

' ── 繁忙期閑散期要員計画シミュレーション補助 ──
Public Sub 繁忙期閑散期要員計画シミュレーション補助()
    Dim workload As Variant: workload = Application.InputBox("予測業務量", "要員シミュレーション", 100, , , , , 1)
    Dim productivity As Variant: productivity = Application.InputBox("1人あたり処理能力", "要員シミュレーション", 10, , , , , 1)
    If workload = False Or productivity = False Then Exit Sub
    Dim required As Double
    required = CDbl(workload) / CDbl(productivity)
    ActiveSheet.Range("A1").Value = "必要要員数"
    ActiveSheet.Range("B1").Value = Application.WorksheetFunction.RoundUp(required, 0)
End Sub




