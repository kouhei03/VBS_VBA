Attribute VB_Name = "Mod190_配布ブック"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 配布用ブック数式値化不要シート削除入力欄ロックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「配布用ブック数式値化不要シート削除入力欄ロック」を選択して実行する。

' ── 配布用ブック数式値化不要シート削除入力欄ロック ──
Public Sub 配布用ブック数式値化不要シート削除入力欄ロック()
    If MsgBox("配布用ブックを作成します。元ブックは変更しませんか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Dim copyWb As Workbook: Set copyWb = Workbooks.Add(1)
    ThisWorkbook.ActiveSheet.Copy Before:=copyWb.Worksheets(1)
    copyWb.Worksheets(1).UsedRange.Value = copyWb.Worksheets(1).UsedRange.Value
    copyWb.Worksheets(copyWb.Worksheets.Count).Delete
    copyWb.Worksheets(1).Cells.Locked = False
    copyWb.Worksheets(1).UsedRange.Locked = True
    copyWb.Worksheets(1).Protect Password:="dist"
    Dim path As String
    path = Application.GetSaveAsFilename("配布用ブック.xlsx", "Excelブック (*.xlsx),*.xlsx")
    If path <> "False" Then copyWb.SaveAs path, xlOpenXMLWorkbook
End Sub




