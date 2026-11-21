Attribute VB_Name = "Mod06_ステータスバー"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 進捗表示を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「進捗表示」を選択して実行する。

' ── 進捗表示 ──
Public Sub 進捗表示(ByVal currentValue As Long, ByVal totalValue As Long)
    If totalValue <= 0 Then
        Application.StatusBar = "処理中..."
    Else
        Application.StatusBar = "処理中: " & currentValue & " / " & totalValue & _
                                " (" & Format$(currentValue / totalValue, "0%") & ")"
    End If
End Sub

' ── 進捗表示終了 ──
Public Sub 進捗表示終了()
    Application.StatusBar = False
End Sub




