Attribute VB_Name = "Mod126_HTMLメール"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod126_HTMLメール
' - ファイル名: Mod126_HTMLメール.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub HTML形式メール作成装飾挿入() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub HTML形式メール作成装飾挿入()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("宛先を入力してください", "HTMLメール")
    mail.Subject = InputBox("件名を入力してください", "HTMLメール")
    mail.HTMLBody = "<html><body><h2 style='color:#2F5597'>" & _
                    Replace(mail.Subject, "&", "&amp;") & _
                    "</h2><table border='1' cellpadding='6' style='border-collapse:collapse'>" & _
                    "<tr style='background:#D9EAF7'><th>項目</th><th>内容</th></tr>" & _
                    "<tr><td>本文</td><td>" & Replace(InputBox("本文を入力", "HTMLメール"), vbCrLf, "<br>") & _
                    "</td></tr></table></body></html>"
    mail.Save
End Sub

