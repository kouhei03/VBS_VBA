Attribute VB_Name = "Mod38_定数列挙型"
Option Explicit

Public Const 最大行数 As Long = 1048576
Public Const 設定シート名 As String = "設定"

Public Enum 処理状態
    未実行 = 0
    実行中 = 1
    完了 = 2
    エラー = 3
End Enum
