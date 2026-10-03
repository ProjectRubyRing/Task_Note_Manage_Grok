Option Explicit

' 画面のボタンとマクロ一覧から呼ぶ入口です。
' 処理本体は modCore にあります。

Public Sub 付箋を貼る()
    modCore.AddSticky
End Sub

Public Sub この付箋を削除()
    modCore.DeleteSticky
End Sub

Public Sub 完了付箋をノートに戻す()
    modCore.RestoreSticky
End Sub
