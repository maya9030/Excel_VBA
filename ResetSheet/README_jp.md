# ResetSheet マクロ

このVBAマクロは、**アクティブなワークシートを初期状態にリセット**するためのものです。画像やセル内容、罫線を一括で削除・クリアしたい場面で使用できます。

## 機能概要

- **実行前に確認ダイアログ**を表示
- シート上の**すべての画像（`msoPicture`）を削除**
- セルの**罫線をすべて削除**
- セルの**内容をすべてクリア**
- 完了後に**通知メッセージを表示**

## 処理の流れ

1. ユーザーに確認ダイアログを表示  
   - メッセージ: `"Reset to default settings?"`
   - 「キャンセル」を選択した場合、処理は中断

2. 「OK」の場合、以下を実行:
   - ワークシート上のすべての図形を調査
   - `msoPicture` 型の図形（画像）のみ削除
   - セルの罫線をすべて除去
   - セル内容を完全にクリア

3. 最後にメッセージボックスで完了通知:
   - `"Reset complete."`

## 使用方法

1. Excelで `Alt + F11` を押して **VBAエディタ**を開く
2. モジュールに以下のコードを貼り付ける
3. リセットしたいシートをアクティブにして、`ResetSheet` を実行

## 注意事項

- **実行すると元に戻せません**
- 必要なデータは事前に保存・バックアップしてください

## サンプルコード

```vba
Sub ResetSheet()
    Dim shp As Shape
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim result As VbMsgBoxResult
    result = MsgBox("Reset to default settings?", vbOKCancel + vbQuestion, "OK")
    If result = vbCancel Then Exit Sub

    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then shp.Delete
    Next shp

    ws.Cells.Borders.LineStyle = xlNone
    ws.Cells.Clear

    MsgBox "Reset complete."
End Sub