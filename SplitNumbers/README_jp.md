# SplitNumbers マクロ

このVBAマクロは、記憶競技（メモリースポーツ）などで用いられる **2桁数字×単語対応表（10×10マトリクス）** を活用し、任意の長い数字列を分割・変換・出力するツールです。

---

## 用途

このマクロは、以下の目的で使用されます：

- **0〜9の数値を縦軸・横軸に配置した10×10マスのマトリクス**を作成（例：`B2:L12`）
- 2桁の数字（例：23、87など）を各マスに対応させ、**数字と単語の対応表として活用**
- 任意の連続した数字列を入力し、以下の処理を実行：
  1. **4桁ずつに分割**
  2. **前2桁・後2桁**に分割
  3. 各2桁をキーとして、マトリクスから対応する単語を **`VLOOKUP` 関数で自動参照**
  4. 結果をセルに整形して出力

### 想定される活用シーン：

- **数字を単語に変換**し、イメージで記憶する訓練
- メモリースポーツにおける**練習用リストの自動生成**
- 数字列からの**連想イメージ構築**の補助

---

## 処理の流れ

1. ユーザーに以下のセルを選択させる：
   - 分割対象の数字が含まれるセル
   - 出力を開始するセル

2. 正規表現で余分な空白などを除去した後、4桁ごとに分割

3. 各4桁を前半2桁・後半2桁に分け、それぞれを使ってマトリクスから `VLOOKUP` で単語を参照

4. 出力セルの周囲に以下の情報を整形して表示：
   - 元の4桁数値
   - 前2桁・後2桁
   - それぞれに対応する単語（`VLOOKUP` 式）

---

## 使用方法

1. `Alt + F11` で VBA エディタを開く
2. 本コードを標準モジュールに貼り付ける
3. Excel上でマクロ `SplitNumbers` を実行
4. メッセージに従ってセルを選択し、処理を実行

---

## 注意事項

- マトリクスの範囲は `B2:L12` に固定されているため、対応表をそこに構築しておく必要があります
- 入力された数字列は、**空白・改行なしの連続した数字列**であることが前提です
- 4桁未満の末尾データは無視されます

---

## サンプルコード

```vba
Option Explicit

Sub SplitNumbers()
    Dim targetCell As Range
    Dim startCell As Range
    Dim outputCell As Range
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

    With reg
        .Pattern = "[\s]+"
        .Global = True
    End With

    Dim number As String
    Dim i As Long, chunk As String
    Dim matrix As Range
    Set matrix = Range("B2:L12")

    Set targetCell = selectCell("Please select the target cell.")
    Set startCell = selectCell("Please select the starting cell for output.")

    number = reg.Replace(targetCell.Value, "")

    For i = 1 To Len(number) Step 4
        chunk = Mid$(number, i, 4)
        If Len(chunk) < 4 Then Exit For

        Set outputCell = startCell.Offset((i - 1) \ 4, 0)
        With outputCell
            .NumberFormat = "@"
            Union(.Offset(0, 1), .Offset(0, 3)).NumberFormat = "@"
            .Value = Format(chunk, "0000")
            .Offset(0, -1).Value = i + 5
            .Offset(0, 1).Value = Format$(Left(chunk, 2), "00")
            .Offset(0, 3).Value = Format$(Right(chunk, 2), "00")
            .Offset(0, 2).Formula = "=VLOOKUP(VALUE(LEFT(" & _
                .Offset(0, 1).Address & ",1)," & matrix.Address & _
                ",VALUE(RIGHT(" & .Offset(0, 1).Address & ",1)+2),FALSE)"
            .Offset(0, 4).Formula = "=VLOOKUP(VALUE(LEFT(" & _
                .Offset(0, 3).Address & ",1)," & matrix.Address & _
                ",VALUE(RIGHT(" & .Offset(0, 3).Address & ",1)+2),FALSE)"
        End With
    Next i

    MsgBox "Split completed"
End Sub

Function selectCell(text As String) As Range
    Dim returnCell As Range
    On Error Resume Next
    Set returnCell = Application.InputBox(text, Type:=8)
    On Error GoTo 0

    If returnCell Is Nothing Then
        MsgBox "Invalid", vbExclamation
        Exit Function
    End If

    Set selectCell = returnCell
End Function