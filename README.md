# VBAマクロツール集 / VBA Tools Collection

このリポジトリは、個人的に開発した **Excel用VBAマクロ** をまとめたものです。  
This repository is a collection of **VBA macros for Excel** that I created as personal projects.

各マクロは個別のディレクトリに分かれており、それぞれにコードとREADMEが含まれています。  
Each macro is organized into its own directory with code and a dedicated README.

---

## 収録内容 / Contents

### 1. `ResetSheet/`

アクティブなワークシートを**初期状態にリセット**するマクロです。テンプレートの再利用やシートの初期化に便利です。  
A macro that **resets the currently active worksheet** to a clean state. Useful for template reuse or fresh starts.

**主な機能 / Key Features:**
- 実行前に確認ダイアログを表示  
  Prompts user for confirmation
- すべての画像（`msoPicture`）を削除  
  Deletes all images (shapes of type `msoPicture`)
- セルの内容と罫線をすべてクリア  
  Clears all cell contents and borders
- 完了メッセージを表示  
  Displays a confirmation message

**想定用途 / Use case:**  
データを再入力する前にワークシートを完全にリセットしたいとき  
Resetting a worksheet before data re-entry

---

### 2. `SplitNumbers/`

**記憶競技（メモリースポーツ）用**のマクロです。数字列を4桁ごとに分割し、2桁×2の数字を使って、10×10のマトリクスから単語を参照します。  
A macro designed for **memory athletes and trainers**. It splits a number string into 4-digit chunks, then uses the 2-digit pairs to reference words from a 10×10 lookup matrix.

**主な機能 / Key Features:**
- 数字列を4桁ずつ分割  
  Splits number string into 4-digit chunks
- 各4桁を2桁×2に分割  
  Further splits each chunk into two 2-digit numbers
- 2桁数字を使ってマトリクスから単語を検索  
  Uses 2-digit pairs to perform `VLOOKUP` on the matrix
- 出力開始セルを選択可能  
  Interactive output cell selection

**想定用途 / Use case:**  
記憶練習、数字→イメージ変換、視覚的な連想支援  
Memory training, number-to-image conversion, visual association

---

## 使い方 / How to Use

1. 使用するマクロのディレクトリ（`ResetSheet` または `SplitNumbers`）に移動  
   Navigate to the desired macro directory (`ResetSheet` or `SplitNumbers`)
2. 各ディレクトリ内の `README.md` を確認  
   Review the `README.md` inside each directory
3. ExcelでVBAエディタを開き、コードを貼り付けて実行  
   Open Excel VBA editor, paste the code, and run the macro

---

## 動作環境 / Requirements

- Microsoft Excel（VBAが使用可能な環境）  
  Microsoft Excel (VBA-enabled)
- マクロの基本的な実行知識  
  Basic familiarity with running macros
- `SplitNumbers` 使用時は、`B2:L12` に単語マトリクスが必要  
  For `SplitNumbers`, a 10×10 word matrix in `B2:L12` is required

---

## ライセンス / License

MIT License

---

ご自由にフォーク・改良してお使いください。  
Feel free to fork, modify, and use these tools as needed.
