# OnlineVST-NJ8

# VST-NJ8 オンライン語彙サイズテスト

VST-NJ8はHamada et al. (2021) で開発された日本人英語学習者向けの英語語彙サイズを測定するテストです。項目反応理論(IRT)に基づいた精密な測定により、学習者の英語語彙力を8つのレベルで評価します。こちらはオンライン版です。

## 機能

- **8レベル構成**: 1000語レベルから8000語レベルまでの語彙を測定（各レベル20問）
- **項目反応理論**: 3パラメータロジスティックモデル(3PLM)を使用した正確な能力推定
- **インターフェース**:
  - キーボードおよびマウスによる回答入力
  - レベル間の休憩機能
  - 「わからない」オプションによる推測の排除
- **結果の表示**:
  - レベル別正答率の表示
  - 平均反応時間の分析
  - 視覚的なグラフ表示
- **データエクスポート**:
  - JSON形式でのダウンロード
  - Excel形式でのダウンロード (SheetJS使用)
  - 印刷機能

## 使用方法

1. 「テストを開始する」ボタンをクリックしてテストを開始します
2. 日本語単語に対応する英単語を選択肢から選びます:
   - マウスでクリック、または
   - キーボードの1〜4キー（選択肢選択）および5キー（わからない）を使用
3. 各レベル（20問）が終了すると休憩画面が表示されます
4. 全160問が終了すると結果画面が表示されます
5. 結果はExcel、JSONでダウンロードまたは印刷できます

## ファイル構成

- `index.html` - HTMLマークアップとCSSスタイル
- `script.js` - アプリケーションのロジックとテストデータ

## テストデータのカスタマイズ

`script.js`ファイルの先頭部分にあるテストデータ配列を編集することで、問題を変更できます:

```javascript
const testData = [
    { 
        level: 1, 
        item: "売り場", 
        partOfSpeech: "noun", 
        correctAnswer: "department", 
        distractors: ["boy", "meeting", "town"],
        discrimination: 1.98, 
        difficulty: -2.38, 
        guessing: 0.13 
    },
    // 追加の問題...
];
```

各問題には以下の属性が必要です:
- `level`: 難易度レベル（1〜8）
- `item`: 日本語の単語
- `partOfSpeech`: 品詞（noun, verb, adjective, adverb）
- `correctAnswer`: 正解となる英単語
- `distractors`: 誤答選択肢の配列（3つ）
- `discrimination`, `difficulty`, `guessing`: IRTパラメータ

## 技術的詳細

### 項目反応理論（IRT）の実装

このアプリケーションでは、3パラメータロジスティックモデル（3PLM）を使用して受験者の能力値（θ）を推定しています:

```
P(θ) = c + (1-c) / (1 + exp(-a(θ-b)))
```

ここで:
- P(θ)は能力値θの受験者が正答する確率
- a は識別力パラメータ
- b は困難度パラメータ
- c は推測パラメータ

能力値の推定には最尤推定法とニュートン・ラフソン法を使用しています。

### ライブラリ

- [SheetJS](https://sheetjs.com/) - Excel形式でのデータエクスポート

## 作者

[小室竜也] - [komuro.4121☆gmail.com]
☆を@に変更してください。

## 謝辞

このアプリケーションはHamada et al. (2021): https://doi.org/10.32234/jacetjournal.65.0_23 を基に開発されました。
