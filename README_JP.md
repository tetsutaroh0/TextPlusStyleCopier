# TextPlusStyleCopier for DaVinci Resolve

クリップカラーを使って、複数の Text+ クリップへ一括でスタイルをコピーできる DaVinci Resolve 用スクリプトです。

テキスト内容やレイアウトを維持しながら、効率よく Text+ のデザインを統一できます。

---

# 主な機能

* 複数の Text+ クリップへスタイルを一括適用
* クリップカラーで対象クリップを抽出
* トラック限定適用

  * 全トラック
  * V1 / V2 / V3 など個別指定
* 以下の項目を保持可能

  * テキスト本文
  * 位置
  * サイズ
  * 回転
  * ピボット
* 常駐型UI
* Power Bin / Media Pool の Text+ を参照元に使用可能
* 再生ヘッド位置を維持
* 一時参照クリップを自動削除
* Resolve を再起動せず繰り返し実行可能

---

# 必要環境

* DaVinci Resolve
* Python 3.x
* DaVinci Resolve Scripting 有効化

動作確認:

* Windows
* DaVinci Resolve 19.x

---

# インストール

## 1. Python をインストール

Python 3.x をインストールしてください。

https://www.python.org/downloads/

インストール時に:

```text id="6xukol"

Add Python to PATH

```

へチェックを入れてください。

---

## 2. スクリプトを配置

以下へスクリプトをコピーします。

```text id="jvklyp"

C:\\ProgramData\\Blackmagic Design\\DaVinci Resolve\\Fusion\\Scripts\\Utility

```

---

# 使い方

## 基本手順

1. Power Bin / Media Pool に参照元 Text+ を用意
2. 参照元 Text+ を選択
3. タイムライン上の対象 Text+ にクリップカラーを設定
4. スクリプトを実行
5. 以下を設定

   * 対象クリップカラー
   * 対象トラック
   * 保持項目
6. 「実行」をクリック

---

# 保持項目

スタイル適用時に、以下を保持できます。

| 項目   | 内容           |
| ---- | ------------ |
| 本文   | テキスト内容を維持    |
| 位置   | Text+ の位置を維持 |
| サイズ  | テキストサイズを維持   |
| 回転   | 回転を維持        |
| ピボット | Pivot設定を維持   |

---

# トラック限定機能

対象トラックを指定できます。

| 値 | 対象    |
| - | ----- |
| 0 | 全トラック |
| 1 | V1    |
| 2 | V2    |
| 3 | V3    |

---

# 注意事項

* Text+ 専用です
* 一部の Fusion Title では動作しない場合があります
* 処理中、一時的に参照元クリップをタイムラインへ追加します
* 一時クリップは処理後に自動削除されます

---

# 同梱ファイル

| ファイル                      | 内容     |
| ------------------------- | ------ |
| TextPlusStyleCopier_JP.py | 日本語UI版 |
| TextPlusStyleCopier.py    | 英語UI版  |

---

# ライセンス

MIT License

---

# Author

Tetsu
