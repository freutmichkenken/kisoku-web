# 回帰テスト用の文章

整形の不具合やレビューの指摘を受けて作った、就業規則の文章の例を置く。
今後 `py/` や編集画面を直したときに、同じ失敗が戻っていないかを確かめるのに使う。

## ファイルの形
- 1ファイル1例。拡張子は `.txt`（UTF-8）
- `#` で始まる行は説明（目的・期待する結果）。それ以外の行が、docx の1段落ずつの本文
- docx にして解析するには `tests/outline_util.py` の `make_docx` と `parse_path` を使う

```
python3 tests/regression_texts/show.py tests/regression_texts/<名前>.txt
```

で、解析した骨組みを表示する。

## 一覧
| ファイル | 確かめること |
|---|---|
| branch_hyphen.txt | 枝番をハイフンで書いた条見出し（第６条－２）が独立した条になる |
| loose_article_sharp.txt | 想定外の枝番（第６条＃２）の行が、独立した項として残り、要確認の印が付く |
| loose_article_inline.txt | 文の途中に結合された疑いの行（＃３）にも印が付く |
| loose_article_front.txt | 文書の最初の条が疑いの行のとき、前付けに印が付く |
| editor_split.txt | 画面の「ここから新しい条にする」で、印の付いた第3項から分けるとタイトルが取り出される |
| heading_like_references.txt | 見出しに似た本文中の参照。分割しても本文から文言を消さない（1行ずつ別の例） |
