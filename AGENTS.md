# AGENTS.md

## 就業規則の整形が失敗したときの対応

このアプリの主目的は、就業規則（docx）を整形することです。
ある就業規則で整形が失敗した場合は、次のように対応してください。

- その就業規則ファイルに合わせた修正だけで終わらせない。
- 失敗の原因（条文の書き方、見出しや番号の形式など）を一般化し、ほかの就業規則でも同じ失敗が起きないように整形プログラム（`py/`）を直す。
- その修正によって、これまで正しく整形できていた就業規則の結果が変わるおそれがある場合は、修正する前にユーザーへ可否を確認する。

## 回帰テスト

`py/` を修正したり機能を追加したりしたら、次を実行して既存の整形結果が壊れていないか確認する。

```
python3 py/test_parser.py
python3 py/test_custom_template.py
python3 py/test_crossref.py
python3 py/test_gemini_agent.py
python3 py/test_comparison.py
python3 py/test_editing.py
python3 tests/regression.py
python3 tests/test_misdetection.py
python3 tests/test_format_variation.py
node tests/test_refs.js
node tests/test_editor_split.js
```

- GitHub Actions（`.github/workflows/test.yml`）でも、PR と main への push のたびに同じテストを実行する。
- `tests/samples/` の就業規則サンプル（`messy_kisoku.docx` は `tests/make_messy_sample.py` で作った書式の乱れたもの）を解析・整形（テンプレ t1〜t4）し、`tests/expected/` の基準と比べる。サンプルの記述内容そのものは検証しない。
- 誤判定例（`tests/test_misdetection.py`）：`tests/cases/misdetection_cases.json` の例を、期待する構造（`expected_outline`）と比べる。既知の失敗（`known_failure`）は数えるだけで落とさない。直ったら表示されるので `known_failure` を外す。仕様未決（`pending_spec`）は比べない。
- 参照の書き換え（`tests/test_refs.js`）：`refs.js` が、条の追加・削除で本文中の「第◯条」を新しい番号に書き換え、法令・ほかの規程への参照は書き換えないことを確かめる。
- 条の分割・まとめ（`tests/test_editor_split.js`）：`editor.js` の「ここから新しい条にする」「前の条にまとめる」が、項を号・表ごと移し、条見出しで始まる項だけタイトルを取り出し、まとめるときはタイトルを本文の先頭に残すことを確かめる。
- 書式ゆらぎ（`tests/test_format_variation.py`）：サンプルの標準形と、書式を崩した形（折り返し・数字の空白・タイトル別行・項の「２．」）との違いを条ごとに数える。`tests/cases/format_variation_known.json` に無い不一致が出たら失敗。
- 差分が出た場合、意図した変更かどうかをユーザーに確認し、了承を得てから `python3 tests/regression.py --update` で基準を更新する。

## アプリ地図

`docs/APP_MAP.md` に、画面の役割・共通部品・データの持ち方・UIの約束をまとめている。
- 画面、UI、共通部品、データの持ち方に触る変更の前に読み、置き場所や既存部品との重複を確かめる
- 変更で画面・共通部品・UIの約束が変わったら、同じ変更の中で地図も直す
- 末尾の機械抽出の節は手で直さず、sagy スキルの `scripts/app-inventory.py` で作り直す(手順は `references/app-map.md`)
