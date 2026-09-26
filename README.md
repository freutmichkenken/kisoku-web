# 就業規則 整形（ブラウザ版）

Colab 版の就業規則整形ツールを、ブラウザだけで動くようにしたものです。
Python は Pyodide（ブラウザ内で動く Python）で実行するため、
PC にもスマホにも何もインストールする必要はありません。

就業規則の docx は、開いた端末のブラウザ内だけで処理されます。
サーバーには送られません（Gemini を使う場合のみ、判定対象の文章が Google に送られます）。

## フォルダ構成

```
index.html      画面
worker.js       裏で Python を動かす処理
py/             整形プログラム（Colab 版と同じもの）
  web_entry.py    ブラウザ版の入口（Colab 版の kisoku_app.py に相当）
  gemini_helper.py  ブラウザからも Gemini を呼べるよう改修済み
templates/      就業規則テンプレ1〜4.docx
wheels/         python-docx 1.2.0（固定版）
.nojekyll       GitHub Pages の加工を止めるための空ファイル
```

## GitHub Pages で公開する手順（初回のみ）

1. GitHub で新しいリポジトリを作る（例：`kisoku-web`）。
   無料プランの GitHub Pages は **Public（公開）リポジトリ** が必要です。
2. リポジトリの画面で「Add file」→「Upload files」を選び、
   この zip を展開したフォルダの**中身**をすべてドラッグして「Commit changes」。
   - `.nojekyll` は隠しファイルのため、ドラッグで漏れることがあります。
     漏れた場合は「Add file」→「Create new file」で `.nojekyll` という名前の空ファイルを作ってください。
3. 「Settings」→「Pages」を開き、Source を「Deploy from a branch」、
   Branch を「main」「/ (root)」にして「Save」。
4. 1〜2分後、同じ画面に `https://＜ユーザー名＞.github.io/kisoku-web/` が表示されます。
   これをスマホ・PC のブラウザでブックマークすれば完了です。
   - Android の Chrome では「︙」→「ホーム画面に追加」でアイコンを置けます。

## 公開される範囲について

リポジトリが Public なので、**プログラムとテンプレート docx は誰でも見られます**。
顧問先のファイルはリポジトリに入らないので公開されません。
`index.html` には `noindex` を付けていないため、検索エンジンに載ることがあります。

## プログラムを直したとき

- `py/` の中の .py を差し替えて、GitHub にアップロード（上書き）するだけです。
- **ファイルを新しく増やした場合**は、`worker.js` の `PY_FILES` にファイル名を追加してください。
  テンプレートを増やした場合も同様に `TEMPLATE_FILES` への追加と、
  `py/web_entry.py` の `TEMPLATES`、`index.html` のテンプレート選択欄への追加が必要です。
- 反映まで数分かかることがあります。反映されない場合はページを再読み込みしてください。

## Colab 版との関係

- `py/` の整形プログラムは Colab 版と同じもので、出力結果も同じです。
- `gemini_helper.py` だけ、ブラウザ版向けに改修しています
  （ライブラリが無い環境では Gemini の REST API を直接呼ぶ）。
  Colab で `google-generativeai` が入っていれば従来どおりの動作になるので、
  Drive 上の Colab 版にこの改修版を置いても問題ありません。

## 動作の目安

- 初回は Python 実行環境（約 15MB）をダウンロードするため、数十秒かかります。
  2回目以降はブラウザのキャッシュが効いて速くなります。
- Gemini の API キーは、「この端末のブラウザに保存する」にチェックした場合のみ、
  その端末のブラウザ内に保存されます。

## 回帰テスト

`tests/samples/` の就業規則サンプルで、修正後も整形結果が変わっていないかを確認できます（python-docx が必要：`pip install wheels/*.whl`）。

```
python3 tests/regression.py            # 基準（tests/expected/）と比較
python3 tests/regression.py --update   # 意図した変更のあと基準を作り直す
```
