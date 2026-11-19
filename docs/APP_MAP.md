# アプリ地図

最終確認: 2026-11-19 / コミット 643ed15

## あるべき姿
<!-- 案：コードの現状から起こした。利用者の確認待ち -->
- 目的: 社労士が、就業規則の docx をブラウザ内（Pyodide）で条・項・号に解析し、テンプレートで整形し直し、新旧対照表を作り、改訂の版を管理する。ファイルは外へ出さない（Gemini 利用時の判定対象の文章だけ）。
- 主な操作の流れ:
  - 規程一覧 →（docx を読み込む）→ 整形の設定ダイアログ →（この設定で開く）→ 編集画面
  - 編集画面：左の新旧対照表で改正案を編集 →「変更履歴」で版を確定 →「出力」でダウンロード／対照表（PDF・docx）
  - 編集画面 →「チェック」で参照（第◯条）の切れを直す
- UIの約束:
  - 主ボタンの位置と文言: `.btn.primary`。ダイアログはフッター右端に主ボタン、その左に副ボタン。
  - エラーの出し方: その場の `.hint` に `err`／`ok` を付けて1行で出す。編集中の条の中だけ `.ed-err`。`alert()` は使わない。
  - 確認ダイアログ: データが消える操作だけ `confirm()`。削除は「元に戻せません」を付ける。
  - 色・余白・文字の大きさ: index.html の `<style>` の CSS 変数。ダークは `prefers-color-scheme` と `data-theme="dark"`。
  - 文言の語調: 説明は です・ます、ラベルは体言止め。空の状態は「〜はまだありません。」の1行。

## 画面
| 画面 | 役割（1行） | 載っている機能 | 置かないもの | 根拠 | 確認 |
|---|---|---|---|---|---|
| 規程一覧（#viewList） | 保存した作業の一覧と開始点 | docx 読み込み、作業（.json）読み込み、並べ替え、開く・名前変更・削除・書き出し | 整形・編集 | index.html:834 | 643ed15 |
| 編集画面のヘッダ | 戻る、タイトル、Python エンジンの状態 | 「← 規程一覧」、エンジン状態 | — | index.html:855 | 643ed15 |
| 左：新旧対照表（section#h4） | 現行と改正案を並べ、改正案を条・章ごとに編集 | 変更点表示、別の docx と比較、元に戻す、下書き削除、目次、変更のある条だけ表示 | 出力 | index.html:874、editor.js | 643ed15 |
| タブ「本文」 | 狭い画面で左の対照表を出す | 左欄の切り替え | 広い画面では非表示 | index.html:913 | 643ed15 |
| タブ「変更履歴」（#panHist） | 版の確定・一覧・比較 | 施行日・メモ、版の確定、版の比較、削除 | — | index.html:919 | 643ed15 |
| タブ「出力」（#panOut） | 整形した docx と対照表を作る | ダウンロード、中止、処理の記録、対照表（PDF・docx） | 設定 | index.html:952 | 643ed15 |
| タブ「設定」（#panSet） | 読み込んだ docx、オプション、テンプレート | ファイル選択、オプション（Gemini・表記ゆれ等）、テンプレ1〜4とカスタム、プレビュー | — | index.html:983、custom.js | 643ed15 |
| タブ「チェック」（#panChk） | 参照のチェック | 参照先なし等の一覧 | — | index.html:975 | 643ed15 |
| ダイアログ 新旧対照表（#tblDlg） | 対照表の確認と出力 | 全文／差分のみ、備考、PDF、docx | — | index.html:1116 | 643ed15 |
| ダイアログ 整形の設定（#openDlg） | docx を開くときの設定 | テンプレート・オプション、キャンセル／この設定で開く | — | index.html:1140 | 643ed15 |
| サイドメニュー（aside#side） | 端末全体の設定 | テンプレート登録、Gemini API、既定のオプション、データ管理、表示、使い方 | 作業単位の操作 | index.html:744 | 643ed15 |

## 共通部品・共通関数
| 名前 | 用途 | 場所 | 使っている画面 | 確認 |
|---|---|---|---|---|
| `store` | localStorage の get/set/del（例外を握りつぶす） | index.html:1158 | サイド、備考、既定、表示 | 643ed15 |
| `idb` | IndexedDB の drafts / works | index.html:1424 | 一覧、編集画面 | 643ed15 |
| `call` / worker | Python の呼び出し | index.html:1181、worker.js:199 | 全画面 | 643ed15 |
| `setEngine` / `log` | エンジン状態と処理の記録 | index.html:1199 | ヘッダ、出力 | 643ed15 |
| `verHint` ほか5つの `*Hint` | `.hint` の class と文を切り替える | index.html:1859, 2405 ほか | 履歴、対照表、カスタム、テンプレ、API、データ | 643ed15 |
| `showFormError` | 編集中の条にエラーを出す | index.html:2174 | 左の対照表 | 643ed15 |
| `selectTab` / `requestTab` | タブ切替（編集中の確定確認付き） | index.html:1613 | 編集画面 | 643ed15 |
| `KisokuRefs`（refs.js） | 「第◯条」参照の書き換え | refs.js | 対照表の編集、チェック | 643ed15 |
| `KisokuCustom`（custom.js） | カスタム書式の設定とプレビュー | custom.js:450 | 設定タブ | 643ed15 |
| `KisokuEditor`（editor.js） | 条・章・附則の編集フォーム | editor.js:616 | 左の対照表 | 643ed15 |
| `_check_tree` 系 | 画面から来た木の検証（上限値も） | py/web_entry.py:153 | 保存・整形 | 643ed15 |

## データと状態
| データ | 持つ場所 | 読み書きする画面 | 確認 |
|---|---|---|---|
| 作業（docx 本体・設定・版） | IndexedDB `kisoku` の works（index.html:1429） | 一覧、編集画面 | 643ed15 |
| 編集の下書き | IndexedDB の drafts（index.html:1423） | 左の対照表 | 643ed15 |
| 対照表の備考 | localStorage `kisoku.notes.<id>`（index.html:1402） | 対照表 | 643ed15 |
| オプション・テンプレ選択 | localStorage `kisoku.prefs`（index.html:1214） | 設定タブ | 643ed15 |
| Gemini API キー | localStorage `kisoku.geminiKey`（index.html:1163） | サイド | 643ed15 |
| 登録テンプレート | localStorage `kisoku.customTpls`（index.html:3048） | サイド、設定タブ | 643ed15 |
| 既定のオプション | localStorage `kisoku.defaults`（index.html:3208） | サイド、docx 読み込み | 643ed15 |
| 表示（配色・文字サイズ） | localStorage `kisoku.view`（index.html:10） | サイド | 643ed15 |
| カスタム設定 | localStorage `kisoku.custom`（custom.js:6） | 設定タブ | 643ed15 |
| 解析木・一時ファイル | Pyodide の仮想 FS（py/web_entry.py:334） | 裏側 | 643ed15 |

## 既知の重複・不統一
- 番号の整形 `formatNumber` が3か所で別実装（custom.js:126、editor.js:27、py/comparison.py:81）。対応する書式と、20を超える丸数字の扱いが違う — （直す案: 画面側は refs.js のように共有ファイルにする。先に出力差を確かめる）
- 画像・表・数式の上限値を画面と Python で二重に持つ — editor.js:139, 152, 196、py/web_entry.py:153（直す案: 1か所から渡す。現状はコメントで同期を明記）
- `.hint` の切り替え関数が6つ別々で、cmpHint だけ `err` しか受けない — index.html:1859, 2405, 2708, 3028, 3187, 3228（直す案: `setHint(id, kind, text)` に寄せる）
- エラーの出し方が3種（`.hint`、`.ed-err`、画像だけ `alert()`） — editor.js:299, 310（直す案: フォーム内の表示に揃える）
- 削除の確認文に「元に戻せません」があるものと無いものが混在。キャンセルで別の動きをする確認もある — index.html:1681, 3137, 2547、editor.js:276（直す案: 削除の文を1つに決め、意味が違う確認は文言で明示）
- `.btn.danger` をラベルが「削除」かどうかの文字比較で付けている — index.html:1594, 2025（直す案: 引数で渡す）
- ダークの CSS 変数を2か所に同じ値で持つ — index.html:50, 64（CSS だけでは1か所にしづらいので分けたまま）
- 手順番号 ①〜⑦ が画面の並びと合わない — index.html:880, 923, 955, 969, 989（番号を外すか振り直す。要確認）
- custom.js に別の localStorage ラッパーがある — custom.js:187（読み込み順の都合で分けたまま）
- `fit`（custom.js:428 と editor.js:115）は名前が同じだけで別物（分けたまま）

## 見ていない範囲
- tests/、templates/、wheels/、docs/ は読んでいない。
- index.html の JS 本体（約1190〜3299行）は関数一覧と要所だけ。renderComparison、整形処理、openTable／printTable、custom.js と editor.js のフォーム生成本体は読んでいない。
- py/ は web_entry.py の定義一覧だけ。kisoku_parser.py ほかの解析ロジックの重複は調べていない。
- CSS 本体の余白・文字サイズの一貫性、実際の画面での見た目は確かめていない。

<!-- ここから下は app-inventory.py の出力をそのまま貼る -->

<!-- app-inventory.py で生成（コミット 643ed15）。手で書き換えず、作り直す。 -->
## 画面・入口の一覧（機械抽出）

### HTML のファイル（1件）
- index.html

## 同じ名前の関数が複数のファイルにある箇所（重複実装の候補）

- `_p_text` — py/crossref.py:72, py/verify_report.py:95
- `_run` — py/test_crossref.py:9, py/test_custom_template.py:71
- `add` — py/comparison.py:207, py/kisoku_parser.py:1967
- `art` — py/test_comparison.py:18, tests/outline_util.py:47, tests/test_refs.js:6
- `article` — index.html:1831, py/test_parser.py:38, py/web_entry.py:278, tests/test_format_variation.py:66
- `clone` — editor.js:8, tests/test_refs.js:9
- `compare` — py/comparison.py:715, py/verify_report.py:159
- `fill` — index.html:2032, py/comparison.py:784
- `find` — index.html:2282, py/hyoki_check.py:369
- `fit` — custom.js:428, editor.js:115
- `flush` — py/comparison.py:660, py/kisoku_parser.py:630
- `formatNumber` — custom.js:126, editor.js:27
- `items` — py/test_parser.py:33, tests/outline_util.py:42, tests/test_format_variation.py:60
- `normalize` — py/custom_template.py:106, py/verify_report.py:67
- `parse` — index.html:1282, py/web_entry.py:55
- `quiet` — py/test_editing.py:30, py/test_gemini_agent.py:69
- `run` — index.html:1443, py/apply_style.py:929, tests/test_format_variation.py:202
- `summary_lines` — py/hyoki_check.py:553, py/verify_report.py:271
- `text` — py/test_comparison.py:37, py/web_entry.py:242
- `units` — index.html:1827, py/comparison.py:321, py/web_entry.py:250
- `walk` — py/verify_report.py:205, refs.js:144
