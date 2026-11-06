# -*- coding: utf-8 -*-
"""
ブラウザ版（Pyodide）から呼ばれる入口。

Colab 版の kisoku_app.py から、画面（ipywidgets）部分を取り除き、
「解析」「整形」の2つの処理だけを関数として公開したもの。
画面は index.html / worker.js 側が担当する。

戻り値は JavaScript へ渡しやすいよう JSON 文字列にしている。
"""

import os
import sys
import json

import gemini_helper
from kisoku_parser import parse_kisoku
from apply_style import apply_style, TEMPLATE_MAP  # noqa: F401
import custom_template
import comparison
import settings_reader

TEMPLATE_DIR = "/app/templates"

TEMPLATES = {
    "t1": "就業規則テンプレ1.docx",
    "t2": "就業規則テンプレ2.docx",
    "t3": "就業規則テンプレ3.docx",
    "t4": "就業規則テンプレ4.docx",
}


def _set_api_key(api_key):
    if api_key:
        os.environ["GEMINI_API_KEY"] = api_key
    else:
        os.environ.pop("GEMINI_API_KEY", None)


def _set_agent_hook():
    """ブラウザでは、まとめて判定の interaction ID を画面（中止ボタン）へ知らせる。"""
    if sys.platform != "emscripten":
        return
    import js
    from pyodide.ffi import to_js

    def hook(interaction_id):
        js.postMessage(to_js({"type": "agentStarted",
                              "interactionId": interaction_id},
                             dict_converter=js.Object.fromEntries))

    gemini_helper.agent_started_hook = hook


def parse(src_path, use_gemini=False, layout_as_figure=False, api_key="",
          gemini_mode="paragraph"):
    """整形前 docx を解析して JSON を保存し、章・条の数を返す。"""
    _set_api_key(api_key if use_gemini else "")
    _set_agent_hook()

    tree = parse_kisoku(
        src_path,
        use_gemini=bool(use_gemini),
        treat_layout_as_figure=bool(layout_as_figure),
        gemini_mode=gemini_mode,
    )

    json_path = os.path.splitext(src_path)[0] + "_parsed.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(tree, f, ensure_ascii=False, indent=2)

    n_ch = sum(1 for n in tree if n.get("type") == "chapter")
    n_art = sum(len(ch.get("articles", []))
                for ch in tree if ch.get("type") == "chapter")
    n_art += sum(1 for n in tree if n.get("type") == "article")

    return json.dumps({"json_path": json_path,
                       "chapters": n_ch, "articles": n_art},
                      ensure_ascii=False)


def _template_paths():
    return {k: os.path.join(TEMPLATE_DIR, f) for k, f in TEMPLATES.items()}


def _load_custom(custom_json):
    try:
        return json.loads(custom_json or "")
    except ValueError:
        raise ValueError("書式の設定を読み取れませんでした")


def _build_custom(custom_json, work_dir):
    """カスタム設定からテンプレート docx を作り、(パス, 対応表) を返す。"""
    path = os.path.join(work_dir, "_カスタムテンプレート.docx")
    M = custom_template.build(_load_custom(custom_json), _template_paths(),
                              path)
    return path, M


def format_docx(src_path, json_path, template_key,
                check_hyoki=False, insert_toc=False, keep_crossref=False,
                custom_json="", edited_json_path=""):
    """
    解析済み JSON にテンプレートを適用し、出力ファイルのパスを返す。
    edited_json_path があれば（画面で編集した木）そちらを整形し、
    照合レポートでは編集による変更を欠落・増加と分けて示す。
    template_key が "custom" のときは custom_json（書式の設定）で
    テンプレートを作ってから適用する。
    """
    work_dir = os.path.dirname(src_path)
    if template_key == "custom":
        template_path, template_map = _build_custom(custom_json, work_dir)
        suffix = "カスタム"
    elif template_key in TEMPLATES:
        template_path = os.path.join(TEMPLATE_DIR, TEMPLATES[template_key])
        template_map = None
        suffix = template_key
    else:
        raise ValueError(f"不明なテンプレートです: {template_key}")

    base = os.path.splitext(os.path.basename(src_path))[0]
    output_path = os.path.join(work_dir, f"{base}_整形済_{suffix}.docx")
    report_path = os.path.splitext(output_path)[0] + "_照合レポート.md"
    for p in (output_path, report_path):      # 同じテンプレでやり直す場合
        if os.path.exists(p):
            os.remove(p)

    apply_style(
        json_path=edited_json_path or json_path,
        base_json_path=json_path if edited_json_path else None,
        template_path=template_path,
        template_key=template_key,
        template_map=template_map,
        output_path=output_path,
        source_docx=src_path,        # 表・前付けの引き継ぎ元
        show_notice=False,           # 注意事項は画面に常時表示している
        check_hyoki=bool(check_hyoki),
        insert_toc=bool(insert_toc),
        keep_crossref=bool(keep_crossref),
    )

    if not os.path.exists(report_path):
        report_path = None

    return json.dumps({"output_path": output_path,
                       "report_path": report_path},
                      ensure_ascii=False)


_TREE_TYPES = {"front_matter", "chapter", "article", "back_matter"}
_UNIT_NAME = {"paragraph": "項", "item": "号", "sub": "号の細分", "sub2": "号の細分の下"}
IMAGE_MAX_BYTES = 5 * 1024 * 1024        # 画面で足す画像1枚の上限（画面側と同じ）
IMAGE_TOTAL_BYTES = 30 * 1024 * 1024     # 規程全体の画像の合計の上限（ブラウザ内の Python のメモリを使い切らないように）
TABLE_MAX_ROWS = 50                      # 画面で足す表の行・列の上限（画面側と同じ）
TABLE_MAX_COLS = 10
TABLE_CELL_MAX = 1000                    # 画面で足す表の1セルの文字数の上限（画面側と同じ）
TABLE_TOTAL_CELLS = 2000                 # 規程全体で、画面で足す表のマスの合計の上限（表を何個も足して重くしないように）
FORMULA_PART_MAX = 200                   # 画面で足す数式の1欄（前・分子・分母・後）の文字数の上限（画面側と同じ）
_IMAGE_MAGIC = {"image/png": b"\x89PNG\r\n\x1a\n", "image/jpeg": b"\xff\xd8\xff"}


def _check_image(dw):
    """画面で足した画像：PNG・JPEG で、宣言した形式と中身が一致し、上限以下であること。"""
    import base64
    import binascii
    mime, data = dw.get("mime"), dw.get("data")
    if mime not in _IMAGE_MAGIC or not isinstance(data, str) or not isinstance(dw.get("name", ""), str):
        raise ValueError("画像は PNG か JPEG にしてください")
    if len(data) > IMAGE_MAX_BYTES * 4 // 3 + 4:
        raise ValueError("画像は 5MB 以下にしてください")
    try:
        raw = base64.b64decode(data, validate=True)
    except (binascii.Error, ValueError):
        raise ValueError("画像が壊れているため読み込めません")
    if len(raw) > IMAGE_MAX_BYTES:
        raise ValueError("画像は 5MB 以下にしてください")
    if not raw.startswith(_IMAGE_MAGIC[mime]):
        raise ValueError("画像の形式と中身が一致しません（拡張子だけ変えた画像は使えません）")
    return len(raw)


def _check_figures(node, used=None):
    """表・図・画像の一覧の形を確かめる（表・図は元の文書の位置 index を持つ）。"""
    for key in ("tables", "drawings"):
        lst = node.get(key, [])
        if not isinstance(lst, list) or not all(isinstance(x, dict) for x in lst):
            raise ValueError("表・図の一覧の形式が正しくありません")
        for x in lst:
            if key == "drawings" and x.get("type") == "image":
                n = _check_image(x)
                if used is not None:
                    used[0] += n
                    if used[0] > IMAGE_TOTAL_BYTES:
                        raise ValueError("画像の合計は 30MB 以下にしてください")
                continue
            if key == "drawings" and x.get("type") == "formula":
                # 画面で足した数式。各欄は文字列（無ければ空）で、長すぎないこと
                for k in ("before", "num", "den", "after"):
                    v = x.get(k)
                    if not (v is None or isinstance(v, str)):
                        raise ValueError("数式の形式が正しくありません")
                    if v and len(v) > FORMULA_PART_MAX:
                        raise ValueError(f"数式の各欄は{FORMULA_PART_MAX}文字以内にしてください")
                    # 1行の欄なので改行は入らない。制御文字は docx に書けず数式が黙って消えるので止める
                    if v and any(ord(ch) < 0x20 or ord(ch) == 0x7f for ch in v):
                        raise ValueError("数式に改行や制御文字は使えません")
                continue
            idx = x.get("index")
            if not (idx is None or (isinstance(idx, int) and not isinstance(idx, bool) and idx >= 0)):
                raise ValueError("表・図の元の位置（index）が正しくありません")
            if key == "tables":
                rows = x.get("rows", [])
                if not isinstance(rows, list) or not all(
                        isinstance(r, list) and all(c is None or isinstance(c, str) for c in r) for r in rows):
                    raise ValueError("表の中身（rows）の形式が正しくありません")
                if idx is None:
                    # 画面で足した表（整形で rows から作り直す）。大きすぎる表で整形が重くならないように
                    if (not rows or len(rows) > TABLE_MAX_ROWS
                            or not all(1 <= len(r) <= TABLE_MAX_COLS for r in rows)):
                        raise ValueError(f"表は1〜{TABLE_MAX_ROWS}行、1〜{TABLE_MAX_COLS}列にしてください")
                    if any(c is not None and len(c) > TABLE_CELL_MAX for r in rows for c in r):
                        raise ValueError(f"表の1マスは{TABLE_CELL_MAX}文字以内にしてください")
                    if used is not None:
                        used[1] += len(rows) * max(len(r) for r in rows)   # 整形は最大列数で作るので、そのマス数で数える
                        if used[1] > TABLE_TOTAL_CELLS:
                            raise ValueError(f"追加した表のマスは合計{TABLE_TOTAL_CELLS}までにしてください")
            elif not (x.get("text") is None or isinstance(x.get("text"), str)):
                raise ValueError("図の文字（text）の形式が正しくありません")


def _s(v):
    """エラーの場所の表示に使う文字列（文字列でなければ空。前後の空白は除く）。"""
    return v.strip() if isinstance(v, str) else ""


def _check_tree(tree):
    """画面から受け取った木の形を確かめる（壊れた下書きで整形が落ちないように）。"""
    if not isinstance(tree, list):
        raise ValueError("編集内容の形式が正しくありません")

    def text(v):
        return v is None or isinstance(v, str)

    used = [0, 0]                    # 画像の合計の大きさ、画面で足した表のマスの合計

    def figures(node):
        _check_figures(node, used)

    def units(lst, kind):
        name = _UNIT_NAME[kind]
        if not isinstance(lst, list):
            raise ValueError(f"{name}の一覧の形式が正しくありません")
        for i, u in enumerate(lst, 1):
            if not isinstance(u, dict) or not text(u.get("body")):
                raise ValueError(f"{i}番目の{name}の本文の形式が正しくありません")
            if not text(u.get("mark")) or u.get("level") not in (None, "multi", "single"):
                raise ValueError(f"{i}番目の{name}の記号の形式が正しくありません")
            try:
                figures(u)
                if kind == "paragraph":
                    units(u.get("items", []), "item")
                elif kind == "item":
                    units(u.get("sub_items", []), "sub")
                elif kind == "sub":
                    # 解析結果では sub_items2 が null のことがある
                    v = u.get("sub_items2")
                    units([] if v is None else v, "sub2")
            except ValueError as e:
                raise ValueError(f"{i}番目の{name}：{e}") from e

    def numbered(u):
        # 番号は整数、元の番号の文字列は文字列（比較・整形で文字列として扱うため）
        n = u.get("number", 0)
        if (n is not None and (isinstance(n, bool) or not isinstance(n, int))) or not text(u.get("number_str")):
            raise ValueError(f"番号の形式が正しくありません（number={n!r}, number_str={u.get('number_str')!r}）")

    def article(a):
        if not isinstance(a, dict) or a.get("type") != "article":
            raise ValueError("条の形式が正しくありません")
        if not text(a.get("title")):
            raise ValueError("条タイトルの形式が正しくありません")
        numbered(a)
        figures(a)
        units(a.get("paragraphs", []), "paragraph")

    def located(fn, where):
        try:
            fn()
        except ValueError as e:
            raise ValueError(f"{where}：{e}") from e

    for ti, node in enumerate(tree, 1):
        if not isinstance(node, dict):
            raise ValueError(f"先頭から{ti}番目の要素の形式が正しくありません")
        kind = node.get("type")
        if not isinstance(kind, str) or not kind:
            raise ValueError(f"先頭から{ti}番目の要素の種類が正しくありません")
        if kind not in _TREE_TYPES:
            # 章・条の外にある本文（節の行など）・表・図。画面では編集しないので形だけ確かめる。
            # それ以外の種類も、整形は読み飛ばすので止めない（解析で種類が増えても編集できなくならないように）
            if kind == "plain" and not text(node.get("body")):
                raise ValueError(f"先頭から{ti}番目の本文の形式が正しくありません")
            # 附則（画面で版の施行日から足したもの）。見出しと本文の2つの文字列
            if kind == "supplement" and not (text(node.get("title")) and text(node.get("body"))
                                             and len(node["title"]) <= 200 and len(node["body"]) <= 5000):
                raise ValueError(f"先頭から{ti}番目の附則の形式が正しくありません")
            if kind in ("table", "drawing"):
                located(lambda node=node, kind=kind: _check_figures({kind + "s": [node]}, used),
                        f"先頭から{ti}番目の{'表' if kind == 'table' else '図'}")
            continue
        if node["type"] == "chapter":
            where = f"{_s(node.get('number_str'))}{_s(node.get('title'))}".strip() or f"{ti}番目の章"

            def chapter_head(node=node):
                if not text(node.get("title")) or not isinstance(node.get("articles", []), list):
                    raise ValueError("章タイトルか条の一覧の形式が正しくありません")
                numbered(node)
                figures(node)
            located(chapter_head, where)
            for ai, a in enumerate(node.get("articles", []), 1):
                aw = (_s(a.get("number_str")) or _s(a.get("title")) or f"{ai}番目の条") if isinstance(a, dict) else f"{ai}番目の条"
                located(lambda a=a: article(a), f"{where}の{aw}")
        elif node["type"] == "article":
            located(lambda node=node: article(node), _s(node.get("number_str")) or _s(node.get("title")) or f"{ti}番目の条")
        else:
            # 前付け・後付けは編集しないが、壊れた下書きで整形が落ちないよう形だけ確かめる
            idx, texts = node.get("indices", []), node.get("texts", [])
            if (not isinstance(idx, list) or not all(isinstance(i, int) and not isinstance(i, bool) for i in idx)
                    or not isinstance(texts, list) or not all(isinstance(s, str) for s in texts)):
                raise ValueError("表紙・附則など（前付け・後付け）の形式が正しくありません")


def read_tree(json_path):
    """解析済みの木を JSON 文字列で返す（画面で編集を始めるときに使う）。"""
    with open(json_path, encoding="utf-8") as f:
        return f.read()


def write_tree(path, tree_json):
    """画面から受け取った木を確かめてから保存し、パスを返す。"""
    try:
        tree = json.loads(tree_json)
    except ValueError:
        raise ValueError("編集内容の形式が正しくありません")
    _check_tree(tree)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(tree, f, ensure_ascii=False)
    return path


_marks_cache = {}


def _marks_spec(template_key, custom_json, work_dir):
    """
    ③で選んだテンプレートの、階層ごとの番号の書き方（新旧対照表の改正案の表示に使う）。
    読み取れなければ空（これまでの表示になる）。カスタムは設定から作ったテンプレートを読む。
    """
    if not template_key:
        return {}
    ck = (template_key, custom_json if template_key == "custom" else "")
    if ck in _marks_cache:
        return _marks_cache[ck]
    try:
        if template_key == "custom":
            path, M = _build_custom(custom_json, work_dir)
        elif template_key in TEMPLATES:
            path, M = os.path.join(TEMPLATE_DIR, TEMPLATES[template_key]), TEMPLATE_MAP[template_key]
        else:
            return {}
        spec = comparison.marks_spec_from_template(path, M)
    except Exception as e:      # 読めなくても新旧対照表は作る（記号はこれまでの表示）。失敗は覚えない
        print(f"  ※ テンプレートの番号の書き方を読み取れませんでした: {e}")
        return {}
    if len(_marks_cache) >= 20:     # カスタムの設定を変えるたびに増えるので、古いものから捨てる
        _marks_cache.pop(next(iter(_marks_cache)))
    _marks_cache[ck] = spec
    return spec


def compare_rows(json_path, edited_json_path="", template_key="", custom_json=""):
    """
    新旧対照表の画面用データ {"blocks"（全文の並び）, "rows"（変更のあった章・条）, "marks"} を JSON 文字列で返す。
    改正案は、画面で編集した木（edited_json_path）があればそれ、無ければ解析結果そのもの。
    改正案の記号は、③で選んだテンプレート（template_key・custom_json）の書き方で表示する。
    """
    spec = _marks_spec(template_key, custom_json, os.path.dirname(json_path))
    return json.dumps(comparison.view_files(json_path, edited_json_path or None, spec),
                      ensure_ascii=False)


def make_comparison(src_path, json_path, notes_json="", edited_json_path="",
                    template_key="", custom_json="", labels_json=""):
    """新旧対照表の docx を作り、パスを返す。notes_json は {行の id: 備考}。
    labels_json は表の見出し ["（第1版）", "改正案"]。正しくなければ「現行」「改正案」。"""
    try:
        notes = json.loads(notes_json or "{}")
    except ValueError:
        raise ValueError("備考の内容を読み取れませんでした")
    if not isinstance(notes, dict):
        raise ValueError("備考の内容を読み取れませんでした")
    notes = {str(k): str(v) for k, v in notes.items()}
    base = os.path.splitext(os.path.basename(src_path))[0]
    out = os.path.join(os.path.dirname(src_path), f"{base}_新旧対照表.docx")
    spec = _marks_spec(template_key, custom_json, os.path.dirname(src_path))
    rows = comparison.compare_files(json_path, edited_json_path or None, spec)
    try:
        labels = json.loads(labels_json or "[]")
    except ValueError:
        labels = []
    if not (isinstance(labels, list) and len(labels) == 2
            and all(isinstance(x, str) and 0 < len(x) <= 200 for x in labels)):
        labels = ["現行", "改正案"]
    return comparison.build_docx(rows, notes, out,
                                 subtitle=os.path.basename(src_path),
                                 old_label=labels[0], new_label=labels[1])


def make_sample(custom_json, tree_json, work_dir):
    """
    書式の見本 docx（画面のプレビューと同じ例文）を作り、パスを返す。
    設定が埋め込まれるので、この docx を読み込めば設定を復元できる。
    """
    os.makedirs(work_dir, exist_ok=True)
    tree = json.loads(tree_json)
    if not isinstance(tree, list):
        raise ValueError("見本の例文の形式が正しくありません")
    json_path = os.path.join(work_dir, "_見本.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(tree, f, ensure_ascii=False)
    template_path, template_map = _build_custom(custom_json, work_dir)
    output_path = os.path.join(work_dir, "就業規則_書式見本.docx")
    if os.path.exists(output_path):
        os.remove(output_path)
    apply_style(json_path=json_path, template_path=template_path,
                template_key="custom", template_map=template_map,
                output_path=output_path, show_notice=False,
                keep_crossref=False, verify=False)
    return output_path


def read_custom(docx_path):
    """
    整形済み docx（テンプレ1〜4・カスタム。Word で手直ししたものを含む）から
    書式の設定を読み取り、{"settings": 設定, "notes": 注意書き} の JSON 文字列で返す。
    """
    return json.dumps(settings_reader.read_from_docx(docx_path),
                      ensure_ascii=False)
