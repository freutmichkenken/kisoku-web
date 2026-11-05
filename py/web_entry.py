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


def _check_tree(tree):
    """画面から受け取った木の形を確かめる（壊れた下書きで整形が落ちないように）。"""
    if not isinstance(tree, list):
        raise ValueError("編集内容の形式が正しくありません")

    def text(v):
        return v is None or isinstance(v, str)

    def units(lst, kind):
        if not isinstance(lst, list):
            raise ValueError("編集内容の形式が正しくありません")
        for u in lst:
            if not isinstance(u, dict) or not text(u.get("body")):
                raise ValueError("編集内容の形式が正しくありません")
            if kind == "paragraph":
                units(u.get("items", []), "item")
            elif kind == "item":
                units(u.get("sub_items", []), "sub")
            elif kind == "sub":
                # 解析結果では sub_items2 が null のことがある
                v = u.get("sub_items2")
                units([] if v is None else v, "sub2")

    def numbered(u):
        # 番号は整数、元の番号の文字列は文字列（比較・整形で文字列として扱うため）
        n = u.get("number", 0)
        if (n is not None and (isinstance(n, bool) or not isinstance(n, int))) or not text(u.get("number_str")):
            raise ValueError("編集内容の形式が正しくありません")

    def article(a):
        if not isinstance(a, dict) or a.get("type") != "article" or not text(a.get("title")):
            raise ValueError("編集内容の形式が正しくありません")
        numbered(a)
        units(a.get("paragraphs", []), "paragraph")

    for node in tree:
        if not isinstance(node, dict) or node.get("type") not in _TREE_TYPES:
            raise ValueError("編集内容の形式が正しくありません")
        if node["type"] == "chapter":
            if not text(node.get("title")) or not isinstance(node.get("articles", []), list):
                raise ValueError("編集内容の形式が正しくありません")
            numbered(node)
            for a in node.get("articles", []):
                article(a)
        elif node["type"] == "article":
            article(node)
        else:
            # 前付け・後付けは編集しないが、壊れた下書きで整形が落ちないよう形だけ確かめる
            idx, texts = node.get("indices", []), node.get("texts", [])
            if (not isinstance(idx, list) or not all(isinstance(i, int) and not isinstance(i, bool) for i in idx)
                    or not isinstance(texts, list) or not all(isinstance(s, str) for s in texts)):
                raise ValueError("編集内容の形式が正しくありません")


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


def compare_rows(json_path, edited_json_path=""):
    """
    新旧対照表の画面用データ {"blocks"（全文の並び）, "rows"（変更のあった章・条）} を JSON 文字列で返す。
    改正案は、画面で編集した木（edited_json_path）があればそれ、無ければ解析結果そのもの。
    """
    return json.dumps(comparison.view_files(json_path, edited_json_path or None),
                      ensure_ascii=False)


def make_comparison(src_path, json_path, notes_json="", edited_json_path=""):
    """新旧対照表の docx を作り、パスを返す。notes_json は {行の id: 備考}。"""
    try:
        notes = json.loads(notes_json or "{}")
    except ValueError:
        raise ValueError("備考の内容を読み取れませんでした")
    if not isinstance(notes, dict):
        raise ValueError("備考の内容を読み取れませんでした")
    notes = {str(k): str(v) for k, v in notes.items()}
    base = os.path.splitext(os.path.basename(src_path))[0]
    out = os.path.join(os.path.dirname(src_path), f"{base}_新旧対照表.docx")
    rows = comparison.compare_files(json_path, edited_json_path or None)
    return comparison.build_docx(rows, notes, out,
                                 subtitle=os.path.basename(src_path))


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
