# -*- coding: utf-8 -*-
"""
ブラウザ版（Pyodide）から呼ばれる入口。

Colab 版の kisoku_app.py から、画面（ipywidgets）部分を取り除き、
「解析」「整形」の2つの処理だけを関数として公開したもの。
画面は index.html / worker.js 側が担当する。

戻り値は JavaScript へ渡しやすいよう JSON 文字列にしている。
"""

import os
import json

from kisoku_parser import parse_kisoku
from apply_style import apply_style, TEMPLATE_MAP  # noqa: F401
import custom_template

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


def parse(src_path, use_gemini=False, layout_as_figure=False, api_key=""):
    """整形前 docx を解析して JSON を保存し、章・条の数を返す。"""
    _set_api_key(api_key if use_gemini else "")

    tree = parse_kisoku(
        src_path,
        use_gemini=bool(use_gemini),
        treat_layout_as_figure=bool(layout_as_figure),
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
                custom_json=""):
    """
    解析済み JSON にテンプレートを適用し、出力ファイルのパスを返す。
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
        json_path=json_path,
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
    """見本・整形済み docx に埋め込んだ書式の設定を JSON 文字列で返す。"""
    return json.dumps(custom_template.read_settings(docx_path),
                      ensure_ascii=False)
