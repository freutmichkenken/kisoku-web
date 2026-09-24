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


def format_docx(src_path, json_path, template_key,
                check_hyoki=False, insert_toc=False):
    """解析済み JSON にテンプレートを適用し、出力ファイルのパスを返す。"""
    if template_key not in TEMPLATES:
        raise ValueError(f"不明なテンプレートです: {template_key}")

    work_dir = os.path.dirname(src_path)
    base = os.path.splitext(os.path.basename(src_path))[0]
    output_path = os.path.join(work_dir, f"{base}_整形済_{template_key}.docx")
    report_path = os.path.splitext(output_path)[0] + "_照合レポート.md"
    for p in (output_path, report_path):      # 同じテンプレでやり直す場合
        if os.path.exists(p):
            os.remove(p)

    apply_style(
        json_path=json_path,
        template_path=os.path.join(TEMPLATE_DIR, TEMPLATES[template_key]),
        template_key=template_key,
        output_path=output_path,
        source_docx=src_path,        # 表・前付けの引き継ぎ元
        show_notice=False,           # 注意事項は画面に常時表示している
        check_hyoki=bool(check_hyoki),
        insert_toc=bool(insert_toc),
    )

    if not os.path.exists(report_path):
        report_path = None

    return json.dumps({"output_path": output_path,
                       "report_path": report_path},
                      ensure_ascii=False)
