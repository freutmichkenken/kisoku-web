# -*- coding: utf-8 -*-
"""
画面での編集（改正案の木）まわりの検証。 python3 py/test_editing.py で実行。
- web_entry.write_tree の形の確認
- 編集した木の整形と、照合レポートでの「編集で消した／足した文章」の切り分け
- 新旧対照表の block の位置（new_path）と、編集しても変わらない行の id
"""

import contextlib
import copy
import io
import json
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import comparison as c  # noqa: E402
import web_entry  # noqa: E402
from apply_style import apply_style  # noqa: E402
from kisoku_parser import parse_kisoku  # noqa: E402

SAMPLE = os.path.join(ROOT, "tests", "samples", "messy_kisoku.docx")
TEMPLATE = os.path.join(ROOT, "templates", "就業規則テンプレ1.docx")


def quiet(fn, *a, **k):
    with contextlib.redirect_stdout(io.StringIO()):
        return fn(*a, **k)


tree = quiet(parse_kisoku, SAMPLE)
chapters = [i for i, n in enumerate(tree) if n["type"] == "chapter"]
assert chapters, "サンプルに章がありません"
ci = chapters[0]

with tempfile.TemporaryDirectory() as d:
    # --- write_tree：正しい木は保存でき、壊れた木は保存しない ---
    p = web_entry.write_tree(os.path.join(d, "ok.json"), json.dumps(tree, ensure_ascii=False))
    with open(p, encoding="utf-8") as f:
        assert json.load(f) == tree
    for bad in ("not json", "{}", "[1]", '[{"type": "x"}]',
                '[{"type": "chapter", "title": 1, "articles": []}]',
                '[{"type": "chapter", "title": "a", "articles": [{"type": "paragraph"}]}]',
                '[{"type": "article", "title": "a", "paragraphs": [{"body": 3}]}]',
                '[{"type": "article", "title": "a", "paragraphs": [{"body": "x", "items": "y"}]}]',
                '[{"type": "article", "title": "a", "number": "5", "paragraphs": []}]',
                '[{"type": "article", "title": "a", "number_str": 5, "paragraphs": []}]',
                '[{"type": "chapter", "title": "a", "number": null, "number_str": [], "articles": []}]',
                '[{"type": "front_matter", "indices": ["x"], "texts": []}]',
                '[{"type": "back_matter", "indices": [1], "texts": [2]}]'):
        try:
            web_entry.write_tree(os.path.join(d, "bad.json"), bad)
        except ValueError:
            continue
        raise AssertionError(f"壊れた木を保存した: {bad}")
    assert not os.path.exists(os.path.join(d, "bad.json"))

    # --- 編集：第1章の第1条の第1項を書き換え、条を1つ足す ---
    edited = copy.deepcopy(tree)
    art0 = edited[ci]["articles"][0]
    old_body = art0["paragraphs"][0]["body"]
    art0["paragraphs"][0]["body"] = "この規則は、編集で書き換えた本文である。"
    edited[ci]["articles"].insert(1, {
        "type": "article", "number": 0, "branch": 0, "number_str": "", "title": "新設の条",
        "title_note": "", "paragraphs": [{"type": "paragraph", "number": 1,
                                          "body": "編集で足した条の本文である。", "items": []}]})
    base_path = os.path.join(d, "base.json")
    edit_path = os.path.join(d, "edited.json")
    for path, t in ((base_path, tree), (edit_path, edited)):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(t, f, ensure_ascii=False)

    # --- 整形：編集した本文が出力され、照合レポートでは編集の分として分ける ---
    out = os.path.join(d, "out.docx")
    quiet(apply_style, json_path=edit_path, template_path=TEMPLATE, template_key="t1",
          output_path=out, source_docx=SAMPLE, show_notice=False, base_json_path=base_path)
    from docx import Document  # noqa: E402
    texts = [para.text for para in Document(out).paragraphs]
    assert "この規則は、編集で書き換えた本文である。" in texts
    assert "編集で足した条の本文である。" in texts
    report = open(os.path.splitext(out)[0] + "_照合レポート.md", encoding="utf-8").read()
    assert "## 画面の編集で消した文章（1件）" in report, report
    assert old_body.strip()[:10] in report.split("## 画面の編集で消した文章")[1]
    assert "## 画面の編集で足した文章（" in report
    missing_sec = report.split("## 原本にあって出力に無い文章")[1].split("##")[0]
    assert old_body.strip()[:10] not in missing_sec, missing_sec

    # 編集しないとき（base_json_path なし）は、今までどおり編集の欄を出さない
    out2 = os.path.join(d, "out2.docx")
    quiet(apply_style, json_path=base_path, template_path=TEMPLATE, template_key="t1",
          output_path=out2, source_docx=SAMPLE, show_notice=False)
    assert "画面の編集" not in open(os.path.splitext(out2)[0] + "_照合レポート.md", encoding="utf-8").read()

# --- 照合レポートの切り分け：本当の欠落は編集の分に入れない ---
import verify_report as vr  # noqa: E402

N = vr.normalize
base_t = [{"type": "article", "title": "目的", "paragraphs": [
    {"type": "paragraph", "body": "この規則は、従業員の就業に関する事項を定める。前半の文はそのまま残す。", "items": []},
    {"type": "paragraph", "body": "別表のとおりとする。", "items": []}]}]
edit_t = copy.deepcopy(base_t)
edit_t[0]["paragraphs"][0]["body"] = "この規則は、従業員の就業に関する事項を定める。前半の文はそのまま残す。後半を足した。"
missing = [("第1条この規則は、従業員の就業に関する事項を定める。", N("第1条この規則は、従業員の就業に関する事項を定める。")),
           ("目的", N("目的")),                                   # 短い語句（編集で消した長文に含まれるが偶然）
           ("前半の文はそのまま残す。", N("前半の文はそのまま残す。"))]  # 編集後にも残っている文章
res = vr.split_edited({"missing": list(missing), "added": []}, base_t, edit_t)
assert res["edit_removed"] == [], res["edit_removed"]          # どれも編集後の木に残っている → 本当の欠落
assert len(res["missing"]) == 3
edit_t2 = copy.deepcopy(base_t)
edit_t2[0]["paragraphs"].pop(1)                                 # 「別表のとおりとする。」を消した
res = vr.split_edited({"missing": [("別表のとおりとする。", N("別表のとおりとする。"))], "added": []}, base_t, edit_t2)
assert [r for r, _ in res["edit_removed"]] == ["別表のとおりとする。"] and res["missing"] == []
assert vr._HEAD_LABEL.sub("", "第十二条の二会社は") == "会社は"

# --- 新旧対照表：block の new_path は編集した木の条・章を指す ---
v = c.view(tree, edited)
for b in v["blocks"]:
    if b["new_path"] is None:
        continue
    node = edited[b["new_path"][0]] if len(b["new_path"]) == 1 else \
        edited[b["new_path"][0]]["articles"][b["new_path"][1]]
    assert node["type"] == b["unit"], (b["new_path"], node["type"], b["unit"])
added = [b for b in v["blocks"] if b["kind"] == "added"]
assert len(added) == 1 and added[0]["new_path"] == [ci, 1]
changed = [b for b in v["blocks"] if b["kind"] == "changed"]
assert changed and changed[0]["new_path"] == [ci, 0] and changed[0]["old_path"] == [ci, 0]

# --- 行の id は現行の見出しで決まり、編集で改正案の中身が変わっても同じ ---
ids_before = {r["old_heading"]: r["id"] for r in c.view(tree, tree)["rows"]}
for r in v["rows"]:
    if r["kind"] in ("changed", "renumbered") and r["old_heading"] in ids_before:
        assert r["id"] == ids_before[r["old_heading"]], (r["id"], ids_before[r["old_heading"]])
assert all(r["id"].startswith(("o:", "n:", "summary:")) for r in v["rows"])

print("OK")
