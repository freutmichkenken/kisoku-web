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
    # 知らない種類の要素（{"type": "x"}）は、整形が読み飛ばすので受け付ける（下の「章・条の外」の検証を参照）
    for bad in ("not json", "{}", "[1]",
                '[{"type": "chapter", "title": 1, "articles": []}]',
                '[{"type": "chapter", "title": "a", "articles": [{"type": "paragraph"}]}]',
                '[{"type": "article", "title": "a", "paragraphs": [{"body": 3}]}]',
                '[{"type": "article", "title": "a", "paragraphs": [{"body": "x", "items": "y"}]}]',
                '[{"type": "article", "title": "a", "number": "5", "paragraphs": []}]',
                '[{"type": "article", "title": "a", "paragraphs": [{"body": "x", "mark": 3}]}]',
                '[{"type": "article", "title": "a", "paragraphs": [{"body": "x", "items": [{"body": "y", "sub_items": [{"body": "z", "level": "x"}]}]}]}]',
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

# --- 画像の挿入・表の移動と削除（2回目）---
import base64  # noqa: E402
import struct  # noqa: E402
import zlib  # noqa: E402
from docx import Document  # noqa: E402,F811
from docx.oxml.ns import qn  # noqa: E402


def png(w, h):
    """w×h の白い PNG を作る（テスト用）。"""
    raw = b"".join(b"\x00" + b"\xff\xff\xff" * w for _ in range(h))
    chunk = lambda t, d: struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d))  # noqa: E731
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


wide = base64.b64encode(png(3000, 300)).decode()      # 本文の幅より広い → 縮める
img_tree = copy.deepcopy(tree)
img_tree[ci]["articles"][0]["paragraphs"][0]["drawings"] = [
    {"type": "image", "name": "図1.png", "mime": "image/png", "data": wide}]
with tempfile.TemporaryDirectory() as d:
    p = web_entry.write_tree(os.path.join(d, "img.json"), json.dumps(img_tree, ensure_ascii=False))
    out = os.path.join(d, "img.docx")
    quiet(apply_style, json_path=p, template_path=TEMPLATE, template_key="t1",
          output_path=out, source_docx=SAMPLE, show_notice=False)
    doc = Document(out)
    sec = doc.sections[-1]
    shapes = list(doc.inline_shapes)
    assert len(shapes) == 1, len(shapes)
    assert shapes[0].width <= sec.page_width - sec.left_margin - sec.right_margin
    assert abs(shapes[0].width / shapes[0].height - 10) < 0.1        # 縦横比を保つ
    # 画像は第1項の本文の直後に出る
    texts = [para.text for para in doc.paragraphs]
    k = texts.index(tree[ci]["articles"][0]["paragraphs"][0]["body"].strip()) \
        if tree[ci]["articles"][0]["paragraphs"][0]["body"].strip() in texts else None
    pic_at = next(i for i, para in enumerate(doc.paragraphs) if para._p.xpath(".//pic:pic"))
    assert k is None or pic_at == k + 1, (k, pic_at)

    # 壊れた画像・形式の違う画像・大きすぎる画像は保存しない
    def with_img(dw):
        t = copy.deepcopy(tree)
        t[ci]["articles"][0]["paragraphs"][0]["drawings"] = [dw]
        return json.dumps(t, ensure_ascii=False)
    jpeg_as_png = base64.b64encode(b"\xff\xd8\xff\xe0" + b"0" * 100).decode()
    for bad in ({"type": "image", "mime": "image/png", "data": jpeg_as_png},          # 中身が JPEG
                {"type": "image", "mime": "image/gif", "data": wide},                 # 形式が対象外
                {"type": "image", "mime": "image/png", "data": "!!!"},                # base64 でない
                {"type": "image", "mime": "image/png",
                 "data": base64.b64encode(b"\x89PNG\r\n\x1a\n" + b"0" * (web_entry.IMAGE_MAX_BYTES + 1)).decode()},
                {"type": "drawing", "index": "x"},
                {"type": "drawing", "index": -1},
                {"type": "drawing", "index": 3, "text": 5}):
        try:
            web_entry.write_tree(os.path.join(d, "bad.json"), with_img(bad))
        except ValueError:
            continue
        raise AssertionError(f"壊れた画像を保存した: {str(bad)[:60]}")

    # 表の rows が壊れていれば保存しない（index が正しくても rows を確かめる）
    for rows in ("x", [1], [["a", 2]]):
        t = copy.deepcopy(tree)
        t[ci]["articles"][0]["paragraphs"][0]["tables"] = [{"type": "table", "index": 3, "rows": rows}]
        try:
            web_entry.write_tree(os.path.join(d, "bad.json"), json.dumps(t, ensure_ascii=False))
        except ValueError:
            continue
        raise AssertionError(f"壊れた表を保存した: {rows}")

    # 画像の合計が上限を超えれば保存しない（1枚の上限より小さい画像を何枚も）
    old_total = web_entry.IMAGE_TOTAL_BYTES
    web_entry.IMAGE_TOTAL_BYTES = len(base64.b64decode(wide)) * 2
    try:
        t = copy.deepcopy(tree)
        t[ci]["articles"][0]["paragraphs"][0]["drawings"] = [
            {"type": "image", "name": f"{k}.png", "mime": "image/png", "data": wide} for k in range(3)]
        try:
            web_entry.write_tree(os.path.join(d, "bad.json"), json.dumps(t, ensure_ascii=False))
            raise AssertionError("画像の合計の上限を確かめていない")
        except ValueError as e:
            assert "合計" in str(e), e
    finally:
        web_entry.IMAGE_TOTAL_BYTES = old_total

    # 縦長の画像は本文の高さに収める
    tall = copy.deepcopy(tree)
    tall[ci]["articles"][0]["paragraphs"][0]["drawings"] = [
        {"type": "image", "name": "縦長.png", "mime": "image/png", "data": base64.b64encode(png(200, 4000)).decode()}]
    p = web_entry.write_tree(os.path.join(d, "tall.json"), json.dumps(tall, ensure_ascii=False))
    quiet(apply_style, json_path=p, template_path=TEMPLATE, template_key="t1",
          output_path=os.path.join(d, "tall.docx"), source_docx=SAMPLE, show_notice=False)
    doc = Document(os.path.join(d, "tall.docx"))
    sec = doc.sections[-1]
    sh = list(doc.inline_shapes)[0]
    assert sh.height <= sec.page_height - sec.top_margin - sec.bottom_margin
    assert abs(sh.height / sh.width - 20) < 0.2

    # 新旧対照表には［画像］として出る
    v = c.view(tree, img_tree)
    ch_rows = [r for r in v["rows"] if r["kind"] == "changed"]
    assert ch_rows and any("［画像］　図1.png" in "".join(s for s, _ in line) for line in ch_rows[0]["new"])

# 表のある規程で、表を次の項へ移す・消す
TSAMPLE = os.path.join(ROOT, "tests", "samples", "01_01_202601.docx")
ttree = quiet(parse_kisoku, TSAMPLE)
loc = None
for ti, n in enumerate(ttree):
    for ai, a in enumerate(n.get("articles", [])):
        ps = a.get("paragraphs", [])
        for pi, para in enumerate(ps[:-1]):
            if para.get("tables") and not para.get("items"):
                loc = (ti, ai, pi)
                break
        if loc:
            break
    if loc:
        break
assert loc, "表のある項（次の項もある）が見つかりません"
ti, ai, pi = loc
with tempfile.TemporaryDirectory() as d:
    def body_order(t):
        p = web_entry.write_tree(os.path.join(d, "t.json"), json.dumps(t, ensure_ascii=False))
        out = os.path.join(d, "t.docx")
        quiet(apply_style, json_path=p, template_path=os.path.join(ROOT, "templates", "就業規則テンプレ1.docx"),
              template_key="t1", output_path=out, source_docx=TSAMPLE, show_notice=False)
        doc = Document(out)
        seq = []
        for el in doc.element.body.iterchildren():
            tag = el.tag.split("}")[1]
            seq.append("T" if tag == "tbl" else "".join(el.itertext()).strip() if tag == "p" else "")
        return seq
    next_body = ttree[ti]["articles"][ai]["paragraphs"][pi + 1]["body"].strip()
    before = body_order(ttree)
    moved = copy.deepcopy(ttree)
    ps = moved[ti]["articles"][ai]["paragraphs"]
    ps[pi + 1]["tables"] = ps[pi].pop("tables") + ps[pi + 1].get("tables", [])
    after = body_order(moved)
    assert before.count("T") == after.count("T")
    key = "".join(next_body.split())[:20]
    find = lambda seq: next(i for i, s in enumerate(seq) if key in "".join(s.split()))  # noqa: E731
    i_next = find(after)
    assert after[i_next + 1] == "T", after[i_next - 2:i_next + 3]          # 表は次の項の直後に出る
    assert before[find(before) - 1] == "T", before[find(before) - 2:find(before) + 2]  # 元は前の項の後ろ
    deleted = copy.deepcopy(ttree)
    deleted[ti]["articles"][ai]["paragraphs"][pi]["tables"] = []
    n_del = len(ttree[ti]["articles"][ai]["paragraphs"][pi]["tables"])
    assert body_order(deleted).count("T") == before.count("T") - n_del

# --- 章・条の外にある本文・表・図（木の一番上の段）があっても、編集した木を受け付ける ---
# 節の行（第1節）のあとに条が来ないまま章が変わると、節の行は章の外の本文（plain）になる
with tempfile.TemporaryDirectory() as d:
    src = os.path.join(d, "plain.docx")
    pd = Document()
    for line in ("就業規則", "第1章　総則", "（目的）", "第1条　この規則は、従業員の就業に関する事項を定める。",
                 "第1節　補則", "第2章　勤務", "（労働時間）", "第2条　所定労働時間は8時間とする。"):
        pd.add_paragraph(line)
    pd.save(src)
    ptree = quiet(parse_kisoku, src)
    assert "plain" in [n["type"] for n in ptree], [n["type"] for n in ptree]
    pedit = copy.deepcopy(ptree)
    next(n for n in pedit if n["type"] == "chapter")["articles"][0]["paragraphs"][0]["body"] = "編集した本文。"
    path = web_entry.write_tree(os.path.join(d, "e.json"), json.dumps(pedit, ensure_ascii=False))
    quiet(apply_style, json_path=path, template_path=TEMPLATE, template_key="t1",
          output_path=os.path.join(d, "o.docx"), source_docx=src, show_notice=False)
    assert "編集した本文。" in [p.text for p in Document(os.path.join(d, "o.docx")).paragraphs]
    # 一番上の段の表・図、解析で今後増えるかもしれない種類も受け付ける
    extra = [{"type": "plain", "body": "章の外の文"}, {"type": "table", "index": 2, "rows": [["a", "b"]]},
             {"type": "drawing", "index": 3, "text": ""}, {"type": "future_kind", "x": 1}] + pedit
    web_entry.write_tree(os.path.join(d, "x.json"), json.dumps(extra, ensure_ascii=False))
    for bad in ({"type": "plain", "body": 3}, {"type": "table", "index": 1, "rows": "x"},
                {"type": "drawing", "index": -2}, "文字列", {"body": "種類が無い"}, {"type": 5}):
        try:
            web_entry.write_tree(os.path.join(d, "bad.json"), json.dumps([bad] + pedit, ensure_ascii=False))
        except ValueError:
            continue
        raise AssertionError(f"壊れた要素を保存した: {bad}")

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

# --- 画面で足した表（index なし）：rows から作り、項の本文の直後に出る ---
tbl_tree = copy.deepcopy(tree)
p0 = tbl_tree[ci]["articles"][0]["paragraphs"][0]
p0.setdefault("tables", []).append(
    {"type": "table", "rows": [["区分", "割増率"], ["時間外労働", "２５％"], ["休日労働", "３５％\n（法定休日）"]]})
with tempfile.TemporaryDirectory() as d:
    p = web_entry.write_tree(os.path.join(d, "tbl.json"), json.dumps(tbl_tree, ensure_ascii=False))
    for key in ("t1", "t2", "t3", "t4"):
        tmpl = os.path.join(ROOT, "templates", f"就業規則テンプレ{key[1]}.docx")
        out = os.path.join(d, f"tbl_{key}.docx")
        quiet(apply_style, json_path=p, template_path=tmpl, template_key=key,
              output_path=out, source_docx=SAMPLE, show_notice=False)
        doc = Document(out)
        found = [t for t in doc.tables if t.cell(0, 0).text == "区分"]
        assert len(found) == 1, (key, len(found))
        t = found[0]
        assert [[cl.text for cl in r.cells] for r in t.rows] == \
            [["区分", "割増率"], ["時間外労働", "２５％"], ["休日労働", "３５％\n（法定休日）"]], key
        # 罫線がある（テンプレートに Table Grid が無ければ直接付ける）。tblBorders は tblLook より前に置く
        tp = t._tbl.tblPr
        names = [ch.tag.split('}')[1] for ch in tp]
        assert "tblBorders" in names or tp.find(qn('w:tblStyle')) is not None, (key, names)
        if "tblBorders" in names and "tblLook" in names:
            assert names.index("tblBorders") < names.index("tblLook"), (key, names)
        # 表は第1項の本文の直後（間に段落を挟まない）
        prev = t._tbl.getprevious()
        first = p0["body"].strip()
        assert prev is not None and first and first in "".join(prev.itertext()), (key, "".join(prev.itertext())[:30])

    # 大きすぎる表・空の表は保存しない（元の文書の表＝index ありには上限をかけない）
    def with_tbl(tb):
        t = copy.deepcopy(tree)
        t[ci]["articles"][0]["paragraphs"][0]["tables"] = [tb]
        return json.dumps(t, ensure_ascii=False)
    R, C = web_entry.TABLE_MAX_ROWS, web_entry.TABLE_MAX_COLS
    for bad in ({"type": "table", "rows": []},
                {"type": "table", "rows": [[]]},
                {"type": "table", "rows": [[""]] * (R + 1)},
                {"type": "table", "rows": [[""] * (C + 1)]},
                {"type": "table", "rows": [["あ" * (web_entry.TABLE_CELL_MAX + 1)]]}):
        try:
            web_entry.write_tree(os.path.join(d, "bad.json"), with_tbl(bad))
        except ValueError:
            continue
        raise AssertionError(f"大きすぎる・空の表を保存した: {str(bad)[:60]}")
    web_entry.write_tree(os.path.join(d, "ok.json"), with_tbl({"type": "table", "rows": [[""] * C for _ in range(R)]}))
    web_entry.write_tree(os.path.join(d, "ok.json"), with_tbl({"type": "table", "index": 0, "rows": [[""] * (C + 5)]}))
    # 画面で足した表のマスの合計：上限ちょうどは保存し、超えたら保存しない（元の表のマスは数えない）
    full = {"type": "table", "rows": [[""] * C for _ in range(R)]}               # 1つで R×C マス
    n_ok = web_entry.TABLE_TOTAL_CELLS // (R * C)
    t = copy.deepcopy(tree)
    units_ = [u for a in t[ci]["articles"] for u in a["paragraphs"]]
    assert len(units_) > n_ok, "サンプルの項が足りません"
    for u in units_[:n_ok]:
        u["tables"] = [copy.deepcopy(full)]
    units_[0]["tables"].append({"type": "table", "index": 0, "rows": [[""] * C for _ in range(R)]})
    web_entry.write_tree(os.path.join(d, "ok.json"), json.dumps(t, ensure_ascii=False))
    units_[n_ok]["tables"] = [{"type": "table", "rows": [[""]]}]
    try:
        web_entry.write_tree(os.path.join(d, "bad.json"), json.dumps(t, ensure_ascii=False))
        raise AssertionError("表のマスの合計の上限を確かめていない")
    except ValueError as e:
        assert "合計" in str(e), e

# 新旧対照表には［表］とセルの文字で出る
v = c.view(tree, tbl_tree)
ch_rows = [r for r in v["rows"] if r["kind"] == "changed"]
assert ch_rows and any("［表］　区分／割増率" in "".join(s for s, _ in line) for line in ch_rows[0]["new"])

# 照合レポート：画面で足した表のセルは「編集で足した文章」、消した元の表のセルは「編集で消した文章」になる
base_texts, tbl_texts = vr.tree_texts(tree), vr.tree_texts(tbl_tree)
assert {vr.normalize(s) for s in ("区分", "時間外労働", "３５％", "（法定休日）")} <= tbl_texts - base_texts
res = vr.split_edited({"missing": [], "added": [("時間外労働", vr.normalize("時間外労働")),
                                                 ("（法定休日）", vr.normalize("（法定休日）"))]}, tree, tbl_tree)
assert res["added"] == [] and len(res["edit_added"]) == 2, res
# 元の文書の表（index あり）を画面で消したとき（サンプルの第1章には表が無いので、木に元の表を置いて確かめる）
with_src = copy.deepcopy(tree)
with_src[ci]["articles"][0]["paragraphs"][0]["tables"] = [{"type": "table", "index": 0, "rows": [["元の表のセル", "賃金区分表"]]}]
res = vr.split_edited({"missing": [("賃金区分表", vr.normalize("賃金区分表"))], "added": []}, with_src, tree)
assert res["missing"] == [] and len(res["edit_removed"]) == 1, res
# 残っている表のセルは、出力で欠けていれば本当の欠落のまま（編集の分にしない）
res = vr.split_edited({"missing": [("賃金区分表", vr.normalize("賃金区分表"))], "added": []}, with_src, with_src)
assert len(res["missing"]) == 1 and res["edit_removed"] == [], res

# --- 画面で足した数式：Word の数式（分数）で、項の本文の直後に出る ---
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
fml = {"type": "formula", "before": "時間外労働割増賃金 ＝", "num": "基本給＋〇〇手当",
       "den": "1か月平均所定労働時間数", "after": "× 1.25 × 時間外労働時間数"}
f_tree = copy.deepcopy(tree)
fp0 = f_tree[ci]["articles"][0]["paragraphs"][0]
fp0["drawings"] = [dict(fml), {"type": "formula", "before": "", "num": "", "den": "", "after": ""},
                   {"type": "formula", "before": "月平均 ＝ 160時間"},
                   {"type": "formula", "before": "賃金 ＝", "num": "基本給", "den": "", "after": "× 1.25"},
                   {"type": "formula", "num": " ", "den": "12"}]
with tempfile.TemporaryDirectory() as d:
    p = web_entry.write_tree(os.path.join(d, "f.json"), json.dumps(f_tree, ensure_ascii=False))
    for key in ("t1", "t2", "t3", "t4"):
        out = os.path.join(d, f"f_{key}.docx")
        quiet(apply_style, json_path=p, template_path=os.path.join(ROOT, "templates", f"就業規則テンプレ{key[1]}.docx"),
              template_key=key, output_path=out, source_docx=SAMPLE, show_notice=False)
        doc = Document(out)
        maths = [para for para in doc.paragraphs if para._p.find(M + "oMathPara") is not None]
        assert len(maths) == 4, (key, len(maths))                 # 空の数式は出さない
        assert [t.text for t in maths[3]._p.iter(M + "t")] == ["12"], key   # 分子が空白だけなら分母の文字を出す
        # 分母が空なら分数にせず、文字だけを並べる
        assert maths[2]._p.find(".//" + M + "f") is None \
            and [t.text for t in maths[2]._p.iter(M + "t")] == ["賃金 ＝", "基本給", "× 1.25"], key
        m0 = maths[0]._p
        texts = lambda e: [t.text for t in e.iter(M + "t")]     # noqa: E731
        f = m0.find(".//" + M + "f")
        assert f is not None and texts(f.find(M + "num")) == ["基本給＋〇〇手当"] \
            and texts(f.find(M + "den")) == ["1か月平均所定労働時間数"], key
        assert texts(m0) == ["時間外労働割増賃金 ＝", "基本給＋〇〇手当", "1か月平均所定労働時間数",
                             "× 1.25 × 時間外労働時間数"], (key, texts(m0))
        assert all(r.find(M + "rPr/" + M + "nor") is not None for r in m0.iter(M + "r"))   # 標準テキスト
        assert maths[1]._p.find(".//" + M + "f") is None and texts(maths[1]._p) == ["月平均 ＝ 160時間"]
        # 本文の直後に出る（間に段落を挟まない）
        prev = maths[0]._p.getprevious()
        assert fp0["body"].strip() in "".join(prev.itertext()), key

    # 形の壊れた数式・長すぎる数式は保存しない
    for bad in ({"type": "formula", "num": 3},
                {"type": "formula", "den": ["x"]},
                {"type": "formula", "after": "あ" * (web_entry.FORMULA_PART_MAX + 1)},
                {"type": "formula", "num": "基本給\n＋手当"},
                {"type": "formula", "den": "時間\x00数"}):
        t = copy.deepcopy(tree)
        t[ci]["articles"][0]["paragraphs"][0]["drawings"] = [bad]
        try:
            web_entry.write_tree(os.path.join(d, "bad.json"), json.dumps(t, ensure_ascii=False))
        except ValueError:
            continue
        raise AssertionError(f"壊れた数式を保存した: {str(bad)[:60]}")

# 1行の文字（新旧対照表に出す書き方）
assert c.formula_text(fml) == "時間外労働割増賃金 ＝ （基本給＋〇〇手当）÷1か月平均所定労働時間数 × 1.25 × 時間外労働時間数"
assert c.formula_text({"num": "（365 － 年間所定休日日数）× 8", "den": "12"}) == "（（365 － 年間所定休日日数）× 8）÷12"
assert c.formula_text({"num": "（a＋b）", "den": "c"}) == "（a＋b）÷c"
assert c.formula_text({"num": "(a+b)", "den": "c"}) == "(a+b)÷c"                     # 半角の括弧も1組と見る
assert c.formula_text({"num": "（a＋b）＋（c＋d）", "den": "e"}) == "（（a＋b）＋（c＋d））÷e"   # 先頭の括弧が途中で閉じる
assert c.formula_text({"num": "基本給", "den": "", "after": "× 1.25"}) == "基本給 × 1.25"
assert c.formula_text({"before": "x", "num": "", "den": "", "after": ""}) == "x"
assert c.formula_text({}) == ""
v = c.view(tree, f_tree)
ch_rows = [r for r in v["rows"] if r["kind"] == "changed"]
assert ch_rows and any("［数式］　時間外労働割増賃金 ＝ （基本給＋〇〇手当）÷" in "".join(s for s, _ in line)
                       for line in ch_rows[0]["new"])
# 数式の中身だけ変えても、新旧対照表の変更として出る
f2 = copy.deepcopy(f_tree)
f2[ci]["articles"][0]["paragraphs"][0]["drawings"][0]["after"] = "× 1.30 × 時間外労働時間数"
assert [r for r in c.view(f_tree, f2)["rows"] if r["kind"] == "changed"], "数式の変更が対照表に出ない"

print("OK")
