# -*- coding: utf-8 -*-
"""新旧対照表（comparison.py）の検証。 python3 py/test_comparison.py で実行。"""

import copy
import glob
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import comparison as c  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def art(num, title, *bodies, number_str=None, para_nums=None, items=None):
    paras = []
    for i, b in enumerate(bodies, 1):
        p = {"type": "paragraph", "number": (para_nums or {}).get(i, i), "body": b,
             "items": []}
        if items and i in items:
            p["items"] = [{"type": "item", "number": k, "body": t}
                          for k, t in enumerate(items[i], 1)]
        paras.append(p)
    return {"type": "article", "number": num, "branch": 0,
            "number_str": number_str if number_str is not None else f"第{num}条",
            "title": title, "title_note": "", "paragraphs": paras}


def chapter(num, title, *arts):
    return {"type": "chapter", "number": num, "number_str": f"第{num}章",
            "title": title, "articles": list(arts)}


def text(lines):
    return "\n".join("".join(s for s, _ in segs) for segs in lines)


def underlined(lines):
    return "".join(s for segs in lines for s, ul in segs if ul)


BASE = [chapter(1, "総則",
                art(1, "目的", "この規則は、就業に関する事項を定める。"),
                art(2, "適用範囲", "この規則は、従業員に適用する。", "パートタイマーは別に定める。")),
        chapter(2, "勤務",
                art(3, "労働時間", "所定労働時間は8時間とする。", items={1: ["始業 9時", "終業 18時"]}),
                art(4, "休日", "休日は土曜日及び日曜日とする。"),
                art(5, "休暇", "年次有給休暇を与える。"),
                art(6, "遅刻", "遅刻は届け出る。"),
                art(7, "早退", "早退は届け出る。"))]

# --- 番号どおりの規則は、変更なし ---
assert c.compare(BASE, BASE) == []

# --- 半角・全角の番号の違いは変更にしない ---
zen = copy.deepcopy(BASE)
for ch in zen:
    for a in ch["articles"]:
        a["number_str"] = a["number_str"].translate(str.maketrans("0123456789", "０１２３４５６７８９"))
assert c.compare(zen, BASE) == []

# --- 条番号の重複：第3条が2つあると、以降の番号が1つずつずれる ---
dup = copy.deepcopy(BASE)
dup[1]["articles"][1]["number"] = 3          # 第4条（休日）を「第3条」と書き間違えた
dup[1]["articles"][1]["number_str"] = "第3条"
for k, a in enumerate(dup[1]["articles"][2:], 4):  # 以降も1つずつ小さい
    a["number"], a["number_str"] = k, f"第{k}条"
rows = c.compare(dup, dup)
assert [r["kind"] for r in rows] == ["renumbered", "renumbered", "summary"], rows
assert rows[0]["old_heading"] == "第３条（休日）" and rows[0]["new_heading"] == "第４条（休日）"
assert underlined(rows[0]["old"]) == "３" and underlined(rows[0]["new"]) == "４"
assert rows[2]["old_heading"] == "第５条～第６条" and rows[2]["new_heading"] == "第６条～第７条"
assert rows[2]["note"] == "条番号の繰り下げのみ（内容の変更なし、以降同様）"
assert text(rows[0]["new"]).startswith("（休日）\n第４条　休日は")

# --- 残りが1条だけなら、まとめずに載せる ---
rows = c.compare(dup[:1] + [dict(dup[1], articles=dup[1]["articles"][:4])],
                 dup[:1] + [dict(dup[1], articles=dup[1]["articles"][:4])])
assert [r["kind"] for r in rows] == ["renumbered"] * 3, rows

# --- 条番号の欠落：第4条が無いと、以降が繰り上がる ---
gap = copy.deepcopy(BASE)
for k, a in enumerate(gap[1]["articles"][1:], 5):
    a["number"], a["number_str"] = k, f"第{k}条"
rows = c.compare(gap, gap)
assert rows[-1]["kind"] == "summary" and "繰り上げ" in rows[-1]["note"], rows[-1]

# --- 中身の変更で区間が切れる（番号のみ → 変更 → 番号のみ）---
new = copy.deepcopy(dup)
new[1]["articles"][3]["paragraphs"][0]["body"] = "遅刻は事前に届け出る。"
rows = c.compare(dup, new)
assert [r["kind"] for r in rows] == ["renumbered", "renumbered", "changed", "renumbered"], rows
ch_row = rows[2]
assert ch_row["old_heading"] == "第５条（遅刻）" and ch_row["new_heading"] == "第６条（遅刻）"
assert "事前に" in underlined(ch_row["new"]) and "事前に" not in text(ch_row["old"])

# --- 項の番号だけ違う（「2」が2つ）→ 番号のみ。項の番号に下線 ---
pn = copy.deepcopy(BASE)
pn[0]["articles"][1]["paragraphs"][1]["number"] = 3
rows = c.compare(pn, BASE)
assert len(rows) == 1 and rows[0]["kind"] == "renumbered"
assert underlined(rows[0]["old"]) == "３" and underlined(rows[0]["new"]) == "２"

# --- 新設・削除 ---
added = copy.deepcopy(BASE)
added[1]["articles"].insert(2, art(0, "特別休暇", "慶弔の際に特別休暇を与える。", number_str=""))
rows = c.compare(BASE, added)
assert rows[0]["kind"] == "added" and rows[0]["new_heading"] == "第５条（特別休暇）", rows[0]
assert text(rows[0]["old"]) == "（新設）" and underlined(rows[0]["new"]).startswith("（特別休暇）第５条　慶弔")
assert all(r["kind"] in ("renumbered", "summary") for r in rows[1:])
rows = c.compare(added, BASE)
assert rows[0]["kind"] == "deleted" and text(rows[0]["new"]) == "（削除）"

# --- 号の変更：号の記号はそろえて表示 ---
it = copy.deepcopy(BASE)
it[1]["articles"][0]["paragraphs"][0]["items"][1]["body"] = "終業 17時"
rows = c.compare(BASE, it)
assert len(rows) == 1 and rows[0]["kind"] == "changed"
assert "（２）終業 17時" in text(rows[0]["new"]) and underlined(rows[0]["new"]) == "7"

# --- 番号も中身も同じ条が来たら区間を切る（まとめの範囲に変わっていない条を含めない）---
mix = [art(1, "a", "本文a"), art(2, "b", "本文b", para_nums={1: 1}),
       art(3, "c", "本文c", "二項c"), art(4, "d", "本文d"), art(5, "e", "本文e", "二項e"),
       art(6, "f", "本文f", "二項f"), art(7, "g", "本文g", "二項g")]
old_mix = copy.deepcopy(mix)
for k in (2, 4, 5, 6):                     # 第3・5・6・7条の第2項を「3」と書き間違えた
    old_mix[k]["paragraphs"][1]["number"] = 3
rows = c.compare(old_mix, mix)
assert [r["kind"] for r in rows] == ["renumbered"] * 4, [(r["kind"], r["old_heading"]) for r in rows]

# --- 区間の中の章見出しは、載せる2条に数えない ---
chs = [chapter(1, "A", art(1, "a", "本文a")), chapter(2, "B", art(2, "b", "本文b")),
       chapter(3, "C", art(3, "c", "本文c"), art(4, "d", "本文d"), art(5, "e", "本文e"))]
old_chs = copy.deepcopy(chs)
old_chs[1]["number"], old_chs[1]["number_str"] = 1, "第1章"      # 第2章を「第1章」と書き間違えた
for ch, n in ((old_chs[2], 2),):
    ch["number"], ch["number_str"] = n, f"第{n}章"
for a, n in zip([a for ch in old_chs for a in ch["articles"]][1:], range(1, 5)):
    a["number"], a["number_str"] = n, f"第{n}条"                   # 第2条以降も1つずつ小さい
rows = c.compare(old_chs, chs)
kinds = [(r["kind"], r["new_heading"]) for r in rows]
assert kinds[:4] == [("renumbered", "第２章　B"), ("renumbered", "第２条（b）"),
                     ("renumbered", "第３章　C"), ("renumbered", "第３条（c）")], kinds
assert rows[4]["kind"] == "summary" and rows[4]["new_heading"] == "第４条～第５条", kinds

# --- 下線は行ごと：項を1つ足すと、足した行だけに下線（他の行の「の」「、」は巻き込まない）---
p_add = copy.deepcopy(BASE)
p_add[1]["articles"][1]["paragraphs"].append(
    {"type": "paragraph", "number": 2, "body": "会社は、業務の都合により休日を振り替えることがある。", "items": []})
rows = c.compare(BASE, p_add)
assert len(rows) == 1 and rows[0]["kind"] == "changed"
assert underlined(rows[0]["new"]) == "２　会社は、業務の都合により休日を振り替えることがある。"
assert underlined(rows[0]["old"]) == ""

# --- 行を足し、その後ろの行も少し直した：後ろの行は直した文字だけに下線（行のずれに引きずられない）---
shift = copy.deepcopy(BASE)
paras = shift[0]["articles"][1]["paragraphs"]
paras.insert(1, {"type": "paragraph", "number": 2, "body": "嘱託社員は、別に定める。", "items": []})
paras[2]["body"] = "パートタイマーは、別に定める。"
rows = c.compare(BASE, shift)
assert len(rows) == 1 and rows[0]["kind"] == "changed"
new_lines = rows[0]["new"]
assert underlined([new_lines[2]]) == "２　嘱託社員は、別に定める。", new_lines
assert underlined([new_lines[3]]) == "３、", new_lines[3]
assert underlined([rows[0]["old"][2]]) == "２", rows[0]["old"][2]

# --- 組にする処理：前の行が後ろの似た行を先取りしない（削除＋後ろの行の一部改正）---
olds = ["短い文", "重要な規定A", "重要な規定B", "ただし、前項の規定は適用しない。"]
news = ["重要な規定A（改）", "重要な規定B（改）", "ただし、前項の規定は適用しない。短い文"]
assert c._pair_similar(olds, news, c._ratio) == [(0, None), (1, 0), (2, 1), (3, 2)]
# 1行目が2行目の改正案（似ている度合い0.93）を先に取ると、組が崩れる例
assert c._pair_similar(["賞与は支給する", "賞与は支給する。ただし業績による"],
                       ["賞与は年2回支給する", "賞与は支給する。"], c._ratio) == [(0, 0), (1, 1)]

# --- 本文を大きく書き換えても「削除＋新設」に分けず、内容変更にする ---
rw = copy.deepcopy(BASE)
rw[1]["articles"][1]["paragraphs"][0]["body"] = "まったく別の文章に書き換えた。"
rows = c.compare(BASE, rw)
assert [r["kind"] for r in rows] == ["changed"], [(r["kind"], r["old_heading"]) for r in rows]
# 2つの条を続けて書き換えても、タイトルが同じなら対応する条どうしで内容変更
rw2 = copy.deepcopy(BASE)
rw2[1]["articles"][1]["paragraphs"][0]["body"] = "一つ目を書き換えた。"
rw2[1]["articles"][2]["paragraphs"][0]["body"] = "二つ目も書き換えた。"
rows = c.compare(BASE, rw2)
assert [(r["kind"], r["new_heading"]) for r in rows] == [("changed", "第４条（休日）"), ("changed", "第５条（休暇）")], rows
# タイトルも中身も違う条に入れ替えた（休日を廃止して副業を新設）ときは、削除＋新設のまま
sw = copy.deepcopy(BASE)
sw[1]["articles"][1] = art(4, "副業", "副業は届け出る。")
rows = c.compare(BASE, sw)
assert sorted(r["kind"] for r in rows) == ["added", "deleted"], rows
# タイトルの無い条は、その位置で1対1に書き換えたら内容変更
nt_old = [art(1, "", "最初の本文である。"), art(2, "", "二つ目の本文。")]
nt_new = copy.deepcopy(nt_old)
nt_new[0]["paragraphs"][0]["body"] = "まるごと書き換えた。"
rows = c.compare(nt_old, nt_new)
assert [r["kind"] for r in rows] == ["changed"], rows

# --- 同じ見出しの行は id を分ける（備考が共有されないように）---
twin = [art(1, "x", "同じ"), art(1, "x", "同じ"), art(1, "x", "同じ")]
rows = c.compare(twin, twin)
ids = [r["id"] for r in rows]
assert len(ids) == len(set(ids)), ids

# --- 章の無い規則、表・図は置き場所だけ ---
flat = [art(1, "目的", "目的を定める。"), art(1, "定義", "用語を定める。")]
flat[1]["paragraphs"][0]["tables"] = [{"type": "table", "rows": [["区分", "日数"]]}]
rows = c.compare(flat, flat)
assert rows[0]["new_heading"] == "第２条（定義）" and "［表］　区分／日数" in text(rows[0]["new"])

# --- Word の表：3列・見出し行の繰り返し・下線・備考 ---
from docx import Document  # noqa: E402

rows = c.compare(dup, new)
with tempfile.TemporaryDirectory() as d:
    path = c.build_docx(rows, {rows[2]["id"]: "届出の時期を明確にした"},
                        os.path.join(d, "t.docx"), subtitle="sample.docx")
    doc = Document(path)
    tbl = doc.tables[0]
    assert [cell.text for cell in tbl.rows[0].cells] == ["現行", "改正案", "備考"]
    assert len(tbl.rows) == len(rows) + 1
    assert tbl.rows[0]._tr.trPr.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblHeader") is not None
    assert tbl.rows[3].cells[2].text == "届出の時期を明確にした"
    assert any(r.underline and "事前に" in r.text
               for p in tbl.rows[3].cells[1].paragraphs for r in p.runs)
    assert "以降同様" not in tbl.rows[4].cells[2].text   # 最後は1条だけなので個別の行
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    assert tbl._tbl.tblPr.find(w + "tblLayout").get(w + "type") == "fixed"
    assert [round(col.width.mm) for col in tbl.columns] == [72, 72, 26]
    # 見出しは版の表示に変えられる（画面と同じ「（第1版）」など）
    path3 = c.build_docx(rows, {}, os.path.join(d, "v.docx"), old_label="（第1版）", new_label="（第2版）")
    assert [cell.text for cell in Document(path3).tables[0].rows[0].cells] == ["（第1版）", "（第2版）", "備考"]
    path2 = c.build_docx([], {}, os.path.join(d, "e.docx"))
    d2 = Document(path2)
    assert "変更点はありません。" in [p.text for p in d2.paragraphs] and not d2.tables

# --- view()：画面用の全文の並び ---
def check_view(v, old_tree, new_tree, label=""):
    blocks, rows = v["blocks"], v["rows"]
    assert len(blocks) == len(c._align(c.units(old_tree, True), c.units(new_tree, False))), label
    for b in blocks:
        for side in ("old", "new"):
            if b[side] is None:
                assert b[side + "_cls"] is None, (label, b)
            else:
                assert len(b[side]) == len(b[side + "_cls"]), (label, b)
        if b["kind"] == "same":
            assert not any(ul for side in ("old", "new") for segs in b[side] for _, ul in segs), (label, b)
        if b["kind"] == "added":
            assert b["old"] is None and b["new"] is not None
        if b["kind"] == "deleted":
            assert b["new"] is None and b["old"] is not None
    got = [b["row_id"] for b in blocks if b["row_id"]]
    assert sorted(got) == sorted(r["id"] for r in rows), (label, got, [r["id"] for r in rows])
    assert len(got) == len(set(got)), label


v = c.view(dup, dup)
check_view(v, dup, dup, "dup")
assert v["rows"] == c.compare(dup, dup)
sm = [r for r in v["rows"] if r["kind"] == "summary"][0]
hits = [i for i, b in enumerate(v["blocks"]) if b["row_id"] == sm["id"]]
assert len(hits) == 1
b = v["blocks"][hits[0]]
assert b["row_label"] == "Word では「第５条～第６条」を1行にまとめます" and b["row_note"] == sm["note"], b
assert v["blocks"][hits[0] + 1]["row_id"] is None and v["blocks"][hits[0] + 1]["row_note"] == ""
# まとめの先頭は区間の3条目（第５条＝遅刻）：old の1行目の見出しで確かめる
assert text(b["old"]).startswith("（遅刻）\n第５条　遅刻は"), text(b["old"])
check_view(c.view(BASE, added), BASE, added, "added")
check_view(c.view(added, BASE), added, BASE, "deleted")
check_view(c.view(dup, new), dup, new, "changed")
check_view(c.view(old_chs, chs), old_chs, chs, "chapters")
check_view(c.view(twin, twin), twin, twin, "twin")
cls = c.units(BASE, False)[4]["cls"]          # 第3条：（労働時間）・第1項・号2つ
assert cls == ["title", "para", "item", "item"], cls
sub = [art(1, "x", "本文", items={1: ["号"]})]
sub[0]["paragraphs"][0]["items"][0]["sub_items"] = [
    {"body": "イ", "number": 1, "sub_items2": [{"body": "ロ", "number": 1}]}]
sub[0]["paragraphs"][0]["tables"] = [{"type": "table", "rows": [["a"]]}]
sub[0]["lead"] = ["第１節　節"]
assert c.units(sub, False)[0]["cls"] == ["lead", "title", "para", "mark", "item", "sub", "sub2"]
assert c.units([art(1, "", "本文")], False)[0]["lines"] == ["第１条　本文"]
assert c.units([art(1, "x")], False)[0]["lines"] == ["（x）", "第１条"]

# --- サンプル：落ちずに作れる。番号どおりのサンプルは変更なし ---
for p in sorted(glob.glob(os.path.join(ROOT, "tests", "expected", "*.parsed.json"))):
    with open(p, encoding="utf-8") as f:
        tree = json.load(f)
    check_view(c.view(tree, tree), tree, tree, p)
    rows = c.compare_files(p)
    assert all(r["kind"] in ("renumbered", "summary") for r in rows), (p, rows[:1])
    if "messy_kisoku" in p:
        assert rows == [], rows

# ============================================================
# 記号：現行は元の文書の記号、改正案は③のテンプレートの書き方。記号の書き方だけの違いは変更にしない
# ============================================================
import apply_style  # noqa: E402

spec_t1 = c.marks_spec_from_template(os.path.join(ROOT, "templates", "就業規則テンプレ1.docx"),
                                     apply_style.TEMPLATE_MAP["t1"])
assert {"chapter", "article", "para", "item", "sub", "sub_single", "sub2"} <= set(spec_t1), spec_t1
m1 = c.Marks(spec_t1)
assert (m1.mark("article", 5), m1.mark("para", 2), m1.mark("item", 1), m1.mark("sub", 2), m1.mark("sub2", 1)) \
    == ("第５条", "2　", "(1)　", "2.　", "(ア)　"), (m1.mark("article", 5), m1.mark("para", 2), m1.mark("item", 1))
assert c.format_number("decimalEnclosedCircle", 3) == "③" and c.format_number("japaneseCounting", 12) == "十二"

marked = [chapter(1, "総則", art(1, "提出書類", "次の書類を出す。", "前項の書類に変更があれば届け出る。",
                                 items={1: ["履歴書", "住民票"]}))]
marked[0]["articles"][0]["paragraphs"][1]["mark"] = "２．　"          # 元の文書の記号
for it, mk in zip(marked[0]["articles"][0]["paragraphs"][0]["items"], ("①　", "②　")):
    it["mark"] = mk
v = c.view(marked, marked, spec_t1)
b = [x for x in v["blocks"] if x["unit"] == "article"][0]
assert b["kind"] == "same", b["kind"]                                # 記号の書き方だけの違いは変更なし
assert text(b["old"]).split("\n")[1:4] == ["第１条　次の書類を出す。", "①　履歴書", "②　住民票"], text(b["old"])
assert text(b["new"]).split("\n")[1:4] == ["第１条　次の書類を出す。", "(1)　履歴書", "(2)　住民票"], text(b["new"])
assert "２．　前項の" in text(b["old"]) and "2　前項の" in text(b["new"])
assert v["rows"] == [] and v["marks"] == spec_t1

# 番号が変わったときは記号に下線（書き方が違っても）。本文の変更は文字だけに下線
moved = copy.deepcopy(marked)
moved[0]["articles"][0]["paragraphs"][0]["items"].insert(0, {"type": "item", "number": 0, "body": "誓約書"})
v = c.view(marked, moved, spec_t1)
row = v["rows"][0]
assert row["kind"] == "changed"
new_lines = row["new"]
assert any(segs == [["(1)　誓約書", True]] for segs in new_lines), new_lines      # 足した号は全部
assert any(segs[:1] == [["(2)　", True]] and "".join(s for s, _ in segs) == "(2)　履歴書" for segs in new_lines), new_lines

# 記号と本文は別々に比べる：本文の先頭が記号と同じ文字でも、記号に下線が化けない
cross = [chapter(1, "総則", art(1, "書類", "次の書類を出す。", items={1: ["（１）の書類を出す"]}))]
cross_new = copy.deepcopy(cross)
cross_new[0]["articles"][0]["paragraphs"][0]["items"][0]["body"] = "所定の書類を出す"
cross[0]["articles"][0]["paragraphs"][0]["items"][0]["mark"] = "①　"
row = c.view(cross, cross_new, spec_t1)["rows"][0]
old_item = [segs for segs in row["old"] if "書類を出す" in "".join(s for s, _ in segs) and "次の" not in "".join(s for s, _ in segs)][0]
assert old_item[0] == ["①　", False], old_item                   # 番号は同じなので記号に下線なし
assert underlined([old_item]) == "（１）", old_item                 # 本文の消した「（１）」だけ（「の」は新しい本文にもある）

# 現行で元の記号が無い番号なしの細分（1つだけの号の下位）は、比べる用の「＊」を出さない
single = [chapter(1, "総則", art(1, "定義", "用語は次のとおり。", items={1: ["従業員"]}))]
single[0]["articles"][0]["paragraphs"][0]["items"][0]["sub_items"] = [{"type": "item_sub", "number": 1, "body": "正社員"}]
old_lines = text([x for x in c.view(single, single)["blocks"] if x["unit"] == "article"][0]["old"]).split("\n")
assert "＊" not in "".join(old_lines) and "　正社員" in old_lines, old_lines

# 号の下位その1／その2：level があればそのとおり、無ければ件数で（2つ以上ならその1、1つならその2）
assert c.sub_levels([{}]) == ["single"] and c.sub_levels([{}, {}]) == ["multi", "multi"]
assert c.sub_levels([{"level": "single"}, {"level": "multi"}, {"level": "multi"}]) == ["single", "multi", "multi"]
subs = [chapter(1, "総則", art(1, "定義", "用語は次のとおり。", items={1: ["従業員"]}))]
it = subs[0]["articles"][0]["paragraphs"][0]["items"][0]
it["sub_items"] = [{"type": "item_sub", "number": 1, "body": "正社員", "level": "multi"},
                   {"type": "item_sub", "number": 2, "body": "注記", "level": "single"},
                   {"type": "item_sub", "number": 3, "body": "契約社員", "level": "multi",
                    "sub_items2": [{"type": "item_sub2", "number": 1, "body": "有期"}]}]
lines = text([x for x in c.view(subs, subs, spec_t1)["blocks"] if x["unit"] == "article"][0]["new"]).split("\n")
assert lines[3:7] == ["　1.　正社員", "　……　注記", "　2.　契約社員", "　　(ア)　有期"], lines

# 整形でも level のとおりの書式で出す（その1・その2・その下）。level の無い細分は今までどおり
with tempfile.TemporaryDirectory() as d:
    jp = os.path.join(d, "s.json")
    with open(jp, "w", encoding="utf-8") as f:
        json.dump(subs, f, ensure_ascii=False)
    out = os.path.join(d, "s.docx")
    import contextlib  # noqa: E402
    import io  # noqa: E402
    with contextlib.redirect_stdout(io.StringIO()):
        apply_style.apply_style(json_path=jp, template_path=os.path.join(ROOT, "templates", "就業規則テンプレ1.docx"),
                                template_key="t1", output_path=out, show_notice=False, verify=False)
    from docx import Document  # noqa: E402,F811
    doc = Document(out)
    style_of = {p.text: p.style.style_id for p in doc.paragraphs if p.text in ("正社員", "注記", "契約社員", "有期")}
    M = apply_style.TEMPLATE_MAP["t1"]
    assert style_of["正社員"] == style_of["契約社員"] == M["item_sub_multi"][0], style_of
    assert style_of["注記"] == M["item_sub_single"][0] and style_of["有期"] == M["item_sub2"][0], style_of

# --- 条を足して、本文中の参照（第6条→第7条）を直しただけの条は「番号のみ」でまとめる ---
refs_old = [art(i, f"条{i}", f"第{i + 4}条の規定による。") for i in range(1, 7)] + \
    [art(i, f"条{i}", "定める。") for i in range(7, 11)]
refs_new = [art(1, "新設", "新しい規定。")] + \
    [art(i + 1, f"条{i}", f"第{i + 5}条の規定による。") for i in range(1, 7)] + \
    [art(i + 1, f"条{i}", "定める。") for i in range(7, 11)]
kinds = [r["kind"] for r in c.compare(refs_old, refs_new)]
assert kinds == ["added", "renumbered", "renumbered", "summary"], kinds
# 参照先が付け替えと合わない（中身の変更）なら「変更」のまま
bad = copy.deepcopy(refs_new)
bad[1]["paragraphs"][0]["body"] = "第9条の規定による。"
assert c.compare(refs_old, bad)[1]["kind"] == "changed"

# --- 附則：元の附則（後付け）は変更なしで表示だけ、足した附則は「追加」の行になる ---
back = {"type": "back_matter", "indices": [], "texts": ["附　則", "この規則は令和元年から施行する。"]}
sup = {"type": "supplement", "title": "附　則（令和8年4月1日）", "body": "この規則は、令和8年4月1日から施行する。"}
v = c.view([art(1, "目的", "a"), back], [art(1, "目的", "a"), back, sup])
assert [(b["unit"], b["kind"]) for b in v["blocks"]] == [("article", "same"), ("back", "same"), ("supplement", "added")], v["blocks"]
assert v["blocks"][1]["new_path"] is None and v["blocks"][2]["new_path"] == [2]
assert [r["kind"] for r in v["rows"]] == ["added"] and v["rows"][0]["new_heading"] == sup["title"]

# 整形した docx では、足した附則が最後に出る
with tempfile.TemporaryDirectory() as d:
    jp = os.path.join(d, "sup.json")
    with open(jp, "w", encoding="utf-8") as f:
        json.dump([art(1, "目的", "a"), sup], f, ensure_ascii=False)
    out = os.path.join(d, "sup.docx")
    with contextlib.redirect_stdout(io.StringIO()):
        apply_style.apply_style(json_path=jp, template_path=os.path.join(ROOT, "templates", "就業規則テンプレ1.docx"),
                                template_key="t1", output_path=out, show_notice=False, verify=False)
    texts = [p.text for p in Document(out).paragraphs if p.text]
    assert texts[-2:] == [sup["title"], sup["body"]], texts[-4:]

print("OK")
