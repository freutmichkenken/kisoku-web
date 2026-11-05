# -*- coding: utf-8 -*-
"""
新旧対照表

現行（読み込んだ docx を解析した木）と改正案（整形で出力する木）を、章・条の単位で突き合わせ、
変更のあった章・条を「現行／改正案／備考」の3列の Word の表にする。

番号の付け方
------------
- 現行: 元の文書に書かれていた番号（解析時に読み取った number / number_str）。
- 改正案: 整形で振り直す番号（apply_style と同じく、章・条は文書全体の通し番号、
  項・号は条・項ごとに1から）。
  ※ ページ内の編集機能ができたら、改正案の木を編集後の木に差し替える。

突き合わせ
----------
- 章・条を「番号を除いた中身」（タイトルと本文）で並べて対応付ける（difflib）。
  番号がずれていても、中身が同じなら同じ条とみなす。
- 中身が同じで番号だけ違う条は「番号のみ」。続く区間は最初の2条だけ載せ、
  残りは1行（第○条～第○条）にまとめる。中身の変わった条が来たら区切る。
- 変わった文字には、現行・改正案の両方で下線を引く。
- 号の記号は元の記号（①・(1)など）を木に残していないので、両方の列で同じ形
  （号は（１）、その下はア、さらに下は（ア））にそろえて表示する。
"""

import difflib
import json

# 表示と突き合わせに使う定数
_ZEN = str.maketrans("0123456789", "０１２３４５６７８９")
_KANA = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
_SHOW_IN_RUN = 2           # 番号のみの区間で、まとめずに載せる条の数
_PAIR_RATIO = 0.5          # 置き換えの区間で、同じ条の書き換えとみなす似ている度合い


def _zen(n):
    return str(n).translate(_ZEN)


def _kana(n):
    return _KANA[n - 1] if 1 <= n <= len(_KANA) else _zen(n)


def _label(number_str):
    """元の「第5条の2」などを、数字を全角・空白なしにそろえて返す（半角・全角の違いは変更にしない）。"""
    return "".join((number_str or "").translate(_ZEN).split())


def _num(value, fallback):
    """読み取った番号。0 や欠けは、直前から数えた番号で代える。"""
    try:
        n = int(value)
    except (TypeError, ValueError):
        n = 0
    return n if n > 0 else fallback


# ============================================================
# 木 → 章・条の単位（表示用の行と、突き合わせ用の中身）
# ============================================================

def _article_lines(art, number_label, old):
    """
    条を表示用の行にする。old=True なら元の番号を使う。
    返り値: (行のリスト, 同じ長さのクラス名のリスト)。
    行の形は整形後の書き方（節の行／（タイトル）／第1項「第○条　本文」／２　本文／（１）…）。
    """
    lines, cls = [], []

    def add(text, c):
        lines.append(text)
        cls.append(c)

    def marks(node):
        for m in _container_marks(node):
            add(m, "mark")

    title = (art.get("title") or "").strip()
    note = (art.get("title_note") or "").strip()
    for t in art.get("lead", []):
        if t:
            add(t, "lead")
    if title:
        add(f"（{title}）" + (f"　{note}" if note else ""), "title")
    elif note:
        add(note, "title")
    paras = art.get("paragraphs", [])
    if not paras:
        add(number_label, "para")
    for pi, para in enumerate(paras, 1):
        body = para.get("body", "")
        if pi == 1:
            # 第1項は条番号の行（突き合わせ用に番号なしのときは本文だけ）
            add(f"{number_label}　{body}" if number_label and body else (number_label or body), "para")
        else:
            pn = _num(para.get("number"), pi) if old else pi
            add(f"{_zen(pn)}　{body}", "para")
        marks(para)
        for ii, it in enumerate(para.get("items", []), 1):
            n = _num(it.get("number"), ii) if old else ii
            add(f"（{_zen(n)}）{it.get('body', '')}", "item")
            marks(it)
            for si, su in enumerate(it.get("sub_items", []), 1):
                n = _num(su.get("number"), si) if old else si
                add(f"　{_kana(n)}　{su.get('body', '')}", "sub")
                marks(su)
                for s2i, s2 in enumerate(su.get("sub_items2") or [], 1):
                    n = _num(s2.get("number"), s2i) if old else s2i
                    add(f"　　（{_kana(n)}）{s2.get('body', '')}", "sub2")
                    marks(s2)
    marks(art)
    return lines, cls


def _container_marks(node):
    """表・図は中身を比べず、置き場所だけを示す。"""
    marks = []
    for tbl in node.get("tables", []):
        cells = [c for row in tbl.get("rows", []) for c in row if c]
        marks.append("［表］" + ("　" + "／".join(cells)[:40] if cells else ""))
    for dw in node.get("drawings", []):
        if dw.get("type") == "image":           # 画面で足した画像
            marks.append("［画像］" + ("　" + dw["name"][:40] if dw.get("name") else ""))
        else:
            marks.append("［図］" + ("　" + dw["text"][:40] if dw.get("text") else ""))
    return marks


def units(tree, old):
    """
    木を章・条の単位の列にする。
    各単位: {kind, number, label, title, sig, lines, cls}
      cls: lines と同じ長さのクラス名（chapter/lead/title/para/item/sub/sub2/mark）。
      number: 比べるための番号（int）。label: 「第５条」「第２章」。
      sig: 番号を除いた中身（突き合わせに使う）。lines: 表示用の行。
    """
    out = []
    n_ch = n_art = 0

    def add_article(art, path):
        nonlocal n_art
        n_art += 1
        if old:
            number = _num(art.get("number"), n_art)
            label = _label(art.get("number_str")) or f"第{_zen(number)}条"
            # ponytail: 番号の無い条（Wordの自動番号など）は、直前の条から数えた番号とみなす。
            # 元の自動番号を正確に出すなら、解析時に numPr から表示番号を計算して木に残す。
            n_art = number
        else:
            number = n_art
            label = f"第{_zen(number)}条"
        lines, cls = _article_lines(art, label, old)
        sig_lines, _ = _article_lines(art, "", False)
        out.append({"kind": "article", "number": number, "label": label,
                    "title": (art.get("title") or "").strip(),
                    "sig": "\n".join(sig_lines), "lines": lines, "cls": cls,
                    "path": path})

    for ti, node in enumerate(tree):
        t = node.get("type")
        if t == "chapter":
            n_ch += 1
            if old:
                number = _num(node.get("number"), n_ch)
                label = _label(node.get("number_str")) or f"第{_zen(number)}章"
                n_ch = number
            else:
                number = n_ch
                label = f"第{_zen(number)}章"
            title = (node.get("title") or "").strip()
            out.append({"kind": "chapter", "number": number, "label": label,
                        "title": title, "sig": "章:" + title,
                        "lines": [f"{label}　{title}".rstrip()], "cls": ["chapter"],
                        "path": [ti]})
            for ai, art in enumerate(node.get("articles", [])):
                add_article(art, [ti, ai])
        elif t == "article":
            add_article(node, [ti])
    return out


# ============================================================
# 突き合わせ
# ============================================================

def _ratio(a, b):
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


def _align(old_units, new_units):
    """(old or None, new or None) の組を、文書の順に返す。"""
    sm = difflib.SequenceMatcher(None, [u["sig"] for u in old_units],
                                 [u["sig"] for u in new_units], autojunk=False)
    pairs = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            pairs += [(old_units[i1 + k], new_units[j1 + k]) for k in range(i2 - i1)]
            continue
        # 置き換え・追加・削除：似ている条どうしを順番を保って組にし、残りは新設・削除
        olds, news = old_units[i1:i2], new_units[j1:j2]
        one_to_one = len(olds) == 1 and len(news) == 1
        pairs += [(None if i is None else olds[i], None if j is None else news[j])
                  for i, j in _pair_similar(olds, news, lambda o, n: _unit_score(o, n, one_to_one))]
    return pairs


def _unit_score(o, n, one_to_one):
    """
    置き換えの区間で、現行と改正案の章・条を同じものとみなす度合い。
    本文を大きく書き換えても「削除＋新設」に分かれないよう、次のときは同じ条（内容変更）とみなす。
    - 条タイトル（章なら章タイトル）が同じで空でない
    - どちらもタイトルが無く、区間に現行・改正案が1つずつしかない（その位置の条を書き換えた）
    タイトルが違う条どうしは、本文が似ていなければ「削除＋新設」のまま（廃止と新設を改正扱いにしない）。
    """
    if o["kind"] != n["kind"]:
        return 0.0
    r = _ratio(o["sig"], n["sig"])
    if (o["title"] and o["title"] == n["title"]) or (one_to_one and not o["title"] and not n["title"]):
        r = max(r, _PAIR_RATIO)
    return r


def _pair_similar(olds, news, score):
    """
    似ているものどうしを、順番を保って組にする。
    score(o, n) が _PAIR_RATIO 以上の組だけを候補にし、組の似ている度合いの合計が
    最大になるように選ぶ（動的計画法。前の行が後ろの似た行を先取りしてずれるのを防ぐ）。
    返り値: (i or None, j or None) のリスト（文書の順）。
    """
    # ponytail: 計算量は len(olds)×len(news)。置き換えの区間（条の中の行、連続する条）だけに使うので小さい。
    # 何千条もの区間を一度に比べて遅くなったら、区間を章ごとに分ける。
    n_o, n_n = len(olds), len(news)
    sc = [[score(o, n) for n in news] for o in olds]
    best = [[0.0] * (n_n + 1) for _ in range(n_o + 1)]
    for i in range(n_o - 1, -1, -1):
        for j in range(n_n - 1, -1, -1):
            v = max(best[i + 1][j], best[i][j + 1])
            if sc[i][j] >= _PAIR_RATIO:
                v = max(v, sc[i][j] + best[i + 1][j + 1])
            best[i][j] = v
    out, i, j = [], 0, 0
    while i < n_o and j < n_n:
        if sc[i][j] >= _PAIR_RATIO and best[i][j] == sc[i][j] + best[i + 1][j + 1]:
            out.append((i, j))
            i, j = i + 1, j + 1
        elif best[i][j] == best[i + 1][j]:
            out.append((i, None))
            i += 1
        else:
            out.append((None, j))
            j += 1
    out += [(k, None) for k in range(i, n_o)] + [(None, k) for k in range(j, n_n)]
    return out


def _marked(old_lines, new_lines):
    """
    変わった文字に下線の印を付ける。
    先に行どうしを対応させ、書き換わった行の組の中だけを文字単位で比べる
    （全文を文字単位で比べると、別の行の「の」「、」が一致して下線が虫食いになる）。
    返り値: (old, new)。各列は行のリスト、各行は [[文字列, 下線か], ...]。
    """
    old, new = [], []
    sm = difflib.SequenceMatcher(None, old_lines, new_lines, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            old += _plain(old_lines[i1:i2])
            new += _plain(new_lines[j1:j2])
            continue
        a, b = old_lines[i1:i2], new_lines[j1:j2]
        for i, j in _pair_similar(a, b, _ratio):
            if i is not None and j is not None:
                oa, nb = _char_marked(a[i], b[j])
                old.append(oa)
                new.append(nb)
            elif i is not None:
                old += _plain([a[i]], True)
            else:
                new += _plain([b[j]], True)
    return old, new


def _char_marked(a, b):
    ma, mb = [False] * len(a), [False] * len(b)
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op != "equal":
            ma[i1:i2] = [True] * (i2 - i1)
            mb[j1:j2] = [True] * (j2 - j1)
    return _segments(a, ma), _segments(b, mb)


def _segments(line, marks):
    cur = []
    for ch, ul in zip(line, marks):
        if cur and cur[-1][1] == ul:
            cur[-1][0] += ch
        else:
            cur.append([ch, ul])
    return cur


def _plain(lines, underline=False):
    return [[[ln, underline]] if ln else [] for ln in lines]


def _heading(u):
    if u is None:
        return ""
    if u["kind"] == "chapter":
        return f"{u['label']}　{u['title']}".rstrip()
    return u["label"] + (f"（{u['title']}）" if u["title"] else "")


def _row(kind, o, n, note=""):
    if kind == "added":
        old, new = [[["（新設）", False]]], _plain(n["lines"], True)
    elif kind == "deleted":
        old, new = _plain(o["lines"], True), [[["（削除）", False]]]
    else:
        old, new = _marked(o["lines"], n["lines"])
    # 備考の対応付けに使う id。編集で改正案の見出しや変化の種類が変わっても外れないよう、
    # 現行の見出しがあればそれだけで作る（新設は改正案の見出し）
    rid = f"o:{_heading(o)}" if o is not None else f"n:{_heading(n)}"
    return {"id": rid, "kind": kind,
            "old_heading": _heading(o), "new_heading": _heading(n),
            "old": old, "new": new, "note": note}


def _summary_row(run):
    """番号のみの区間の残りを1行にまとめる。"""
    arts = [(o, n) for o, n in run if o["kind"] == "article"]
    (o1, n1), (o2, n2) = arts[0], arts[-1]

    def rng(a, b):
        return a["label"] if a is b else f"{a['label']}～{b['label']}"

    diffs = [n["number"] - o["number"] for o, n in arts]
    if any(o["label"] != n["label"] for o, n in arts):
        what = ("繰り下げ" if all(d > 0 for d in diffs)
                else "繰り上げ" if all(d < 0 for d in diffs) else "変更")
        note = f"条番号の{what}のみ（内容の変更なし、以降同様）"
    else:
        note = "項・号の番号の変更のみ（内容の変更なし、以降同様）"
    if any(o["kind"] == "chapter" for o, _ in run):
        note += "。章番号の変更を含む"
    old_r, new_r = rng(o1, o2), rng(n1, n2)
    return {"id": f"summary:{old_r}", "kind": "summary",
            "old_heading": old_r, "new_heading": new_r,
            "old": [[[old_r, False]]], "new": [[[new_r, False]]], "note": note}


def view(old_tree, new_tree):
    """
    画面用の全文の並びと、Word の表の行を返す: {"blocks": [...], "rows": [...]}。
    rows は変更のあった章・条だけ（kind: changed / renumbered / summary / added / deleted）。
    blocks は章・条の組を文書の順にすべて（変更の無いものも）並べたもの。
    """
    pairs = _align(units(old_tree, True), units(new_tree, False))
    blocks = []
    for o, n in pairs:
        if o is None:
            kind, old, new = "added", None, _plain(n["lines"], True)
        elif n is None:
            kind, old, new = "deleted", _plain(o["lines"], True), None
        elif o["sig"] == n["sig"]:
            if o["lines"] == n["lines"]:
                kind, old, new = "same", _plain(o["lines"]), _plain(n["lines"])
            else:
                kind = "renumbered"
                old, new = _marked(o["lines"], n["lines"])
        else:
            kind = "changed"
            old, new = _marked(o["lines"], n["lines"])
        blocks.append({"kind": kind, "unit": (o or n)["kind"], "old": old, "new": new,
                       "old_cls": o["cls"] if o else None, "new_cls": n["cls"] if n else None,
                       # 木の中の位置（[章の位置, 条の位置] か [位置]）。画面で改正案を編集するときに使う
                       "old_path": o["path"] if o else None, "new_path": n["path"] if n else None,
                       "row_id": None, "row_note": "", "row_label": ""})

    rows, run = [], []
    owner = []                    # (行の位置, block の位置, summary か)

    def flush():
        # 条を _SHOW_IN_RUN 個載せるまでは、間の章も載せる。残りをまとめる。
        shown = 0
        cut = len(run)
        for k, (_, o, _n) in enumerate(run):
            if shown == _SHOW_IN_RUN:
                cut = k
                break
            if o["kind"] == "article":
                shown += 1
        for bi, o, n in run[:cut]:
            owner.append((len(rows), bi, False))
            rows.append(_row("renumbered", o, n))
        rest = run[cut:]
        n_art = sum(1 for _, o, _n in rest if o["kind"] == "article")
        if n_art >= 2:
            first = next(bi for bi, o, _n in rest if o["kind"] == "article")
            owner.append((len(rows), first, True))
            rows.append(_summary_row([(o, n) for _, o, n in rest]))
        else:
            for bi, o, n in rest:
                owner.append((len(rows), bi, False))
                rows.append(_row("renumbered", o, n))
        run.clear()

    for bi, (o, n) in enumerate(pairs):
        kind = blocks[bi]["kind"]
        if kind == "same":
            if o["kind"] == "article":
                flush()                  # 番号も中身も同じ条で区間を切る（章見出しでは切らない）
            continue
        if kind == "renumbered":
            run.append((bi, o, n))       # 番号のみ
            continue
        flush()
        owner.append((len(rows), bi, False))
        rows.append(_row(kind, o, n))
    flush()
    # 同じ見出しの行（条番号の重複など）で備考が共有されないよう、2つ目以降の id に番号を付ける
    seen = {}
    for r in rows:
        k = seen.get(r["id"], 0) + 1
        seen[r["id"]] = k
        if k > 1:
            r["id"] += f"#{k}"
    for ri, bi, is_summary in owner:
        b, r = blocks[bi], rows[ri]
        b["row_id"] = r["id"]
        if is_summary:
            b["row_note"] = r["note"]
            b["row_label"] = f"Word では「{r['old_heading']}」を1行にまとめます"
    return {"blocks": blocks, "rows": rows}


def compare(old_tree, new_tree):
    """新旧対照表の行のリストを返す（変更の無い章・条は含めない）。view() の rows。"""
    return view(old_tree, new_tree)["rows"]


# ============================================================
# Word の表にする
# ============================================================

def build_docx(rows, notes, out_path, title="新旧対照表", subtitle=""):
    """
    rows: compare() の結果。notes: {行の id: 備考の文字列}（画面で入力したもの）。
    まとめ行の備考は、自動の説明のあとに入力した備考を続ける。
    """
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt

    notes = notes or {}
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Mm(20))

    normal = doc.styles["Normal"]
    normal.font.size = Pt(9.5)
    normal.font.name = "ＭＳ 明朝"
    rpr = normal.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    fonts.set(qn("w:eastAsia"), "ＭＳ 明朝")
    normal.paragraph_format.space_after = Pt(0)

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(title)
    r.bold = True
    r.font.size = Pt(14)
    if subtitle:
        s = doc.add_paragraph(subtitle)
        s.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    if not rows:
        doc.add_paragraph("変更点はありません。")
        doc.save(out_path)
        return out_path

    widths = (Mm(72), Mm(72), Mm(26))
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False                      # Word の自動調整で列幅が崩れないように
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    head = table.rows[0]
    trpr = head._tr.get_or_add_trPr()
    rep = OxmlElement("w:tblHeader")           # ページをまたいでも見出し行を繰り返す
    rep.set(qn("w:val"), "true")
    trpr.append(rep)
    for cell, text in zip(head.cells, ("現行", "改正案", "備考")):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(text).bold = True

    def fill(cell, lines):
        first = True
        for segs in lines or [[]]:
            p = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            for text, ul in segs:
                run = p.add_run(text)
                if ul:
                    run.underline = True

    for row in rows:
        cells = table.add_row().cells
        fill(cells[0], row["old"])
        fill(cells[1], row["new"])
        note = (notes.get(row["id"]) or "").strip()
        auto = row.get("note", "")
        text = "\n".join(t for t in (auto, note) if t)
        fill(cells[2], [[[t, False]] for t in text.split("\n")] if text else [])

    for col, w in zip(table.columns, widths):
        col.width = w
    for row in table.rows:
        for cell, w in zip(row.cells, widths):
            cell.width = w

    doc.save(out_path)
    return out_path


def compare_files(old_json_path, new_json_path=None):
    """解析済み JSON（改正案を省くと同じ木）から行を作る。"""
    return compare(*_load_trees(old_json_path, new_json_path))


def view_files(old_json_path, new_json_path=None):
    """解析済み JSON から画面用の {"blocks", "rows"} を作る。"""
    return view(*_load_trees(old_json_path, new_json_path))


def _load_trees(old_json_path, new_json_path=None):
    with open(old_json_path, encoding="utf-8") as f:
        old_tree = json.load(f)
    if new_json_path and new_json_path != old_json_path:
        with open(new_json_path, encoding="utf-8") as f:
            new_tree = json.load(f)
    else:
        new_tree = old_tree
    return old_tree, new_tree
