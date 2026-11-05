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
- 記号の表示：現行は元の文書の記号（解析で残した mark）、改正案は③で選んだテンプレートの書き方。
  変更かどうかは、両方の列で同じ形（第５条・２・（１）・ア・（ア））にそろえた「比べる用の記号」で決めるので、
  記号の書き方（①と(1) など）だけの違いは変更にしない。
"""

import difflib
import json
import re

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


# ============================================================
# 番号の書き方（改正案は③で選んだテンプレートの書き方で表示する）
# ============================================================

_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
_IROHA = "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス"
_HALF_KANA = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ"
_HALF_IROHA = "ｲﾛﾊﾆﾎﾍﾄﾁﾘﾇﾙｦﾜｶﾖﾀﾚｿﾂﾈﾅﾗﾑｳｲﾉｵｸﾔﾏｹﾌｺｴﾃｱｻｷﾕﾒﾐｼｴﾋﾓｾｽ"
_KANJI = "〇一二三四五六七八九"


def _kanji_num(n):
    if n < 10:
        return _KANJI[n]
    if n < 100:
        t, o = divmod(n, 10)
        return ("" if t == 1 else _KANJI[t]) + "十" + (_KANJI[o] if o else "")
    return str(n)


def _roman(n):
    out = ""
    for v, s in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                 (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while n >= v:
            out, n = out + s, n - v
    return out


def format_number(fmt, n):
    """Word の numFmt の書き方で番号 n を文字にする（画面の editor.js にも同じ表がある）。"""
    seq = {"decimalEnclosedCircle": _CIRCLED, "aiueoFullWidth": _KANA, "irohaFullWidth": _IROHA,
           "aiueo": _HALF_KANA, "iroha": _HALF_IROHA}.get(fmt)
    if seq is not None:
        return seq[n - 1] if 1 <= n <= len(seq) else _zen(n)
    if fmt in ("decimalFullWidth", "decimalFullWidth2"):
        return _zen(n)
    if fmt in ("japaneseCounting", "ideographTraditional", "japaneseDigitalTenThousand"):
        return _kanji_num(n)
    if fmt == "upperLetter":
        return chr(64 + n) if 1 <= n <= 26 else str(n)
    if fmt == "lowerLetter":
        return chr(96 + n) if 1 <= n <= 26 else str(n)
    if fmt == "upperRoman":
        return _roman(n)
    if fmt == "lowerRoman":
        return _roman(n).lower()
    if fmt in ("none", "bullet"):
        return ""
    return str(n)


_SUFF = {"tab": "　", "space": " ", "nothing": ""}


class Marks:
    """
    階層ごとの番号の書き方。spec: {階層: {"fmt", "text", "ilvl", "suff"}}（marks_spec_from_template の結果）。
    階層: chapter / article / para / item / sub（号の下位その1）/ sub_single（その2）/ sub2（号の下位の下）
    spec に無い階層は、これまでの表示（第１条・２・（１）・ア・（ア））にする。
    """
    DEFAULT = {"chapter": "第{}章", "article": "第{}条", "para": "{}　", "item": "（{}）",
               "sub": "{}　", "sub_single": "", "sub2": "（{}）"}

    def __init__(self, spec=None):
        self.spec = spec or {}

    def mark(self, level, n):
        s = self.spec.get(level)
        if not s:
            d = self.DEFAULT[level]
            num = _kana(n) if level in ("sub", "sub2") else _zen(n)
            return d.format(num) if "{}" in d else d
        text = s.get("text") or ""
        own = f"%{int(s.get('ilvl', 0)) + 1}"
        out = text.replace(own, format_number(s.get("fmt"), n))
        out = re.sub(r"%[1-9]", "", out)          # 他の階層の %N は消す
        sep = "" if level in ("chapter", "article") else _SUFF.get(s.get("suff", "tab"), "　")
        return out + sep


def marks_spec_from_template(template_path, template_map):
    """テンプレート（docx）の段落番号の定義から、階層ごとの番号の書き方を読み取る。"""
    from docx import Document
    from docx.oxml.ns import qn
    import apply_style
    doc = Document(template_path)

    def level(key):
        if key not in template_map:
            return None
        _sid, num, ilvl = template_map[key]
        if not num:
            return None
        lv = apply_style._numbering_level(doc, num, ilvl)
        if lv is None:
            return None
        get = lambda tag: (lv.find(qn(tag)).get(qn("w:val")) if lv.find(qn(tag)) is not None else None)  # noqa: E731
        return {"fmt": get("w:numFmt") or "decimal", "text": get("w:lvlText") or "",
                "ilvl": int(lv.get(qn("w:ilvl")) or ilvl or 0), "suff": get("w:suff") or "tab"}

    spec = {}
    for name, keys in (("chapter", ["chapter"]), ("para", ["paragraph_n"]), ("item", ["item"]),
                       ("sub", ["item_sub_multi"]), ("sub_single", ["item_sub_single"]),
                       ("sub2", ["item_sub2"])):
        for k in keys:
            s = level(k)
            if s:
                spec[name] = s
                break
    # 条番号は、条見出しか第1項のどちらかの段落番号に「条」を含む形で付く（テンプレートによって違う）
    for k in ("article_title", "paragraph_1"):
        s = level(k)
        if s and "条" in s["text"]:
            spec["article"] = s
            break
    return spec


def sub_levels(subs):
    """
    号の下位それぞれが、号の下位その1（multi、番号付き）かその2（single、記号のみ）か。
    整形（apply_style）と同じ決め方：画面で選んだ level があればそれ、無ければ件数で
    （2つ以上ならその1、1つならその2）。
    """
    legacy = "multi" if len(subs) >= 2 else "single"
    explicit = any(su.get("level") in ("multi", "single") for su in subs)
    return [(su.get("level") if su.get("level") in ("multi", "single") else legacy) if explicit else legacy
            for su in subs]


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

def _article_lines(art, number_label, old, marks=None, disp_label=None):
    """
    条を行にする。old=True なら元の番号を使う。
    返り値: 行のリスト。各行は (クラス名, 比べる用の記号, 表示する記号, 本文)。
      比べる用の記号：これまでの表示（第５条・２・（１）・ア・（ア））。現行・改正案で同じ形なので、
        記号の書き方（①と(1) など）だけの違いを変更にしない。
      表示する記号：現行は元の文書の記号（解析で残した mark）、改正案はテンプレートの書き方（marks）。
    行の形は整形後の書き方（節の行／（タイトル）／第1項「第○条　本文」／２　本文／（１）…）。
    """
    out = []

    def add(c, cm, dm, text):
        out.append((c, cm, dm, text))

    def figs(node):
        for m in _container_marks(node):
            add("mark", "", "", m)

    def shown(node, canon, made, plain=None):
        """
        表示する記号：現行は元の記号（無ければ plain、plain も無ければ比べる用と同じ）、
        改正案はテンプレートの書き方。
        """
        if old:
            return node.get("mark") or (plain if plain is not None else canon)
        return made if made is not None else canon

    title = (art.get("title") or "").strip()
    note = (art.get("title_note") or "").strip()
    for t in art.get("lead", []):
        if t:
            add("lead", "", "", t)
    if title:
        add("title", "", "", f"（{title}）" + (f"　{note}" if note else ""))
    elif note:
        add("title", "", "", note)
    paras = art.get("paragraphs", [])
    shown_label = disp_label if disp_label is not None else number_label
    if not paras:
        add("para", number_label, shown_label, "")
    for pi, para in enumerate(paras, 1):
        body = para.get("body", "")
        if pi == 1:
            # 第1項は条番号の行（突き合わせ用に番号なしのときは本文だけ）
            if number_label:
                add("para", number_label + ("　" if body else ""), shown_label + ("　" if body else ""), body)
            else:
                add("para", "", "", body)
        else:
            pn = _num(para.get("number"), pi) if old else pi
            canon = f"{_zen(pn)}　"
            add("para", canon, shown(para, canon, marks.mark("para", pi) if marks else None), body)
        figs(para)
        for ii, it in enumerate(para.get("items", []), 1):
            n = _num(it.get("number"), ii) if old else ii
            canon = f"（{_zen(n)}）"
            add("item", canon, shown(it, canon, marks.mark("item", ii) if marks else None), it.get("body", ""))
            figs(it)
            subs = it.get("sub_items", [])
            multi_no = 0
            for si, (su, lvl) in enumerate(zip(subs, sub_levels(subs)), 1):
                if lvl == "multi":
                    multi_no += 1
                    n = _num(su.get("number"), multi_no) if old else multi_no
                    canon = f"　{_kana(n)}　"
                    made = marks.mark("sub", multi_no) if marks else None
                else:
                    canon = "　＊　"                     # 号の下位その2（番号なし。＊は比べる用で、表示しない）
                    made = marks.mark("sub_single", 1) if marks else ""
                plain = "" if lvl == "single" else None   # 現行で元の記号が無い番号なしの細分は、記号を出さない
                # 字下げは画面の CSS で付けるが、行頭の全角空白も揃えておく（比べる用と同じ形）
                add("sub", canon, "　" + shown(su, canon.lstrip("　"), made, plain).lstrip("　"), su.get("body", ""))
                figs(su)
                for s2i, s2 in enumerate(su.get("sub_items2") or [], 1):
                    n = _num(s2.get("number"), s2i) if old else s2i
                    canon = f"　　（{_kana(n)}）"
                    made = marks.mark("sub2", s2i) if marks else None
                    add("sub2", canon, "　　" + shown(s2, canon.lstrip("　"), made).lstrip("　"), s2.get("body", ""))
                    figs(s2)
    figs(art)
    return out


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


def units(tree, old, marks=None):
    """
    木を章・条の単位の列にする。
    各単位: {kind, number, label, title, sig, lines, cls}
      cls: lines と同じ長さのクラス名（chapter/lead/title/para/item/sub/sub2/mark）。
      number: 比べるための番号（int）。label: 「第５条」「第２章」。
      sig: 番号を除いた中身（突き合わせに使う）。
      lines: 比べる用の行（記号はこれまでの形）。parts: 行ごとの (比べる用の記号, 表示する記号, 本文)。
    marks: 改正案の番号の書き方（Marks）。現行（old=True）では使わない。
    """
    if old:
        marks = None
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
        shown = marks.mark("article", number) if marks else label
        rows = _article_lines(art, label, old, marks, shown)
        sig_rows = _article_lines(art, "", False)
        out.append({"kind": "article", "number": number, "label": label,
                    "title": (art.get("title") or "").strip(),
                    "sig": "\n".join(cm + tx for _, cm, _, tx in sig_rows),
                    "lines": [cm + tx for _, cm, _, tx in rows],
                    "parts": [(cm, dm, tx) for _, cm, dm, tx in rows],
                    "cls": [c for c, _, _, _ in rows],
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
            shown = marks.mark("chapter", number) if marks else label
            sep = "　" if title else ""
            out.append({"kind": "chapter", "number": number, "label": label,
                        "title": title, "sig": "章:" + title,
                        "lines": [f"{label}{sep}{title}"], "parts": [(label + sep, shown + sep, title)],
                        "cls": ["chapter"], "path": [ti]})
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


def _disp(parts):
    """表示する行（表示する記号＋本文）。"""
    return [dm + tx for _cm, dm, tx in parts]


def _marked(old_parts, new_parts):
    """
    変わった文字に下線の印を付ける。
    比べる用の行（記号はこれまでの形）で行どうしを対応させ、書き換わった行の組の中だけを文字単位で比べる
    （全文を文字単位で比べると、別の行の「の」「、」が一致して下線が虫食いになる）。
    下線は表示する行に写す：記号は番号が変わったときだけ、本文は変わった文字に引く。
    記号の書き方（①と(1) など）だけの違いには引かない。
    返り値: (old, new)。各列は行のリスト、各行は [[文字列, 下線か], ...]。
    """
    old_lines = [cm + tx for cm, _dm, tx in old_parts]
    new_lines = [cm + tx for cm, _dm, tx in new_parts]
    old, new = [], []
    sm = difflib.SequenceMatcher(None, old_lines, new_lines, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            old += _plain(_disp(old_parts[i1:i2]))
            new += _plain(_disp(new_parts[j1:j2]))
            continue
        a, b = old_lines[i1:i2], new_lines[j1:j2]
        for i, j in _pair_similar(a, b, _ratio):
            if i is not None and j is not None:
                # 記号と本文は別々に比べる（つなげて比べると、記号の「（１）」と本文の「（１）」が一致して下線がずれる）
                (ocm, _odm, otx), (ncm, _ndm, ntx) = old_parts[i1 + i], new_parts[j1 + j]
                cma, cmb = _char_flags(ocm, ncm)
                txa, txb = _char_flags(otx, ntx)
                old.append(_to_disp(old_parts[i1 + i], cma + txa))
                new.append(_to_disp(new_parts[j1 + j], cmb + txb))
            elif i is not None:
                old += _plain(_disp([old_parts[i1 + i]]), True)
            else:
                new += _plain(_disp([new_parts[j1 + j]]), True)
    return old, new


def _char_flags(a, b):
    """2つの文字列の、変わった文字の印（True）の列。"""
    ma, mb = [False] * len(a), [False] * len(b)
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op != "equal":
            ma[i1:i2] = [True] * (i2 - i1)
            mb[j1:j2] = [True] * (j2 - j1)
    return ma, mb


def _to_disp(part, flags):
    """比べる用の行の印を、表示する行（表示する記号＋本文）の下線に写す。"""
    cm, dm, tx = part
    if dm == cm:
        # 表示も比べる用と同じ記号なら、変わった文字だけに引く（「第３条」の「３」だけ、など）
        return _segments(cm + tx, flags)
    segs = [[dm, any(flags[:len(cm)])]] if dm else []
    for seg in _segments(tx, flags[len(cm):]):
        if segs and segs[-1][1] == seg[1]:
            segs[-1][0] += seg[0]
        else:
            segs.append(seg)
    return segs


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
        old, new = [[["（新設）", False]]], _plain(_disp(n["parts"]), True)
    elif kind == "deleted":
        old, new = _plain(_disp(o["parts"]), True), [[["（削除）", False]]]
    else:
        old, new = _marked(o["parts"], n["parts"])
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


def view(old_tree, new_tree, marks_spec=None):
    """
    画面用の全文の並びと、Word の表の行を返す: {"blocks": [...], "rows": [...]}。
    rows は変更のあった章・条だけ（kind: changed / renumbered / summary / added / deleted）。
    blocks は章・条の組を文書の順にすべて（変更の無いものも）並べたもの。
    """
    marks = Marks(marks_spec) if marks_spec else None
    pairs = _align(units(old_tree, True), units(new_tree, False, marks))
    blocks = []
    for o, n in pairs:
        if o is None:
            kind, old, new = "added", None, _plain(_disp(n["parts"]), True)
        elif n is None:
            kind, old, new = "deleted", _plain(_disp(o["parts"]), True), None
        elif o["sig"] == n["sig"]:
            # 比べる用の行（記号はこれまでの形）で決める。記号の書き方だけの違いは「変更なし」
            if o["lines"] == n["lines"]:
                kind, old, new = "same", _plain(_disp(o["parts"])), _plain(_disp(n["parts"]))
            else:
                kind = "renumbered"
                old, new = _marked(o["parts"], n["parts"])
        else:
            kind = "changed"
            old, new = _marked(o["parts"], n["parts"])
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
    # marks: 改正案の番号の書き方（画面の編集欄のラベルにも使う）
    return {"blocks": blocks, "rows": rows, "marks": marks_spec or {}}


def compare(old_tree, new_tree, marks_spec=None):
    """新旧対照表の行のリストを返す（変更の無い章・条は含めない）。view() の rows。"""
    return view(old_tree, new_tree, marks_spec)["rows"]


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


def compare_files(old_json_path, new_json_path=None, marks_spec=None):
    """解析済み JSON（改正案を省くと同じ木）から行を作る。"""
    return compare(*_load_trees(old_json_path, new_json_path), marks_spec)


def view_files(old_json_path, new_json_path=None, marks_spec=None):
    """解析済み JSON から画面用の {"blocks", "rows", "marks"} を作る。"""
    return view(*_load_trees(old_json_path, new_json_path), marks_spec)


def _load_trees(old_json_path, new_json_path=None):
    with open(old_json_path, encoding="utf-8") as f:
        old_tree = json.load(f)
    if new_json_path and new_json_path != old_json_path:
        with open(new_json_path, encoding="utf-8") as f:
            new_tree = json.load(f)
    else:
        new_tree = old_tree
    return old_tree, new_tree
