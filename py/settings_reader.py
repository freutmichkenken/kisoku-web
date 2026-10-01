# -*- coding: utf-8 -*-
"""
整形済みの docx から、カスタムテンプレートの設定を読み取る。

テンプレ1〜4・カスタムで整形したあと Word で手直しした docx から、
章・条・項・号…の段ごとに「Word で実際に表示される書式」を読み、
custom_template の設定（normalize() と同じ形）にする。

  ・段ごとに、その段のスタイルが付いた段落を集め、段落ごとの実際の値
    （段落の直接書式 → 番号定義 → スタイル（basedOn をたどる）→ 文書の既定）
    を求めて、いちばん多い値を採用する。スタイルの変更でも、段落を
    直接変更した場合でも、見た目どおりの値になる。
  ・読めなかった値は設定に入れない。カスタムで整形した docx は埋め込まれた
    設定（文書変数）で、テンプレ1〜4は画面側のプリセットで補う。

使い方:
  from settings_reader import read_from_docx
  r = read_from_docx("整形済.docx")   # {"settings": {...}, "notes": [...]}
"""

import re
from collections import Counter

from docx import Document
from docx.oxml.ns import qn

import custom_template as ct
from apply_style import TEMPLATE_MAP

# カスタム設定の段 → TEMPLATE_MAP のキー
LEVEL_KEYS = {
    "chapter": "chapter", "article": "article_title",
    "para1": "paragraph_1", "paraN": "paragraph_n", "item": "item",
    "sub": "item_sub_multi", "sub_single": "item_sub_single",
    "sub2": "item_sub2",
}
LAYOUT_OF = {"t1": "inline", "t2": "separate", "t3": "separate",
             "t4": "inline"}
_OFF = ("0", "false", "off")


def _val(el, attr="val"):
    return None if el is None else el.get(qn(f"w:{attr}"))


def _int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _mode(values):
    """None を除いた中でいちばん多い値（同数なら先に出たもの）。無ければ None。"""
    c = Counter(v for v in values if v is not None)
    return c.most_common(1)[0][0] if c else None


# ============================================================
# 文書の書式を引く
# ============================================================

class _Doc:
    def __init__(self, doc):
        self.doc = doc
        self.styles = {}
        self.default_pstyle = None
        for st in doc.styles.element.findall(qn("w:style")):
            sid = st.get(qn("w:styleId"))
            self.styles[sid] = st
            if st.get(qn("w:type")) == "paragraph" \
                    and st.get(qn("w:default")) in ("1", "true", "on"):
                self.default_pstyle = sid
        dd = doc.styles.element.find(qn("w:docDefaults"))
        self.def_pPr = None if dd is None else dd.find(
            qn("w:pPrDefault") + "/" + qn("w:pPr"))
        self.def_rPr = None if dd is None else dd.find(
            qn("w:rPrDefault") + "/" + qn("w:rPr"))
        try:
            self.numbering = doc.part.numbering_part.element
        except (KeyError, NotImplementedError):
            self.numbering = None
        grid = doc.element.body.find(qn("w:sectPr") + "/" + qn("w:docGrid"))
        # 「行」単位の段落間隔の1行（twips）。文書の行送りが無ければ Word の既定 18pt
        self.line_pitch = _int(_val(grid, "linePitch")) or 360

    def chain(self, sid):
        """スタイルIDから basedOn をたどった style 要素の列（近い順）。"""
        out, seen = [], set()
        while sid and sid in self.styles and sid not in seen and len(out) < 20:
            seen.add(sid)
            st = self.styles[sid]
            out.append(st)
            sid = _val(st.find(qn("w:basedOn")))
        return out

    # ---------- 番号定義 ----------

    def lvl(self, num_id, ilvl, depth=0):
        if self.numbering is None or not num_id or num_id == "0" or depth > 3:
            return None
        num = next((n for n in self.numbering.findall(qn("w:num"))
                    if n.get(qn("w:numId")) == num_id), None)
        if num is None:
            return None
        for ov in num.findall(qn("w:lvlOverride")):
            if ov.get(qn("w:ilvl")) == str(ilvl):
                lv = ov.find(qn("w:lvl"))
                if lv is not None:
                    return lv
        a_id = _val(num.find(qn("w:abstractNumId")))
        ab = next((a for a in self.numbering.findall(qn("w:abstractNum"))
                   if a.get(qn("w:abstractNumId")) == a_id), None)
        if ab is None:
            return None
        link = _val(ab.find(qn("w:numStyleLink")))
        if link:
            # 番号定義が「リストのスタイル」を指している → そのスタイルの番号へ
            for st in self.chain(link):
                nid = _val(st.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}"))
                if nid:
                    return self.lvl(nid, ilvl, depth + 1)
            return None
        return next((lv for lv in ab.findall(qn("w:lvl"))
                     if lv.get(qn("w:ilvl")) == str(ilvl)), None)

    def style_numpr(self, sid):
        """スタイル（basedOn を含む）が持つ (numId, ilvl)。"""
        num_id = ilvl = None
        for st in self.chain(sid):
            npr = st.find(qn("w:pPr") + "/" + qn("w:numPr"))
            if npr is None:
                continue
            if num_id is None:
                num_id = _val(npr.find(qn("w:numId")))
            if ilvl is None:
                ilvl = _val(npr.find(qn("w:ilvl")))
        return num_id, (_int(ilvl) or 0)


class _Para:
    """1段落の、実際に表示される書式を求める。"""

    def __init__(self, D, p, sid, fallback_num=None):
        self.D, self.p = D, p
        chain = D.chain(sid or D.default_pstyle)
        pPr = None if p is None else p.find(qn("w:pPr"))
        npr = None if pPr is None else pPr.find(qn("w:numPr"))
        s_num, s_ilvl = D.style_numpr(sid)
        if npr is not None and _val(npr.find(qn("w:numId"))) is not None:
            d_ilvl = _val(npr.find(qn("w:ilvl")))
            self.num_id = _val(npr.find(qn("w:numId")))
            self.ilvl = _int(d_ilvl) if d_ilvl is not None else s_ilvl
            direct_num = True
        elif fallback_num is not None:
            self.num_id, self.ilvl = fallback_num
            direct_num = True
        else:
            self.num_id, self.ilvl, direct_num = s_num, s_ilvl, False
        self.lvl = D.lvl(self.num_id, self.ilvl)
        lvl_pPr = None if self.lvl is None else self.lvl.find(qn("w:pPr"))
        s_pPrs = [st.find(qn("w:pPr")) for st in chain]
        # 段落に直接付けた番号は、番号定義の字下げがスタイルより優先される。
        # スタイルに付けた番号なら、スタイルの字下げが優先される（Word と同じ）
        if direct_num:
            layers = [pPr, lvl_pPr] + s_pPrs
        else:
            layers = [pPr] + s_pPrs + [lvl_pPr]
        self.pPrs = [x for x in layers + [D.def_pPr] if x is not None]
        self.s_rPrs = [x for x in [st.find(qn("w:rPr")) for st in chain]
                       + [D.def_rPr] if x is not None]

    def _attr(self, tag, attrs):
        """pPr の層を上から見て、attrs のどれかを持つ最初の要素の値（dict）。"""
        for pPr in self.pPrs:
            el = pPr.find(qn(f"w:{tag}"))
            if el is None:
                continue
            got = {a: el.get(qn(f"w:{a}")) for a in attrs}
            if any(v is not None for v in got.values()):
                return got
        return {a: None for a in attrs}

    def jc(self):
        return self._attr("jc", ("val",))["val"] or "both"

    def spacing(self, which):
        """段落の前・後の間隔（pt）。「行」単位なら文書の行送りで換算する。"""
        a = self._attr("spacing", (which, which + "Lines"))
        lines = _int(a[which + "Lines"])
        if lines:
            return lines / 100 * self.D.line_pitch / 20
        tw = _int(a[which])
        return None if tw is None else tw / 20

    def line(self):
        """固定値・最小値の行間（pt）。倍数指定などは None。"""
        a = self._attr("spacing", ("line", "lineRule"))
        tw = _int(a["line"])
        if tw is None or a["lineRule"] not in ("exact", "atLeast"):
            return None
        return tw / 20

    def indent(self, unit):
        """(1行目, 2行目以降) の開始位置（字）。unit は1字の twips。"""
        left = self._attr("ind", ("left", "start", "leftChars", "startChars"))
        first = self._attr("ind", ("firstLine", "hanging",
                                   "firstLineChars", "hangingChars"))

        def chars(c_vals, tw_vals):
            # 「字」単位（1/100字）があればそちらを使う（0 は未指定扱い：Word と同じ）
            for v in c_vals:
                if _int(v):
                    return _int(v) / 100
            for v in tw_vals:
                if _int(v) is not None:
                    return _int(v) / unit
            return None

        L = chars((left["leftChars"], left["startChars"]),
                  (left["left"], left["start"])) or 0.0
        hang = chars((first["hangingChars"],), (first["hanging"],))
        fl = chars((first["firstLineChars"],), (first["firstLine"],))
        if hang:
            return L - hang, L
        return L + (fl or 0.0), L

    # ---------- 文字 ----------

    def runs(self):
        """(文字数, rPr の層) を本文の run ごとに返す。run が無ければスタイルだけ。"""
        out = []
        if self.p is not None:
            for r in self.p.iter(qn("w:r")):
                n = sum(len(t.text or "") for t in r.findall(qn("w:t")))
                if not n:
                    continue
                layers = []
                rPr = r.find(qn("w:rPr"))
                if rPr is not None:
                    layers.append(rPr)
                    for st in self.D.chain(_val(rPr.find(qn("w:rStyle")))):
                        if st.find(qn("w:rPr")) is not None:
                            layers.append(st.find(qn("w:rPr")))
                out.append((n, layers + self.s_rPrs))
        return out or [(1, self.s_rPrs)]

    @staticmethod
    def run_attr(layers, tag, attr="val"):
        for rPr in layers:
            el = rPr.find(qn(f"w:{tag}"))
            if el is not None and (attr is None or el.get(qn(f"w:{attr}")) is not None):
                return el
        return None

    @staticmethod
    def font_name(layers):
        """和文フォント名。テーマのフォント指定は名前が分からないので None。"""
        for rPr in layers:
            f = rPr.find(qn("w:rFonts"))
            if f is None:
                continue
            if f.get(qn("w:eastAsiaTheme")) is not None:
                return None
            if f.get(qn("w:eastAsia")) is not None:
                return f.get(qn("w:eastAsia"))
        for rPr in layers:        # 和文の指定が無ければ欧文のフォント
            f = rPr.find(qn("w:rFonts"))
            if f is not None and f.get(qn("w:ascii")) is not None:
                return f.get(qn("w:ascii"))
        return None

    def run_values(self):
        """文字数の重みつきで (フォント, サイズpt, 太字) の最多値。"""
        font, size, bold = Counter(), Counter(), Counter()
        for n, layers in self.runs():
            name = self.font_name(layers)
            if name:
                font[name] += n
            sz = _int(_val(self.run_attr(layers, "sz")))
            if sz:
                size[sz / 2] += n
            b = self.run_attr(layers, "b", None)
            bold[b is not None and _val(b) not in _OFF] += n
        top = lambda c: c.most_common(1)[0][0] if c else None  # noqa: E731
        return top(font), top(size), top(bold)


# ============================================================
# どのテンプレで整形したか
# ============================================================

def _pstyle(p):
    return _val(p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}"))


def _detect_template(paras):
    """段落のスタイルの使われ方から t1〜t4 を選ぶ。見つからなければ None。"""
    used = Counter(_pstyle(p) for p in paras)
    best, best_score = None, 0
    for key, M in TEMPLATE_MAP.items():
        if key == "t3":
            continue          # スタイルは t2 と同じ。下で見分ける
        sids = {M[k][0] for k in LEVEL_KEYS.values()}
        score = sum(used[s] for s in sids)
        if score > best_score:
            best, best_score = key, score
    if best is None:
        return None
    M = TEMPLATE_MAP[best]
    if not used[M["chapter"][0]] and not used[M["article_title"][0]]:
        return None
    if best == "t2":
        # テンプレ3は項の行頭に全角スペースを入れている
        sids = {M["paragraph_1"][0], M["paragraph_n"][0]}
        texts = ["".join(t.text or "" for t in p.iter(qn("w:t")))
                 for p in paras if _pstyle(p) in sids]
        texts = [t for t in texts if t]
        if texts and sum(t.startswith("　") for t in texts) * 2 > len(texts):
            best = "t3"
    return best


def _level_paragraphs(paras, M):
    """段 → その段の段落の列。第1項と第2項以降が同じスタイルなら位置で分ける。"""
    sid_of = {k: M[LEVEL_KEYS[k]][0] for k in LEVEL_KEYS}
    out = {k: [] for k in LEVEL_KEYS}
    by_sid = {}
    for k, sid in sid_of.items():
        by_sid.setdefault(sid, []).append(k)
    prev = None
    for p in paras:
        sid = _pstyle(p)
        keys = by_sid.get(sid)
        if keys:
            if len(keys) == 1:
                out[keys[0]].append(p)
            elif set(keys) == {"para1", "paraN"}:
                out["para1" if prev == sid_of["article"] else "paraN"].append(p)
            else:
                for k in keys:
                    out[k].append(p)
        if sid is not None:
            prev = sid
    return out


# ============================================================
# 番号定義 → 設定
# ============================================================

def _numbering(lvl, symbol=False, num_id=None):
    """
    番号定義を {fmt, pre, suf, sep} にする。読めなければ (None, 理由)。
    symbol=True は番号の代わりに記号だけを出す段（号の下位その2）。
    num_id は段落の番号の ID。無い・"0"（Word で番号を「なし」にした）なら番号なし。
    """
    if lvl is None:
        if num_id in (None, "0"):
            return ({"fmt": "none", "pre": "", "suf": ""} if symbol
                    else {"fmt": "none"}), None
        return None, "番号の定義が見つかりません"
    fmt = _val(lvl.find(qn("w:numFmt"))) or "decimal"
    text = _val(lvl.find(qn("w:lvlText"))) or ""
    suff = _val(lvl.find(qn("w:suff"))) or "tab"
    out = {}
    if text.endswith("　"):
        text, out["sep"] = text[:-1], "zen"
    else:
        out["sep"] = {"tab": "tab", "space": "space"}.get(suff, "none")
    marks = re.findall(r"%(\d)", text)
    if symbol or fmt == "none":
        # 番号を出さない段。%n は表示されないので取り除いた残りが記号
        rest = re.sub(r"%\d", "", text)
        if symbol:
            if fmt != "none" and marks:
                return None, "番号ではなく記号だけの形式にしてください"
            out.update(fmt="none", pre=rest, suf="")
        else:
            if rest:
                return None, f"番号なしで「{rest}」を出す形式は設定できません"
            out["fmt"] = "none"
        return out, None
    if fmt not in ct.NUM_FORMATS:
        return None, f"番号の形式（{fmt}）は設定にありません"
    ilvl = _int(lvl.get(qn("w:ilvl")))
    if len(marks) != 1 or ilvl is None or marks[0] != str(ilvl + 1):
        return None, "上の段の番号を含む形式（例：1-1）は設定できません"
    pre, suf = text.split(f"%{marks[0]}")
    if len(pre) > 8 or len(suf) > 8:
        return None, "番号の前後の文字が長すぎます（8文字まで）"
    out.update(fmt=fmt, pre=pre, suf=suf)
    return out, None


# ============================================================
# 読み取り
# ============================================================

def _snap(v, step):
    return round(round(v / step) * step, 2)


def _put(lv, key, v, lo, hi, label, notes, step=0.5):
    if v is None:
        return
    v = _snap(v, step)
    if not lo <= v <= hi:
        notes.append(f"{label}を {lo}〜{hi} の範囲に収めました（読み取った値: {v}）")
        v = min(hi, max(lo, v))
    lv[key] = v


def _font_ok(name):
    return isinstance(name, str) and 0 < len(name.strip()) <= 40 \
        and not ct._CTRL.search(name)


def read_from_docx(docx_path):
    """
    整形済みの docx から設定を読み取る。
    戻り値 {"settings": 設定（読めた項目だけ）, "notes": 注意書きの列}。
    整形した docx と分からなければ ValueError。
    """
    try:
        doc = Document(docx_path)
    except Exception:
        raise ValueError("docx として開けませんでした")
    D = _Doc(doc)
    paras = list(doc.element.body.iter(qn("w:p")))
    notes = []

    embedded = None
    try:
        embedded = ct.read_settings(docx_path)
    except ValueError:
        pass

    if embedded is not None and "KW-1" in D.styles:
        base, layout = embedded["base"], embedded["layout"]
        M = ct.template_map(None)
        fallback = None        # KW スタイル自身が番号を持っている
        settings = embedded
    else:
        base = _detect_template(paras)
        if base is None:
            raise ValueError(
                "このアプリで整形した docx として読み取れませんでした"
                "（テンプレ1〜4・カスタムで整形した docx を選んでください）")
        layout = LAYOUT_OF[base]
        M = TEMPLATE_MAP[base]
        fallback = M
        settings = {"version": ct.VERSION, "base": base, "layout": layout,
                    "levels": {k: {} for k in ct.LEVELS}}

    if base == "t3" and embedded is None:
        notes.append("テンプレ3の項の行頭の全角スペースは設定に無いため、引き継がれません")
    groups = _level_paragraphs(paras, M)

    def para_objs(key):
        sid, num_id, ilvl = M[LEVEL_KEYS[key]]
        fb = None if fallback is None or num_id is None else (num_id, int(ilvl))
        ps = groups[key] or [None]       # 使われていない段はスタイルだけで読む
        return [_Para(D, p, sid, fb if p is None else None) for p in ps]

    objs = {k: para_objs(k) for k in ct.LEVELS}

    # --- 本文（第1項）の文字・行間を全体の値にする ---
    body = objs["para1"]
    font = _mode(o.run_values()[0] for o in body)
    size = _mode(o.run_values()[1] for o in body)
    line = _mode(o.line() for o in body)
    if _font_ok(font):
        settings["font"] = font.strip()
    elif font:
        notes.append("本文のフォント名が読み取れませんでした")
    _put(settings, "size", size, 6, 36, "本文の文字サイズ", notes)
    if line is None:
        notes.append("行間が「固定値」ではないため、読み取れませんでした")
    _put(settings, "line", line, 6, 72, "行間", notes)
    table_sid = M["table"][0]
    if table_sid in D.styles:
        _put(settings, "table_size", _Para(D, None, table_sid).run_values()[1],
             6, 36, "表の文字サイズ", notes)
    body_font = settings.get("font")
    unit = (settings.get("size") or 10.5) * 20

    for key in ct.LEVELS:
        label = ct.LEVEL_LABELS[key]
        lv = settings["levels"][key]
        os_ = objs[key]
        vals = [o.run_values() for o in os_]
        f = _mode(v[0] for v in vals)
        if _font_ok(f):
            lv["font"] = "" if f.strip() == body_font else f.strip()
        _put(lv, "size", _mode(v[1] for v in vals), 6, 36,
             f"{label}の文字サイズ", notes)
        b = _mode(v[2] for v in vals)
        if b is not None:
            lv["bold"] = b
        jc = _mode(o.jc() for o in os_)
        if jc is not None:
            lv["align"] = "center" if jc == "center" else "left"
        for which in ("before", "after"):
            _put(lv, which, _mode(o.spacing(which) for o in os_), 0, 72,
                 f"{label}の段落の{'前' if which == 'before' else '後'}の間隔",
                 notes, step=1)
        ind = _mode(o.indent(unit) for o in os_)
        if ind is not None:
            first, left = ind
            _put(lv, "first", first, 0, 30, f"{label}の1行目の開始位置", notes)
            _put(lv, "left", left, 0, 30, f"{label}の2行目以降の開始位置", notes)

        # 番号。分離型では条番号が第1項の行頭に出るので、第1項の番号定義を読む
        if layout == "separate" and key == "para1":
            continue
        src = objs["para1"] if layout == "separate" and key == "article" else os_
        got = [_numbering(o.lvl, key == "sub_single", o.num_id) for o in src]
        num = _mode(tuple(sorted(g[0].items())) for g in got if g[0])
        if num is not None:
            lv.update(dict(num))
        else:
            why = _mode(g[1] for g in got)
            if why:
                notes.append(f"{label}の番号：{why}")

    if embedded is not None:
        settings = ct.normalize(settings)
    return {"settings": settings, "notes": notes}
