# -*- coding: utf-8 -*-
"""
カスタムテンプレート（画面で入力した書式）を作る。

既存テンプレ（t1〜t4）の docx を土台にし、そこへ
  ・章〜(ア) の番号定義（numbering.xml に新しい abstractNum / num）
  ・段ごとの段落スタイル（styles.xml に KW-1〜KW-9）
を書き足して、apply_style が使えるテンプレート docx と対応表を作る。
土台から引き継ぐのは用紙・余白・ページ番号など。土台の既存スタイルと
番号定義は変えない（目次スタイル toc 2 の番号だけ、条の組み方に合わせる）。

設定値は文書変数（settings.xml の w:docVars）に JSON で埋め込む。
Word で開いて保存しても残り、read_settings() で読み戻せる。

使い方:
  from custom_template import build
  M = build(settings, {"t1": "就業規則テンプレ1.docx", ...}, "custom.docx")
  apply_style(..., template_path="custom.docx", template_key="custom",
              template_map=M)
"""

import json
import math
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DOCVAR_NAME = "kisokuWebCustomTemplate"
VERSION = 1

BASES = ("t1", "t2", "t3", "t4")
LAYOUTS = ("inline", "separate")      # 「第1条（目的）」を1行 ／ 見出しと条番号を分ける
NUM_FORMATS = ("decimalFullWidth", "decimal", "japaneseCounting",
               "decimalEnclosedCircle", "aiueoFullWidth", "irohaFullWidth",
               "none")
SEPS = ("zen", "space", "tab", "none")  # 番号と本文の間
ALIGNS = ("left", "center")
LEVELS = ("chapter", "article", "para1", "paraN", "item", "sub", "sub2")
LEVEL_LABELS = {"chapter": "章", "article": "条", "para1": "第1項",
                "paraN": "第2項以降", "item": "号", "sub": "号の下位",
                "sub2": "号の下位の下"}

# 段 → (スタイルID, スタイル名, ilvl)。名前にカンマを入れない（目次の \t で使う）
STYLE_DEFS = {
    "chapter":    ("KW-1", "KW-1 章", 1),
    "article":    ("KW-2", "KW-2 条見出し", 2),
    "para1":      ("KW-3", "KW-3 第1項", 3),
    "paraN":      ("KW-4", "KW-4 第2項以降", 4),
    "item":       ("KW-5", "KW-5 号", 5),
    "sub":        ("KW-6", "KW-6 号の下位", 6),
    "sub_single": ("KW-7", "KW-7 号の下位（単独）", 7),
    "sub2":       ("KW-8", "KW-8 号の下位の下", 8),
}
TABLE_STYLE = ("KW-9", "KW-9 表")

_CTRL = re.compile(r"[\x00-\x1f\x7f]")


# ============================================================
# 設定値の検証（画面・読み込んだ docx のどちらから来ても通す）
# ============================================================

def _number(v, lo, hi, name):
    if isinstance(v, bool) or not isinstance(v, (int, float)) \
            or not math.isfinite(v):
        raise ValueError(f"{name} は数値で指定してください")
    if not lo <= v <= hi:
        raise ValueError(f"{name} は {lo}〜{hi} の範囲で指定してください（指定値: {v}）")
    return round(float(v), 2)


def _text(v, max_len, name, allow_empty=True):
    if not isinstance(v, str):
        raise ValueError(f"{name} は文字で指定してください")
    v = v.strip()
    if not v and not allow_empty:
        raise ValueError(f"{name} を入力してください")
    if len(v) > max_len:
        raise ValueError(f"{name} は {max_len} 文字以内にしてください")
    if _CTRL.search(v):
        raise ValueError(f"{name} に使えない文字が含まれています")
    return v


def _choice(v, choices, name):
    if v not in choices:
        raise ValueError(f"{name} の指定が正しくありません（指定値: {v}）")
    return v


def _bool(v, name):
    if not isinstance(v, bool):
        raise ValueError(f"{name} は true / false で指定してください")
    return v


def _get(d, key, where):
    if not isinstance(d, dict) or key not in d:
        raise ValueError(f"設定に {where}{key} がありません")
    return d[key]


def normalize(settings):
    """設定値を検証し、範囲内に整えた新しい dict を返す。不正なら ValueError。"""
    if not isinstance(settings, dict):
        raise ValueError("設定の形式が正しくありません")
    version = _get(settings, "version", "")
    if isinstance(version, bool) or version != VERSION:
        raise ValueError("この設定は対応していない版のものです")
    s = {
        "version": VERSION,
        "base": _choice(_get(settings, "base", ""), BASES, "土台のテンプレ"),
        "layout": _choice(_get(settings, "layout", ""), LAYOUTS, "条の組み方"),
        "font": _text(_get(settings, "font", ""), 40, "本文のフォント",
                      allow_empty=False),
        "size": _number(_get(settings, "size", ""), 6, 36, "本文の文字サイズ"),
        "line": _number(_get(settings, "line", ""), 6, 72, "行間"),
        "table_size": _number(_get(settings, "table_size", ""), 6, 36,
                              "表の文字サイズ"),
        "levels": {},
    }
    levels = _get(settings, "levels", "")
    for key in LEVELS:
        lv = _get(levels, key, "levels.")
        label = LEVEL_LABELS[key]
        pre = _text(_get(lv, "pre", f"{key}."), 8, f"{label}の前に付ける文字")
        suf = _text(_get(lv, "suf", f"{key}."), 8, f"{label}の後に付ける文字")
        if "%" in pre + suf:
            # lvlText の %1〜%9 と衝突するため
            raise ValueError(f"{label}の番号の前後に「%」は使えません")
        s["levels"][key] = {
            "fmt": _choice(_get(lv, "fmt", f"{key}."), NUM_FORMATS,
                           f"{label}の番号の形式"),
            "pre": pre,
            "suf": suf,
            "sep": _choice(_get(lv, "sep", f"{key}."), SEPS,
                           f"{label}の番号と本文の間"),
            "first": _number(_get(lv, "first", f"{key}."), 0, 30,
                             f"{label}の1行目の開始位置"),
            "left": _number(_get(lv, "left", f"{key}."), 0, 30,
                            f"{label}の2行目以降の開始位置"),
            "size": _number(_get(lv, "size", f"{key}."), 6, 36,
                            f"{label}の文字サイズ"),
            "bold": _bool(_get(lv, "bold", f"{key}."), f"{label}の太字"),
            "font": _text(_get(lv, "font", f"{key}."), 40, f"{label}のフォント"),
            "align": _choice(_get(lv, "align", f"{key}."), ALIGNS,
                             f"{label}の配置"),
            "before": _number(_get(lv, "before", f"{key}."), 0, 72,
                              f"{label}の段落の前の間隔"),
            "after": _number(_get(lv, "after", f"{key}."), 0, 72,
                             f"{label}の段落の後の間隔"),
        }
    return s


def line_height(s, lv):
    """
    行の高さ（pt・固定値）。文字が行より大きいと上下が欠けるので、
    文字サイズの1.25倍を下限にする。画面のプレビューも同じ式で計算する。
    """
    return max(s["line"], lv["size"] * 1.25)


# ============================================================
# 番号定義
# ============================================================

def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), str(v))
    return e


def _twips_per_char(s):
    # 字下げの「1字」は本文の文字サイズで換算する（1pt = 20twips）
    return s["size"] * 20


def _ind(first, left, u):
    """1行目・2行目以降の開始位置（字）から w:ind を作る。"""
    L, F = round(left * u), round(first * u)
    ind = _el("w:ind", left=L)
    if F < L:
        ind.set(qn("w:hanging"), str(L - F))
    else:
        ind.set(qn("w:firstLine"), str(F - L))
    return ind


def _lvl_text(lv, ilvl):
    """(lvlText, suff)。番号なしなら前後の文字も出さない。"""
    if lv["fmt"] == "none":
        return "", "nothing"
    text = f"{lv['pre']}%{ilvl + 1}{lv['suf']}"
    if lv["sep"] == "zen":
        return text + "　", "nothing"
    return text, {"space": "space", "tab": "tab", "none": "nothing"}[lv["sep"]]


def _make_lvl(ilvl, fmt, text, suff, ind_lv, u, start=1, restart=None):
    lvl = _el("w:lvl", ilvl=ilvl)
    lvl.append(_el("w:start", val=start))
    lvl.append(_el("w:numFmt", val=fmt))
    if restart is not None:
        # 0 = 振り直さない（1以上は「何段目が出たら振り直すか」の1始まり番号）
        lvl.append(_el("w:lvlRestart", val=restart))
    lvl.append(_el("w:suff", val=suff))
    lvl.append(_el("w:lvlText", val=text))
    lvl.append(_el("w:lvlJc", val="left"))
    pPr = OxmlElement("w:pPr")
    if suff == "tab" and ind_lv is not None:
        # タブは2行目以降の開始位置まで進める
        tabs = OxmlElement("w:tabs")
        tabs.append(_el("w:tab", val="left", pos=round(ind_lv["left"] * u)))
        pPr.append(tabs)
    if ind_lv is not None:
        pPr.append(_ind(ind_lv["first"], ind_lv["left"], u))
    else:
        pPr.append(_el("w:ind", left=0, firstLine=0))
    lvl.append(pPr)
    rPr = OxmlElement("w:rPr")
    rPr.append(_el("w:rFonts", hint="eastAsia"))
    lvl.append(rPr)
    return lvl


def _next_ids(nel):
    a_ids = [int(a.get(qn("w:abstractNumId")))
             for a in nel.findall(qn("w:abstractNum"))]
    n_ids = [int(n.get(qn("w:numId"))) for n in nel.findall(qn("w:num"))]
    return (max(a_ids, default=-1) + 1, max(n_ids, default=0) + 1)


def _insert_numbering(nel, abstract, nsid):
    """abstractNum と、それを指す num を追加し、numId を返す。"""
    a_id, n_id = _next_ids(nel)
    abstract.set(qn("w:abstractNumId"), str(a_id))
    abstract.insert(0, _el("w:nsid", val=nsid))
    # スキーマ上、abstractNum はすべての num より前に置く
    nums = nel.findall(qn("w:num"))
    if nums:
        nums[0].addprevious(abstract)
    else:
        nel.append(abstract)
    num = _el("w:num", numId=n_id)
    num.append(_el("w:abstractNumId", val=a_id))
    nums = nel.findall(qn("w:num"))
    if nums:
        nums[-1].addnext(num)
    else:
        abstract.addnext(num)
    return str(n_id)


def _add_numbering(doc, s):
    """章〜(ア) の9段の番号定義を追加し、numId を返す。"""
    L = s["levels"]
    u = _twips_per_char(s)
    sep = s["layout"] == "separate"
    ab = OxmlElement("w:abstractNum")
    ab.append(_el("w:multiLevelType", val="multilevel"))

    ab.append(_make_lvl(0, "none", "", "nothing", None, u))
    t, sf = _lvl_text(L["chapter"], 1)
    ab.append(_make_lvl(1, L["chapter"]["fmt"], t, sf, L["chapter"], u,
                        restart=0))
    # 条の番号は章をまたいで通し番号にする（lvlRestart=0）
    if sep:
        # 見出し行には番号を出さず、第1項の行頭に条番号を出す
        ab.append(_make_lvl(2, "none", "", "nothing", L["article"], u,
                            restart=0))
        t, sf = _lvl_text(L["article"], 3)
        ab.append(_make_lvl(3, L["article"]["fmt"], t, sf, L["para1"], u,
                            restart=0))
    else:
        t, sf = _lvl_text(L["article"], 2)
        ab.append(_make_lvl(2, L["article"]["fmt"], t, sf, L["article"], u,
                            restart=0))
        t, sf = _lvl_text(L["para1"], 3)
        ab.append(_make_lvl(3, L["para1"]["fmt"], t, sf, L["para1"], u))
    # 第2項以降は 2 から始め、条見出し（3段目＝ilvl 2）が出るたびに振り直す。
    # テンプレ2と同じ指定にそろえている
    t, sf = _lvl_text(L["paraN"], 4)
    ab.append(_make_lvl(4, L["paraN"]["fmt"], t, sf, L["paraN"], u,
                        start=2, restart=3))
    t, sf = _lvl_text(L["item"], 5)
    ab.append(_make_lvl(5, L["item"]["fmt"], t, sf, L["item"], u))
    t, sf = _lvl_text(L["sub"], 6)
    ab.append(_make_lvl(6, L["sub"]["fmt"], t, sf, L["sub"], u))
    # 号の下位が1つだけのときは番号を付けない（字下げは号の下位と同じ）
    ab.append(_make_lvl(7, "none", "", "nothing", L["sub"], u))
    t, sf = _lvl_text(L["sub2"], 8)
    ab.append(_make_lvl(8, L["sub2"]["fmt"], t, sf, L["sub2"], u))
    return _insert_numbering(doc.part.numbering_part.element, ab, "4B570001")


def _add_toc_numbering(doc, s):
    """
    条見出しと条番号を分ける組み方では、目次に載る見出し行に番号が無い。
    テンプレ2と同じく、目次スタイル toc 2 自身に「第%1条」の連番を持たせる。
    """
    lv = s["levels"]["article"]
    t, sf = _lvl_text(lv, 0)
    ab = OxmlElement("w:abstractNum")
    ab.append(_el("w:multiLevelType", val="singleLevel"))
    lvl = _el("w:lvl", ilvl=0)
    lvl.append(_el("w:start", val=1))
    lvl.append(_el("w:numFmt", val=lv["fmt"]))
    lvl.append(_el("w:suff", val=sf))
    lvl.append(_el("w:lvlText", val=t))
    lvl.append(_el("w:lvlJc", val="left"))
    ab.append(lvl)
    return _insert_numbering(doc.part.numbering_part.element, ab, "4B570002")


# ============================================================
# スタイル
# ============================================================

def _find_style(styles_el, style_id=None, name=None):
    for st in styles_el.findall(qn("w:style")):
        if style_id is not None and st.get(qn("w:styleId")) == style_id:
            return st
        nm = st.find(qn("w:name"))
        if name is not None and nm is not None and nm.get(qn("w:val")) == name:
            return st
    return None


def _rpr(font, size, bold):
    rPr = OxmlElement("w:rPr")
    rPr.append(_el("w:rFonts", ascii=font, eastAsia=font, hAnsi=font, cs=font))
    if bold:
        rPr.append(OxmlElement("w:b"))
        rPr.append(OxmlElement("w:bCs"))
    half = round(size * 2)          # w:sz は半ポイント単位
    rPr.append(_el("w:sz", val=half))
    rPr.append(_el("w:szCs", val=half))
    return rPr


def _make_style(style_id, name, lv, s, num_id, ilvl, outline=None,
                keep_next=False):
    u = _twips_per_char(s)
    st = _el("w:style", type="paragraph", customStyle=1, styleId=style_id)
    st.append(_el("w:name", val=name))
    st.append(OxmlElement("w:qFormat"))
    pPr = OxmlElement("w:pPr")       # 子要素の順番はスキーマどおりにする
    if keep_next:
        pPr.append(OxmlElement("w:keepNext"))
    numPr = OxmlElement("w:numPr")
    numPr.append(_el("w:ilvl", val=ilvl))
    numPr.append(_el("w:numId", val=num_id))
    pPr.append(numPr)
    # 文字の格子に吸着させない（プレビューと字の並びを合わせるため）
    pPr.append(_el("w:snapToGrid", val=0))
    pPr.append(_el("w:spacing", before=round(lv["before"] * 20),
                   after=round(lv["after"] * 20),
                   line=round(line_height(s, lv) * 20), lineRule="exact"))
    pPr.append(_ind(lv["first"], lv["left"], u))
    pPr.append(_el("w:jc", val="center" if lv["align"] == "center" else "both"))
    if outline is not None:
        pPr.append(_el("w:outlineLvl", val=outline))
    st.append(pPr)
    st.append(_rpr(lv["font"] or s["font"], lv["size"], lv["bold"]))
    return st


def _add_styles(doc, s, num_id):
    styles_el = doc.styles.element
    for sid, _name, _ilvl in list(STYLE_DEFS.values()) + [TABLE_STYLE + (None,)]:
        if _find_style(styles_el, style_id=sid) is not None:
            raise ValueError(f"土台のテンプレートに同じスタイルID {sid} があります")
    L = s["levels"]
    for key, (sid, name, ilvl) in STYLE_DEFS.items():
        lv = L["sub"] if key == "sub_single" else L[key]
        styles_el.append(_make_style(
            sid, name, lv, s, num_id, ilvl,
            outline={"chapter": 0, "article": 1}.get(key),
            keep_next=key in ("chapter", "article")))
    # 表の中の文字：本文のフォント・表の文字サイズ・番号なし
    st = _el("w:style", type="paragraph", customStyle=1, styleId=TABLE_STYLE[0])
    st.append(_el("w:name", val=TABLE_STYLE[1]))
    st.append(OxmlElement("w:qFormat"))
    pPr = OxmlElement("w:pPr")
    pPr.append(_el("w:snapToGrid", val=0))
    pPr.append(_el("w:jc", val="both"))
    st.append(pPr)
    st.append(_rpr(s["font"], s["table_size"], False))
    styles_el.append(st)


def _set_toc2_numbering(doc, toc_num_id):
    """目次スタイル toc 2 の番号を、条の組み方に合わせて付け外しする。"""
    st = _find_style(doc.styles.element, name="toc 2")
    if st is None:
        return
    pPr = st.find(qn("w:pPr"))
    if pPr is not None:
        for old in pPr.findall(qn("w:numPr")):
            pPr.remove(old)
    if toc_num_id is None:
        return
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        rPr = st.find(qn("w:rPr"))
        if rPr is not None:
            rPr.addprevious(pPr)
        else:
            st.append(pPr)
    numPr = OxmlElement("w:numPr")
    numPr.append(_el("w:ilvl", val=0))
    numPr.append(_el("w:numId", val=toc_num_id))
    # numPr は keepNext〜widowControl の後、tabs などの前
    before = [qn(f"w:{t}") for t in ("pStyle", "keepNext", "keepLines",
                                       "pageBreakBefore", "framePr",
                                       "widowControl")]
    idx = 0
    for i, child in enumerate(pPr):
        if child.tag in before:
            idx = i + 1
    pPr.insert(idx, numPr)


# ============================================================
# 設定の埋め込み・読み出し（文書変数）
# ============================================================

# settings.xml で docVars より後ろに来る要素（スキーマの順番）
_AFTER_DOCVARS = {"rsids", "mathPr", "attachedSchema", "themeFontLang",
                  "clrSchemeMapping", "doNotIncludeSubdocsInStats",
                  "doNotAutoCompressPictures", "forceUpgrade", "captions",
                  "readModeInkLockDown", "smartTagType", "schemaLibrary",
                  "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol",
                  "listSeparator"}
_W_NS = qn("w:x")[:-1]


def write_settings(doc, s):
    """設定を文書変数に書き込む（同名があれば置き換える）。"""
    sel = doc.settings.element
    dvs = sel.find(qn("w:docVars"))
    if dvs is None:
        dvs = OxmlElement("w:docVars")
        # docVars の後ろに来る要素の直前に入れる。どれも無ければ、末尾に
        # まとまっている拡張要素（w14:docId など）の直前に入れる
        children = list(sel)
        anchor = next((c for c in children
                       if c.tag.split("}")[-1] in _AFTER_DOCVARS), None)
        if anchor is None:
            for c in reversed(children):
                if c.tag.startswith(_W_NS):
                    break
                anchor = c
        if anchor is not None:
            anchor.addprevious(dvs)
        else:
            sel.append(dvs)
    for dv in dvs.findall(qn("w:docVar")):
        if dv.get(qn("w:name")) == DOCVAR_NAME:
            dvs.remove(dv)
    dvs.append(_el("w:docVar", name=DOCVAR_NAME,
                   val=json.dumps(s, ensure_ascii=False,
                                  separators=(",", ":"))))


def read_settings(docx_path):
    """docx に埋め込んだ設定を読み出して検証する。無ければ ValueError。"""
    try:
        doc = Document(docx_path)
    except Exception:
        raise ValueError("docx として開けませんでした")
    dvs = doc.settings.element.find(qn("w:docVars"))
    for dv in ([] if dvs is None else dvs.findall(qn("w:docVar"))):
        if dv.get(qn("w:name")) != DOCVAR_NAME:
            continue
        try:
            data = json.loads(dv.get(qn("w:val")) or "")
        except ValueError:
            raise ValueError("埋め込まれた書式の設定が壊れています")
        return normalize(data)
    raise ValueError("この docx には書式の設定が入っていません"
                     "（このアプリで作った見本・整形済みの docx を選んでください）")


# ============================================================
# 組み立て
# ============================================================

def template_map(num_id):
    """apply_style の TEMPLATE_MAP と同じ形の対応表。"""
    d = {k: (sid, num_id, str(ilvl)) for k, (sid, _n, ilvl) in STYLE_DEFS.items()}
    return {
        "name": "カスタム",
        "num_id": num_id,
        "chapter": d["chapter"],
        "article_title": d["article"],
        "paragraph_1": d["para1"],
        "paragraph_n": d["paraN"],
        "item": d["item"],
        "item_sub_multi": d["sub"],
        "item_sub_single": d["sub_single"],
        "item_sub2": d["sub2"],
        "table": (TABLE_STYLE[0], None, None),
        "lead_space": (),
    }


def build(settings, base_paths, out_path):
    """
    設定どおりのテンプレート docx を out_path に作り、対応表を返す。
    base_paths: {"t1": テンプレ1のパス, ...}
    """
    s = normalize(settings)
    doc = Document(base_paths[s["base"]])
    num_id = _add_numbering(doc, s)
    toc_num = _add_toc_numbering(doc, s) if s["layout"] == "separate" else None
    _add_styles(doc, s, num_id)
    _set_toc2_numbering(doc, toc_num)
    write_settings(doc, s)
    doc.save(out_path)
    return template_map(num_id)
