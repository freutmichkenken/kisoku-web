# -*- coding: utf-8 -*-
"""
相互参照（クロスリファレンス）の保持モジュール

問題
----
パーサは段落をプレーンテキスト（`para.text`）として取り出すため、
Wordの相互参照フィールドが失われる。整形後は「第１条」という
ただの文字列になり、条を追加・削除しても追従しなくなる。

    元:  <w:r><w:fldChar begin/></w:r>
         <w:r><w:instrText> REF _Ref236732735 \\n \\h </w:instrText></w:r>
         <w:r><w:fldChar separate/></w:r>
         <w:r><w:t>第１条</w:t></w:r>        ← 表示されているのはこれ
         <w:r><w:fldChar end/></w:r>
    整形後: <w:r><w:t>第１条</w:t></w:r>       ← フィールドが消える

方針
----
パーサ側は変更しない（テキスト処理の途中で文字位置を追跡するのは
壊れやすいため）。代わりに、

    1. 元のdocxから「参照フィールド」と「参照先ブックマーク」を回収する
    2. apply_style が通常どおり整形して段落を組み立てる
    3. 出力の直前に、テキストを手がかりに元の段落を突き止めて
       フィールドとブックマークを復元する

対象は REF / PAGEREF / NOTEREF（相互参照系）のみ。
ページ番号や目次などのフィールドは前付け・後付けの複製で残るため
ここでは扱わない。

復元したフィールドには dirty 属性を付けるので、Wordで開いたときに
参照先の番号が自動で更新される（条の増減に追従する）。
"""

import re
from copy import deepcopy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# 相互参照とみなすフィールド。
# PAGEREF は目次のページ番号で、参照元の段落（目次の行）が本文に
# 存在しないため対象外。目次は Word 側で更新（F9）すれば再生成される。
_REF_INSTR_RE = re.compile(r'^\s*(REF|NOTEREF)\b', re.I)
# REF のあとに続くブックマーク名
_REF_NAME_RE = re.compile(r'^\s*(?:REF|NOTEREF)\s+(\S+)', re.I)
# 目次が使う自動生成ブックマーク（本文に復元しても Word が作り直す）
_TOC_BM_RE = re.compile(r'^_Toc', re.I)
# 段落番号を表示する REF のスイッチ（\n \r \w）
_PARA_NUM_SW_RE = re.compile(r'\\[nrw]\b', re.I)
# 「第N条」を表示している参照
_ART_RESULT_RE = re.compile(r'^第[0-9０-９一二三四五六七八九十百千]+条')
# 「第N条」だけ（「第N条第2項」などを含まない）
_ART_ONLY_RE = re.compile(r'^第[0-9０-９一二三四五六七八九十百千]+条$')
# 段落番号スイッチを足すための REF の頭（REF とブックマーク名）
_REF_HEAD_RE = re.compile(r'^(\s*REF\s+\S+)', re.I)
# 復元した相互参照の文字色（青は表記ゆれ、赤は別用途で使用済み）
XREF_COLOR = "00B050"    # 緑


# ============================================================
# 共通ヘルパー
# ============================================================

def _norm(s):
    """段落の突き合わせ用キー。空白と括弧の全角・半角差を吸収する。"""
    s = re.sub(r'[\s\u3000]', '', s or '')
    return s.replace('（', '(').replace('）', ')')


def _p_text(p):
    """w:p 要素の表示テキスト（instrText は含まない）。"""
    return ''.join(t.text or '' for t in p.iter(qn('w:t')))


def _iter_text_runs(p):
    """
    段落内の「文字を持つラン」を (run要素, w:t要素, 開始位置, 終了位置)
    で順に返す。コメント参照だけのランなど、文字を持たないランは除く。
    """
    pos = 0
    for r in p.findall(qn('w:r')):
        t = r.find(qn('w:t'))
        if t is None:
            continue
        s = t.text or ''
        yield r, t, pos, pos + len(s)
        pos += len(s)


def _split_run_at(r, t, offset_in_run):
    """
    ランを offset の位置で2つに割り、後半のランを返す。
    書式（rPr）は複製する。
    """
    text = t.text or ''
    new_r = deepcopy(r)
    t.text = text[:offset_in_run]
    new_t = new_r.find(qn('w:t'))
    new_t.text = text[offset_in_run:]
    new_t.set(qn('xml:space'), 'preserve')
    t.set(qn('xml:space'), 'preserve')
    r.addnext(new_r)
    return new_r


def _isolate_range(p, start, end):
    """
    段落テキストの [start, end) がちょうど1つ以上のランに収まるように
    ランを分割し、その範囲のランのリストを返す。範囲が取れなければ [].
    """
    if start >= end:
        return []
    # 後ろから処理すると位置がずれないが、分割で構造が変わるため
    # 都度取り直す（段落内のランは高々数十個なので実用上問題ない）
    for _ in range(4):
        runs = list(_iter_text_runs(p))
        target = [x for x in runs if x[3] > start and x[2] < end]
        if not target:
            return []
        first = target[0]
        if first[2] < start:
            _split_run_at(first[0], first[1], start - first[2])
            continue
        last = target[-1]
        if last[3] > end:
            _split_run_at(last[0], last[1], end - last[2])
            continue
        return [x[0] for x in target]
    return []


# ============================================================
# 1. 元ファイルからの回収
# ============================================================

def collect(src_doc):
    """
    元の docx（python-docx の Document）から相互参照の情報を集める。

    戻り値:
      {
        "fields":    [{"instr","result","para","occ","bookmark"}, ...],
        "bookmarks": {name: {"para","start","end","whole"}},
      }
      para : その段落の表示テキスト（突き合わせに使う）
      occ  : その段落内で result が何番目に出現するか（0始まり）
    """
    fields = []
    bookmarks = {}

    for order, p in enumerate(src_doc.element.body.iter(qn('w:p'))):
        ptext = _p_text(p)
        pos = 0
        cur = None
        open_bm = {}       # id -> (name, 開始位置)

        for el in p.iter():
            tag = el.tag
            if tag == qn('w:fldChar'):
                ftype = el.get(qn('w:fldCharType'))
                if ftype == 'begin':
                    cur = {"instr": "", "start": None, "end": None}
                elif ftype == 'separate' and cur is not None:
                    cur["start"] = pos
                elif ftype == 'end' and cur is not None:
                    cur["end"] = pos
                    if (_REF_INSTR_RE.match(cur["instr"])
                            and cur["start"] is not None
                            and cur["end"] > cur["start"]):
                        result = ptext[cur["start"]:cur["end"]]
                        m = _REF_NAME_RE.match(cur["instr"])
                        fields.append({
                            "instr": cur["instr"],
                            "result": result,
                            "para": ptext,
                            # 同じ文字列が複数あるときの順番
                            "occ": ptext[:cur["start"]].count(result),
                            "bookmark": m.group(1) if m else None,
                            "order": order,
                        })
                    cur = None
            elif tag == qn('w:instrText'):
                if cur is not None:
                    cur["instr"] += el.text or ''
            elif tag == qn('w:t'):
                pos += len(el.text or '')
            elif tag == qn('w:bookmarkStart'):
                open_bm[el.get(qn('w:id'))] = (el.get(qn('w:name')), pos)
            elif tag == qn('w:bookmarkEnd'):
                info = open_bm.pop(el.get(qn('w:id')), None)
                if info:
                    name, st = info
                    if (name and not name.startswith('_GoBack')
                            and not _TOC_BM_RE.match(name)):
                        bookmarks[name] = {
                            "para": ptext, "start": st, "end": pos,
                            "whole": (st == 0 and pos == len(ptext)),
                            "order": order,
                        }

    # 参照されているブックマークだけ残す（Word内部の _Toc などを除くため）
    used = {f["bookmark"] for f in fields if f["bookmark"]}
    bookmarks = {k: v for k, v in bookmarks.items() if k in used}
    return {"fields": fields, "bookmarks": bookmarks}


# ============================================================
# 2. 出力側での復元
# ============================================================

def _find_paragraph(paragraphs, orig_text, must_contain=None, min_index=0):
    """
    整形後の段落から、元段落 orig_text に対応するものを探す。

    誤マッチを避けるため、次の順に絞り込む。
      1. 正規化して完全一致するもの（最優先）
      2. 一方が他方を含むもの（整形で「２　」などの番号が落ちるため）
    さらに、元ファイルと出力で本文の並び順は変わらないので、
    「直前に確定した段落より後ろ」を優先して選ぶ。

    戻り値: (段落インデックス or None, 候補数, 完全一致だったか)
    """
    key = _norm(orig_text)
    if not key:
        return None, 0, False
    exact, loose = [], []
    for i, p in enumerate(paragraphs):
        gk = _norm(p.text)
        if not gk:
            continue
        if must_contain and must_contain not in p.text:
            continue
        if gk == key:
            exact.append(i)
        elif (gk in key or key in gk) and min(len(gk), len(key)) >= 4:
            # 極端に短い一致は誤マッチのもとなので採らない
            loose.append(i)
    for cand, is_exact in ((exact, True), (loose, False)):
        if not cand:
            continue
        ahead = [i for i in cand if i >= min_index]
        return (ahead[0] if ahead else cand[0]), len(cand), is_exact
    return None, 0, False


class _NumLookup:
    """段落の自動番号を (レベル, 番号書式 lvlText) で引く。番号なしは None。"""

    def __init__(self, doc):
        self.styles = {}
        for s in doc.styles.element.iter(qn('w:style')):
            self.styles[s.get(qn('w:styleId'))] = s
        self.abstract = {}
        self.nums = {}
        try:
            numbering = doc.part.numbering_part.element
        except Exception:
            return          # 自動番号の定義が無い文書
        for a in numbering.iter(qn('w:abstractNum')):
            self.abstract[a.get(qn('w:abstractNumId'))] = a
        for n in numbering.iter(qn('w:num')):
            self.nums[n.get(qn('w:numId'))] = n

    def _style_numpr(self, style_id):
        """スタイル（basedOn をたどる）に設定された numPr と、そのスタイルID。"""
        seen = set()
        while style_id and style_id not in seen:
            seen.add(style_id)
            s = self.styles.get(style_id)
            if s is None:
                return None, None
            numPr = s.find(qn('w:pPr') + '/' + qn('w:numPr'))
            if numPr is not None:
                return numPr, style_id
            based = s.find(qn('w:basedOn'))
            style_id = based.get(qn('w:val')) if based is not None else None
        return None, None

    def _lvl(self, num_id, ilvl, style_id):
        num = self.nums.get(num_id)
        if num is None:
            return None
        for ov in num.findall(qn('w:lvlOverride')):
            lvl = ov.find(qn('w:lvl'))
            if lvl is not None and ov.get(qn('w:ilvl')) == str(ilvl):
                return lvl
        aid = num.find(qn('w:abstractNumId'))
        a = self.abstract.get(aid.get(qn('w:val')) if aid is not None else None)
        if a is None:
            return None
        lvls = a.findall(qn('w:lvl'))
        if ilvl is None:
            # スタイル経由でレベル指定が無い場合は、そのスタイルに結び付いたレベル
            for lvl in lvls:
                ps = lvl.find(qn('w:pStyle'))
                if ps is not None and ps.get(qn('w:val')) == style_id:
                    return lvl
            ilvl = 0
        for lvl in lvls:
            if lvl.get(qn('w:ilvl')) == str(ilvl):
                return lvl
        return None

    def get(self, p):
        pPr = p.find(qn('w:pPr'))
        own = pPr.find(qn('w:numPr')) if pPr is not None else None
        ps = pPr.find(qn('w:pStyle')) if pPr is not None else None
        style_numPr, style_id = self._style_numpr(
            ps.get(qn('w:val')) if ps is not None else None)

        def val(numPr, tag):
            el = numPr.find(qn(tag)) if numPr is not None else None
            return el.get(qn('w:val')) if el is not None else None

        if val(own, 'w:numId'):
            # 段落に直接の番号指定があれば、レベルもスタイルと混ぜない
            num_id, ilvl = val(own, 'w:numId'), val(own, 'w:ilvl')
        else:
            num_id = val(style_numPr, 'w:numId')
            ilvl = val(own, 'w:ilvl') or val(style_numPr, 'w:ilvl')
        if not num_id or num_id == '0':
            return None
        lvl = self._lvl(num_id, int(ilvl) if ilvl else None, style_id)
        if lvl is None:
            return None
        text = lvl.find(qn('w:lvlText'))
        return (int(lvl.get(qn('w:ilvl')) or ilvl or 0),
                (text.get(qn('w:val')) or '') if text is not None else '')


def _article_number_para(paragraphs, idx, nums):
    """
    「第N条」の自動番号を持つ段落のインデックスを返す。

    テンプレートによって「第N条」の番号が付く段落が違う
    （見出し「（〇〇）」に付くもの／第1項の本文に付くもの）。
    元ファイルと同じ段落にブックマークを置くと、テンプレによっては
    項番号（「1」）を参照してしまうため、同じ条の隣の段落へ移す。
    見出しは第1項より上位のレベルなので、レベルの上下で同じ条かを確かめ、
    隣の条の段落へ移らないようにしている。
    """
    def info(i):
        return nums.get(paragraphs[i]._p) if 0 <= i < len(paragraphs) \
            else None

    own = info(idx)
    if own is None:
        # 番号の無い段落は見出しか本文か判別できず、隣の条へ移るおそれがあるため動かさない
        return idx
    if '条' in own[1]:
        return idx
    lvl = own[0]
    prev = info(idx - 1)
    if prev and '条' in prev[1] and prev[0] < lvl:     # 見出しに番号がある
        return idx - 1
    nxt = info(idx + 1)
    if nxt and '条' in nxt[1] and nxt[0] > lvl:        # 第1項に番号がある
        return idx + 1
    return idx


def _next_bookmark_id(doc):
    ids = [int(e.get(qn('w:id')))
           for e in doc.element.body.iter(qn('w:bookmarkStart'))
           if (e.get(qn('w:id')) or '').isdigit()]
    return (max(ids) + 1) if ids else 1000


def _existing_bookmark_names(doc):
    return {e.get(qn('w:name'))
            for e in doc.element.body.iter(qn('w:bookmarkStart'))}


# w:rPr の中で w:color より後ろに置くべき要素（スキーマの順序）
_AFTER_COLOR = {qn('w:' + t) for t in (
    'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u',
    'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
    'eastAsianLayout', 'specVanish', 'oMath')}


def _set_color(r, hex_value):
    """ラン要素に文字色を直接設定する（順序を守らないと Word が修復を求める）。"""
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r.insert(0, rPr)
    for old in rPr.findall(qn('w:color')):
        rPr.remove(old)
    c = OxmlElement('w:color')
    c.set(qn('w:val'), hex_value)
    after = next((e for e in rPr if e.tag in _AFTER_COLOR), None)
    if after is not None:
        after.addprevious(c)
    else:
        rPr.append(c)


def _make_fld_run(kind, dirty=False, instr=None):
    r = OxmlElement('w:r')
    if instr is not None:
        el = OxmlElement('w:instrText')
        el.set(qn('xml:space'), 'preserve')
        el.text = instr
    else:
        el = OxmlElement('w:fldChar')
        el.set(qn('w:fldCharType'), kind)
        if dirty:
            # Wordで開いたときに参照先の番号を更新させる
            el.set(qn('w:dirty'), 'true')
    r.append(el)
    return r


def _convert_typed_refs(info):
    """
    手打ちの条番号（「第10条」という文字）をブックマークで囲み、その文字を
    表示している相互参照を、段落番号を表示する参照（\\n 付き）に作り替える。

    整形で手打ちの条番号はテンプレの自動番号に置き換わり、囲んでいた文字が
    無くなるため、そのままでは参照が段落全体の文章を表示してしまう。
    作り替えた参照は、以降の処理で「第N条」の番号を持つ段落を指す。

    戻り値: (作り替えたフィールドのリスト, 作り替えられなかった旨のメッセージ)
    """
    fields, skipped = [], []
    for f in info["fields"]:
        f = dict(f)
        bm = info["bookmarks"].get(f["bookmark"])
        if (bm and not _PARA_NUM_SW_RE.search(f["instr"])
                and _ART_RESULT_RE.match(f["result"])
                and _REF_HEAD_RE.match(f["instr"])):
            bm_text = _norm(bm["para"][bm["start"]:bm["end"]])
            # 段落の先頭の条番号だけを対象にする（本文中の「第2条による」を
            # 囲んだ参照まで作り替えると、その段落の条番号を指してしまう）
            at_head = not _norm(bm["para"][:bm["start"]])
            if at_head and _ART_ONLY_RE.match(_norm(f["result"])) \
                    and bm_text == _norm(f["result"]):
                f["instr"] = _REF_HEAD_RE.sub(r'\1 \\n', f["instr"], count=1)
            else:
                skipped.append(f"条番号の参照に作り替えられません: {f['result']}"
                               f"（{f['para'][:20]}）")
        fields.append(f)
    return fields, skipped


def restore(doc, info, verbose=True, dirty=True, color=XREF_COLOR):
    """
    整形後の Document に相互参照を復元する。

    color を指定すると、復元した参照の文字をその色にする（整形後の確認用）。
    None なら色を付けない。

    dirty=True（既定）にすると、Wordで開いたときに参照先の条番号が
    自動更新される。その代わり、Wordが毎回「他のファイルを参照する
    フィールドが含まれています…」という確認ダイアログを出す。

    戻り値: (復元したフィールド数, 復元したブックマーク数, 失敗リスト)
    """
    if not info or not info["fields"]:
        return 0, 0, []

    paragraphs = doc.paragraphs
    failed = []
    fields, warned = _convert_typed_refs(info)
    landed = []       # (ブックマーク名, 着地した段落のテキスト)

    # --- ブックマーク（参照先）を先に復元する ---
    bm_id = _next_bookmark_id(doc)
    exists = _existing_bookmark_names(doc)
    n_bm = 0
    # 元ファイルでの並び順に処理し、出力側でも順序が前後しないようにする
    bm_items = sorted(info["bookmarks"].items(),
                      key=lambda kv: kv[1].get("order", 0))
    # 「第N条」を段落番号として参照されているブックマーク
    art_refs = {f["bookmark"] for f in fields
                if f["bookmark"] and _PARA_NUM_SW_RE.search(f["instr"])
                and _ART_RESULT_RE.match(f["result"])}
    nums = _NumLookup(doc) if art_refs else None
    min_index = 0
    for name, bm in bm_items:
        if name in exists:
            continue          # 前付け・表などの複製で既に残っている
        idx, n_cand, is_exact = _find_paragraph(paragraphs, bm["para"],
                                                min_index=min_index)
        if idx is None:
            failed.append(f"参照先が見つかりません: {name}"
                          f"（{bm['para'][:20]}）")
            continue
        if n_cand > 1:
            warned.append(f"参照先の候補が{n_cand}件ありました: "
                          f"{bm['para'][:20]}")
        if name in art_refs:
            moved = _article_number_para(paragraphs, idx, nums)
            if moved != idx:
                idx, is_exact = moved, False   # 移した先は段落全体を範囲にする
        p = paragraphs[idx]
        landed.append((name, p.text))
        min_index = idx        # 以降のブックマークはここより後ろを優先

        start = OxmlElement('w:bookmarkStart')
        start.set(qn('w:id'), str(bm_id))
        start.set(qn('w:name'), name)
        end = OxmlElement('w:bookmarkEnd')
        end.set(qn('w:id'), str(bm_id))
        bm_id += 1

        runs = [] if (bm["whole"] or not is_exact) else \
            _isolate_range(p._p, bm["start"], bm["end"])
        if runs:
            runs[0].addprevious(start)
            runs[-1].addnext(end)
        else:
            # 段落全体を範囲にする（w:pPr は先頭でなければならない）
            pPr = p._p.find(qn('w:pPr'))
            if pPr is not None:
                pPr.addnext(start)
            else:
                p._p.insert(0, start)
            p._p.append(end)
        n_bm += 1

    # --- 参照側のフィールドを復元する ---
    # 表・前付け・後付けは元の XML を複製するのでフィールドが既に残っている。
    # それらを二重に復元したり「見つかりません」と誤って警告したりしないよう、
    # 出力に同じ参照が残っている段落を控えておく（表の中も含めて全段落を見る）
    kept = {}
    for p in doc.element.body.iter(qn('w:p')):
        instr = ''.join(e.text or '' for e in p.iter(qn('w:instrText')))
        if instr:
            kept.setdefault(_norm(_p_text(p)), []).append(_norm(instr))

    n_fld = 0
    min_index = 0
    for f in sorted(fields, key=lambda x: x.get("order", 0)):
        if any(_norm(f["instr"]) in s for s in kept.get(_norm(f["para"]), ())):
            continue
        idx, n_cand, _ = _find_paragraph(paragraphs, f["para"],
                                         must_contain=f["result"],
                                         min_index=min_index)
        if idx is None:
            failed.append(f"参照元が見つかりません: {f['result']}"
                          f"（{f['para'][:20]}）")
            continue
        p = paragraphs[idx]
        min_index = idx
        text = p.text
        # 同じ文字列が複数ある場合は、元と同じ順番のものを選ぶ
        idx, found = -1, -1
        for _ in range(f["occ"] + 1):
            idx = text.find(f["result"], idx + 1)
            if idx < 0:
                break
            found = idx
        if found < 0:
            found = text.find(f["result"])
        if found < 0:
            failed.append(f"表示位置が見つかりません: {f['result']}")
            continue

        runs = _isolate_range(p._p, found, found + len(f["result"]))
        if not runs:
            failed.append(f"ランの分割に失敗: {f['result']}")
            continue

        instr = f["instr"]
        if color and 'MERGEFORMAT' not in instr.upper():
            # 更新後も表示文字の書式（緑字）を保たせる
            instr = instr.rstrip() + ' \\* MERGEFORMAT '
        fld_runs = [_make_fld_run('begin', dirty=dirty),
                    _make_fld_run(None, instr=instr),
                    _make_fld_run('separate')]
        for r in fld_runs:
            runs[0].addprevious(r)
        fld_runs.append(_make_fld_run('end'))
        runs[-1].addnext(fld_runs[-1])
        if color:
            # 表示文字だけでなくフィールドコードにも付け、更新後も色が残るようにする
            for r in [fld_runs[1]] + runs:
                _set_color(r, color)
        n_fld += 1

    if verbose:
        if n_fld or n_bm:
            print(f"  相互参照: {n_fld}件を復元"
                  f"（参照先ブックマーク {n_bm}件）"
                  + ("・緑字で表示" if color else ""))
            # どこに着地したかを出す（飛び先が正しいか目視できるように）
            for name, txt in landed[:20]:
                print(f"    {name} → {txt[:28]}")
            if len(landed) > 20:
                print(f"    …ほか {len(landed) - 20}件")
        for msg in warned:
            print(f"  ※ 相互参照の照合が曖昧です → {msg}")
        for msg in failed:
            print(f"  ※ 相互参照を復元できませんでした → {msg}")
    return n_fld, n_bm, failed
