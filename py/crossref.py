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


def _next_bookmark_id(doc):
    ids = [int(e.get(qn('w:id')))
           for e in doc.element.body.iter(qn('w:bookmarkStart'))
           if (e.get(qn('w:id')) or '').isdigit()]
    return (max(ids) + 1) if ids else 1000


def _existing_bookmark_names(doc):
    return {e.get(qn('w:name'))
            for e in doc.element.body.iter(qn('w:bookmarkStart'))}


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


def restore(doc, info, verbose=True, dirty=True):
    """
    整形後の Document に相互参照を復元する。

    dirty=True（既定）にすると、Wordで開いたときに参照先の条番号が
    自動更新される。その代わり、Wordが毎回「他のファイルを参照する
    フィールドが含まれています…」という確認ダイアログを出す。

    戻り値: (復元したフィールド数, 復元したブックマーク数, 失敗リスト)
    """
    if not info or not info["fields"]:
        return 0, 0, []

    paragraphs = doc.paragraphs
    failed = []
    warned = []
    landed = []       # (ブックマーク名, 着地した段落のテキスト)

    # --- ブックマーク（参照先）を先に復元する ---
    bm_id = _next_bookmark_id(doc)
    exists = _existing_bookmark_names(doc)
    n_bm = 0
    # 元ファイルでの並び順に処理し、出力側でも順序が前後しないようにする
    bm_items = sorted(info["bookmarks"].items(),
                      key=lambda kv: kv[1].get("order", 0))
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
    n_fld = 0
    min_index = 0
    for f in sorted(info["fields"], key=lambda x: x.get("order", 0)):
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

        runs[0].addprevious(_make_fld_run('begin', dirty=dirty))
        runs[0].addprevious(_make_fld_run(None, instr=f["instr"]))
        runs[0].addprevious(_make_fld_run('separate'))
        runs[-1].addnext(_make_fld_run('end'))
        n_fld += 1

    if verbose:
        if n_fld or n_bm:
            print(f"  相互参照: {n_fld}件を復元"
                  f"（参照先ブックマーク {n_bm}件）")
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
