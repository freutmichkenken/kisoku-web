# -*- coding: utf-8 -*-
"""
就業規則 スタイル適用スクリプト（3テンプレ対応・既存テンプレ流し込み方式）

kisoku_parser.py が生成した JSON を読み込み、
指定テンプレート（テンプレ1 / テンプレ2 / テンプレ3）の
スタイルを適用した docx を生成します。

Google Colab で動作します。
  pip install python-docx

使い方:
  from apply_style import apply_style
  apply_style(
      json_path="kisoku_parsed.json",
      template_path="就業規則テンプレ3.docx",
      template_key="t3",              # "t1" / "t2" / "t3"
      output_path="整形済.docx",
  )

番号（第X条・①など）はテンプレートのスタイルが自動採番するため、
JSON の body（番号を除いた本文）だけを流し込みます。
"""

import json
import copy
import re
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# Word コメント管理
#   python-docx はコメントを標準サポートしないため OOXML を直接操作する。
# ============================================================

class CommentManager:
    """出力docxにWordコメントを追加する。"""

    def __init__(self, doc):
        self.doc = doc
        self.comments = []   # (id, author, text)
        self._next_id = 0

    def add_comment_to_paragraph(self, paragraph, text, author="整形ツール"):
        """指定段落の全ランを範囲としてコメントを付ける。"""
        cid = self._next_id
        self._next_id += 1
        self.comments.append((cid, author, text))

        p = paragraph._p
        start = OxmlElement('w:commentRangeStart')
        start.set(qn('w:id'), str(cid))
        # w:pPr は段落の先頭でなければならないため、その直後に挿入する
        pPr = p.find(qn('w:pPr'))
        if pPr is not None:
            pPr.addnext(start)
        else:
            p.insert(0, start)
        end = OxmlElement('w:commentRangeEnd')
        end.set(qn('w:id'), str(cid))
        p.append(end)
        run = OxmlElement('w:r')
        ref = OxmlElement('w:commentReference')
        ref.set(qn('w:id'), str(cid))
        run.append(ref)
        p.append(run)

    def add_comment_to_run(self, run, text, author="整形ツール"):
        """特定のラン（語）だけを範囲としてコメントを付ける。"""
        cid = self._next_id
        self._next_id += 1
        self.comments.append((cid, author, text))

        r = run._r
        start = OxmlElement('w:commentRangeStart')
        start.set(qn('w:id'), str(cid))
        r.addprevious(start)
        end = OxmlElement('w:commentRangeEnd')
        end.set(qn('w:id'), str(cid))
        r.addnext(end)
        ref_run = OxmlElement('w:r')
        ref = OxmlElement('w:commentReference')
        ref.set(qn('w:id'), str(cid))
        ref_run.append(ref)
        end.addnext(ref_run)

    def finalize(self):
        """comments.xml パートを作成して document に関連付ける。"""
        if not self.comments:
            return
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        from docx.opc.constants import RELATIONSHIP_TYPE as RT

        W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        parts_xml = [f'<w:comments xmlns:w="{W}">']
        for cid, author, text in self.comments:
            initials = author[:2]
            # 改行を <w:br/> に変換して Word 上で改行表示させる
            runs = []
            lines = text.split('\n')
            for li, line in enumerate(lines):
                safe = (line.replace('&', '&amp;').replace('<', '&lt;')
                            .replace('>', '&gt;'))
                if li > 0:
                    runs.append('<w:r><w:br/></w:r>')
                runs.append(f'<w:r><w:t xml:space="preserve">{safe}</w:t></w:r>')
            parts_xml.append(
                f'<w:comment w:id="{cid}" w:author="{author}" '
                f'w:date="2000-01-01T00:00:00Z" w:initials="{initials}">'
                f'<w:p>{"".join(runs)}</w:p>'
                f'</w:comment>'
            )
        parts_xml.append('</w:comments>')
        comments_xml = ''.join(parts_xml).encode('utf-8')

        partname = PackURI('/word/comments.xml')
        content_type = ('application/vnd.openxmlformats-officedocument.'
                        'wordprocessingml.comments+xml')
        comments_part = Part(partname, content_type, comments_xml,
                             self.doc.part.package)
        self.doc.part.relate_to(comments_part, RT.COMMENTS)


# ============================================================
# テンプレートごとのスタイル対応表
# JSON要素 → (style_id, numId, ilvl)
#   numId/ilvl が None のスタイルは numPr を設定しない（style だけ）
# ============================================================

TEMPLATE_MAP = {
    # --- テンプレ1（就業規則テンプレ1.docx / numId=16） ---
    "t1": {
        "name": "テンプレ1",
        "num_id": "16",
        "chapter":         ("02-",   "16", "1"),   # 章
        "article_title":   ("03-",   "16", "2"),   # 条見出し
        "paragraph_1":     ("04-",   "16", "3"),   # 第1項
        "paragraph_n":     ("04-",   "16", "3"),   # 第2項以降
        "item":            ("05-1",  "16", "4"),   # 号 (1)
        "item_sub_multi":  ("06-10", "16", "5"),   # 号の下位 1.
        "item_sub_single": ("06-",   "16", "6"),   # 号の下位（……）
        "item_sub2":       ("07-0",  "16", "7"),   # さらに下 (ア)
        "table":           ("10-",   None, None),  # 表中の文字
        # 行頭に全角スペースを入れる段落種別（字下げ用）。
        # テンプレ1にも入れる余地があるので、必要になったら
        # ("paragraph_1", "paragraph_n") にする。
        "lead_space":      (),
    },
    # --- テンプレ2（条の後ろに1項 / numId=37） ---
    "t2": {
        "name": "テンプレ2",
        "num_id": "37",
        "chapter":         ("02-",  "37", "1"),
        "article_title":   ("03-",  "37", "2"),    # 条見出し（番号なし）
        "paragraph_1":     ("04-",  "37", "3"),    # 第1項（条番号を自動付与）
        "paragraph_n":     ("05-",  "37", "4"),    # 第2項以降
        "item":            ("06-1", "37", "5"),    # 号 (1)
        "item_sub_multi":  ("07-1", "37", "6"),    # 号の下位 1.
        "item_sub_single": ("07-",  "37", "7"),    # 号の下位（……）
        "item_sub2":       ("08-",  "37", "8"),    # さらに下 (ア)
        "table":           ("10-",  None, None),
        "lead_space":      (),
    },
    # --- テンプレ3（条の後ろに1項・項を1字下げ / numId=37） ---
    #   スタイル定義はテンプレ2と同一。違いは項の行頭に全角スペースを
    #   入れること。インデントでは表現できないため、文字として入れる。
    "t3": {
        "name": "テンプレ3",
        "num_id": "37",
        "chapter":         ("02-",  "37", "1"),
        "article_title":   ("03-",  "37", "2"),    # 条見出し（番号なし）
        "paragraph_1":     ("04-",  "37", "3"),    # 第1項（条番号を自動付与）
        "paragraph_n":     ("05-",  "37", "4"),    # 第2項以降
        "item":            ("06-1", "37", "5"),    # 号 (1)
        "item_sub_multi":  ("07-1", "37", "6"),    # 号の下位 1.
        "item_sub_single": ("07-",  "37", "7"),    # 号の下位（……）
        "item_sub2":       ("08-",  "37", "8"),    # さらに下 (ア)
        "table":           ("10-",  None, None),
        # 項だけ1字下げる（号・表には入れない）
        "lead_space":      ("paragraph_1", "paragraph_n"),
    },
    # --- テンプレ4（第1項/第2項分離 / numId=16）※旧テンプレ3 ---
    "t4": {
        "name": "テンプレ4",
        "num_id": "16",
        "chapter":         ("02-",   "16", "1"),
        "article_title":   ("03-",   "16", "2"),   # 条見出し
        "paragraph_1":     ("04-1",  "16", "3"),   # 第1項（番号表示なし）
        "paragraph_n":     ("04-2",  "16", "4"),   # 第2項以降
        "item":            ("05-1",  "16", "5"),   # 号 (1)
        "item_sub_multi":  ("06-10", "16", "6"),   # 号の下位 1.
        "item_sub_single": ("06-",   "16", "7"),   # 号の下位（……）
        "item_sub2":       ("07-0",  "16", "8"),   # さらに下 (ア)
        "table":           ("10-",   None, None),
        # インデントで字下げを表現できているので全角スペースは入れない
        "lead_space":      (),
    },
}


# ============================================================
# 段落生成のヘルパー
# ============================================================

def _set_numbering(paragraph, num_id, ilvl):
    """段落のpPrにnumPr(numId, ilvl)を設定する。python-docxのXML直接操作。"""
    if num_id is None:
        return
    pPr = paragraph._p.get_or_add_pPr()
    # 既存のnumPrを除去
    for old in pPr.findall(qn('w:numPr')):
        pPr.remove(old)
    numPr = OxmlElement('w:numPr')
    ilvl_el = OxmlElement('w:ilvl')
    ilvl_el.set(qn('w:val'), str(ilvl if ilvl is not None else 0))
    numId_el = OxmlElement('w:numId')
    numId_el.set(qn('w:val'), str(num_id))
    numPr.append(ilvl_el)
    numPr.append(numId_el)
    pPr.append(numPr)


def _set_pstyle_direct(paragraph, style_id):
    """pPr に w:pStyle を直接書き込む（styleId で確実に指定）。"""
    pPr = paragraph._p.get_or_add_pPr()
    for old in pPr.findall(qn('w:pStyle')):
        pPr.remove(old)
    ps = OxmlElement('w:pStyle')
    ps.set(qn('w:val'), style_id)
    pPr.insert(0, ps)


# ============================================================
# 元docxからの引き継ぎ（表のXML複製・文字書式の復元）
# ============================================================

def _load_source_tables(source_docx):
    """元docxの表要素を出現順に取得する。失敗したら空リスト。"""
    if not source_docx:
        return []
    try:
        src = Document(source_docx)
        return [c for c in _iter_block_elements(src.element.body)
                if c.tag == qn('w:tbl')]
    except Exception as e:
        print(f"  ※ 元docxの表を読めませんでした（再構築で代替）: {e}")
        return []


def _sanitize_copied_table(tbl_el):
    """
    複製した表XMLから、出力文書と衝突しうる要素を取り除く。
    （コメント参照・ブックマークはID重複の原因になるため）
    """
    for tag in ('w:commentRangeStart', 'w:commentRangeEnd',
                'w:commentReference', 'w:bookmarkStart', 'w:bookmarkEnd'):
        for el in tbl_el.findall('.//' + qn(tag)):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
    return tbl_el



HYOKI_COLOR = "0000FF"   # 表記ゆれの着色（青）
LEAD_SPACE = "\u3000"     # 字下げ用の全角スペース


def _style_name_by_id(doc, style_id):
    """styleId から Word 上の表示名（w:name）を引く。無ければ None。"""
    if not style_id:
        return None
    for st in doc.styles.element.findall(qn('w:style')):
        if st.get(qn('w:styleId')) != style_id:
            continue
        nm = st.find(qn('w:name'))
        if nm is not None:
            return nm.get(qn('w:val'))
    return None


def _toc_style_spec(doc, M):
    """
    目次に載せるスタイルを TOC フィールドの \\t 用の文字列にする。
    章 → レベル1、条見出し → レベル2。

    アウトラインレベル（\\o）ではなくスタイル指定（\\t）を使う理由:
      テンプレ2は条見出し（03-）にアウトラインレベルが無く、
      代わりに第1項（04-）がレベル1を持っている。\\o で拾うと
      第1項の本文がまるごと目次に入ってしまうため、
      TEMPLATE_MAP の章・条見出しのスタイルを名指しする。
    """
    pairs = []
    for key, level in (("chapter", "1"), ("article_title", "2")):
        sid = M.get(key, (None,))[0]
        name = _style_name_by_id(doc, sid)
        if not name:
            continue
        if ',' in name or '"' in name:
            # \t スイッチはカンマ区切りなので、含む名前は使えない
            print(f"  ※ スタイル名にカンマが含まれるため目次から除外: {name}")
            continue
        pairs.append(f"{name},{level}")
    return ",".join(pairs)


def _toc_style_numbering(doc, style_name):
    """
    目次スタイル（toc 1 / toc 2）自身が自動番号を持っているかを調べる。

    テンプレ2は条見出し（03-）が番号を持たない代わりに、
    目次スタイル toc 2 に「第%1条」の連番を割り当てている。
    この場合、目次の条番号はWordが目次を作るときに付く。
    """
    for st in doc.styles.element.findall(qn('w:style')):
        nm = st.find(qn('w:name'))
        if nm is None or nm.get(qn('w:val')) != style_name:
            continue
        npr = st.find(qn('w:pPr') + '/' + qn('w:numPr'))
        if npr is None:
            pPr = st.find(qn('w:pPr'))
            npr = pPr.find(qn('w:numPr')) if pPr is not None else None
        if npr is None:
            return None
        ni = npr.find(qn('w:numId'))
        return ni.get(qn('w:val')) if ni is not None else None
    return None


def _numbering_of(doc, num_id, ilvl):
    """
    numbering.xml から (numFmt, lvlText) を引く。見つからなければ (None, None)。
    """
    if num_id is None:
        return None, None
    try:
        nel = doc.part.numbering_part.element
    except Exception:
        return None, None
    abst = None
    for num in nel.findall(qn('w:num')):
        if num.get(qn('w:numId')) != str(num_id):
            continue
        a = num.find(qn('w:abstractNumId'))
        abst = a.get(qn('w:val')) if a is not None else None
        break
    if abst is None:
        return None, None
    for a in nel.findall(qn('w:abstractNum')):
        if a.get(qn('w:abstractNumId')) != abst:
            continue
        for lvl in a.findall(qn('w:lvl')):
            if lvl.get(qn('w:ilvl')) != str(ilvl):
                continue
            nf = lvl.find(qn('w:numFmt'))
            lt = lvl.find(qn('w:lvlText'))
            return (nf.get(qn('w:val')) if nf is not None else None,
                    lt.get(qn('w:val')) if lt is not None else None)
    return None, None


def _shows_number(numfmt, lvltext):
    """その番号設定が実際に文字を表示するか（テンプレ2の条見出しは表示しない）。"""
    if numfmt == 'none':
        return False
    if not lvltext:
        return False
    # %1 などのプレースホルダと空白を除いて何か残るか
    return bool(re.sub(r'%\d|[\s\u3000]', '', lvltext))


def _insert_toc(doc, M, page_break=True, heading="目　次", dirty=True):
    """
    目次（TOCフィールド）を文書のこの位置に追加する。
    dirty=True（既定）にすると、Wordで開いたときにフィールドが自動更新される。
    その代わり、Wordが毎回「他のファイルを参照するフィールドが
    含まれています…」という確認ダイアログを出す（「はい」で更新）。
    """
    spec = _toc_style_spec(doc, M)
    if not spec:
        print("  ※ 目次に載せるスタイルが特定できないため、目次を作成しませんでした")
        return False

    # --- 見出し「目　次」 ---
    if heading:
        hp = doc.add_paragraph()
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hr = hp.add_run(heading)
        hr.bold = True

    # --- TOCフィールド本体 ---
    p = doc.add_paragraph()
    instr = f' TOC \\h \\z \\t "{spec}" '

    def _fld(kind=None, instr_text=None, dirty=False):
        r = OxmlElement('w:r')
        if instr_text is not None:
            el = OxmlElement('w:instrText')
            el.set(qn('xml:space'), 'preserve')
            el.text = instr_text
        else:
            el = OxmlElement('w:fldChar')
            el.set(qn('w:fldCharType'), kind)
            if dirty:
                el.set(qn('w:dirty'), 'true')
        r.append(el)
        return r

    p._p.append(_fld('begin', dirty=dirty))
    p._p.append(_fld(instr_text=instr))
    p._p.append(_fld('separate'))
    p.add_run("【目次】Wordで開いたときに自動作成されます"
              "（作成されない場合は Ctrl+A → F9）")
    p._p.append(_fld('end'))

    # --- 目次の次のページから本文を始める ---
    if page_break:
        bp = doc.add_paragraph()
        br = OxmlElement('w:r')
        brk = OxmlElement('w:br')
        brk.set(qn('w:type'), 'page')
        br.append(brk)
        bp._p.append(br)

    print(f"  目次を挿入しました（対象スタイル: {spec}）")
    # --- 目次に条番号が出るかを事前に確認する ---
    toc2_num = _toc_style_numbering(doc, "toc 2")
    a_sid, a_num, a_ilvl = M.get("article_title", (None, None, None))
    fmt, lvltext = _numbering_of(doc, a_num, a_ilvl)
    if toc2_num:
        print(f"    条番号は目次スタイル toc 2 の連番（numId={toc2_num}）で付きます")
    elif _shows_number(fmt, lvltext):
        print(f"    条番号は条見出しスタイル {a_sid} の番号（{lvltext}）で付きます")
    else:
        print(f"    ※ このテンプレートは条見出しスタイル {a_sid} が番号を表示せず"
              f"（numFmt={fmt}）、")
        print("       目次スタイル toc 2 にも連番がありません。"
              "目次に条番号が付きません。")
        print("       テンプレート側で toc 2 スタイルに"
              "「第%1条」の連番を設定してください。")
    return True


def _set_run_color(run, hex_value=HYOKI_COLOR):
    """ランに文字色を直接設定する。"""
    rPr = run._r.get_or_add_rPr()
    for old in rPr.findall(qn('w:color')):
        rPr.remove(old)
    c = OxmlElement('w:color')
    c.set(qn('w:val'), hex_value)
    rPr.append(c)


def _add_runs_with_spans(p, text, spans, out_runs=None):
    """
    text を spans（[{start,end},...]）の位置で分割してランを作り、
    該当部分だけ色を付ける。spans は位置順・重なりなしであること。
    out_runs にリストを渡すと、着色したランが spans と同じ順で入る
    （語単位でコメントを付けるため）。
    """
    pos = 0
    for s in spans:
        st, en = s["start"], s["end"]
        if st < pos:          # 念のための保険
            if out_runs is not None:
                out_runs.append(None)
            continue
        if st > pos:
            p.add_run(text[pos:st])
        run = p.add_run(text[st:en])
        _set_run_color(run)
        if out_runs is not None:
            out_runs.append(run)
        pos = en
    if pos < len(text):
        p.add_run(text[pos:])


def _add_paragraph(doc, text, style_id, num_id, ilvl, spans=None,
                   out_runs=None):
    """指定スタイル(styleId)＋numPrで段落を1つ追加する。"""
    p = doc.add_paragraph()
    # スタイルは styleId で直接指定（表示名に依存しない）
    _set_pstyle_direct(p, style_id)
    if text:
        if spans:
            _add_runs_with_spans(p, text, spans, out_runs)
        else:
            p.add_run(text)
    _set_numbering(p, num_id, ilvl)
    return p


# ============================================================
# メイン処理
# ============================================================

def _print_notice(source_docx=None):
    """整形前に表示する注意事項。"""
    print("─" * 56)
    print("【整形を開始します。次の点にご注意ください】")
    print("  ・元ファイルの文字書式（フォント・太字・文字色）は")
    print("    テンプレートのスタイルに置き換わり、引き継がれません。")
    print("    強調表示など重要な書式は、整形後に付け直してください。")
    if not source_docx:
        print("  ・元ファイルが指定されていないため、表を再構築します。")
        print("    セル結合のある表は崩れることがあります。")
    print("  ・文書の書き方によっては、条・項・号の判定を誤ることがあります。")
    print("    整形結果は必ず元ファイルと照合してください。")
    print("─" * 56)
    print()


def _append_block(doc, element):
    """
    body の本文末尾に要素を追加する。
    sectPr（セクション設定）がある場合はその手前に挿入する。
    （単純な append だと sectPr の後ろに置かれ、文書末尾に飛んでしまう）
    """
    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    if sectPr is not None:
        sectPr.addprevious(element)
    else:
        body.append(element)


def _iter_block_elements(body):
    """
    本文のブロック要素（段落・表）を出現順に取り出す。
    コンテンツコントロール（w:sdt）等の入れ物は中身を展開する。
    kisoku_parser 側と同じ並びにするため、必ず同じ規則で走査する。
    """
    for child in body.iterchildren():
        tag = child.tag
        if tag in (qn('w:p'), qn('w:tbl')):
            yield child
        elif tag in (qn('w:sdt'), qn('w:customXml')):
            for sub in child.iterchildren():
                if sub.tag in (qn('w:sdtContent'), qn('w:customXmlPr')):
                    for x in _iter_block_elements(sub):
                        yield x
                elif sub.tag in (qn('w:p'), qn('w:tbl')):
                    yield sub
                elif sub.tag in (qn('w:sdt'), qn('w:customXml')):
                    for x in _iter_block_elements(sub):
                        yield x


def _load_source_body(source_docx):
    """
    元docxの Document と body 直下要素を返す。
    返り値: (Document または None, 要素リスト)
    """
    if not source_docx:
        return None, []
    try:
        src = Document(source_docx)
        return src, list(_iter_block_elements(src.element.body))
    except Exception as e:
        print(f"  ※ 元docxを読めませんでした: {e}")
        return None, []


_W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _style_maps(doc):
    """styleId -> (表示名, pPr要素, rPr要素) の辞書を作る。"""
    out = {}
    if doc is None:
        return out
    try:
        for st in doc.styles.element.findall(qn('w:style')):
            sid = st.get(qn('w:styleId'))
            if not sid:
                continue
            nm = st.find(qn('w:name'))
            out[sid] = (nm.get(qn('w:val')) if nm is not None else None,
                        st.find(qn('w:pPr')), st.find(qn('w:rPr')),
                        st.find(qn('w:tblPr')), st.get(qn('w:type')))
    except Exception:
        pass
    return out


def _fix_copied_tblstyle(tbl, src_styles, out_styles):
    """
    複製した表の tblStyle を安全にする。

    表スタイルIDも文書ごとに意味が違う（例: 元ファイルの "a5" は
    「Table Grid（罫線あり）」だが、テンプレートの "a5" は
    段落スタイルの「footer」）。そのまま残すと罫線が失われる。

    → テンプレートに同じIDかつ同じ名前の表スタイルがある場合だけ残し、
      それ以外は tblStyle を外して、元スタイルの罫線を表に直接書き込む。
    """
    import copy as _copy
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        return
    ts = tblPr.find(qn('w:tblStyle'))
    if ts is None:
        return
    sid = ts.get(qn('w:val'))
    src = src_styles.get(sid)
    out = out_styles.get(sid)
    src_name = src[0] if src else None
    out_name = out[0] if out else None
    out_type = out[4] if out and len(out) > 4 else None
    if out_name is not None and out_name == src_name and out_type == 'table':
        return          # テンプレにも同じ表スタイルがある → そのまま
    tblPr.remove(ts)
    # 元の表スタイルの罫線・セル余白を直接書き込む
    s_tblPr = src[3] if src and len(src) > 3 else None
    if s_tblPr is not None:
        for tag in ('w:tblBorders', 'w:tblCellMar', 'w:tblLayout'):
            if tblPr.find(qn(tag)) is None:
                e = s_tblPr.find(qn(tag))
                if e is not None:
                    tblPr.append(_copy.deepcopy(e))


def _fix_copied_pstyles(el, src_styles, out_styles):
    """
    複製した段落の pStyle を安全にする。

    スタイルIDは文書ごとに意味が違う（例: 元ファイルの "2" は
    「本文インデント 2」だが、テンプレートの "2" は「見出し2」で
    自動採番付き）。同じIDでも別物なら、そのまま残すと
    「第二節」などの誤った番号が振られてしまう。

    → テンプレートに同じIDかつ同じ名前のスタイルがある場合だけ残し、
      それ以外は pStyle を外して、元スタイルの見た目
      （インデント・配置・行間・フォント）を段落へ直接書き込む。
    """
    import copy as _copy
    for para in el.iter(qn('w:p')):
        pPr = para.find(qn('w:pPr'))
        if pPr is None:
            continue
        ps = pPr.find(qn('w:pStyle'))
        if ps is None:
            continue
        sid = ps.get(qn('w:val'))
        src_name = src_styles.get(sid, (None,))[0]
        out_name = out_styles.get(sid, (None,))[0]
        if out_name is not None and out_name == src_name:
            continue          # テンプレにも同じスタイルがある → そのまま
        # 別物なので外す
        pPr.remove(ps)
        # 元スタイルの見た目を段落に直接書き込む
        _, s_pPr, s_rPr = src_styles.get(sid, (None, None, None, None, None))[:3]
        if s_pPr is not None:
            for tag in ('w:ind', 'w:jc', 'w:spacing'):
                if pPr.find(qn(tag)) is None:
                    src_el = s_pPr.find(qn(tag))
                    if src_el is not None:
                        pPr.append(_copy.deepcopy(src_el))
        if s_rPr is not None:
            for run in para.findall(qn('w:r')):
                rPr = run.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    run.insert(0, rPr)
                for tag in ('w:rFonts', 'w:sz', 'w:szCs'):
                    if rPr.find(qn(tag)) is None:
                        src_el = s_rPr.find(qn(tag))
                        if src_el is not None:
                            rPr.append(_copy.deepcopy(src_el))


def _sanitize_copied_para(el):
    """
    複製した段落XMLから、出力文書と衝突しうる要素を取り除く。
    numPr は元文書の採番定義を参照しており、テンプレートでは
    別の番号が振られてしまうため除去する。
    """
    for tag in ('w:commentRangeStart', 'w:commentRangeEnd',
                'w:commentReference', 'w:bookmarkStart', 'w:bookmarkEnd'):
        for e in el.findall('.//' + qn(tag)):
            parent = e.getparent()
            if parent is not None:
                parent.remove(e)
    for npr in el.findall('.//' + qn('w:numPr')):
        parent = npr.getparent()
        if parent is not None:
            parent.remove(npr)
    return el


def _emit_verbatim(doc, node, src_body, src_styles=None, out_styles=None):
    """
    前付け（表紙・前文）／後付け（附則・改版履歴）を出力する。
    元docxがあれば該当要素のXMLをそのまま複製し、
    センタリング・フォント・改ページなどの体裁を保持する。
    整形は行わない（元のまま保持する箇所）。
    """
    import copy as _copy
    indices = node.get('indices') or []
    n = 0
    if src_body and indices:
        for idx in indices:
            if not (0 <= idx < len(src_body)):
                continue
            try:
                el = _copy.deepcopy(src_body[idx])
                if el.tag == qn('w:p'):
                    _sanitize_copied_para(el)
                elif el.tag == qn('w:tbl'):
                    _sanitize_copied_table(el)
                    if src_styles is not None:
                        _fix_copied_tblstyle(el, src_styles, out_styles or {})
                else:
                    continue
                if src_styles is not None:
                    _fix_copied_pstyles(el, src_styles, out_styles or {})
                _append_block(doc, el)
                n += 1
            except Exception as e:
                print(f"  ※ 前付け要素{idx}の複製に失敗: {e}")
        return n

    # フォールバック: テキストだけ段落として出力
    for t in node.get('texts') or []:
        p = doc.add_paragraph()
        p.add_run(t)
        n += 1
    return n


_R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def _remap_relationships(el, src_doc, out_doc):
    """
    複製した要素の中にある関係ID（画像などへの参照）を、
    出力文書側の関係IDに貼り替える。
    図形のみ（参照なし）の場合は何もしない。
    """
    if src_doc is None:
        return
    try:
        src_part = src_doc.part
        out_part = out_doc.part
    except Exception:
        return
    for e in el.iter():
        for key in list(e.attrib.keys()):
            if not key.startswith(_R_NS):
                continue
            old_rid = e.attrib[key]
            if not old_rid:
                continue
            try:
                rel = src_part.rels[old_rid]
            except Exception:
                continue
            try:
                if rel.is_external:
                    new_rid = out_part.relate_to(
                        rel.target_ref, rel.reltype, is_external=True)
                else:
                    new_rid = out_part.relate_to(rel.target_part, rel.reltype)
                e.set(key, new_rid)
            except Exception:
                pass


def _emit_drawing(doc, node, src_body, src_doc=None,
                  src_styles=None, out_styles=None):
    """
    図形・画像を含む段落を、元のXMLをそのまま複製して出力する。
    （テキストに置き換えると図形が失われるため）
    """
    import copy as _copy
    idx = node.get('index')
    if src_body and isinstance(idx, int) and 0 <= idx < len(src_body):
        try:
            el = _copy.deepcopy(src_body[idx])
            _sanitize_copied_para(el)
            if src_styles is not None:
                _fix_copied_pstyles(el, src_styles, out_styles or {})
            _remap_relationships(el, src_doc, doc)
            _append_block(doc, el)
            return True
        except Exception as e:
            print(f"  ※ 図形{idx}の複製に失敗: {e}")
    # フォールバック: テキストだけ出力（図形は失われる）
    t = (node.get('text') or '').strip()
    if t:
        p = doc.add_paragraph()
        p.add_run(t)
        return True
    return False


def _add_table(doc, tbl_node, style_id, src_tables=None,
               src_styles=None, out_styles=None):
    """
    表を出力する。
    元docxの表XMLが使える場合はそれを複製し、セル結合・罫線・
    セル内の書式をそのまま保持する。使えない場合のみ再構築する。
    """
    import copy as _copy
    idx = tbl_node.get('index') if isinstance(tbl_node, dict) else None
    # 1) 元の表XMLをそのまま複製（推奨）
    if src_tables and isinstance(idx, int) and 0 <= idx < len(src_tables):
        try:
            new_tbl = _copy.deepcopy(src_tables[idx])
            _sanitize_copied_table(new_tbl)
            if src_styles is not None:
                _fix_copied_tblstyle(new_tbl, src_styles, out_styles or {})
            # セル内の段落にテンプレートの表用スタイル（10-）を適用する。
            # セル結合や罫線は元のまま保ちつつ、文字体裁だけ揃える。
            if style_id:
                for cell_p in new_tbl.findall('.//' + qn('w:p')):
                    pPr = cell_p.find(qn('w:pPr'))
                    if pPr is None:
                        pPr = OxmlElement('w:pPr')
                        cell_p.insert(0, pPr)
                    for old in pPr.findall(qn('w:pStyle')):
                        pPr.remove(old)
                    ps = OxmlElement('w:pStyle')
                    ps.set(qn('w:val'), style_id)
                    pPr.insert(0, ps)
            _append_block(doc, new_tbl)
            return True
        except Exception as e:
            print(f"  ※ 表{idx}の複製に失敗、再構築します: {e}")

    # 2) フォールバック: rows から再構築（セル結合は再現できない）
    rows = tbl_node.get('rows', []) if isinstance(tbl_node, dict) else tbl_node
    if not rows:
        return False
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)
    table = doc.add_table(rows=n_rows, cols=n_cols)
    try:
        table.style = doc.styles["Table Grid"]
    except KeyError:
        pass
    for i, row in enumerate(rows):
        for j in range(n_cols):
            text = row[j] if j < len(row) else ""
            cell = table.cell(i, j)
            p = cell.paragraphs[0]
            if style_id:
                try:
                    _set_pstyle_direct(p, style_id)
                except Exception:
                    pass
            if text:
                p.add_run(text)
    return True


def _wrap_title(title):
    """
    条タイトルを全角括弧で囲む。
    既に全角/半角の括弧で囲まれている場合は全角に正規化して二重付与を防ぐ。
    空文字の場合は空のまま返す。
    """
    if not title:
        return ""
    t = title.strip()
    # 既存の外側の括弧（全角・半角）を剥がす
    while len(t) >= 2 and t[0] in "（(" and t[-1] in "）)":
        t = t[1:-1].strip()
    if not t:
        return ""
    return f"（{t}）"


def apply_style(json_path, template_path, template_key, output_path,
                source_docx=None, show_notice=True,
                check_hyoki=False, hyoki_options=None,
                keep_crossref=True, insert_toc=False,
                auto_update_fields=True, verify=True):
    """
    JSON構造をテンプレートに流し込んでスタイル適用済みdocxを生成する。

    json_path     : kisoku_parser.py が出力した JSON
    template_path : テンプレートdocxのパス
    template_key  : "t1" / "t2" / "t3"
    output_path   : 出力docxのパス
    source_docx   : 元docx。指定すると表を原本から複製し、セル結合や
                    セル内書式を保持できる（推奨）
    show_notice   : 整形前の注意事項を表示するか。アプリ側の画面で
                    既に表示している場合は False にする
    auto_update_fields :
                    True にすると、目次・相互参照のフィールドに dirty を付け、
                    Wordで開いたときに自動更新させる。ただしWordが毎回
                    「他のファイルを参照するフィールドが含まれています」
                    という確認ダイアログを出す（「はい」で更新される）。
                    既定は True。ダイアログを出したくない場合は False にし、
                    Wordで Ctrl+A → F9 で更新する。
    verify        : True にすると、出力後に原本と突き合わせて欠落・増加を
                    検出し、照合レポート（.md）を出力先と同じ場所に作る。
                    source_docx の指定が必要。
    insert_toc    : True にすると、前付け（表紙）の直後に目次（TOCフィールド）を
                    挿入する。章・条見出しのスタイルを名指しで拾うので、
                    テンプレ2でも第1項の本文が目次に混ざらない。
    keep_crossref : True にすると、元docxのWord相互参照（REFフィールド）を
                    整形後の文書に復元する。source_docx の指定が必要。
    check_hyoki   : True にすると表記ゆれを検出し、該当語を青字にして
                    段落ごとにWordコメントを付ける（本文のみ・自動修正はしない）
    hyoki_options : hyoki_check.analyze() に渡す辞書。例:
                    {"use_auto": False,
                     "extra_groups": [{"category":"用語",
                                       "forms":["役職者","管理職"]}],
                     "exclude_forms": ["社員"]}
    """
    if template_key not in TEMPLATE_MAP:
        raise ValueError(f"template_key は t1/t2/t3 のいずれか。指定値: {template_key}")
    M = TEMPLATE_MAP[template_key]

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # --- 表記ゆれの解析（本文のみ・出力前に文書全体を見る必要がある） ---
    hyoki_report = None
    if check_hyoki:
        try:
            from hyoki_check import analyze as _hyoki_analyze
            hyoki_report = _hyoki_analyze(data, **(hyoki_options or {}))
        except Exception as e:
            print(f"  ※ 表記ゆれチェックをスキップしました: {e}")
            hyoki_report = None

    # テンプレートを開き、本文（body内の段落・表）を全削除
    doc = Document(template_path)
    body = doc.element.body
    # sectPr（セクション設定）は残す。それ以外の子要素を削除。
    sectPr = body.find(qn('w:sectPr'))
    for child in list(body):
        if child is sectPr:
            continue
        body.remove(child)

    # JSON を走査して段落を積む
    n_chapter = n_article = n_para = n_item = n_table = 0
    n_comment = 0

    table_style_id = M.get("table", (None, None, None))[0]

    # --- 整形を始める前に、必ず読んでいただきたい注意事項を表示 ---
    # （アプリ側の画面で既に表示している場合は show_notice=False で抑制）
    if show_notice:
        _print_notice(source_docx)

    # 元docxの表XMLを引き継ぐ準備（セル結合・書式を保持するため）
    src_tables = _load_source_tables(source_docx)
    # 前付け（表紙・前文）の複製元
    src_doc, src_body = _load_source_body(source_docx)

    # --- 元docxから相互参照（REFフィールド）を回収しておく ---
    xref_info = None
    if keep_crossref and src_doc is not None:
        try:
            import crossref
            xref_info = crossref.collect(src_doc)
            if not xref_info["fields"]:
                xref_info = None
        except Exception as e:
            print(f"  ※ 相互参照の回収をスキップしました: {e}")
            xref_info = None
    # スタイルIDの意味が元ファイルとテンプレートで食い違う問題への対策
    src_styles = _style_maps(src_doc)
    out_styles = _style_maps(doc)
    n_front = 0
    n_back = 0
    n_draw = 0

    comment_mgr = CommentManager(doc)
    n_hyoki_comment = 0  # 表記ゆれコメントの件数（＝ゆれの種類数）
    n_hyoki_word = 0     # 青字にした語の数
    commented_gids = set()   # コメント済みの表記ゆれグループ

    def _add_p(text, key):
        """
        テンプレのスタイルで段落を1つ追加する。
        表記ゆれチェックが有効なら、確定した文字列を再スキャンして
        該当語を青字にする。コメントは1件のゆれにつき1個だけ、
        文書内で最初に出てきた箇所に付ける。
        （同じ段落が複数のゆれの初出になる場合は、ゆれごとに別コメント）
        """
        nonlocal n_hyoki_comment, n_hyoki_word, n_comment
        # テンプレの仕様で行頭を1字下げる（インデントで表現できないため
        # 全角スペースを文字として入れる）。JSONは常にスペース無しで保持し、
        # 出力時にだけ付けるので、二重付与にはならない。
        if text and key in M.get("lead_space", ()) \
                and not text.startswith(LEAD_SPACE):
            text = LEAD_SPACE + text
        spans = hyoki_report.scan(text) if (hyoki_report and text) else None
        runs = [] if spans else None
        p = _add_paragraph(doc, text, *M[key], spans=spans, out_runs=runs)
        if spans:
            n_hyoki_word += len(spans)
            for s, run in zip(spans, runs):
                if s["gid"] in commented_gids or run is None:
                    continue
                commented_gids.add(s["gid"])
                # その語だけを範囲にしてコメントを付ける
                comment_mgr.add_comment_to_run(
                    run, hyoki_report.comment_text([s]), author="表記ゆれ")
                n_hyoki_comment += 1
                n_comment += 1
        return p

    def _split_hint_text(hint):
        """_split_hint の内容から、コメント本文を作る。"""
        lines = ["【要確認】この項目は複数に分割すべき可能性があります。",
                 "AIによる分割候補:"]
        for i, part in enumerate(hint, 1):
            snippet = part if len(part) <= 40 else part[:40] + "…"
            lines.append(f"{i}. {snippet}")
        return "\n".join(lines)

    def _restructured_text(reasons):
        """_ai_restructured の内容から、コメント本文を作る。"""
        lines = ["【AI再構成】この条は階層構造が疑わしいと判定され、",
                 "AIが項・号の階層を組み直しました。内容をご確認ください。",
                 "検出理由:"]
        for r in reasons[:5]:
            lines.append(f"・{r}")
        return "\n".join(lines)

    def emit_tables(container):
        """dict内の tables を出力する"""
        nonlocal n_table
        for tbl in container.get("tables", []):
            _add_table(doc, tbl, table_style_id, src_tables,
                       src_styles, out_styles)
            n_table += 1

    def emit_drawings(container):
        """dict内の drawings（図形・画像）を元のまま出力する"""
        nonlocal n_draw
        for dw in container.get("drawings", []):
            if _emit_drawing(doc, dw, src_body, src_doc,
                             src_styles, out_styles):
                n_draw += 1

    def emit_article(art):
        nonlocal n_article, n_para, n_item, n_comment
        n_article += 1
        # 条見出し（タイトル）— 全角括弧で囲んで出力
        # 「（目的）　※〇〇パターン」のような末尾注釈があれば括弧の外に付ける
        title = _wrap_title(art.get("title", ""))
        note = (art.get("title_note") or "").strip()
        if note:
            title = f"{title}　{note}" if title else note
        _add_p(title, "article_title")
        # 項
        for pi, para in enumerate(art.get("paragraphs", [])):
            n_para += 1
            key = "paragraph_1" if para.get("number", 1) == 1 else "paragraph_n"
            body_text = para.get("body", "")
            p_obj = _add_p(body_text, key)
            # 項に紐づく表・図は、号より前（本文の直後）に出す
            emit_tables(para)
            emit_drawings(para)
            # 分割要確認フラグがあればコメントを付ける
            if para.get("_split_hint"):
                comment_mgr.add_comment_to_paragraph(
                    p_obj, _split_hint_text(para["_split_hint"]))
                n_comment += 1
            # AIが条を再構成した場合もコメントで知らせる
            if para.get("_ai_restructured"):
                comment_mgr.add_comment_to_paragraph(
                    p_obj, _restructured_text(para["_ai_restructured"]))
                n_comment += 1
            # 号
            items = para.get("items", [])
            for it in items:
                n_item += 1
                it_obj = _add_p(it.get("body", ""), "item")
                # 号に紐づく表・図は、下位項目より前に出す
                emit_tables(it)
                emit_drawings(it)
                if it.get("_split_hint"):
                    comment_mgr.add_comment_to_paragraph(
                        it_obj, _split_hint_text(it["_split_hint"]))
                    n_comment += 1
                # 号の下位項目（複数なら item_sub_multi、単体なら item_sub_single）
                subs = it.get("sub_items", [])
                if subs:
                    sub_key = "item_sub_multi" if len(subs) >= 2 else "item_sub_single"
                    if sub_key not in M:
                        sub_key = "item_sub_multi"
                    if sub_key in M:
                        for su in subs:
                            _add_p(su.get("body", ""), sub_key)
                            # 号の下位に紐づく表・図（さらに下より前）
                            emit_tables(su)
                            emit_drawings(su)
                            # 号の下位のさらに下（(ア)(イ)…）
                            for su2 in (su.get("sub_items2") or []):
                                if "item_sub2" in M:
                                    _add_p(su2.get("body", ""), "item_sub2")
                                else:
                                    _add_p(su2.get("body", ""), sub_key)
                                # さらに下に紐づく表・図・数式
                                emit_tables(su2)
                                emit_drawings(su2)

    toc_done = False

    def _maybe_insert_toc():
        """前付けを出し終えて本文に入る直前に、一度だけ目次を挿入する。"""
        nonlocal toc_done
        if insert_toc and not toc_done:
            _insert_toc(doc, M, dirty=auto_update_fields)
            toc_done = True

    for node in data:
        t = node.get("type")
        if t == "front_matter":
            n_front += _emit_verbatim(doc, node, src_body, src_styles, out_styles)
            continue
        if t != "back_matter":
            # 前付けが終わって本文に入るところで目次を挟む
            _maybe_insert_toc()
        if t == "back_matter":
            n_back += _emit_verbatim(doc, node, src_body, src_styles, out_styles)
        elif t == "chapter":
            n_chapter += 1
            _add_p(node.get("title", ""), "chapter")
            for art in node.get("articles", []):
                emit_article(art)
            emit_tables(node)   # 章直下の表
            emit_drawings(node)
        elif t == "article":
            emit_article(node)
        elif t == "table":
            _add_table(doc, node, table_style_id, src_tables,
                       src_styles, out_styles)
            n_table += 1
        elif t == "drawing":
            if _emit_drawing(doc, node, src_body, src_doc,
                             src_styles, out_styles):
                n_draw += 1

    # コメントを確定（comments.xml パートを作成）
    # --- 相互参照を復元する（ランの分割を伴うので最後に行う） ---
    if xref_info:
        try:
            import crossref
            crossref.restore(doc, xref_info, dirty=auto_update_fields)
        except Exception as e:
            print(f"  ※ 相互参照の復元に失敗しました: {e}")

    comment_mgr.finalize()

    doc.save(output_path)
    print(f"[{M['name']}] 出力完了: {output_path}")

    # --- 原本との照合（保存後のファイルを読み直して突き合わせる） ---
    report_path = None
    if verify and source_docx:
        try:
            import os
            import verify_report
            _src = Document(source_docx)
            _out = Document(output_path)
            result = verify_report.compare(_src, _out)
            base = os.path.splitext(output_path)[0]
            report_path = base + "_照合レポート.md"
            verify_report.write_report(
                result, report_path,
                src_name=os.path.basename(source_docx),
                out_name=os.path.basename(output_path))
        except Exception as e:
            print(f"  ※ 原本との照合に失敗しました: {e}")
            result = None
    print(f"  章 {n_chapter} / 条 {n_article} / 項 {n_para} / 号 {n_item} / 表 {n_table}")
    if n_comment:
        print(f"  要確認コメント: {n_comment}件（Wordのコメント欄に表示されます）")
    if report_path:
        print("  ── 原本との照合 ──")
        for line in verify_report.summary_lines(result):
            print(line)
        print(f"  レポート: {report_path}")
    if (insert_toc or xref_info) and not auto_update_fields:
        print("  ★ Wordで開いたら Ctrl+A → F9 を押してください"
              "（目次の作成・相互参照の更新）")
    if hyoki_report is not None:
        print("  ── 表記ゆれチェック ──")
        for line in hyoki_report.summary_lines():
            print(line)
        if n_hyoki_word:
            print(f"  → 青字にした語 {n_hyoki_word}箇所 ／ "
                  f"コメント {n_hyoki_comment}件（ゆれ1件につき初出の1箇所のみ）")
            print("    ※ 自動修正はしていません。統一するかはご判断ください。")
    if n_front:
        print(f"  前付け（表紙・前文）を {n_front}要素 引き継ぎました")
    if n_back:
        print(f"  後付け（附則等）を {n_back}要素 引き継ぎました（整形なし）")
    if n_draw:
        print(f"  図形・画像を {n_draw}件 引き継ぎました")
    if src_tables:
        print(f"  表は元ファイルのものを複製（セル結合・書式を保持）")

    return output_path


# ============================================================
# 実行例（Colab）
# ============================================================

if __name__ == "__main__":
    # Colabではファイル名を実際のものに合わせてください
    apply_style(
        json_path="kisoku_parsed.json",
        template_path="就業規則テンプレ3.docx",
        template_key="t3",
        output_path="整形済_テンプレ3.docx",
    )
