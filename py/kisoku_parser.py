# -*- coding: utf-8 -*-
"""
就業規則 JSON パーサー
Google Colab で動作します。
python-docx のみ必要（pip install python-docx）

出力JSON構造:
[
  {
    "type": "chapter",
    "number": "第１章",
    "title": "総則",
    "articles": [
      {
        "type": "article",
        "number": "第１条",
        "title": "目的",          # 条タイトルがある場合
        "paragraphs": [
          {
            "type": "paragraph",
            "number": 1,           # 第1項は番号なし表示だが内部的に1
            "body": "この就業規則は…",
            "items": [
              {"type": "item", "number": 1, "body": "履歴書"}
            ]
          },
          {
            "type": "paragraph",
            "number": 2,
            "body": "前項の定めにより…",
            "items": []
          }
        ]
      }
    ]
  }
]
"""

import re
import json
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

import functools

# 記号階層の動的判定（ブラウザ版でも必ず読み込む。worker.js の PY_FILES）
from marker_hierarchy import (identify_symbol, _SYMBOL_TYPES, kanji_to_int,
                              REF_AFTER_MARKER as _REF_AFTER_MARKER)


# ============================================================
# スタイル名から要素種別を判定するマップ
#   numbering.xml が無くても、テンプレ由来のスタイル名が付いていれば
#   構造を確定できる。テンプレ1/2/3 のスタイルIDを網羅。
# ============================================================

STYLE_KIND_MAP = {
    # 章
    '02-': 'chapter',
    # 条見出し
    '03-': 'article',
    # 第1項
    '04-1': 'paragraph',
    '04-':  'paragraph',   # テンプレ2の第1項
    # 第2項以降
    '04-2': 'paragraph',
    '05-':  'paragraph',   # テンプレ2の第2項以降
    # 号
    '05-1': 'item',
    '06-1': 'item_sub',    # 号の下位
    '06-10':'item_sub',    # 号の下位（テンプレ1/3の新スタイル）
    '07-1': 'item_sub',    # テンプレ2の号の下位
    '06-':  'item_sub',
    '07-':  'item_sub',
    # 号の下位のさらに下（(ア)(イ)(ウ)）
    '07-0': 'item_sub2',   # テンプレ1/3
    '08-':  'item_sub2',   # テンプレ2
    # 本文（番号なし）
    '00-':  'body_plain',
}


def get_style_id(para):
    """段落に直接指定された pStyle の styleId を返す。"""
    ppr = para._p.find(qn('w:pPr'))
    if ppr is None:
        return None
    ps = ppr.find(qn('w:pStyle'))
    if ps is None:
        return None
    return ps.get(qn('w:val'))


def get_style_name(para, doc):
    """段落に適用されているスタイルの「表示名」を返す（例: 'toc 1'）。"""
    style_id = get_style_id(para)
    if not style_id:
        return None
    for s in doc.styles.element.findall(qn('w:style')):
        if s.get(qn('w:styleId')) != style_id:
            continue
        nm = s.find(qn('w:name'))
        if nm is not None:
            return nm.get(qn('w:val'))
        break
    return None


_MC_NS = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
# Word の数式（OMML）の名前空間
_MATH_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'


def iter_block_elements(body):
    """
    本文のブロック要素（段落・表）を、文書に現れる順に取り出す。

    コンテンツコントロール（w:sdt）やカスタムXML（w:customXml）は、
    それ自体が段落や表を内包する「入れ物」なので、中身を展開して
    同じ並びに含める。展開しないと、その中の条文が丸ごと失われる。
    """
    for child in body.iterchildren():
        tag = child.tag
        if tag in (qn('w:p'), qn('w:tbl')):
            yield child
        elif tag in (qn('w:sdt'), qn('w:customXml')):
            # 中身（sdtContent / customXml 直下）を再帰的に展開
            for sub in child.iterchildren():
                if sub.tag in (qn('w:sdtContent'), qn('w:customXmlPr')):
                    for x in iter_block_elements(sub):
                        yield x
                elif sub.tag in (qn('w:p'), qn('w:tbl')):
                    yield sub
                elif sub.tag in (qn('w:sdt'), qn('w:customXml')):
                    for x in iter_block_elements(sub):
                        yield x


def has_drawing(para):
    """
    段落に図形・画像・テキストボックス・数式が含まれるか。

    これらはテキストに置き換えると失われるため、
    元のXMLをそのまま複製して保持する対象とする。
    """
    el = para._p
    return bool(
        el.findall('.//' + qn('w:drawing'))
        or el.findall('.//' + qn('w:pict'))
        or el.findall('.//' + _MC_NS + 'AlternateContent')
        or el.findall('.//' + qn('w:object'))
        # Word の数式ブロック（oMath / oMathPara）
        or el.findall('.//' + _MATH_NS + 'oMath')
        or el.findall('.//' + _MATH_NS + 'oMathPara')
    )


def has_underline(para):
    """段落内に下線（フォントの下線書式）が引かれた文字があるか。"""
    for u in para._p.findall('.//' + qn('w:u')):
        val = u.get(qn('w:val'))
        if val is None or val != 'none':
            return True
    return False


def is_layout_paragraph(text):
    """
    スペースやタブで位置合わせした「疑似的な図・レイアウト行」かを判定する。
    （例:「      賃金      　    手  当             役付手当」）

    これらは本文として結合するとレイアウトが崩れるため、
    元のXMLをそのまま複製して保持する。

    誤検出を避けるため、次はすべて除外する:
      - 章・条・条タイトル・項の行
      - 行頭に番号記号がある行（号などの列揃えと区別するため）
      - 句点で終わる行（通常の文）
    """
    s = (text or '').strip()
    if len(s) < 2:
        return False
    if s.endswith(('。', '、', '．')):
        return False
    ln = line_info(s)
    if ln.chapter or ln.article or ln.title or ln.sym_type is not None:
        return False
    if line_info(text).para:
        return False
    # (a) 内部に「半角3個以上」または「全角2個以上」の空白があり、
    #     それで区切られた塊が2つ以上ある＝列揃えされたレイアウト行
    chunks = [c for c in re.split(r'[ \t]{3,}|\u3000{2,}', s) if c.strip()]
    if len(chunks) >= 2:
        return True

    # (a2) アンダーバー（_＿）が並ぶ罫線・記入欄の行
    if len(re.findall(r'[_＿]{3,}', s)) >= 1:
        return True

    # (b) 深く字下げされた短い断片＝図の中のラベル
    #     （例:「                　                 通勤手当」）
    #     本文の折り返し行と区別するため、短さを条件にする
    lead = 0
    for c in (text or ''):
        if c == '\u3000': lead += 2
        elif c in ' \t': lead += 1
        else: break
    if lead >= 8 and len(s) <= 20:
        return True

    return False


def is_toc_paragraph(para, doc):
    """
    目次（Table of Contents）の段落かどうか判定する。
    Wordの目次は 'toc 1' 'toc 2' ... や 'TOC Heading' というスタイル名を持つ。
    目次を本文と誤認識すると、章条見出しが二重に取り込まれるため除外する。
    """
    name = get_style_name(para, doc)
    if not name:
        return False
    low = name.lower().replace('-', ' ').strip()
    return low.startswith('toc') or low.startswith('table of contents')


# 手打ちの目次の行
#  a. 行末が「リーダー記号2文字以上 or タブ」+ ページ番号
#  b. 章・条の見出しで、行末が「（第1条～第4条）」のような条の範囲の括弧
_NUM_CH = '[0-9０-９一二三四五六七八九十百]+'
#  ページ番号は「６～７」「10～14」のような範囲でもよい
_RE_TOC_LEADER_PAGE = re.compile(
    r'(?:[…‥・．.]{2,}|\t)[\s　]*[－-]?[\s　]*[0-9０-９]+[\s　]*[－-]?'
    r'(?:[\s　]*[～〜~][\s　]*[－-]?[\s　]*[0-9０-９]+[\s　]*[－-]?)?[\s　]*$')
# 目次の見出し（「目次」「目　次」「【目次】」）
_RE_TOC_TITLE = re.compile(
    r'^[【［\[〈＜<（(]?[\s　]*目[\s　]*次[\s　]*[】］\]〉＞>）)]?$')
_RE_TOC_ARTICLE_RANGE = re.compile(
    r'[（(][\s　]*第[\s　]*' + _NUM_CH + r'[\s　]*条'
    r'(?:[\s　]*[～〜~－ー-][\s　]*第?[\s　]*' + _NUM_CH + r'[\s　]*条)?'
    r'[\s　]*[）)][\s　]*$')


def is_manual_toc_line(text):
    """手打ちの目次の行か（本文が始まる前の段落にだけ使う）。"""
    s = text.strip()
    if not s:
        return False
    if _RE_TOC_LEADER_PAGE.search(text.rstrip('\r\n')):
        return True
    st = text.lstrip('　 \t').rstrip()
    ln = line_info(st)
    if (ln.chapter or ln.article) and _RE_TOC_ARTICLE_RANGE.search(st):
        return True
    return False


def _toc_heading_number(st):
    """章・条の見出しなら ('chapter'|'article', (番号, 枝番))、それ以外は None。
    枝番まで比べるので、目次の「第2条」のあとの「第2条の2」は番号が戻ったとみなさない。"""
    ln = line_info(st)
    if ln.chapter:
        return 'chapter', (zenkaku_to_int(ln.chapter.group(1)), 0)
    if ln.article:
        br = ln.article.group('branch')
        return 'article', (zenkaku_to_int(ln.article.group('num')), zenkaku_to_int(br) if br else 0)
    return None


def find_manual_toc(paras):
    """
    手打ちの目次の段落の位置（child_idx の集合）を返す。
    paras: 目次スタイルでない段落の (child_idx, テキスト) を文書の順に並べたもの。

    行末で目次と分かる行（リーダー＋ページ番号、条の範囲）だけでなく、
    ページ番号の無い目次の行（「第１条（目的）」だけの行）も拾うため、
    「目次」の見出しか目次の行から始まるまとまりを、
    番号が戻る章（第１章がもう一度出る）、または条だけの目次のあとの章・番号が戻る条
    で終わるとみなす。
    途中に句点で終わる行（＝本文）があるか、番号が戻らないまま終わったときは、
    まとまりを目次とみなさず、本文の始まりより前の、行末で目次と分かる行だけを返す。
    """
    # 本文の始まり（最初の章・条の見出し）より前にある、行末で目次と分かる行
    base = set()
    for idx, text in paras:
        st = text.lstrip('　 \t').rstrip()
        if is_manual_toc_line(text):
            base.add(idx)
        elif _toc_heading_number(st):
            break

    start = None
    for k, (idx, text) in enumerate(paras):
        st = text.lstrip('　 \t').rstrip()
        if _RE_TOC_TITLE.match(st) or is_manual_toc_line(text):
            start = k
            break
        if _toc_heading_number(st):
            return base          # 目次より前に本文が始まった
    if start is None:
        return base

    block = set()
    last = {'chapter': None, 'article': None}
    for idx, text in paras[start:]:
        st = text.lstrip('　 \t').rstrip()
        h = _toc_heading_number(st)
        if st and not is_manual_toc_line(text) and not _RE_TOC_TITLE.match(st):
            if h:
                kind, n = h
                if kind == 'chapter' and (
                        (last['chapter'] is not None and n <= last['chapter'])
                        or (last['chapter'] is None and last['article'] is not None)):
                    return base | block      # ここから本文
                if (kind == 'article' and last['chapter'] is None
                        and last['article'] is not None and n <= last['article']):
                    return base | block      # 条だけの目次で番号が戻った
            if st.endswith('。'):
                return base              # 文がある → 目次ではない
        if h:
            last[h[0]] = h[1]
        block.add(idx)
    return base                          # 本文の始まりが見つからない


# ============================================================
# STEP 0: 元ファイルの自動番号（numPr）定義を読み取る
# ============================================================

def build_numpr_map(doc):
    """
    元docxの numbering.xml を読み、
    (numId, ilvl) -> 要素種別 のマップを作る。

    Wordの自動段落番号が設定されている場合、番号テキストは
    p.text に現れないため、この情報で種別を判定する。

    返り値: {(numId, ilvl): 'chapter'|'article'|'paragraph'|'item'|'item_sub'}
    """
    mapping = {}
    try:
        num_part = doc.part.numbering_part
    except (NotImplementedError, AttributeError, KeyError):
        return mapping   # numbering.xml が無い文書
    root = num_part.element

    # numId -> abstractNumId
    num2abst = {}
    for num in root.findall(qn('w:num')):
        nid = num.get(qn('w:numId'))
        a = num.find(qn('w:abstractNumId'))
        if a is not None:
            num2abst[nid] = a.get(qn('w:val'))

    # abstractNumId -> {ilvl: (numFmt, lvlText)}
    abst_levels = {}
    for abst in root.findall(qn('w:abstractNum')):
        aid = abst.get(qn('w:abstractNumId'))
        levels = {}
        for lvl in abst.findall(qn('w:lvl')):
            il = lvl.get(qn('w:ilvl'))
            nf = lvl.find(qn('w:numFmt'))
            lt = lvl.find(qn('w:lvlText'))
            ps = lvl.find(qn('w:pStyle'))
            levels[il] = (
                nf.get(qn('w:val')) if nf is not None else None,
                lt.get(qn('w:val')) if lt is not None else '',
                ps.get(qn('w:val')) if ps is not None else None,
            )
        abst_levels[aid] = levels

    # (numId, ilvl) -> 種別 を推定
    # numFmt と lvlText に加え、pStyle名も判定材料にする。
    for nid, aid in num2abst.items():
        levels = abst_levels.get(aid, {})
        for il, (numfmt, lvltext, pstyle) in levels.items():
            kind = _guess_kind(numfmt, lvltext, pstyle)
            if kind:
                mapping[(nid, il)] = kind
    return mapping


def _guess_kind(numfmt, lvltext, pstyle=None):
    """numFmt・lvlText・pStyle から要素種別を推定する。"""
    lt = lvltext or ''
    ps = pstyle or ''

    # --- pStyle が標準命名なら最優先で信頼する ---
    #   02-=章, 03-=条, 04-=項(第1項), 05-=項(第2項)/号, 06-1=号/号の下位, 07-=号の下位
    #   （文書によって割り当てが違うため、minor と numFmt で補正）
    if ps:
        m = re.match(r'^0(\d)-(\d)?', ps)
        if m:
            major = m.group(1)
            minor = m.group(2)
            if major == '2':
                return 'chapter'
            if major == '3':
                return 'article'
            if major == '4':
                # 04-1/04- は第1項、04-2 は第2項。番号形式に関わらず項
                return 'paragraph_1' if minor != '2' else 'paragraph'
            if major == '5':
                # 05-1 は号、05- は項
                return 'item' if minor == '1' else 'paragraph'
            if major == '6':
                # 06-1 は号（丸数字）または号の下位
                if numfmt == 'decimalEnclosedCircle':
                    return 'item'
                return 'item_sub'
            if major == '7':
                return 'item_sub'

    # --- pStyle が無い/非標準なら lvlText と numFmt で判断 ---
    if '章' in lt:
        return 'chapter'
    if '条' in lt:
        return 'article'
    if numfmt == 'decimalEnclosedCircle':
        return 'item'
    if numfmt in ('decimal', 'decimalFullWidth') and re.search(r'[（(]%\d+[）)]', lt):
        return 'item'
    if numfmt == 'none':
        return 'paragraph_1'
    if numfmt in ('decimal', 'decimalFullWidth'):
        return 'paragraph'
    return None


def get_paragraph_numpr(para, doc):
    """
    段落の numPr (numId, ilvl) を取得する。
    段落に直接指定がなければ、適用スタイルの定義から継承する。
    返り値: (numId, ilvl) または (None, None)
    """
    # 1) 段落自身の pPr/numPr
    ppr = para._p.find(qn('w:pPr'))
    if ppr is not None:
        npr = ppr.find(qn('w:numPr'))
        if npr is not None:
            ni = npr.find(qn('w:numId'))
            il = npr.find(qn('w:ilvl'))
            num_id = ni.get(qn('w:val')) if ni is not None else None
            ilvl = il.get(qn('w:val')) if il is not None else '0'
            if num_id is not None:
                return num_id, ilvl

    # 2) スタイル定義から継承
    style_id = None
    if ppr is not None:
        ps = ppr.find(qn('w:pStyle'))
        if ps is not None:
            style_id = ps.get(qn('w:val'))
    if not style_id:
        return None, None

    for s in doc.styles.element.findall(qn('w:style')):
        if s.get(qn('w:styleId')) != style_id:
            continue
        spr = s.find(qn('w:pPr'))
        if spr is None:
            break
        npr = spr.find(qn('w:numPr'))
        if npr is None:
            break
        ni = npr.find(qn('w:numId'))
        il = npr.find(qn('w:ilvl'))
        num_id = ni.get(qn('w:val')) if ni is not None else None
        ilvl = il.get(qn('w:val')) if il is not None else '0'
        if num_id is not None:
            return num_id, ilvl
        break
    return None, None


# ============================================================
# STEP 1: 生テキストの取得と途中改行の結合
# ============================================================


class Line:
    """
    1行（段落、または段落内の改行で分けた1行）の見た目の判定結果。
    行頭の章・条・条タイトル・項・号の記号とその番号、行末の文字、字下げを、
    line_info() で行ごとに一度だけ求める。結合・分類・記号の学習・図の判定は
    これを使い、結合した文字列に見出し・記号の正規表現をかけ直さない
    （結合後は、先頭の行の判定結果と、後ろの行をつないだ文字列から作る）。
    """
    __slots__ = ('text', 's', 'indent', 'chapter', 'article', 'title', 'para',
                 'decimal', 'paren_only', 'sym_type', 'sym_num', 'sym_body', 'sym_char',
                 'mark_head', 'is_marker', 'last', 'section')

    def __init__(self, text):
        self.text = text
        s = text.lstrip('　 \t')
        self.s = s
        w = 0
        for c in text:
            if c == '　':
                w += 2
            elif c == ' ':
                w += 1
            elif c == '\t':
                w += 4
            else:
                break
        self.indent = w                           # 行頭空白の視覚幅
        self.chapter = RE_CHAPTER.match(s)        # 「第1章　総則」
        self.article = RE_ARTICLE.match(s)        # 「第1条（目的）　本文」
        self.section = RE_SECTION.match(s)        # 「第1節　労働時間」
        self.title = RE_TITLE.match(s)            # 「（目的）」だけの行
        self.para = RE_PARA.match(text)           # 第2項以降「２　本文」
        self.decimal = _RE_DECIMAL_HEAD.match(s)  # 「1.25倍」
        # 括弧だけの行（結合の区切りに使う。RE_TITLE より広い）
        self.paren_only = bool(_RE_PAREN_ONLY.match(s))
        # 号の記号（marker_hierarchy の記号の種類・番号・後ろの本文。本文は前後の空白込み）
        self.sym_type, self.sym_num, self.sym_body, self.sym_char = None, 0, s, ''
        for type_name, pat, extractor in _SYMBOL_TYPES:
            m = pat.match(s)
            if m:
                self.sym_type = type_name
                self.sym_num, self.sym_body = extractor(m)
                if type_name in _KANA_TYPES:
                    self.sym_char = s.lstrip('（(')[0]   # 記号の文字（イロハ順の判定に使う）
                break
        self.mark_head = _RE_MARK_HEAD.match(s)   # 「（１）」「①」「（一）」
        self.is_marker = bool(self.mark_head) or self.sym_type is not None
        t = text.rstrip('　 \t')
        self.last = t[-1] if t else ''           # 行末の文字


@functools.lru_cache(maxsize=8192)
def line_info(text):
    """行の見た目の判定結果（Line）。同じ文字列は一度だけ判定する。"""
    return Line(text)


class Unit:
    """
    結合した1つの単位（条見出しと第1項、折り返した文など）。
    lines は結合した行の Line、text は結合した文字列（先頭の行は字下げ込み、
    2行目以降は字下げを除いて区切り sep を挟んだもの）。
    判定は先頭の行の Line と、2行目以降をつないだ文字列（tail）で行う。
    """
    __slots__ = ('lines', 'text')

    def __init__(self, line):
        self.lines = [line]
        self.text = line.text

    def append(self, line, sep=''):
        self.lines.append(line)
        self.text += sep + line.s

    @property
    def first(self):
        return self.lines[0]

    @property
    def tail(self):
        """先頭の行より後ろの文字列（区切りを含む）。"""
        return self.text[len(self.lines[0].text):]

    @property
    def last(self):
        """結合した文字列の行末の文字。"""
        return self.lines[-1].last

    def rest(self, m, group, skip='　 \t\n'):
        """先頭の行の match の最後のグループ（行末までの部分）に、後ろの行をつないだもの。
        結合した文字列に同じ正規表現をかけたときのグループと同じになる
        （グループが空なら、その前の区切り skip の文字は飛ばす）。"""
        g = m.group(group) or ''
        return g + self.tail if g else self.tail.lstrip(skip)


def _as_unit(x):
    """文字列・Line・Unit を Unit にする（テストや単独の行からも呼べるように）。"""
    if isinstance(x, Unit):
        return x
    return Unit(x if isinstance(x, Line) else line_info(x))


def merge_lines(items):
    """
    複数の p に分割された1文を結合して返す。
    items: [('text', 文字列 or Line) or ('table', テーブルdict), ...] の順序付きリスト
    返り値: [('text', Unit) or ('table', dict), ...]
    """
    merged = []
    buf = None
    # いまの条番号と項番号（次に来るはずの条・項の番号を見るため）
    # prev_items: いまの条の、前の項までに出てきた号の (記号の種類, 番号)
    # cur_last:   いまの項で、記号の種類ごとに最後に出てきた号の番号
    # （「前項に定める書類のうち／（１）の書類は」「…上記（１）及び／（２）の」の
    #   号の参照を見分けるため）
    ctx = {'chapter': None, 'article': None, 'para': 1,
           'prev_items': set(), 'cur_items': set(), 'cur_last': {}}

    def count_unit(ln):
        if ln.chapter:
            ctx['chapter'] = zenkaku_to_int(ln.chapter.group(1))
        elif ln.article:
            # 範囲の条（「第５条～第７条　削除」）は終わりの番号
            ctx['article'] = zenkaku_to_int(ln.article.group('to') or ln.article.group('num'))
            _br = ln.article.group('branch')
            ctx['branch'] = zenkaku_to_int(_br) if _br else 0
            ctx['para'] = 1
            ctx['prev_items'], ctx['cur_items'], ctx['cur_last'] = set(), set(), {}
            count_inline(ln.s)
        elif ln.para:
            ctx['para'] = zenkaku_to_int(ln.para.group('num'))
            ctx['prev_items'] |= ctx['cur_items']
            ctx['cur_items'], ctx['cur_last'] = set(), {}
            count_inline(ln.s)
        elif ln.sym_type is not None:
            ctx['cur_items'].add((ln.sym_type, ln.sym_num))
            ctx['cur_last'][ln.sym_type] = ln.sym_num

    def count_inline(s):
        """条・項の行の途中（読点の後ろ）に書かれた号（「従業員は、（１）…に努め、」）も数える。"""
        for mm in re.finditer(r'[、，]（([０-９0-9]+)）', s):
            n = zenkaku_to_int(mm.group(1))
            ctx['cur_items'].add(('paren_zen_num', n))
            ctx['cur_last']['paren_zen_num'] = n

    def flush():
        nonlocal buf
        if buf is not None:
            merged.append(('text', buf))
            buf = None

    def next_line(i):
        """items[i] より後ろの、空でないテキストの行（Line）。表・自動番号段落が先なら None。"""
        for kind2, val2 in items[i + 1:]:
            if kind2 != 'text':
                return None
            ln2 = val2 if isinstance(val2, Line) else line_info(val2)
            if ln2.s.strip():
                return ln2
        return None

    for i, (kind, val) in enumerate(items):
        if kind in ('table', 'auto', 'drawing'):
            # 表・自動番号段落は結合対象外。それまでのテキストを確定してから追加
            flush()
            merged.append((kind, val))
            if kind == 'auto':
                count_unit(line_info(val[1]))
            continue

        ln = val if isinstance(val, Line) else line_info(val)
        if not ln.s.strip():
            continue
        if buf is None:
            buf = Unit(ln)
            count_unit(ln)
        else:
            stripped = ln.s
            head = _heading_state(buf)
            if (head == 'article_bare' and _RE_PARA1_NUM.match(stripped)
                    and not _is_short_dot_item(stripped)):
                # 見出しだけの条の次の「１　」「１．」は第1項の番号
                is_new = False
            elif (head == 'article_bare' and ln.title and not _is_note_paren(ln.title.group(1))
                  and not _split_article_title(buf.first.article.group('rest').strip())[0]
                  and len(buf.lines) == 1 and not (next_line(i) and next_line(i).article)):
                # タイトルの無い「第1条」だけの行の次の「（目的）」は、この条のタイトル
                # （次の行が条見出しなら、次の条のタイトルなので結合しない）
                is_new = False
            else:
                is_new = _starts_new_unit(ln, buf, ctx)
            if (not is_new and head in ('chapter_titled', 'section')
                    and not _chapter_title_continues(buf.text, stripped)):
                # 章・節のタイトルのあとの本文（章の前文など）はタイトルに混ぜない
                is_new = True
            if (not is_new and head == 'article_short' and not _RE_REF_START.match(stripped)
                    and not (merged and merged[-1][0] == 'text' and _unit_title(merged[-1][1]))):
                # 括弧の無い条タイトル（S2）：「第1条　目的」の次の行から本文が始まるなら、
                # 「目的」を条タイトルにする（「第1条（目的）」として結合する）。
                # 直前の行が括弧の条タイトルなら、その条はタイトルを持つので使わない
                f = buf.first
                phrase = f.article.group('rest').strip()
                head_text = f.text[:f.text.rindex(phrase)].rstrip('　 \t')
                buf = Unit(line_info(f'{head_text}（{phrase}）'))
                head = 'article_bare'
            if not is_new and head == 'chapter_bare' and ln.last == '。':
                # 「第2章」だけの行の次の文は、章タイトルではなく前文
                # （結合すると句点で終わる章タイトルになり、章と判定できなくなる）
                is_new = True
            if is_new:
                merged.append(('text', buf))
                buf = Unit(ln)
                count_unit(ln)
            else:
                # 段落内の折り返しとして結合する。見出しだけの行と本文を
                # 結合するときは、間に空白を入れる（「第1条この規則は」のように
                # 詰めると、条見出しと判定できなくなるため）
                buf.append(ln, '　' if head in ('chapter_bare', 'article_bare') else '')
    flush()
    return merged


def _heading_state(unit):
    """
    結合中の単位が章・条の見出しかを返す。
      'chapter_bare'   : 「第1章」だけ（章タイトルは次の行）
      'chapter_titled' : 「第1章　総則」
      'article_bare'   : 「第1条」「第1条（目的）」だけで本文が無い
      'section'        : 「第1節　労働時間」（節の行）
      'article_short'  : 「第1条　目的」（括弧の無い条タイトルになりうる短い語句だけ。S2）
      None             : それ以外（条見出しに本文が続く行も含む）
    """
    f = unit.first
    if f.chapter:
        return 'chapter_titled' if unit.rest(f.chapter, 2).strip() else 'chapter_bare'
    if f.section:
        return 'section'
    if f.article:
        title, rest = _split_article_title(unit.rest(f.article, 'rest').strip())
        if not rest:
            return 'article_bare'
        if not title and len(unit.lines) == 1 and _is_bare_title_phrase(rest):
            return 'article_short'
    return None


def _is_bare_title_phrase(s):
    """括弧の無い条タイトル（S2）になりうる語句か。「第1条　目的」の「目的」。
    句読点・「は・が・を」の無い15字以内で、漢字・カタカナ・括弧で終わり、「削除」でないもの。
    （「第3条　従業員は」のような文の始まりや、PDFのように25〜40字で折り返した
    本文の1行目（「第12条　職務上の地位や人間関係などの職場内の優越」）を除くため、
    仕様の20字より短くしている）"""
    return (len(s) <= 15 and not re.search(r'[、。，．,.　 はがを]', s)
            and bool(re.search(r'[一-龯々ァ-ヶー）)]$', s))
            and not _RE_DELETED.match(s))


# 次の行が助詞・読点で始まるなら、前の行の続き（括弧の無い条タイトルにしない）
_RE_REF_START = re.compile(r'^(?:の|に|を|は|が|と|で|へ|も|や|から|まで|より|及び|又は|並びに|若しくは|[、，。])')


_RE_MARK_HEAD = re.compile(
    r'^(?:[①-⑳㉑-㊿]|（[０-９0-9一二三四五六七八九十]+）|\([０-９0-9一二三四五六七八九十]+\))[　\s]*')
# 行頭の小数（「1.25倍」）
_RE_DECIMAL_HEAD = re.compile(r'^([０-９0-9]+)[．.][０-９0-9]')


def _is_short_dot_item(s):
    """「1. 履歴書」のような、短い名詞だけの号か（第1項の文ではない）。
    句読点や、文を作る助詞（は・を・が）を含めば文とみなす
    （「1. 試用期間は採用の日から3か月間とし」は第1項）。"""
    m = _RE_PARA1_NUM.match(s)
    if not m or not m.group('dot'):
        return False
    rest = s[m.end():]
    return not re.search(r'[。、はをが]', rest) and len(rest) <= 30


def _chapter_title_continues(buf, line):
    """章タイトルが行の途中で折り返されているか。
      - タイトルが「、・」や助詞・「及び」等で終わる（「第2章　育児・介護休業等に／関する事項」）
      - 次の行が、句点の無い短い語句（「第2章　労働時間、休憩／休日及び休暇」）
    次の行が括弧付きの見出し（「〔Ａ例〕…」「［例１］…」）や、空白で区切った行
    （「Ａ  週40時間労働制の規定例」）なら、章の前文とみなす。"""
    t = buf.rstrip('　 \t')
    if t and (t[-1] in '、，・にをはがでへのとや' or t.endswith(
            ('及び', '並びに', '又は', '若しくは'))):
        return True
    s = line.strip()
    return (bool(s) and len(s) <= 30 and '。' not in s
            and not re.match(r'^[〔［【＜<※・]', s)
            and not re.search(r'[ 　]{2,}', s))


def _para_num_is_continuation(buf, m, next_para=None):
    """RE_PARA に一致した行が、項の始まりではなく前の文の続きか。
    番号の並びで決め、単位の語（_QUANTITY_UNIT）は番号で決まらないときの補助に使う。
      - 前の行が文として終わっていれば項（番号が飛んでいても項とする）
      - 番号が次に来るはずの項番号（next_para）と同じなら項
        （「（２）特殊健康診断、／2 年に1回以上…」の 2 は第2項）。
        ただし前の行が助詞で終わるときは、半角空白＋単位なら数量
        （「契約期間は、／2 年以内」）
      - 番号が次の項番号と違い、前の行が文の途中で終わっていて、半角空白で区切られて
        いれば数量（PDF由来の「満／60 歳」「…のうち／5 項目」）
      - 番号が次の項番号と違い、前の行が句点や名詞（漢字・カタカナ）で終わっておらず、
        番号の後ろが単位の語なら、全角空白でも数量（「（１）…連続する／５　日以内の休暇」
        「…以下の者…／６　時間」）。前の行が名詞で終わる
        （「（２）その他会社が認めた場合／３　日以上の…」）なら項
      - 「ただし助詞で終わるとき」の助詞には、接続の語（「から」「及び」「につき」など。
        行末の読点は除いて見る）も含む"""
    num = zenkaku_to_int(m.group('num'))
    if (next_para is not None and num != next_para
            and buf.last not in '。．」）)' and not re.match(r'[一-龯々ァ-ヶー]', buf.last)
            and _QUANTITY_UNIT.match(m.group('body'))):
        return True
    if not _ends_mid_sentence(buf) or m.group('ws') != ' ':
        return False
    if next_para is not None and num == next_para:
        if not _ends_with_connective(buf.text):
            return False
        return bool(_QUANTITY_UNIT.match(m.group('body')))
    # 「2 . 前項の…」のように番号の後ろに「.」があれば項番号の書き方
    if re.search(r'[．.]', m.string[m.end('ws'):m.start('body')]):
        return False
    return True


def _has_inline_prev_marker(buf, cur_t, cur_n):
    """前の行の途中（読点・空白の後ろ）に、同じ記号で1つ前の番号の号があるか
    （「従業員は、（１）職場の整理整頓に努め、／（２）のれんや…」）。"""
    for i in range(1, len(buf)):
        if buf[i - 1] not in '、，　 ':
            continue
        try:
            t, n, _b = identify_symbol(buf[i:])
        except Exception:
            return False
        if t == cur_t and n == cur_n - 1:
            return True
    return False


def _starts_new_unit(ln, buf, ctx=None):
    """行 ln（Line）の行頭が章・条・条タイトル・第2項以降・号の記号で、
    結合中の単位 buf（Unit）とは別の、新しい段落の始まりか。
    ctx: merge_lines が数える、いまの条番号（'article'）と項番号（'para'）。
    番号が次に来るはずの条・項と同じかを、言葉の規則より先に見る。"""
    ctx = ctx or {}
    next_para = ctx.get('para', 1) + 1
    if ln.section:
        # 節（「第1節　労働時間」）。文の途中の行のあとの「第1節の規定」などは参照
        return not (_ends_mid_sentence(buf, kanji=False)
                    and (_REF_AFTER_HEADING.match(ln.section.group(2).strip())
                         or _RE_WEAK_REF_AFTER_HEADING.match(ln.section.group(2).strip())))
    if ln.chapter or ln.article:
        # 前の行が文として終わっていれば見出し
        if not _ends_mid_sentence(buf, kanji=False):
            return True
        # 文の途中で折り返した行の先頭に来た見出しの形の行は、番号の並びで決める。
        # 「…その後30日間は、／第10条（解雇）の規定にかかわらず」のように
        # 助詞が直接続く形は、番号にかかわらず参照
        nxt = _is_next_heading(ln, ctx)
        if _heading_is_reference(ln):
            # 次に来るはずの条番号の見出しは、助詞が直接続く参照でなければ見出し
            # （「…証明書、／第11条（準用）　第1項の規定は…」）
            return bool(nxt and ln.article
                        and not _heading_is_reference(ln, particle_only=True))
        # 「は・が・と・より」などが続く形（「…については、／第3章は適用しない」）は、
        # 見出しの本文が平仮名で始まる場合（「第2条（がん検診）　がん検診を」）と
        # 言葉では見分けられないので、番号が次に来るはずの章・条でなければ参照とする
        return nxt or not _RE_WEAK_REF_AFTER_HEADING.match(_heading_after(ln))
    # 「通常の賃金の／1.25倍とする」: 文の途中で小数が行頭に来た
    # 「…休憩は1時間／2.1か月単位の」のように、整数部が次の項番号なら第2項
    md = ln.decimal
    if md and (_ends_mid_sentence(buf, kanji=False)
               or (_ends_mid_sentence(buf) and zenkaku_to_int(md.group(1)) != next_para)):
        return False
    if ln.paren_only:                                            # 条タイトル
        return True
    if ln.para:                                                  # 第2項以降
        return not _para_num_is_continuation(buf, ln.para, next_para)
    # 号の記号（上のパターンに無い ア．イ．／a)／一、 なども含む）
    if ln.mark_head:
        body = ln.s[ln.mark_head.end():]
    elif ln.sym_type is not None:
        body = ln.sym_body
    else:
        return False
    # 記号の直後が助詞で、次のどちらかなら、新しい号ではなく、出てきた号を指す本文中の参照
    #  - 同じ条の前の項に、同じ番号の号がもう出ている
    #    （「前項に定める書類のうち／（１）の書類は…」）
    #  - 同じ項の中で、同じ記号の号の番号が戻っている
    #    （「（３）…上記（１）及び／（２）のいずれにも…」）
    if (ln.sym_type is not None and _REF_AFTER_MARKER.match(body.lstrip('　 \t'))
            and ((ln.sym_type, ln.sym_num) in ctx.get('prev_items', ())
                 or ln.sym_num < ctx.get('cur_last', {}).get(ln.sym_type, 0))):
        return False
    # 「…にかかわらず、／（１）に掲げる場合は」: 読点で終わった文の続きで、
    # 記号の直後が助詞なら号ではなく本文中の参照。読点で終わっていなければ、
    # 「（５）より良い職場…」のように本文が平仮名で始まる号とみなす
    if buf.last and buf.last in '、，' and _REF_AFTER_MARKER.match(body.lstrip('　 \t')):
        prev_t, prev_n = buf.first.sym_type, buf.first.sym_num
        cur_t, cur_n = ln.sym_type, ln.sym_num
        # 前の行が同じ記号の続き番号の号（「（１）履歴書、／（２）より…」）なら号
        if prev_t is not None and prev_t == cur_t and cur_n == prev_n + 1:
            return True
        # 前の行と違う記号の号は、号の細分の始まりか、親の号への戻り
        # （「（１）…に関し、／①のれん…」「②…しないこと、／（２）からかい…」）
        if prev_t is not None and prev_t != cur_t:
            return True
        # 前の行が号でなければ、第1号は号の列の始まり
        # （「職務の遂行にあたり、／（１）のれんや看板を…」）
        if prev_t is None and cur_n == 1:
            return True
        # 前の行の途中に1つ前の番号の号がある（本文の途中から号が始まる）
        if prev_t is None and _has_inline_prev_marker(buf.text, cur_t, cur_n):
            return True
        return False
    return True


# ============================================================
# STEP 2: 1行を分類・パース
# ============================================================

# 各パターン
# 「第 1 章」「第 24 条」のように番号の前後に空白が入った書き方も
# 見出しとして認める（PDF変換や手打ちの文書で多い）。
_NUM_K = r'[０-９0-9〇一二三四五六七八九十百千万]+'
RE_CHAPTER = re.compile(r'^第[　\s]*(' + _NUM_K + r')[　\s]*章[　\s]*(.*)$',
                        re.DOTALL)
# 節（章と条の間の見出し）。整形では見出しにせず、元の文字のまま段落として残す（S4）
RE_SECTION = re.compile(r'^第[　\s]*(' + _NUM_K + r')[　\s]*節(?=[　\s]|$)[　\s]*(.*)$',
                        re.DOTALL)
# 「第９条」「第９条の２」「第４条の２の２」（枝番号 = branch、2つ目以降 = branch2）、
# 「第５条～第７条　削除」（範囲の終わり = to）に対応する。
# 「第９条に定める…」のような本文中の条文参照を条見出しと
# 誤判定しないよう、条番号の直後は「空白・括弧（（【［）・コロン・行末」に限る。
RE_ARTICLE = re.compile(
    r'^第[　\s]*(?P<num>' + _NUM_K + r')[　\s]*条'
    r'(?:[　\s]*の[　\s]*(?P<branch>' + _NUM_K + r')'
    r'(?P<branch2>(?:[　\s]*の[　\s]*' + _NUM_K + r')*))?'
    r'(?:[　\s]*[～〜~][　\s]*第[　\s]*(?P<to>' + _NUM_K + r')[　\s]*条)?'
    r'(?=[　\s（(【［\[：:]|$)'
    r'[　\s]*(?P<rest>.*)$',
    re.DOTALL
)
# 括弧の中身（1段の入れ子まで。「（出向（在籍出向・転籍出向））」）
_PAREN_INNER = r'(?:[^（()）]|[（(][^（()）]*[）)])+'
# 条タイトル。「（目的）」のほか、モデル就業規則等でよくある
# 「（目的）　※〇〇パターン」のような末尾注釈も許容する。
# 全角・半角の括弧の混在（「（遵守義務)」）と、1段の入れ子も許す。
# 括弧内が句点で終わるもの（例:「（…により計算する。）」）は
# 条タイトルではなく注記なので除外する
RE_TITLE   = re.compile(r'^[（(](' + _PAREN_INNER + r')(?<![。．])[）)][\s　]*([※＊*].*)?$',
                        re.DOTALL)
# 括弧だけの行（結合の区切りに使う。RE_TITLE と違い、句点で終わる注記も含む）。
# 「（目的）　※注釈」「(目的)」「（遵守義務)」「（休職（私傷病等によるもの））」
_RE_PAREN_ONLY = re.compile(r'^[（(]' + _PAREN_INNER + r'[）)][\s　]*([※＊*].*)?$')
# 条番号の後ろの条タイトル：「（目的）」「【目的】」「：目的」
_RE_ART_TITLE_PAREN = re.compile(r'^[（(](' + _PAREN_INNER + r')[）)][　\s]*(.*)$', re.DOTALL)
_RE_ART_TITLE_BRACKET = re.compile(r'^[【［\[]([^】］\]]+)[】］\]][　\s]*(.*)$', re.DOTALL)
_RE_ART_TITLE_COLON = re.compile(r'^[：:][　\s]*([^　\s、。，．]{1,20})(?:[　\s]+(.*))?$', re.DOTALL)


def _split_article_title(rest):
    """条番号の後ろ（rest）を (条タイトル, タイトルの後ろの本文) に分ける。
    タイトルが無ければ ('', rest)。"""
    for pat in (_RE_ART_TITLE_PAREN, _RE_ART_TITLE_BRACKET, _RE_ART_TITLE_COLON):
        m = pat.match(rest)
        if m:
            return m.group(1).strip(), (m.group(2) or '').strip()
    return '', rest


def _article_number_str(m):
    """RE_ARTICLE の match から、空白を除いた条番号の文字列（「第４条の２の２」「第５条～第７条」）。"""
    return re.sub(r'[　\s]', '', m.string[:m.start('rest')])


def _article_node(m, rest):
    """RE_ARTICLE の match と、番号の後ろの文字列（rest）から条のノードを作る。
    「第1条（目的）本文…」「第2条【目的】本文」「第3条：目的　本文」のタイトルを取り出し、
    条の直後に続く「１　」（第1項の番号）を取り除く。
    範囲の条（「第５条～第７条　削除」）は range_to に終わりの番号を入れる（build_tree が展開する）。"""
    branch = m.group('branch')          # 「の２」の "２" 部分（なければ None）
    title, body = _split_article_title(rest)
    node = {'type': 'article', 'number_str': _article_number_str(m),
            'number': zenkaku_to_int(m.group('num')),
            'branch': zenkaku_to_int(branch) if branch else 0,
            'title': title, 'body': _strip_para1_number(body)}
    if m.group('to'):
        node['range_to'] = zenkaku_to_int(m.group('to'))
    return node
# 第2項以降（先頭インデント + 2以上の数字 + 区切り）。タブ区切りにも対応。
# 「１０　」以降の項も認める。
# 「３．１ヶ月および…」のように本文が数字で始まる項もあるので、ここでは
# 小数（「2.5時間」）を除かない。折り返しで行頭に来た小数は merge_lines が結合する。
RE_PARA    = re.compile(r'^[　\s]*(?P<num>(?:[２-９2-9]|[１1](?=[０-９0-9]))[０-９0-9]*)'
                        r'(?:(?P<dot>[．.])|(?P<ws>[　\s]))[　\s．.]*'
                        r'(?P<body>.*)$',
                        re.DOTALL)
# 任意の数字で始まる項（スタイルで項と分かっている場合に使う）。1も許容
RE_PARA_ANY = re.compile(r'^[　\s]*([０-９0-9]+)[　\s．.\t]+(.*)$', re.DOTALL)
RE_ITEM_MARU = re.compile(r'^[　\s]*([①-⑳㉑-㊿])[　\s\t]*(.*)$', re.DOTALL)  # ①②
RE_ITEM_PAREN = re.compile(r'^[　\s]*[（(]([０-９0-9一二三四五六七八九十]+)[）)][　\s\t]*(.*)$', re.DOTALL)  # (1)
# 号の下位（1. / 1． 形式）
RE_ITEM_SUB = re.compile(r'^[　\s]*([０-９0-9]+)[．.][　\s\t]*(.*)$', re.DOTALL)

# 丸数字を数値に変換
MARU_MAP = {chr(0x2460+i): i+1 for i in range(20)}   # ①-⑳
MARU_MAP.update({chr(0x3251+i): i+21 for i in range(15)})  # ㉑-㉟ 等

# 章・条の番号のすぐ後ろ（空白・括弧書きを除く）がこれらで始まる行は、
# 見出しではなく本文中の参照とみなす。PDF由来の文書では行が途中で
# 折り返されるので、行頭に参照が来ることがある。
#   例:「第９条に定める」「第３章の規定」「第10条（解雇）の規定」「第 9 条 に定める」
# 「が・は・へ・と」は、見出しの本文やタイトルが平仮名の語で始まる場合
# （「第2条（がん検診）　がん検診の…」「第2章　と畜場勤務者の特則」）と
# 区別できないので含めない。
_REF_AFTER_HEADING = re.compile(
    r'^(?:の(?![０-９0-9一二三四五六七八九十])|に|を|で(?!き)'
    r'|から|まで|及び|又は|並びに|若しくは|ないし|乃至'
    r'|第[　\s]*[０-９0-9一二三四五六七八九十]+[　\s]*[項号節款]|[、，。・])')


def _heading_is_reference(ln, particle_only=False):
    """章・条の見出しの形の行（Line）で、番号の後ろ（括弧書きを除く）が助詞・読点か
    （「第3章の規定」「第10条（解雇）の規定」）。見出しの本文が「第1項の規定は…」
    「のべ5日…」で始まることもあるので、前の行が文の途中で終わっている
    （行の折り返し）ときにだけ使う（merge_lines）。
    particle_only=True なら、「第1項」のような項・号の参照が続く形は数えない。"""
    if not isinstance(ln, Line):
        ln = line_info(ln)
    if not (ln.chapter or ln.article):
        return False
    hit = _REF_AFTER_HEADING.match(_heading_after(ln))
    if hit and particle_only and hit.group(0).startswith('第'):
        return False
    return bool(hit)


# 見出しの番号の後ろがこれらで始まり、番号が次に来るはずの章・条でなければ参照
# （「第3章は適用しない」「第10条（定年）は適用しない」「第5条が定める」）
_RE_WEAK_REF_AFTER_HEADING = re.compile(r'^(?:は|が|と|へ|も|より)')


def _heading_after(ln):
    """章・条の見出しの形の行（Line）の、番号の後ろ（条は括弧書きを除く）の文字列。"""
    if ln.chapter:
        return ln.chapter.group(2).strip()
    return _split_article_title(ln.article.group('rest'))[1]


def _is_next_heading(ln, ctx):
    """章・条の見出しの形の行（Line）の番号が、次に来るはずの章・条の番号か。
    条は「いまの条番号＋1」か、いまの条の枝番（「第9条」のあとの「第9条の2」）。
    章は「いまの章番号＋1」（章がまだ無ければ第1章）。"""
    if ln.chapter:
        cur = ctx.get('chapter')
        return zenkaku_to_int(ln.chapter.group(1)) == (cur or 0) + 1
    cur = ctx.get('article')
    if cur is None:
        return False
    n = zenkaku_to_int(ln.article.group('num'))
    br = ln.article.group('branch')
    # 枝番は、いまの条の枝番より大きいときだけ次の条（「第9条の2」の中の「第9条の2は」は参照）
    return n == cur + 1 or (bool(br) and n == cur and zenkaku_to_int(br) > ctx.get('branch', 0))


def _ends_mid_sentence(x, kanji=True):
    """行（文字列・Line）か結合中の単位（Unit）が、文の途中で終わっているか
    （読点・助詞で終わる。PDF由来の折り返し）。
    「…とする」「…による」のように言い切る形や、句点・括弧で終わる行は含めない。
    kanji=True なら漢字で終わる行も含める（「定年は満／60 歳」）。ただし号の行は
    「（２）健康診断書」「（２）国民の祝日及び年末年始」のように名詞で終わるのが
    ふつうなので、号の行では漢字を含めない。次の行が見出しかを見るときも
    kanji=False にする。"""
    if isinstance(x, Unit):
        c, head, text = x.last, x.first, x.text
    else:
        head = x if isinstance(x, Line) else line_info(x or '')
        c, text = head.last, head.text
    if not c:
        return False
    if c in '、，のにをはがでとへやも×＋':
        return True
    if _ends_with_connective(text):
        return True
    return kanji and bool(re.match(r'[一-龯々]', c)) and not head.is_marker


# 行末がこれらの語なら、文の途中（「採用の日から／2 か月」「第10条及び／第11条の規定」）
_MID_ENDINGS = ('から', 'より', 'まで', 'につき', 'として', 'について', 'において', 'あたり',
                '及び', '並びに', '又は', '若しくは', 'ないし', '乃至', 'または', 'および')


def _ends_with_connective(text):
    """行末（読点を除く）が助詞・接続の語（「から」「及び」「につき」など）か。"""
    t = (text or '').rstrip('　 \t').rstrip('、，')
    return bool(t) and (t[-1] in 'のにをはがでとへやも' or t.endswith(_MID_ENDINGS))


# 数字の後ろがこれらの単位なら、項番号ではなく数量（「60 歳」「1 日」「6 割」）。
# 主に半角空白で区切られているときに使う（PDF由来の文書）。全角空白でも、番号が次の
# 項番号と違い前の行が文の途中なら使う（_para_num_is_continuation）。本文が単位の字で
# 始まる項（「2 年次有給休暇は」「2 年間休日は」「2 時間外労働は」）と
# 区別するため、そうした語は除く。
# 1字の単位（年・日・時・分 など）は、後ろに漢字が続くと別の語になる
# （「日割計算」「時差出勤」「社用車」）。数量の続き方（「1 日の」「2 年以内」
# 「8 時間」「3 日前」）のときだけ数量とみなす。
_KANJI_NEXT = r'(?=[^一-龯々]|$'
_QUANTITY_UNIT = re.compile(
    r'^(?:歳|才|か月|ヶ月|ケ月|カ月|ヵ月|箇月|か年|か所|カ所|ヵ所|箇所|親等|万円|万人'
    r'|倍|円|[%％]|回(?![答収覧数線])|割(?![増引当合振])'
    r'|年' + _KANJI_NEXT + r'|間(?![休所労カ計総単])|以|未満|前|後|目|毎|分|半|程度)'
    r'|月' + _KANJI_NEXT + r'|以|未満|前|後|目|分|半|程度)'
    r'|週' + _KANJI_NEXT + r'|間(?![休所労単])|以|前|後|目)'
    r'|日' + _KANJI_NEXT + r'|間|以|未満|前|後|目|分|程度)'
    r'|時' + _KANJI_NEXT + r'|間(?![外休所労カ計単])|以|前|後|半)'
    r'|分' + _KANJI_NEXT + r'|間|以|未満|前|後|程度)'
    r'|秒|人' + _KANJI_NEXT + r'|以|未満|分|程度)'
    r'|名' + _KANJI_NEXT + r'|以|未満|程度)'
    r'|件' + _KANJI_NEXT + r'|以|未満)'
    r'|社' + _KANJI_NEXT + r'|以|未満))')

# 第1項の番号（「１　」「1.」「１．」）
_RE_PARA1_NUM = re.compile(r'^[１1](?:(?P<dot>[．.])(?![０-９0-9])|(?P<ws>[　\s]))[　\s．.]*')


def _strip_para1_number(body):
    """条見出しと同じ行の本文の先頭にある「１　」（第1項の番号）を取り除く。
    「1 週間」「1.5 時間」のような数量は取り除かない。"""
    m = _RE_PARA1_NUM.match(body)
    if not m:
        return body
    rest = body[m.end():]
    if m.group('ws') == ' ' and _QUANTITY_UNIT.match(rest):
        return body
    return rest


# 括弧だけの行のうち、条タイトルではなく注記とみなすもの
#   例:「（ただし、中途採用者を除く）」「（令和６年４月１日一部改定）」
# 「（パートタイマーを除く）」のような行は、条タイトル（「（管理監督者を除く）」）と
# 見分けられないので、ここでは注記にしない。タイトルとして保留し、次に条が
# 来なければ build_tree が本文に戻す。
# 日付は、すぐ後に改定・施行などが続くとき（「（令和６年４月１日一部改定）」）だけ注記とする。
# 「（令和6年4月1日以後に採用された者の特例）」のような日付入りの条タイトルは注記にしない。
_D = r'[元０-９0-9一二三四五六七八九十]+'
_RE_NOTE_PAREN = re.compile(
    r'(?:(?:令和|平成|昭和)[　\s]*' + _D + r'|[０-９0-9]{4})[　\s]*年'
    r'(?:[　\s]*' + _D + r'[　\s]*月)?(?:[　\s]*' + _D + r'[　\s]*日)?'
    r'[　\s]*(?:から|より|付け?)?[　\s]*(?:一部|全部|全面)?[　\s]*'
    r'(?:改定|改正|改訂|施行|制定|追加|削除|新設|変更|実施|適用)'
    r'|^ただし')


def _is_note_paren(inner):
    return bool(_RE_NOTE_PAREN.search(inner.strip()))


# 停止ワード（附則など）の直後がこれなら、見出しではなく本文中の参照
#   例:「附則に定める経過措置」「附則第2項の規定により」
#   「附則 に定める」（PDF由来の空白）、「附則第1条（施行期日）の規定」も参照
#   「附則第1条（施行期日）第2項の規定」のように条・項が続く参照も同じ
_STOP_WORD_REF = re.compile(
    r'^[　\s]*(?:第[　\s]*[０-９0-9一二三四五六七八九十]+[　\s]*[条項号]'
    r'(?:の[０-９0-9一二三四五六七八九十]+)?(?:[（(][^）)]*[）)])?[　\s]*)*'
    r'(?:の|に|を|で|は|が|と|へ|から|まで|より|及び|又は|並びに|若しくは|[、，。])')

def zenkaku_to_int(s):
    """全角数字・漢数字を含む文字列を整数に変換"""
    s = s.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    try:
        return int(s)
    except ValueError:
        return kanji_to_int(s) or 1

def classify_line(text, depth=None, number=None):
    """
    1つの単位（merge_lines の Unit、または1行の文字列）を分類してdictを返す。
    先頭の行の判定結果（Line）と、2行目以降をつないだ文字列から作る。
    depth: 号の記号の深さ（marker_depths が項ごとの入れ子で決めたもの。
           0=号、1=号の細分、2以上=さらに下）。None なら号とする。
    number: 号の番号の上書き（kana_numbers が決めたカタカナの号の番号）。None なら記号の番号
    keys: type, number, title/body
    """
    unit = _as_unit(text)
    f = unit.first
    stripped = unit.text.lstrip('　 \t')

    if not stripped:
        return None

    # 章
    m = f.chapter
    if m:
        return {'type':'chapter', 'number_str':f'第{m.group(1)}章',
                'number': zenkaku_to_int(m.group(1)), 'title': unit.rest(m, 2).strip()}

    # 節（「第1節　労働時間」）。見出しにせず、元の文字のまま残す（S4）
    if f.section:
        return {'type': 'section', 'body': stripped}

    # 条タイトル（括弧のみの行）。「（中途採用者を除く）」のような注記は本文
    tt = _unit_title(unit)
    if tt:
        if _is_note_paren(tt[0]):
            return {'type':'text', 'body': stripped, 'note': True}
        # _raw: 次が条見出しでなかったときに、本文として戻す元の行（parse_kisoku）
        return {'type':'article_title', 'title': tt[0].strip(),
                'title_note': (tt[1] or '').strip(), '_raw': stripped}

    # 条（+本文が同行にある場合も）。枝番号「第X条のY」にも対応
    m = f.article
    if m:
        return _article_node(m, unit.rest(m, 'rest').strip())

    # 第2項以降（先頭インデント+数字）
    # raw は、build_tree が項ではないと判断したときに本文として戻すための元の行
    m = f.para  # raw（インデント込み）に対してマッチしたもの
    if m:
        num = zenkaku_to_int(m.group('num'))
        if num >= 2:  # 2以上のみ項扱い（1は条本文の続きとして扱わない）
            return {'type':'paragraph', 'number': num,
                    'body': unit.rest(m, 'body', skip='　 \t\n．.').strip(),
                    'sep': 'dot' if m.group('dot') else 'ws', 'raw': stripped}

    # --- 号系記号の判定（深さは marker_depths が項ごとの入れ子で決めたもの）---
    if f.sym_type is not None:
        t = ('item' if not depth else 'item_sub' if depth == 1 else 'item_sub2')
        return {'type': t, 'number': f.sym_num if number is None else number,
                'body': (f.sym_body + unit.tail).strip(),
                'marker': f.sym_type, 'raw': stripped}

    # その他（本文の続き、表など）
    # 括弧だけの行（「（ただし、…を除く。）」のように句点で終わり条タイトルに
    # ならないもの）は注記として、直前のいちばん深い項目に付ける
    if len(stripped) >= 2 and stripped[0] in '（(' and stripped[-1] in '）)':
        return {'type':'text', 'body': stripped, 'note': True}
    return {'type':'text', 'body': stripped}


def marker_depths(merged):
    """
    号の記号（（１）①ア 等）の深さ（0=号、1=号の細分、2=さらに下）を、
    merge_lines の結果の単位ごとに決めて {id(Unit): 深さ} で返す。

    就業規則では、条や項ごとに記号の使い方が違うことがある
    （第1条は（１）の下に①、第2条は①を号に直接使う）。文書全体で記号ごとに
    1つの深さを決めると崩れるので、項ごとに、上から読んだ入れ子で決める。
      - 章・節・条・条タイトル・第2項以降の行（スタイルで分かるものを含む）で入れ子を戻す
      - その項で初めて出た記号は、いま開いている記号の一段下（項の最初の記号は号）
      - 開いている記号に戻ったら、それより深い記号を閉じる
      - 記号の無い行（号の折り返し・注記）では入れ子を戻さない
    """
    depths = {}
    stack = []
    for k, val in merged:
        if k == 'auto':
            if val[0] in ('chapter', 'article', 'paragraph', 'paragraph_1', 'body_plain'):
                stack = []
            continue
        if k != 'text':
            continue
        f = val.first
        if (f.chapter or f.article or f.section or _unit_title(val)
                or (f.para and zenkaku_to_int(f.para.group('num')) >= 2)):
            stack = []
            continue
        if f.sym_type is None:
            continue
        if f.sym_type in stack:
            del stack[stack.index(f.sym_type) + 1:]
        else:
            stack.append(f.sym_type)
        depths[id(val)] = stack.index(f.sym_type)
    return depths


_KANA_TYPES = ('katakana', 'paren_zen_kana', 'paren_han_kana')
_GOJUON = ('アイウエオカキクケコサシスセソタチツテトナニヌネノ'
           'ハヒフヘホマミムメモヤユヨラリルレロワヲン')
_IROHA = 'イロハニホヘトチリヌルヲワカヨタレソツネナラム'


def kana_numbers(merged):
    """
    カタカナの号（ア、（イ） 等）の番号を、{id(Unit): 番号} で返す。
    アイウエオ順かイロハ順かは1文字では決まらない（「イ」はアイウエオ順の2番、
    イロハ順の1番）ので、marker_depths と同じ歩き方で見た、同じ記号の種類の
    連なりの最初の文字が「イ」ならイロハ順、それ以外はアイウエオ順で数える。
    """
    nums = {}
    stack = []   # [記号の種類, 連なりの最初の文字]
    for k, val in merged:
        if k == 'auto':
            if val[0] in ('chapter', 'article', 'paragraph', 'paragraph_1', 'body_plain'):
                stack = []
            continue
        if k != 'text':
            continue
        f = val.first
        if (f.chapter or f.article or f.section or _unit_title(val)
                or (f.para and zenkaku_to_int(f.para.group('num')) >= 2)):
            stack = []
            continue
        if f.sym_type is None:
            continue
        types = [e[0] for e in stack]
        if f.sym_type in types:
            idx = types.index(f.sym_type)
            del stack[idx + 1:]
        else:
            stack.append([f.sym_type, f.sym_char])
            idx = len(stack) - 1
        if f.sym_type in _KANA_TYPES:
            order = _IROHA if stack[idx][1] == 'イ' else _GOJUON
            if f.sym_char in order:
                nums[id(val)] = order.index(f.sym_char) + 1
    return nums


# 条タイトルの括弧の後ろ（空白と「※注釈」だけ）
_RE_TITLE_AFTER = re.compile(r'^[）)][\s　]*([※＊*].*)?$', re.DOTALL)


def _unit_title(unit):
    """単位が括弧だけの行（条タイトル）なら (括弧の中, 注釈 or None)、それ以外は None。
    2行目以降があるときは、括弧の後ろにつないだ文字列が空白と「※注釈」だけのときに限る
    （結合した文字列に RE_TITLE をかけたときと同じ結果）。"""
    m = unit.first.title
    if not m:
        return None
    if len(unit.lines) == 1:
        return m.group(1), m.group(2)
    mm = _RE_TITLE_AFTER.match(unit.first.s[m.end(1):] + unit.tail)
    return (m.group(1), mm.group(1)) if mm else None


def classify_auto(kind, text):
    """
    自動番号（Wordのアウトライン）が設定されている段落を分類する。
    ただしテキストに手打ちの章・条番号がある場合はそちらを優先する
    （スタイルは項なのに本文に「第12条」と手打ちされているケースがあるため）。
    """
    body = text.lstrip('　 \t').strip()
    if not body:
        return None
    ln = line_info(body)

    # --- スタイル/自動番号で「号」「項」と確定している場合は、
    #     テキストが「第９条に定める…」のような条文参照で始まっていても
    #     条見出しとは扱わない（誤判定防止）---
    if kind in ('item', 'item_sub', 'paragraph', 'paragraph_1', 'body_plain', 'list_number'):
        pass   # 下の種別ごとの処理へ進む
    else:
        # --- 章・条が期待される文脈でのみ、テキストの手打ち番号を見る ---
        m = ln.chapter
        if m:
            return {'type':'chapter', 'number_str': f'第{m.group(1)}章',
                    'number': zenkaku_to_int(m.group(1)), 'title': m.group(2).strip()}
        m = ln.article
        if m:
            return _article_node(m, m.group('rest').strip())

    # --- スタイル/自動番号で「条」と分かっている場合 ---
    # 「(目的)」のように括弧だけの行は、条見出し（番号は自動採番）
    if kind == 'article':
        tm = ln.title
        if tm:
            return {'type':'article', 'number_str':'', 'number': 0, 'branch': 0,
                    'title': tm.group(1).strip(),
                    'title_note': (tm.group(2) or '').strip(), 'body': ''}
        return {'type':'article', 'number_str':'', 'number':0, 'branch':0,
                'title':'', 'body': body}

    # 括弧だけの行は条タイトル（スタイル情報が無い場合）
    if kind not in ('item', 'item_sub', 'paragraph', 'paragraph_1', 'body_plain', 'list_number'):
        m = ln.title
        if m:
            return {'type':'article_title', 'title': m.group(1).strip(),
                'title_note': (m.group(2) or '').strip()}

    if kind == 'chapter':
        return {'type':'chapter', 'number_str':'', 'number': 0, 'title': body}

    if kind in ('paragraph', 'paragraph_1'):
        # スタイルで項と分かっているので、1始まりの番号も剥がす
        m = RE_PARA_ANY.match(text)
        if m:
            return {'type':'paragraph', 'number': zenkaku_to_int(m.group(1)),
                    'body': m.group(2).strip()}
        # 番号なし = 自動採番（build_tree側で連番付与）
        return {'type':'paragraph', 'number': 0, 'body': body}

    if kind == 'list_number':
        # Word の段落番号。項か号かは、直前の項の予告文を見て build_tree が決める
        return {'type':'paragraph', 'number': 0, 'body': body, '_list': True}

    if kind == 'item':
        m = RE_ITEM_MARU.match(text)
        if m:
            return {'type':'item', 'number': MARU_MAP.get(m.group(1), 0),
                    'body': m.group(2).strip()}
        m = RE_ITEM_PAREN.match(text)
        if m:
            return {'type':'item', 'number': zenkaku_to_int(m.group(1)),
                    'body': m.group(2).strip()}
        # 番号なし = 自動採番
        return {'type':'item', 'number': 0, 'body': body}

    if kind == 'item_sub2':
        # 号の下位のさらに下（(ア)(イ) 等）
        m = RE_ITEM_PAREN.match(text)
        if m:
            return {'type':'item_sub2', 'number': zenkaku_to_int(m.group(1)),
                    'body': m.group(2).strip()}
        if ln.sym_type:
            return {'type':'item_sub2', 'number': ln.sym_num, 'body': ln.sym_body.strip()}
        return {'type':'item_sub2', 'number': 0, 'body': body}

    if kind == 'item_sub':
        # 号の下位（1. / 1． / ① 等）
        m = RE_ITEM_SUB.match(text)
        if m:
            return {'type':'item_sub', 'number': zenkaku_to_int(m.group(1)),
                    'body': m.group(2).strip()}
        m = RE_ITEM_MARU.match(text)
        if m:
            return {'type':'item_sub', 'number': MARU_MAP.get(m.group(1), 0),
                    'body': m.group(2).strip()}
        return {'type':'item_sub', 'number': 0, 'body': body}

    if kind == 'body_plain':
        # 番号なしの本文（00-）→ 第1項として扱う
        return {'type':'paragraph', 'number': 0, 'body': body}

    return {'type':'text', 'body': body}


# ============================================================
# STEP 3: フラットなdictリストをネスト構造に組み立て
# ============================================================

def _deepest_container(cur_paragraph):
    """
    表・図・数式をぶら下げる先として、現在いちばん深い項目を返す。
    号 → 号の下位 → さらに下 の順に辿る。
    （号の下位に続く数式が、号の末尾にまとめられるのを防ぐ）
    """
    if cur_paragraph is None:
        return None
    items = cur_paragraph.get('items')
    if not items:
        return cur_paragraph
    node = items[-1]
    subs = node.get('sub_items')
    if subs:
        node = subs[-1]
        subs2 = node.get('sub_items2')
        if subs2:
            node = subs2[-1]
    return node


_RE_DELETED = re.compile(r'^[（(]?[　\s]*削[　\s]*除[　\s]*[）)]?[。．]?$')


def _expand_article_ranges(flat):
    """「第５条～第７条　削除」を「第５条　削除」「第６条　削除」「第７条　削除」に分ける。
    整形では条番号を自動採番するので、1つの条のままだと後ろの条番号がずれるため。
    本文が「削除」でない範囲の条は、分けずに1つの条のままにする。"""
    out = []
    for node in flat:
        to = node.get('range_to')
        if (node.get('type') != 'article' or not to or not node['number']
                or not _RE_DELETED.match(node.get('body', '').strip())
                or not (node['number'] < to <= node['number'] + 100)):
            out.append(node)
            continue
        full = bool(re.search(r'[０-９]', node['number_str']))
        for n in range(node['number'], to + 1):
            s = str(n).translate(str.maketrans('0123456789', '０１２３４５６７８９')) if full else str(n)
            out.append(dict(node, number=n, number_str=f'第{s}条', branch=0,
                            title=node['title'] if n == node['number'] else '',
                            range_to=None))
    for node in out:
        node.pop('range_to', None)
    return out



def _mk(node):
    """元の行頭の記号（「２　」「（１）」「①」など）を木の要素に引き継ぐ（新旧対照表の現行の表示用）。"""
    m = node.get('mark') if isinstance(node, dict) else None
    return {'mark': m} if m else {}


def _lead_mark(text, body):
    """
    行の先頭から本文の直前までを、元の行頭の記号として返す（無ければ ''）。
    例：「（１）住民票」「（１）」、「２　前項の…」「２　」、「①　履歴書」「①　」
    """
    full = (text or '').strip()
    b = (body or '').strip()
    if not full or not b:
        return ''
    # まず「行の末尾が本文と同じ」なら、その前を記号とする（本文が短くても記号の中に当たらない）
    if full.endswith(b) and 0 < len(full) - len(b) <= 12:
        return full[:len(full) - len(b)]
    i = full.find(b[:12])
    # 記号は短い。本文が見つからない・遠すぎるときは記号とみなさない
    return full[:i] if 0 < i <= 12 else ''


def build_tree(flat):
    """
    分類済みのフラットなdictリストを、
    章 > 条 > 項 > 号 のネスト構造に変換する。
    """
    flat = _expand_article_ranges(flat)
    result = []
    cur_chapter = None
    cur_article = None
    cur_paragraph = None
    pending_title = None   # 条タイトル（article_titleが先行する場合）
    pending_note = ''      # 条タイトル末尾の注釈（※〜）
    pending_lead = []      # 次の条の前に置く節の行（「第1節　労働時間」）

    last_dot = None        # 直前の「1.」形式の号 {'type', 'number'}（「2.」を号と判断するため）

    def _orphan(text):
        """入れ先の条が無い本文を、消さずに章の前文（番号なし）として残す。"""
        if not text:
            return
        if cur_chapter is not None:
            cur_chapter.setdefault('preamble', []).append(text)
        else:
            result.append({'type': 'plain', 'body': text})

    def _new_paragraph1(body=''):
        """条に第1項が無いとき、第1項を作って返す。"""
        p = {'type': 'paragraph', 'number': 1, 'body': body, 'items': []}
        cur_article['paragraphs'].append(p)
        return p

    def _flush_pending():
        """条が来ないまま残った条タイトルを、本文として戻す（欠落防止）。
        戻し先は直前のいちばん深い項目（号の補足なら、その号の末尾）。"""
        nonlocal pending_title, pending_note, cur_paragraph
        if pending_title:
            txt = f'（{pending_title}）'
            if pending_note:
                txt += f'　{pending_note}'
            if cur_paragraph is not None:
                _deepest_container(cur_paragraph)['body'] += txt
            elif cur_article is not None:
                if cur_article['paragraphs']:
                    _deepest_container(cur_article['paragraphs'][-1])['body'] += txt
                else:
                    cur_paragraph = _new_paragraph1(txt)
            else:
                _orphan(txt)
        pending_title = None
        pending_note = ''
    article_seq = [0]      # 自動採番の条番号カウンタ（章をまたいで連番）

    for node in flat:
        t = node['type']

        if t == 'paragraph' and node.get('raw'):
            if (node.get('sep') == 'dot' and last_dot
                    and node['number'] == last_dot['number'] + 1):
                # 号を「1.」「2.」で書いている → 「2.」も項ではなく号
                node = dict(node, type=last_dot['type'], marker='num_dot')
                t = node['type']
            elif (node['number'] >= 10 and cur_article is not None
                    and len(cur_article['paragraphs']) < 9):
                # 第9項までが無いのに「１０　」→ 項番号ではなく本文の続き
                node = {'type': 'text', 'body': node['raw']}
                t = 'text'
        if (t == 'paragraph' and node.get('_list') and cur_paragraph is not None
                and _has_forecast(cur_paragraph['body'])):
            # Word の段落番号の段落で、直前の項が「次の書類を…」と予告している → 号
            node = {'type': 'item', 'number': 0, 'body': node['body'], **_mk(node)}
            t = 'item'

        if t == 'section':
            _flush_pending()
            pending_lead.append(node['body'])
            continue
        if t == 'chapter' and pending_lead:
            # 条が来ないまま章が変わった節の行は、その位置に段落として残す
            result.extend({'type': 'plain', 'body': x} for x in pending_lead)
            pending_lead = []

        if t == 'chapter':
            _flush_pending()
            n = node['number']
            if not n:   # 0 = 自動採番 → 連番
                n = sum(1 for x in result if x.get('type') == 'chapter') + 1
            cur_chapter = {
                'type': 'chapter',
                'number': n,
                'number_str': node['number_str'] or f'第{n}章',
                'title': node['title'],
                'articles': []
            }
            result.append(cur_chapter)
            cur_article = None
            cur_paragraph = None
            last_dot = None

        elif t == 'article_title':
            # 次の条に付けるタイトルとして保留。前の保留が残っていれば、
            # それはタイトルではなかったので本文に戻す
            _flush_pending()
            pending_title = node['title']
            pending_note = node.get('title_note', '')

        elif t == 'article':
            if node.get('title'):
                # 条が自分のタイトルを持っている → 保留中の括弧行は前の条の本文
                _flush_pending()
            last_dot = None
            n = node['number']
            if not n:   # 0 = 自動採番 → 全体通し番号
                n = article_seq[0] + 1
            # 手打ち番号の場合もカウンタを同期させる
            # （手打ちと自動採番が混在しても番号が飛ばない／重複しない）
            article_seq[0] = max(article_seq[0], n)
            cur_article = {
                'type': 'article',
                'number': n,
                'branch': node.get('branch', 0),   # 枝番号（0=枝番なし）
                'number_str': node['number_str'] or f'第{n}条',
                'title': node.get('title') or pending_title or '',
                'title_note': node.get('title_note') or pending_note or '',
                'paragraphs': []
            }
            if pending_lead:
                # 条の前に置く節の行（整形では条見出しの前に、番号の付かない段落で出す）
                cur_article['lead'] = pending_lead
                pending_lead = []
            pending_title = None
            pending_note = ''      # 注釈はこの条だけのもの（次の条に引き継がない）
            # 第1項（本文が同行）
            if node['body']:
                cur_paragraph = {
                    'type': 'paragraph',
                    'number': 1,
                    'body': node['body'], **_mk(node),
                    'items': []
                }
                cur_article['paragraphs'].append(cur_paragraph)
            else:
                cur_paragraph = None
            # 章がなければトップレベルに追加
            if cur_chapter is not None:
                cur_chapter['articles'].append(cur_article)
            else:
                result.append(cur_article)

        elif t == 'paragraph':
            _flush_pending()
            last_dot = None
            if cur_article is None:
                # 条の外に項が来るケース（稀）→ 消さずに本文として残す
                _orphan(node.get('raw') or node['body'])
            else:
                num = node['number']
                if not num:   # 0 = 自動採番 → 連番を割り当てる
                    num = len(cur_article['paragraphs']) + 1
                cur_paragraph = {
                    'type': 'paragraph',
                    'number': num,
                    'body': node['body'], **_mk(node),
                    'items': []
                }
                cur_article['paragraphs'].append(cur_paragraph)

        elif t == 'item':
            _flush_pending()
            if cur_paragraph is not None:
                num = node['number']
                if not num:   # 0 = 自動採番 → 連番を割り当てる
                    num = len(cur_paragraph['items']) + 1
                cur_paragraph['items'].append({
                    'type': 'item',
                    'number': num,
                    'body': node['body'], **_mk(node),
                    '_marker': node.get('marker'),   # 記号の種類（_maru_items_to_paragraphs 用）
                })
            elif cur_article is not None:
                # 項のない条の直下に号が来る場合 → 第1項を自動生成
                if not cur_article['paragraphs']:
                    cur_paragraph = {
                        'type': 'paragraph', 'number': 1,
                        'body': '', 'items': []
                    }
                    cur_article['paragraphs'].append(cur_paragraph)
                else:
                    cur_paragraph = cur_article['paragraphs'][-1]
                num = node['number']
                if not num:
                    num = len(cur_paragraph['items']) + 1
                cur_paragraph['items'].append({
                    'type': 'item',
                    'number': num,
                    'body': node['body'], **_mk(node),
                    '_marker': node.get('marker'),   # 記号の種類（_maru_items_to_paragraphs 用）
                })
            else:
                _orphan(node.get('raw') or node['body'])
            if node.get('marker') == 'num_dot' and node['number']:
                last_dot = {'type': 'item', 'number': node['number']}
            else:
                # 別の記号の号が来たら、「1.」形式の号（細分を含む）の続きは終わり
                last_dot = None

        elif t == 'item_sub':
            _flush_pending()
            if node.get('marker') == 'num_dot' and node['number']:
                last_dot = {'type': 'item_sub', 'number': node['number']}
            if cur_paragraph is None and cur_article is not None:
                # 本文の無い条の直下に来た → 第1項を作り、下の「号に格上げ」へ
                cur_paragraph = _new_paragraph1()
            # 号の下位項目 → 直前の号にぶら下げる
            # ただし、直前の号が「行き場が無くて号に格上げしたもの」なら、
            # 続く項目も同じ号レベルの兄弟とみなす
            # （①②③ が号として使われている条で、②③ が①の下位に
            #   なってしまうのを防ぐ）
            if (cur_paragraph is not None and cur_paragraph.get('items')
                    and cur_paragraph['items'][-1].get('_promoted')):
                num = node['number'] or (len(cur_paragraph['items']) + 1)
                cur_paragraph['items'].append({
                    'type': 'item', 'number': num,
                    'body': node['body'], **_mk(node), '_promoted': True
                })
            elif cur_paragraph is not None and cur_paragraph.get('items'):
                parent = cur_paragraph['items'][-1]
                num = node['number']
                if not num:
                    num = len(parent.get('sub_items', [])) + 1
                parent.setdefault('sub_items', []).append({
                    'type': 'item_sub',
                    'number': num,
                    'body': node['body'], **_mk(node)
                })
            elif cur_paragraph is not None:
                # 号が無いのに下位項目が来た場合は、本文に混ぜてしまうと
                # 内容が埋もれるため、号に格上げして独立させる
                num = node['number'] or (len(cur_paragraph.get('items', [])) + 1)
                cur_paragraph.setdefault('items', []).append({
                    'type': 'item', 'number': num, 'body': node['body'], **_mk(node),
                    '_promoted': True
                })
            else:
                _orphan(node.get('raw') or node['body'])

        elif t == 'item_sub2':
            _flush_pending()
            # 号の下位のさらに下 → 直前の「号の下位」にぶら下げる
            parent = None
            if cur_paragraph is not None and cur_paragraph.get('items'):
                last_it = cur_paragraph['items'][-1]
                if last_it.get('sub_items'):
                    parent = last_it['sub_items'][-1]
            if parent is not None:
                num = node['number'] or (len(parent.get('sub_items2', [])) + 1)
                parent.setdefault('sub_items2', []).append({
                    'type': 'item_sub2', 'number': num, 'body': node['body'], **_mk(node)
                })
            elif cur_paragraph is not None and cur_paragraph.get('items'):
                # 受け皿が無ければ号の下位に格上げして残す
                it = cur_paragraph['items'][-1]
                num = node['number'] or (len(it.get('sub_items', [])) + 1)
                it.setdefault('sub_items', []).append({
                    'type': 'item_sub', 'number': num, 'body': node['body'], **_mk(node)
                })
            elif cur_paragraph is not None or cur_article is not None:
                # 号も無ければ号に格上げして残す（消さない）
                if cur_paragraph is None:
                    cur_paragraph = _new_paragraph1()
                num = node['number'] or (len(cur_paragraph['items']) + 1)
                cur_paragraph['items'].append({
                    'type': 'item', 'number': num, 'body': node['body'], **_mk(node),
                    '_promoted': True
                })
            else:
                _orphan(node.get('raw') or node['body'])

        elif t == 'drawing':
            _flush_pending()
            # 図形・画像は直前の項（なければ条・章）にぶら下げる
            dw = {'type': 'drawing', 'index': node.get('index'),
                  'text': node.get('text', '')}
            if cur_paragraph is not None:
                # いちばん深い項目にぶら下げる（元の位置を保つため）
                _deepest_container(cur_paragraph).setdefault(
                    'drawings', []).append(dw)
            elif cur_article is not None:
                if not cur_article['paragraphs']:
                    cur_paragraph = {'type':'paragraph','number':1,'body':'',
                                     'items':[],'drawings':[]}
                    cur_article['paragraphs'].append(cur_paragraph)
                else:
                    cur_paragraph = cur_article['paragraphs'][-1]
                cur_paragraph.setdefault('drawings', []).append(dw)
            elif cur_chapter is not None:
                cur_chapter.setdefault('drawings', []).append(dw)
            else:
                result.append(dw)

        elif t == 'table':
            _flush_pending()
            # 表は直前の項（なければ条・章）にぶら下げる
            tbl = {'type': 'table', 'rows': node['rows'],
                   'index': node.get('index')}
            if cur_paragraph is not None:
                # いちばん深い項目にぶら下げる（元の位置を保つため）
                _deepest_container(cur_paragraph).setdefault(
                    'tables', []).append(tbl)
            elif cur_article is not None:
                if not cur_article['paragraphs']:
                    cur_paragraph = {'type':'paragraph','number':1,'body':'',
                                     'items':[],'tables':[]}
                    cur_article['paragraphs'].append(cur_paragraph)
                else:
                    cur_paragraph = cur_article['paragraphs'][-1]
                cur_paragraph.setdefault('tables', []).append(tbl)
            elif cur_chapter is not None:
                cur_chapter.setdefault('tables', []).append(tbl)
            else:
                result.append(tbl)

        elif t == 'text':
            _flush_pending()
            # 本文の続き行 → 直前の項目に結合
            body = node['body']
            if node.get('note') and cur_paragraph:
                # 括弧書きの注記は、直前のいちばん深い項目（号の細分など）の末尾に付ける
                _deepest_container(cur_paragraph)['body'] += body
            elif cur_paragraph and cur_paragraph.get('items'):
                # 直前が号 → いちばん深い項目（号の細分の続きなら、その細分）の本文に追記
                _deepest_container(cur_paragraph)['body'] += body
            elif cur_paragraph:
                cur_paragraph['body'] += body
            elif cur_article:
                # 条の本文が別行にある場合
                if cur_article['paragraphs']:
                    cur_article['paragraphs'][-1]['body'] += body
                else:
                    cur_paragraph = {'type':'paragraph','number':1,'body':body,'items':[]}
                    cur_article['paragraphs'].append(cur_paragraph)
            else:
                # 章の直後など、入れ先の条が無い本文 → 消さずに残す
                _orphan(body)

    _flush_pending()   # 最後まで条が来なかった括弧行も本文に戻す
    result.extend({'type': 'plain', 'body': x} for x in pending_lead)
    for node in result:
        for a in (node.get('articles', []) if node.get('type') == 'chapter'
                  else [node] if node.get('type') == 'article' else []):
            _maru_items_to_paragraphs(a)
            for p in a['paragraphs']:
                for it in p.get('items', []):
                    it.pop('_marker', None)
    return result


# 句点で終わる文か（後ろに括弧書きの注記「（ただし、…を除く）」が付いていてもよい）
_RE_ENDS_SENTENCE = re.compile(r'。[　 \t]*(?:[（(]' + _PAREN_INNER + r'[）)][　 \t]*)*$')


def _maru_items_to_paragraphs(article):
    """
    柱書の無い条で、①②…を項として書いた規則（S1）を項に直す。
      第5条（服務）
      ①　従業員は、誠実に職務を遂行しなければならない。
      ②　従業員は、会社の名誉を害してはならない。
    条見出しの直後に①が来て（第1項の本文が空）、号がすべて丸数字の連番で、
    どれも句点で終わる文のときだけ項とする。①の下の（１）などは、その項の号になる。
    柱書のあとの①（「次の事項を守る。／①…」）は号のまま。
    """
    paras = article['paragraphs']
    if len(paras) != 1 or paras[0]['body'].strip():
        return
    items = paras[0].get('items', [])
    if (len(items) < 2 or any(it.get('_marker') != 'maru' for it in items)
            or [it['number'] for it in items] != list(range(1, len(items) + 1))
            or not all(_RE_ENDS_SENTENCE.search(it['body']) for it in items)
            or paras[0].get('tables') or paras[0].get('drawings')):
        return
    new = []
    for it in items:
        p = {'type': 'paragraph', 'number': it['number'], 'body': it['body'], 'items': [], **_mk(it)}
        for sub in it.get('sub_items', []):
            ni = {k: v for k, v in sub.items() if k != 'sub_items2'}
            ni['type'] = 'item'
            if sub.get('sub_items2'):
                ni['sub_items'] = [dict(s2, type='item_sub') for s2 in sub['sub_items2']]
            p['items'].append(ni)
        for key in ('tables', 'drawings'):
            if it.get(key):
                p[key] = it[key]
        new.append(p)
    article['paragraphs'] = new


# ============================================================
# STEP 4: メイン処理
# ============================================================

# 予告文パターン（「次の各号」「以下のとおり」等）
# これらがあるのに下位項目が無い/少ない場合、枝が親に押し込まれている疑い
_YOKOKU_NOUNS = r'(各号|いずれ|とおり|事項|書類|もの|区分|基準|事由|場合|理由|条件|方法|種類|範囲|項目)'
RE_YOKOKU = re.compile(
    r'次の' + _YOKOKU_NOUNS
    + r'|次に(掲げる|定める|示す|よる|該当)'
    + r'|以下の' + _YOKOKU_NOUNS
    + r'|以下に(掲げる|定める|示す|該当)'
    + r'|下記の' + _YOKOKU_NOUNS
    + r'|左記の' + _YOKOKU_NOUNS
)


def _has_forecast(text):
    """予告文（次の各号・以下のとおり 等）を含むか。"""
    return bool(RE_YOKOKU.search(text or ''))


def _find_split_markers(body):
    """本文中に「分割すべき痕跡」があるか検出する。先頭の番号は除外。"""
    markers = []
    inner = body[2:] if len(body) > 2 else body
    # 句点・括弧の後にタブ付き数字（項番号の押し込み）
    if re.search(r'[。）」]\s*[０-９0-9]+[\t　]', inner):
        markers.append('項番号痕跡')
    # 本文中の丸数字（号の押し込み）
    if re.search(r'[①-⑳]', inner):
        markers.append('丸数字痕跡')
    # 「〜とき」が3回以上（列挙の押し込み）
    if len(re.findall(r'とき[。、）]', inner)) >= 2:
        markers.append('とき連続')
    # 「〜こと」「〜もの」が3回以上
    if len(re.findall(r'(こと|もの)[。、）]', inner)) >= 3:
        markers.append('名詞句連続')
    # 潰れた下位項目の痕跡（表組み的な繰り返し）
    #   「……」「・・・」などのリーダー記号が複数（一覧が潰れている）
    if len(re.findall(r'[…・]{2,}|[‥]{1,}', inner)) >= 2:
        markers.append('一覧潰れ痕跡')
    #   「〜まで」「〜から」の繰り返し（範囲列挙が潰れている）
    if len(re.findall(r'(まで|から)[　\s]', inner)) >= 2:
        markers.append('範囲列挙痕跡')
    #   全角コロン「：」が複数（項目：値 の一覧が潰れている）
    if inner.count('：') >= 2:
        markers.append('項目値一覧痕跡')
    #   全角スペースが2個以上連続する箇所が複数（列揃えの一覧が潰れている）
    #   例:「…出頭する場合　　その日…出頭する場合　　　　その期間」
    if len(re.findall(r'\u3000{2,}', inner)) >= 2:
        markers.append('列揃え痕跡')
    return markers


def find_suspicious_units(tree, length_threshold=100, loose=True):
    """
    分割ミスが疑われる項・号を抽出する。

    loose=True（既定）: 検出を緩くする。「AIに広く見せて判断させる」方針。
      - 予告文（次の各号・以下のとおり 等）があれば、下位項目の有無に
        関わらず候補にする（押し込みは下位が1個でも起こりうるため）。
      - 文字数が閾値以上 かつ 分割痕跡があるものも候補にする。
    loose=False: 予告文があっても下位項目が無い場合のみ候補にする（厳しめ）。

    返り値: [{'article','kind','number','body','reason','ref','container'}...]
    """
    suspects = []
    seen = set()

    def add(article, kind, number, body, reason, ref, container):
        if id(ref) in seen:
            return
        seen.add(id(ref))
        suspects.append({'article': article, 'kind': kind, 'number': number,
                         'body': body, 'reason': reason,
                         'ref': ref, 'container': container})

    for ch in tree:
        arts = ch.get('articles', []) if ch.get('type') == 'chapter' else []
        if ch.get('type') == 'article':
            arts = [ch]
        for a in arts:
            for p in a.get('paragraphs', []):
                body = p.get('body', '')
                n_items = len(p.get('items', []))
                n_tables = len(p.get('tables', []))
                has_fc = _has_forecast(body)
                has_mk = len(body) >= length_threshold and bool(_find_split_markers(body))

                if has_fc:
                    if loose:
                        # 緩い: 予告文があれば下位の有無に関わらず候補
                        # （ただし表が受け皿の場合、下位項目も既にある場合は
                        #   優先度低なので理由に注記）
                        if n_items == 0 and n_tables == 0:
                            add(a['number'], '項', p['number'], body,
                                '予告文があるが下位項目なし', p, a)
                        else:
                            add(a['number'], '項', p['number'], body,
                                '予告文あり（下位の分割漏れ確認）', p, a)
                    else:
                        if n_items == 0 and n_tables == 0:
                            add(a['number'], '項', p['number'], body,
                                '予告文があるが下位項目なし', p, a)
                elif has_mk:
                    add(a['number'], '項', p['number'], body,
                        '長文＋分割痕跡', p, a)

                for it in p.get('items', []):
                    ibody = it.get('body', '')
                    n_sub = len(it.get('sub_items', []))
                    ihas_fc = _has_forecast(ibody)
                    ihas_mk = len(ibody) >= length_threshold and bool(_find_split_markers(ibody))
                    if ihas_fc:
                        if loose:
                            if n_sub == 0:
                                add(a['number'], '号', it['number'], ibody,
                                    '予告文があるが下位項目なし', it, p)
                            else:
                                add(a['number'], '号', it['number'], ibody,
                                    '予告文あり（下位の分割漏れ確認）', it, p)
                        else:
                            if n_sub == 0:
                                add(a['number'], '号', it['number'], ibody,
                                    '予告文があるが下位項目なし', it, p)
                    elif ihas_mk:
                        add(a['number'], '号', it['number'], ibody,
                            '長文＋分割痕跡', it, p)
    return suspects


def _norm_for_compare(s):
    """比較用に空白・記号を除いた文字列を返す。"""
    return re.sub(r'[\s\u3000・･,，.．、。]', '', s or '')


def _structure_signature(paragraphs):
    """項・号・号の下位の「構成と本文」を比較用の署名にする。"""
    sig = []
    for p in paragraphs or []:
        items = []
        for it in (p.get('items') or []):
            subs = [_norm_for_compare(su.get('body', ''))
                    for su in (it.get('sub_items') or [])]
            items.append((_norm_for_compare(it.get('body', '')), tuple(subs)))
        sig.append((_norm_for_compare(p.get('body', '')), tuple(items)))
    return tuple(sig)


def _same_structure(old_paragraphs, new_paragraphs):
    """
    AIが返した構造が、元の構造と実質同じかを判定する。
    同じなら「変更なし」として差し替えを行わない。
    """
    def sig(paras):
        out = []
        for p in paras:
            items = []
            for it in (p.get('items') or []):
                subs = []
                for su in (it.get('sub_items') or []):
                    s2 = tuple(_norm_for_compare(x.get('body', ''))
                               for x in (su.get('sub_items2') or []))
                    subs.append((_norm_for_compare(su.get('body', '')), s2))
                items.append((_norm_for_compare(it.get('body', '')), tuple(subs)))
            out.append((_norm_for_compare(p.get('body', '')), tuple(items)))
        return tuple(out)
    return sig(old_paragraphs) == sig(new_paragraphs)


def _validate_restructure(article, new_paragraphs):
    """
    Gemini が返した構造が、元の条の文言だけで構成されているかを検証する。

    AIが元に無い文言を creat してしまう（例:「次のとおりとする。」しか
    無い条に「基本給」「諸手当」を勝手に追加する）事故を防ぐ。

    返り値: (OK か, 問題の説明)
    """
    # 元の条に含まれる全テキストを連結
    orig_parts = []
    for p in article.get('paragraphs', []):
        orig_parts.append(p.get('body', ''))
        for it in p.get('items', []):
            orig_parts.append(it.get('body', ''))
            for su in it.get('sub_items', []):
                orig_parts.append(su.get('body', ''))
                for s2 in su.get('sub_items2', []):
                    orig_parts.append(s2.get('body', ''))
    orig = _norm_for_compare(''.join(orig_parts))
    if not orig:
        return False, '元の条に本文がありません'

    # 返ってきた各テキストが、元の文言に含まれるか
    new_parts = []
    for p in new_paragraphs:
        new_parts.append(p.get('body', ''))
        for it in (p.get('items') or []):
            new_parts.append(it.get('body', ''))
            for su in (it.get('sub_items') or []):
                new_parts.append(su.get('body', ''))
                for s2 in (su.get('sub_items2') or []):
                    new_parts.append(s2.get('body', ''))

    for t in new_parts:
        n = _norm_for_compare(t)
        if not n:
            continue
        if n not in orig:
            return False, f'元に無い文言が追加されています: {t[:24]}…'

    # 大幅に欠落していないか（分割で記号が落ちる分は許容）
    total_new = sum(len(_norm_for_compare(t)) for t in new_parts)
    if total_new < len(orig) * 0.8:
        return False, f'本文が欠落しています（{total_new}/{len(orig)}文字）'

    return True, ''


def _collect_suspicious_articles(tree, length_threshold=100):
    """
    疑わしい項・号を含む「条」を抽出する。
    返り値: [(article_dict, [理由のリスト]), ...]
    """
    suspects = find_suspicious_units(tree, length_threshold)
    by_article = {}
    order = []
    for s in suspects:
        # 条の実体を特定する（container が条 or 項）
        art = s['container'] if s['kind'] == '項' else None
        key = id(s['ref'])
        # 条を辿るため、tree を再走査して該当条を見つける
        by_article.setdefault(s['article'], []).append(
            f"{s['kind']}{s['number']}: {s['reason']}")
        if s['article'] not in order:
            order.append(s['article'])

    # 条番号 → 条dict のマップを作る
    art_map = {}
    for ch in tree:
        arts = ch.get('articles', []) if ch.get('type') == 'chapter' else []
        if ch.get('type') == 'article':
            arts = [ch]
        for a in arts:
            key = (a.get('number'), a.get('branch', 0))
            art_map.setdefault(a.get('number'), []).append(a)

    result = []
    for num in order:
        for a in art_map.get(num, []):
            result.append((a, by_article[num]))
    return result


def _restore_tables(old_paragraphs, new_paragraphs):
    """
    条の再構成後に、元の項が持っていた表を新しい項へ引き継ぐ。
    位置が近い項に割り当て、余りは最後の項に付ける。
    """
    tables = []
    for i, p in enumerate(old_paragraphs):
        for t in p.get('tables', []):
            tables.append((i, t))
    if not tables:
        return
    n_new = len(new_paragraphs)
    for old_idx, t in tables:
        target = min(old_idx, n_new - 1)
        new_paragraphs[target].setdefault('tables', []).append(t)


def apply_gemini_split(tree, length_threshold=100, verbose=True):
    """
    疑わしい項・号を含む「条」を丸ごと Gemini に送り、
    返ってきた構造で条の中身を差し替える。
    Gemini が使えない場合は何もしない。

    条単位で送ることで、条全体の一貫性から階層を判断できる。

    返り値: 再構成した条の件数
    """
    try:
        from gemini_helper import restructure_article_with_gemini, is_available
    except ImportError:
        if verbose:
            print("Gemini再構成: gemini_helper.py が見つからないためスキップ")
        return 0
    if not is_available():
        if verbose:
            print("Gemini再構成: APIキー未設定またはライブラリ未導入のためスキップ")
        return 0

    targets = _collect_suspicious_articles(tree, length_threshold)
    if not targets:
        if verbose:
            print("Gemini再構成: 疑わしい条は見つかりませんでした")
        return 0

    if verbose:
        print(f"Gemini再構成: 疑わしい箇所を含む条 {len(targets)}件を確認します...")
        for art, reasons in targets:
            print(f"  - {art.get('number_str','')} {art.get('title','')}: "
                  f"{', '.join(reasons[:3])}")

    n_changed = 0
    for art, reasons in targets:
        old_paragraphs = art.get('paragraphs', [])

        # 図形やレイアウト行を含む条は、本文だけを見せると内容が
        # 欠けて見えるため、AIが空欄を推測で埋めてしまう恐れがある。
        # そのような条はAIに送らず、元の構造のまま残す。
        has_fig = any(p.get('drawings') for p in old_paragraphs)
        if has_fig:
            if verbose:
                print(f"  {art.get('number_str','')}: 図形を含むためAI再構成をスキップ")
            continue
        new_simple = restructure_article_with_gemini(art, reasons=reasons)
        if not new_simple:
            continue

        # 返ってきた構造を正式な形に変換（番号は連番で振り直す）
        new_paragraphs = []
        for pi, sp in enumerate(new_simple, start=1):
            np_ = {'type': 'paragraph', 'number': pi,
                   'body': (sp.get('body') or '').strip(), 'items': []}
            for ii, si in enumerate(sp.get('items', []) or [], start=1):
                new_it = {'type': 'item', 'number': ii,
                          'body': (si.get('body') or '').strip()}
                subs = si.get('sub_items') or []
                if subs:
                    sub_list = []
                    for k, su in enumerate(
                            [x for x in subs if (x.get('body') or '').strip()],
                            start=1):
                        d = {'type': 'item_sub', 'number': k,
                             'body': (su.get('body') or '').strip()}
                        s2 = [x for x in (su.get('sub_items2') or [])
                              if (x.get('body') or '').strip()]
                        if s2:
                            d['sub_items2'] = [
                                {'type': 'item_sub2', 'number': m,
                                 'body': (x.get('body') or '').strip()}
                                for m, x in enumerate(s2, start=1)]
                        sub_list.append(d)
                    new_it['sub_items'] = sub_list
                if new_it['body']:
                    np_['items'].append(new_it)
            new_paragraphs.append(np_)

        # 中身が空になってしまった場合は安全のため差し替えない
        if not new_paragraphs or not any(p['body'] or p['items']
                                         for p in new_paragraphs):
            continue

        # AIが「変更不要」と判断した（構造が元と同じ）場合は差し替えない
        if _same_structure(old_paragraphs, new_paragraphs):
            if verbose:
                print(f"  {art.get('number_str','')}: 変更なし")
            continue

        # AIが元に無い文言を作っていないか検証する（捏造防止）
        ok, why = _validate_restructure(art, new_paragraphs)
        if not ok:
            if verbose:
                print(f"  ⚠ {art.get('number_str','')}: AIの結果を却下（{why}）")
            continue

        # AIの結果が元と同じなら、差し替えずコメントも付けない
        if _structure_signature(old_paragraphs) == _structure_signature(new_paragraphs):
            if verbose:
                print(f"  {art.get('number_str','')}: 変更なし")
            continue

        # 元の表を引き継ぐ
        _restore_tables(old_paragraphs, new_paragraphs)

        # AIが再構成したことを示すフラグ（Wordコメント用）
        new_paragraphs[0]['_ai_restructured'] = reasons

        art['paragraphs'] = new_paragraphs
        n_changed += 1
        if verbose:
            n_old_i = sum(len(p.get('items') or []) for p in old_paragraphs)
            n_new_i = sum(len(p.get('items') or []) for p in new_paragraphs)
            print(f"  {art.get('number_str','')}: "
                  f"項{len(old_paragraphs)}→{len(new_paragraphs)}、"
                  f"号{n_old_i}→{n_new_i}に再構成")

    if verbose:
        print(f"Gemini再構成: {n_changed}件の条を再構成しました")
    return n_changed


def parse_kisoku(docx_path, stop_words=None, use_gemini=False,
                 treat_layout_as_figure=False, gemini_mode='paragraph'):
    """
    就業規則docxをパースしてJSON構造を返す。

    stop_words: これらのキーワードを含む段落以降は処理しない
                例: ['附則', '改版履歴', '改訂履歴']
    treat_layout_as_figure:
                True にすると、フォントの下線が引かれた行や、
                アンダーバー（＿）・スペースで位置合わせした行を
                「図」とみなし、整形せず元のまま保持する。
                既定は False（通常どおり整形）。
                ※ 実際の図形（線・テキストボックス等）とその周辺は、
                  この設定に関わらず常に元のまま保持される。
    use_gemini: True にすると、正規表現・スタイルで判定できなかった
                段落だけを Gemini に問い合わせて再判定する（任意・無料枠）。
                APIキー未設定やライブラリ未導入の場合は自動でスキップする。
    gemini_mode: use_gemini のときの判定不能段落の問い合わせ方。
                'paragraph'（既定）: 1段落ずつ前後の文脈とともに送る。
                'agent': 本文全体を Antigravity エージェントに送り、まとめて判定する。
    """
    if stop_words is None:
        stop_words = ['附則', '付則', '改版履歴', '改訂履歴', '制定・改廃履歴', '沿革']

    doc = Document(docx_path)

    # 元ファイルの自動番号定義を読み取る
    numpr_map = build_numpr_map(doc)
    if numpr_map:
        print(f"自動番号定義を検出: {len(numpr_map)}レベル")

    # --- 段落と表を「文書に出現する順番どおり」に読む ---
    # python-docx の doc.paragraphs / doc.tables は別コレクションのため
    # 前後関係が失われる。body直下のXMLを直接たどる。
    body = doc.element.body
    seq = []   # [('text', 文字列) or ('auto', (kind, 文字列)) or ('table', {...}), ...]
    # 停止ワードは文字間に空白が入ることがある（例:「附　則」「改 版 履 歴」）。
    # 各文字の間に空白を許容する正規表現に変換して照合する。
    def _spaced(word):
        return r'[\s　]*'.join(re.escape(c) for c in word)
    stop_re = re.compile('|'.join(_spaced(w) for w in stop_words))
    n_toc_skipped = 0
    tbl_index = 0

    # 前付け（表紙・前文など、最初の章・条より前の内容）を保持する。
    # 元の書式・改ページを保つため、後で元XMLを複製できるよう位置を記録する。
    # 図形段落の前後にある「式の一部」（分子・分母など）も図の一部として
    # 元のまま保持する。treat_layout_as_figure が有効なときだけ働く。
    fig_adjacent = set()
    if True:   # 実際の図形の周辺は、チェックの有無に関わらず常に保持する
        _all = [(i, c) for i, c in enumerate(iter_block_elements(body))
                if c.tag == qn('w:p')]
        _texts = {i: Paragraph(c, doc).text for i, c in _all}
        _draws = [i for i, c in _all if has_drawing(Paragraph(c, doc))]
        _order = [i for i, c in _all]
        _pos = {v: k for k, v in enumerate(_order)}

        def _next_line(idx):
            """idx の次の、空でない段落の Line。"""
            for j in range(_pos[idx] + 1, len(_order)):
                t = _texts[_order[j]]
                if t.strip():
                    return line_info(t)
            return line_info('')

        def _fig_like(idx):
            ln = line_info(_texts.get(idx, ''))
            if not ln.s:
                return True                      # 空行は図の余白
            if ln.s.endswith(('。', '、')):
                return False
            if ln.chapter or ln.article or ln.para or ln.sym_type is not None:
                return False
            # 次が条見出しの括弧だけの行は、図の一部ではなく次の条のタイトル
            if ln.title and _next_line(idx).article:
                return False
            return True

        for d in _draws:
            k = _pos[d]
            for step in (-1, 1):                 # 前後へ最大3段落まで広げる
                for n in range(1, 4):
                    j = k + step * n
                    if j < 0 or j >= len(_order):
                        break
                    idx = _order[j]
                    if not _fig_like(idx):
                        break
                    fig_adjacent.add(idx)

    # 手打ちの目次の段落（目次スタイルの段落は除いて探す）
    manual_toc = find_manual_toc(
        [(i, _texts[i]) for i, c in _all
         if not is_toc_paragraph(Paragraph(c, doc), doc)])

    front_indices = []
    front_texts = []
    front_paras = []       # 前付けの要素 (child_idx, テキスト or 表ならNone, 目次の行か)
    body_started = False
    back_start = None      # 停止ワードを検出した位置（ここ以降は後付け）
    _prev_text = ''        # 直前の空でない段落（停止ワードが行の折り返しかを見る）

    for child_idx, child in enumerate(iter_block_elements(body)):
        if child.tag == qn('w:p'):
            para = Paragraph(child, doc)
            text = para.text

            # 目次（TOC）の段落はスキップする
            # 目次には「第1章 総則 4」のようにページ番号付きで
            # 章・条見出しが並ぶため、本文と誤認識してしまう
            if is_toc_paragraph(para, doc):
                n_toc_skipped += 1
                continue

            # 手打ちの目次の行（本文が始まる前だけ）は前付けにする。
            # 停止ワード（附則など）や本文の始まりの判定より前に見分ける
            if not body_started and child_idx in manual_toc:
                front_indices.append(child_idx)
                front_paras.append((child_idx, text, True))
                if text.strip():
                    front_texts.append(text.strip())
                    _prev_text = text
                continue

            # 停止ワード（附則・改版履歴など）を検出したら、
            # そこから末尾までは「後付け」として整形せず元のまま保持する
            # 行頭一致に限る（本文中の「〇〇法附則第3項」で止まらないように）
            # 「附則に定める…」「附則第2項の規定により…」のような本文中の参照
            # （行の折り返しで行頭に来たもの）では止めない
            # 前の段落が文の途中で終わっているとき（行の折り返し）だけ判定する
            _sm = stop_re.match(text.strip()) if text.strip() else None
            if _sm and not (_ends_mid_sentence(_prev_text, kanji=False)
                            and _STOP_WORD_REF.match(text.strip()[_sm.end():])):
                print(f"停止: {repr(text.strip()[:30])}（以降は整形せず保持）")
                back_start = child_idx
                break
            if text.strip():
                _prev_text = text
            # 図形・画像を含む段落は、テキストに置き換えると失われるため
            # 元のXMLをそのまま複製する対象として記録する
            # 章・条・条タイトルの行は、装飾の図形（見出し下の罫線など）を
            # 含んでいても構造として解析する（図形より構造判定を優先）
            ln = line_info(text)
            # スタイル・自動番号から見た種別を先に求める
            _nid, _il = get_paragraph_numpr(para, doc)
            _kind0 = numpr_map.get((_nid, _il)) if _nid else None
            if _kind0 is None:
                _sid0 = get_style_id(para)
                _kind0 = STYLE_KIND_MAP.get(_sid0) if _sid0 else None
            # Word の箇条書き（List Bullet）は号。段落番号（List Number）は、直前の項が
            # 「次の書類を…」と予告していれば号、そうでなければ項（build_tree が決める）
            _sname = get_style_name(para, doc) or ''
            if _sname.startswith('List Bullet') and _kind0 in (None, 'paragraph', 'paragraph_1'):
                _kind0 = 'item'
            elif _sname.startswith('List Number') and _kind0 in (None, 'paragraph'):
                _kind0 = 'list_number'
            # 章・条は最優先で構造として扱う（図形より優先）
            _is_strong = bool(ln.chapter or ln.article
                              or _kind0 in ('chapter', 'article'))
            # 括弧だけの行は通常は条タイトルだが、
            # 図形に隣接する場合は式の見出しとみなして図の一部にする
            _is_title = bool(ln.title)
            # 常に保持: 実際の図形（線・テキストボックス等）とその周辺
            _always_fig = (child_idx in fig_adjacent
                           or (not _is_title and has_drawing(para)))
            # 任意: フォントの下線が引かれた行・下線文字やスペースの位置合わせ
            _opt_fig = (treat_layout_as_figure and not _is_title
                        and (has_underline(para) or is_layout_paragraph(text)))
            if (not _is_strong) and (_always_fig or _opt_fig):
                seq.append(('drawing', {'index': child_idx,
                                        'text': text.strip()}))
                continue

            # 種別の判定材料を集める
            #  種別判定の優先順位:
            #   1. 自動番号 numPr → numbering定義（_guess_kind が pStyle も加味）
            #   2. スタイル名（STYLE_KIND_MAP）
            #  numbering定義は _guess_kind で pStyle・numFmt・lvlText を総合判断
            #  するため最も正確。無い場合のみ STYLE_KIND_MAP にフォールバック。
            kind = _kind0

            # --- 本文（最初の章・条）が始まったかを判定 ---
            if not body_started:
                if kind in ('chapter', 'article') or ln.chapter or ln.article:
                    body_started = True
                    # 条で本文が始まるとき、直前の括弧だけの行
                    # （第1条のタイトル）は前付けではなく本文に含める
                    if kind == 'article' or ln.article:
                        _last = None
                        for _fp in reversed(front_paras):
                            if _fp[1] is not None and _fp[1].strip():
                                _last = _fp
                                break
                        if (_last is not None and not _last[2]
                                and line_info(_last[1]).title
                                and _last[0] in front_indices):
                            front_indices.remove(_last[0])
                            if front_texts:
                                front_texts.pop()
                            seq.append(('text', line_info(_last[1])))
                else:
                    # まだ本文前 → 前付け（表紙・前文）として記録し、
                    # 構造解析の対象からは外す
                    front_paras.append((child_idx, text, False))
                    front_indices.append(child_idx)
                    if text.strip():
                        front_texts.append(text.strip())
                    continue

            if kind and text.strip():
                seq.append(('auto', (kind, text)))
            else:
                # スタイルもnumPrも無い段落。テキストパターンでも判定できず
                # 'body' 扱いになる場合、後で Gemini に問い合わせる候補とする
                # 段落の途中に改行（Shift+Enter = w:br / w:cr）が
                # 入っている場合、1つの w:p に複数の論理行が同居している。
                # python-docx はこれを "\n" として返すので、ここで分割して
                # 別々の行として扱う。分割しないと、
                #   「（採用手続き）\n第4条　組合は…」
                # のような段落が条見出しとして認識されず、直前の章見出しに
                # 吸収されてしまう（merge_lines の新項目判定が行末アンカーの
                # 正規表現を使っているため）。
                # 文の途中で折り返しているだけの改行は、この後の
                # merge_lines が結合し直すので分割して問題ない。
                for _line in text.split('\n'):
                    if _line.strip():
                        seq.append(('text', line_info(_line)))
        elif child.tag == qn('w:tbl'):
            if not body_started:
                # 本文前の表（表紙内の表など）も前付けに含める
                front_paras.append((child_idx, None, False))
                front_indices.append(child_idx)
                tbl_index += 1
                continue
            table = Table(child, doc)
            rows = []
            for row in table.rows:
                rows.append([cell.text.strip() for cell in row.cells])
            # index は「元docx内で何番目の表か」。整形時に元の表XMLを
            # そのまま複製するために使う（セル結合・罫線・書式を保持）。
            seq.append(('table', {'rows': rows, 'index': tbl_index}))
            tbl_index += 1

    # 途中改行を結合（表・自動番号段落はそのまま通過）
    if n_toc_skipped:
        print(f"目次をスキップ: {n_toc_skipped}段落")
    merged = merge_lines(seq)

    # 各要素を分類
    # まず通常の分類を行い、テキスト段落が 'body'（判定不能）になった場合は
    # Gemini による再判定の候補として位置を記録する。
    flat = []
    gemini_candidates = []   # (flat_index, target_text, seq_index)

    # 先に merged 内のテキスト本文だけを取り出しておく（文脈参照用）
    units = [val for k, val in merged if k == 'text']
    merged_texts = [u.text for u in units]

    # 号の記号の深さ（号・号の細分…）を決める
    depths = marker_depths(merged)
    kana_nums = kana_numbers(merged)

    text_cursor = 0
    for k, val in merged:
        if k == 'table':
            flat.append({'type': 'table', 'rows': val['rows'],
                         'index': val.get('index')})
        elif k == 'drawing':
            flat.append({'type': 'drawing', 'index': val.get('index'),
                         'text': val.get('text', '')})
        elif k == 'auto':
            node = classify_auto(val[0], val[1])
            if node:
                flat.append(node)
        else:  # 'text'
            node = classify_line(val, depth=depths.get(id(val)),
                                 number=kana_nums.get(id(val)))
            if node and node.get('type') in ('paragraph', 'item', 'item_sub', 'item_sub2'):
                # 元の行頭の記号を残す（新旧対照表の現行を、元の文書のまま見せるため）
                mk = _lead_mark(val.text, node.get('body'))
                if mk:
                    node['mark'] = mk
            if node:
                # 'text'（＝どのパターンにも該当せず本文扱い）は判定不能候補。
                # 括弧書きの注記（「（中途採用者を除く）」）は本文と分かっているので除く
                if node.get('type') == 'text' and not node.get('note') and use_gemini:
                    gemini_candidates.append((len(flat), val.text, text_cursor))
                flat.append(node)
            text_cursor += 1

    # 括弧だけの行は、次の単位が条見出しのときだけ条タイトルにする。
    # それ以外（注記・補足）は、直前の項目の本文に付ける（⑤ 図の直後でも同じ）
    for i, node in enumerate(flat):
        if node['type'] == 'article_title' and node.get('_raw') and (
                i + 1 >= len(flat) or flat[i + 1]['type'] != 'article'):
            flat[i] = {'type': 'text', 'body': node['_raw'], 'note': True}

    # --- Gemini による判定不能段落の再判定 ---
    if use_gemini and gemini_candidates:
        try:
            from gemini_helper import (classify_with_gemini, is_available,
                                       classify_document_with_agent)
            if is_available():
                print(f"Gemini判定: 判定不能な{len(gemini_candidates)}段落を問い合わせます...")
                n_resolved = 0
                agent_kinds = {}
                if gemini_mode == 'agent':
                    agent_kinds = classify_document_with_agent(
                        merged_texts, [c[2] for c in gemini_candidates])
                    if agent_kinds is None:
                        agent_kinds = {}
                        print("Antigravity判定: 使えなかったためルールの判定のまま")
                for flat_idx, target, tcur in gemini_candidates:
                    if gemini_mode == 'agent':
                        kind = agent_kinds.get(tcur)
                    else:
                        before = merged_texts[max(0, tcur - 2):tcur]
                        after = merged_texts[tcur + 1:tcur + 3]
                        kind = classify_with_gemini(target, before, after)
                    # 番号の無い行を章・条にすると以降の条番号が全部ずれるため、
                    # 章・条の判定は採用しない（番号付きの章・条は正規表現で拾える）
                    if kind and kind not in ('body', 'chapter', 'article'):
                        # Geminiの判定で置き換える
                        new_node = classify_auto(kind, target)
                        if new_node:
                            flat[flat_idx] = new_node
                            n_resolved += 1
                print(f"Gemini判定: {n_resolved}/{len(gemini_candidates)}段落を再判定しました")
            else:
                print("Gemini判定: APIキー未設定またはライブラリ未導入のためスキップ")
        except ImportError:
            print("Gemini判定: gemini_helper.py が見つからないためスキップ")

    # ネスト構造に組み立て
    tree = build_tree(flat)

    # 前付け（表紙・前文）をツリー先頭に追加する。
    # indices は元docxの body 直下での位置。整形時に元XMLを複製して
    # 書式・改ページを保持するために使う。
    if front_indices:
        tree.insert(0, {
            'type': 'front_matter',
            'indices': front_indices,
            'texts': front_texts,
        })
        print(f"前付け（表紙・前文）を保持: {len(front_indices)}要素")

    # 後付け（附則・改版履歴など）を末尾に追加する。
    # 停止ワード以降は整形せず、元のXMLをそのまま複製して保持する。
    if back_start is not None:
        all_children = list(iter_block_elements(body))
        back_indices = [
            i for i in range(back_start, len(all_children))
            if all_children[i].tag in (qn('w:p'), qn('w:tbl'))
        ]
        if back_indices:
            back_texts = []
            for i in back_indices:
                el = all_children[i]
                if el.tag == qn('w:p'):
                    t = Paragraph(el, doc).text.strip()
                    if t:
                        back_texts.append(t)
            tree.append({
                'type': 'back_matter',
                'indices': back_indices,
                'texts': back_texts,
            })
            print(f"後付け（附則等）を保持: {len(back_indices)}要素")

    # --- Gemini による「押し込まれた項・号」の分割（後処理）---
    if use_gemini:
        apply_gemini_split(tree, length_threshold=100, verbose=True)

    return tree


# ============================================================
# 実行（Colabで使う場合）
# ============================================================

if __name__ == '__main__':
    # Colabではここをアップロードしたファイル名に変更
    PATH = 'タイセツ　就業規則.docx'

    result = parse_kisoku(PATH)

    # JSON出力
    json_str = json.dumps(result, ensure_ascii=False, indent=2)
    print(json_str[:3000])  # 先頭3000文字だけ表示

    # ファイルに保存
    with open('kisoku_parsed.json', 'w', encoding='utf-8') as f:
        f.write(json_str)
    print("\n→ kisoku_parsed.json に保存しました")

    # 統計
    chapters = [n for n in result if n['type']=='chapter']
    articles = [a for ch in chapters for a in ch.get('articles',[])]
    print(f"\n章数: {len(chapters)}, 条数: {len(articles)}")
