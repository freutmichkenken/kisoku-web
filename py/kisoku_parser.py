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

# 記号階層の動的判定（あれば使う。無くても従来動作にフォールバック）
try:
    from marker_hierarchy import MarkerAnalyzer, identify_symbol
    _HAS_MARKER_ANALYZER = True
except ImportError:
    _HAS_MARKER_ANALYZER = False


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
    if RE_CHAPTER.match(s) or RE_ARTICLE.match(s) or RE_TITLE.match(s):
        return False
    if RE_PARA.match(text):
        return False
    if _HAS_MARKER_ANALYZER:
        try:
            if identify_symbol(s)[0] is not None:
                return False
        except Exception:
            pass
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

# 末尾が「続き」と判断する条件
COMPLETE_ENDINGS = set('。」）)．…')
COMPLETE_WORDS = (
    'もの','こと','とき','ため','ない','ある','なる','れる','いる',
    'おく','くる','いく','みる','いう','できる','よる','ける','せる',
    'める','てる','でる','場合','期間','時間','以上','以下','以内',
    '以外','業務','事項','規定','条件','事由','義務','権利','方法',
    '手続','基準','措置','期日','日数',
)
INCOMPLETE_KANA = set('てでにをがはもしたとのなくけれせすつぬふむゆるわばへだ')
KANJI_COMPLETE = set('等他者時中内上下前後')

def is_continuation(text):
    """この行が次の行と結合すべき「途中改行」かどうか"""
    t = text.rstrip('　 \t')
    if not t or len(t) < 2:
        return False
    last = t[-1]
    if last in COMPLETE_ENDINGS:
        return False
    if any(t.endswith(w) for w in COMPLETE_WORDS):
        return False
    if last in ('―','─','‐','–'):
        return True
    if last == '、':
        return True
    if last in INCOMPLETE_KANA:
        return True
    # 漢字で終わるが完結語でない
    if re.search(r'[一-龯]$', last) and last not in KANJI_COMPLETE:
        if len(t) >= 2 and not re.match(r'[一-龯]', t[-2]):
            return True
    return False

def merge_lines(items):
    """
    複数の p に分割された1文を結合して返す。
    items: [('text', 文字列) or ('table', テーブルdict), ...] の順序付きリスト
    返り値: [('text', 結合済み文字列) or ('table', dict), ...]
    """
    merged = []
    buf = ''

    def flush():
        nonlocal buf
        if buf:
            merged.append(('text', buf))
            buf = ''

    for kind, val in items:
        if kind in ('table', 'auto', 'drawing'):
            # 表・自動番号段落は結合対象外。それまでのテキストを確定してから追加
            flush()
            merged.append((kind, val))
            continue

        line = val
        if not line.strip():
            continue
        if buf == '':
            buf = line
        else:
            # 次の行が番号で始まる場合は新しい段落
            stripped = line.lstrip('　 \t')
            is_new = bool(
                re.match(r'^第[　\s]*[０-９0-9一二三四五六七八九十百千]+[　\s]*[章条]', stripped) or
                re.match(r'^（[^）]+）[\s　]*([※＊*].*)?$', stripped) or  # 条タイトル
                re.match(r'^\([^)]+\)\s*$', stripped) or
                re.match(r'^[２-９2-9][０-９0-9]*[　 ．.]', stripped) or  # 第2項以降
                re.match(r'^[①-⑳㉑-㊿]', stripped) or
                re.match(r'^（[０-９0-9一二三四五六七八九十]+）', stripped) or
                re.match(r'^\([０-９0-9一二三四五六七八九十]+\)', stripped)
            )
            # 上のパターンに無い記号（ア．イ．／a)／一、 など）も
            # 行頭にあれば新しい項目とみなす
            if not is_new and _HAS_MARKER_ANALYZER:
                try:
                    if identify_symbol(stripped)[0] is not None:
                        is_new = True
                except Exception:
                    pass
            if is_new:
                merged.append(('text', buf))
                buf = line
            elif is_continuation(buf):
                # 前の行が途中 → 結合（インデントを除去して接続）
                buf = buf + stripped
            else:
                # 前の行が完結しているのに新番号でもない
                # → 段落内の折り返しと判断して結合
                buf = buf + stripped
    flush()
    return merged


# ============================================================
# STEP 2: 1行を分類・パース
# ============================================================

# 各パターン
# 「第 1 章」「第 24 条」のように番号の前後に空白が入った書き方も
# 見出しとして認める（PDF変換や手打ちの文書で多い）。
RE_CHAPTER = re.compile(r'^第[　\s]*([０-９0-9一二三四五六七八九十百千万]+)[　\s]*章[　\s]*(.*)$',
                        re.DOTALL)
# 「第９条」「第９条の２」の両方に対応（枝番号 = branch）
# 「第９条に定める…」のような本文中の条文参照を条見出しと
# 誤判定しないよう、条番号の直後は「空白・括弧・行末」に限る。
RE_ARTICLE = re.compile(
    r'^第[　\s]*([０-９0-9一二三四五六七八九十百千万]+)[　\s]*条'
    r'(?:[　\s]*の[　\s]*([０-９0-9一二三四五六七八九十]+))?'
    r'(?=[　\s（(]|$)'
    r'[　\s]*(.*)$',
    re.DOTALL
)
# 条タイトル。「（目的）」のほか、モデル就業規則等でよくある
# 「（目的）　※〇〇パターン」のような末尾注釈も許容する。
# 括弧内が句点で終わるもの（例:「（…により計算する。）」）は
# 条タイトルではなく注記なので除外する
RE_TITLE   = re.compile(r'^[（(]([^）)]*[^）)。．])[）)][\s　]*([※＊*].*)?$',
                        re.DOTALL)
# 第2項以降（先頭インデント + 2以上の数字 + 区切り）。タブ区切りにも対応
RE_PARA    = re.compile(r'^[　\s]*([２-９2-9][０-９0-9]*)[　\s．.\t]+(.*)$',
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

def zenkaku_to_int(s):
    """全角数字を含む文字列を整数に変換"""
    s = s.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    try:
        return int(s)
    except ValueError:
        # 漢数字
        kanji = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10,
                 '百':100,'千':1000,'万':10000}
        val = 0; tmp = 0
        for c in s:
            if c in kanji:
                k = kanji[c]
                if k >= 10:
                    tmp = (tmp or 1) * k
                    if k >= 1000:
                        val += tmp; tmp = 0
                else:
                    tmp += k
        return val + tmp or 1

def detect_item_hierarchy(merged_texts):
    """
    文書全体の号マーカーの使われ方を分析し、
    丸数字（①②）が「号」か「号の下位」かを判定する。

    括弧数字（１）と丸数字①が両方使われている文書では ①＝号の下位、
    丸数字①しか無ければ ①＝号。
    返り値: bool（True なら丸数字は号の下位）
    """
    has_paren = False
    has_maru = False
    for text in merged_texts:
        s = text.lstrip('　 \t')
        if RE_ITEM_PAREN.match(s):
            has_paren = True
        if RE_ITEM_MARU.match(s):
            has_maru = True
    return has_paren and has_maru


def classify_line(text, maru_is_sub=False, analyzer=None):
    """
    1行を分類してdictを返す。
    analyzer:  MarkerAnalyzer。号系記号（①（１）ア 等）の階層判定に使う。
               None の場合は従来ロジック（maru_is_sub フラグ）にフォールバック。
    maru_is_sub: analyzer が無い場合のみ有効。丸数字を号の下位として扱う。
    keys: type, number, title/body
    """
    raw = text
    stripped = text.lstrip('　 \t')

    if not stripped:
        return None

    # 章
    m = RE_CHAPTER.match(stripped)
    if m:
        return {'type':'chapter', 'number_str':f'第{m.group(1)}章',
                'number': zenkaku_to_int(m.group(1)), 'title': m.group(2).strip()}

    # 条タイトル（括弧のみの行）
    m = RE_TITLE.match(stripped)
    if m:
        return {'type':'article_title', 'title': m.group(1).strip(),
                'title_note': (m.group(2) or '').strip()}

    # 条（+本文が同行にある場合も）。枝番号「第X条のY」にも対応
    m = RE_ARTICLE.match(stripped)
    if m:
        art_num = zenkaku_to_int(m.group(1))
        branch = m.group(2)          # 「の２」の "２" 部分（なければ None）
        body = m.group(3).strip()
        if branch:
            branch_num = zenkaku_to_int(branch)
            number_str = f'第{m.group(1)}条の{branch}'
        else:
            branch_num = 0
            number_str = f'第{m.group(1)}条'
        # 「第1条（目的）本文…」のように括弧タイトルが続く場合は抽出
        title = ''
        tm = re.match(r'^[（(]([^）)]+)[）)]\s*(.*)$', body, re.DOTALL)
        if tm:
            title = tm.group(1).strip()
            body = tm.group(2).strip()
        # 条の直後に続く「１　」は第1項の番号なので取り除く
        body = re.sub(r'^[１1][　\s．.]+', '', body)
        return {'type':'article', 'number_str': number_str,
                'number': art_num, 'branch': branch_num,
                'title': title, 'body': body}

    # 第2項以降（先頭インデント+数字）
    m = RE_PARA.match(text)  # raw（インデント込み）に対してマッチ
    if m:
        num = zenkaku_to_int(m.group(1))
        if num >= 2:  # 2以上のみ項扱い（1は条本文の続きとして扱わない）
            return {'type':'paragraph', 'number': num, 'body': m.group(2).strip()}

    # --- 号系記号の判定（analyzer があれば階層を動的判定）---
    if analyzer is not None:
        kind, num, body = analyzer.classify(stripped)
        if kind is not None:
            if kind == 'item_sub2' or kind.startswith('item_sub2'):
                return {'type':'item_sub2', 'number': num, 'body': body.strip()}
            if kind.startswith('item_sub'):
                return {'type':'item_sub', 'number': num, 'body': body.strip()}
            return {'type':'item', 'number': num, 'body': body.strip()}
    else:
        # analyzer が無い場合の従来ロジック
        m = RE_ITEM_MARU.match(text)
        if m:
            num = MARU_MAP.get(m.group(1), 1)
            if maru_is_sub:
                return {'type':'item_sub', 'number': num, 'body': m.group(2).strip()}
            return {'type':'item', 'number': num, 'body': m.group(2).strip()}
        m = RE_ITEM_PAREN.match(text)
        if m:
            return {'type':'item', 'number': zenkaku_to_int(m.group(1)),
                    'body': m.group(2).strip()}

    # その他（本文の続き、表など）
    return {'type':'text', 'body': stripped}


def classify_auto(kind, text):
    """
    自動番号（Wordのアウトライン）が設定されている段落を分類する。
    ただしテキストに手打ちの章・条番号がある場合はそちらを優先する
    （スタイルは項なのに本文に「第12条」と手打ちされているケースがあるため）。
    """
    body = text.lstrip('　 \t').strip()
    if not body:
        return None

    # --- スタイル/自動番号で「号」「項」と確定している場合は、
    #     テキストが「第９条に定める…」のような条文参照で始まっていても
    #     条見出しとは扱わない（誤判定防止）---
    if kind in ('item', 'item_sub', 'paragraph', 'paragraph_1', 'body_plain'):
        pass   # 下の種別ごとの処理へ進む
    else:
        # --- 章・条が期待される文脈でのみ、テキストの手打ち番号を見る ---
        m = RE_CHAPTER.match(body)
        if m:
            return {'type':'chapter', 'number_str': f'第{m.group(1)}章',
                    'number': zenkaku_to_int(m.group(1)), 'title': m.group(2).strip()}
        m = RE_ARTICLE.match(body)
        if m:
            branch = m.group(2)
            rest = m.group(3).strip()
            # 「第1条（目的）」のように残りが括弧だけならタイトルとして扱う
            title = ''
            tm = re.match(r'^[（(]([^）)]+)[）)]\s*(.*)$', rest, re.DOTALL)
            if tm:
                title = tm.group(1).strip()
                rest = tm.group(2).strip()
            rest = re.sub(r'^[１1][　\s．.]+', '', rest)
            return {'type':'article',
                    'number_str': f'第{m.group(1)}条' + (f'の{branch}' if branch else ''),
                    'number': zenkaku_to_int(m.group(1)),
                    'branch': zenkaku_to_int(branch) if branch else 0,
                    'title': title,
                    'body': rest}

    # --- スタイル/自動番号で「条」と分かっている場合 ---
    # 「(目的)」のように括弧だけの行は、条見出し（番号は自動採番）
    if kind == 'article':
        tm = RE_TITLE.match(body)
        if tm:
            return {'type':'article', 'number_str':'', 'number': 0, 'branch': 0,
                    'title': tm.group(1).strip(),
                    'title_note': (tm.group(2) or '').strip(), 'body': ''}
        return {'type':'article', 'number_str':'', 'number':0, 'branch':0,
                'title':'', 'body': body}

    # 括弧だけの行は条タイトル（スタイル情報が無い場合）
    if kind not in ('item', 'item_sub', 'paragraph', 'paragraph_1', 'body_plain'):
        m = RE_TITLE.match(body)
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
        if _HAS_MARKER_ANALYZER:
            try:
                tn, num, bd = identify_symbol(body)
                if tn:
                    return {'type':'item_sub2', 'number': num, 'body': bd.strip()}
            except Exception:
                pass
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


def build_tree(flat):
    """
    分類済みのフラットなdictリストを、
    章 > 条 > 項 > 号 のネスト構造に変換する。
    """
    result = []
    cur_chapter = None
    cur_article = None
    cur_paragraph = None
    pending_title = None   # 条タイトル（article_titleが先行する場合）
    pending_note = ''      # 条タイトル末尾の注釈（※〜）

    def _flush_pending():
        """条が来ないまま残った条タイトルを、本文として戻す（欠落防止）。"""
        nonlocal pending_title, pending_note
        if pending_title:
            txt = f'（{pending_title}）'
            if pending_note:
                txt += f'　{pending_note}'
            if cur_paragraph is not None:
                cur_paragraph['body'] += txt
            elif cur_article is not None and cur_article['paragraphs']:
                cur_article['paragraphs'][-1]['body'] += txt
        pending_title = None
        pending_note = ''  
    article_seq = [0]      # 自動採番の条番号カウンタ（章をまたいで連番）

    for node in flat:
        t = node['type']

        if t == 'chapter':
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
            pending_title = None

        elif t == 'article_title':
            # 次の条に付けるタイトルとして保留
            pending_title = node['title']
            pending_note = node.get('title_note', '')

        elif t == 'article':
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
            pending_title = None
            # 第1項（本文が同行）
            if node['body']:
                cur_paragraph = {
                    'type': 'paragraph',
                    'number': 1,
                    'body': node['body'],
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
            if cur_article is None:
                # 条の外に項が来るケース（稀）→ textとして扱う
                pass
            else:
                num = node['number']
                if not num:   # 0 = 自動採番 → 連番を割り当てる
                    num = len(cur_article['paragraphs']) + 1
                cur_paragraph = {
                    'type': 'paragraph',
                    'number': num,
                    'body': node['body'],
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
                    'body': node['body']
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
                    'body': node['body']
                })

        elif t == 'item_sub':
            _flush_pending()
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
                    'body': node['body'], '_promoted': True
                })
            elif cur_paragraph is not None and cur_paragraph.get('items'):
                parent = cur_paragraph['items'][-1]
                num = node['number']
                if not num:
                    num = len(parent.get('sub_items', [])) + 1
                parent.setdefault('sub_items', []).append({
                    'type': 'item_sub',
                    'number': num,
                    'body': node['body']
                })
            elif cur_paragraph is not None:
                # 号が無いのに下位項目が来た場合は、本文に混ぜてしまうと
                # 内容が埋もれるため、号に格上げして独立させる
                num = node['number'] or (len(cur_paragraph.get('items', [])) + 1)
                cur_paragraph.setdefault('items', []).append({
                    'type': 'item', 'number': num, 'body': node['body'],
                    '_promoted': True
                })

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
                    'type': 'item_sub2', 'number': num, 'body': node['body']
                })
            elif cur_paragraph is not None and cur_paragraph.get('items'):
                # 受け皿が無ければ号の下位に格上げして残す
                it = cur_paragraph['items'][-1]
                num = node['number'] or (len(it.get('sub_items', [])) + 1)
                it.setdefault('sub_items', []).append({
                    'type': 'item_sub', 'number': num, 'body': node['body']
                })

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
            if cur_paragraph and cur_paragraph.get('items'):
                # 直前が号 → 号の本文に追記
                cur_paragraph['items'][-1]['body'] += body
            elif cur_paragraph:
                cur_paragraph['body'] += body
            elif cur_article:
                # 条の本文が別行にある場合
                if cur_article['paragraphs']:
                    cur_article['paragraphs'][-1]['body'] += body
                else:
                    cur_paragraph = {'type':'paragraph','number':1,'body':body,'items':[]}
                    cur_article['paragraphs'].append(cur_paragraph)

    return result


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
                 treat_layout_as_figure=False):
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

        def _fig_like(idx):
            t = _texts.get(idx, '')
            st = t.lstrip('　 \t')
            if not st:
                return True                      # 空行は図の余白
            if st.endswith(('。', '、')):
                return False
            if (RE_CHAPTER.match(st) or RE_ARTICLE.match(st)
                    or RE_PARA.match(t)):
                return False
            if _HAS_MARKER_ANALYZER:
                try:
                    if identify_symbol(st)[0] is not None:
                        return False
                except Exception:
                    pass
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

    front_indices = []
    front_texts = []
    body_started = False
    back_start = None      # 停止ワードを検出した位置（ここ以降は後付け）

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

            # 停止ワード（附則・改版履歴など）を検出したら、
            # そこから末尾までは「後付け」として整形せず元のまま保持する
            # 行頭一致に限る（本文中の「〇〇法附則第3項」で止まらないように）
            if text.strip() and stop_re.match(text.strip()):
                print(f"停止: {repr(text.strip()[:30])}（以降は整形せず保持）")
                back_start = child_idx
                break
            # 図形・画像を含む段落は、テキストに置き換えると失われるため
            # 元のXMLをそのまま複製する対象として記録する
            # 章・条・条タイトルの行は、装飾の図形（見出し下の罫線など）を
            # 含んでいても構造として解析する（図形より構造判定を優先）
            _st = text.lstrip('　 \t')
            # スタイル・自動番号から見た種別を先に求める
            _nid, _il = get_paragraph_numpr(para, doc)
            _kind0 = numpr_map.get((_nid, _il)) if _nid else None
            if _kind0 is None:
                _sid0 = get_style_id(para)
                _kind0 = STYLE_KIND_MAP.get(_sid0) if _sid0 else None
            # 章・条は最優先で構造として扱う（図形より優先）
            _is_strong = bool(RE_CHAPTER.match(_st) or RE_ARTICLE.match(_st)
                              or _kind0 in ('chapter', 'article'))
            # 括弧だけの行は通常は条タイトルだが、
            # 図形に隣接する場合は式の見出しとみなして図の一部にする
            _is_title = bool(RE_TITLE.match(_st))
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
                stripped_t = text.lstrip('　 \t')
                if (kind in ('chapter', 'article')
                        or RE_CHAPTER.match(stripped_t)
                        or RE_ARTICLE.match(stripped_t)):
                    body_started = True
                else:
                    # まだ本文前 → 前付け（表紙・前文）として記録し、
                    # 構造解析の対象からは外す
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
                        seq.append(('text', _line))
        elif child.tag == qn('w:tbl'):
            if not body_started:
                # 本文前の表（表紙内の表など）も前付けに含める
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
    text_positions = []      # merged内でのテキスト行の並び（前後文脈用）

    # 先に merged 内のテキスト本文だけを取り出しておく（文脈参照用）
    merged_texts = [val for k, val in merged if k == 'text']

    # 記号階層を動的に学習（analyzer があれば号系記号を階層判定）
    analyzer = None
    maru_is_sub = False
    if _HAS_MARKER_ANALYZER:
        analyzer = MarkerAnalyzer()
        # 字下げ量は簡易に「行頭空白の視覚幅」で近似（Wordインデントは結合後に
        # 失われることがあるため、テキストの行頭空白を主に使う）
        learn_lines = []
        for t in merged_texts:
            stripped = t.lstrip('　 \t')
            # 章・条・条タイトル・項（2以降）の番号は「号の記号」ではないので
            # 記号階層の学習から除外する。
            # （例:「２．前項各号の…」の "２．" を号の記号と誤学習するのを防ぐ）
            if (RE_CHAPTER.match(stripped) or RE_ARTICLE.match(stripped)
                    or RE_TITLE.match(stripped)):
                # 条が変わる位置に番兵を入れ、記号の入れ子状態をリセットさせる
                learn_lines.append((None, 0))
                continue
            mp = RE_PARA.match(t)
            if mp and zenkaku_to_int(mp.group(1)) >= 2:
                continue
            w = 0
            for c in t:
                if c == '\u3000': w += 2
                elif c == ' ': w += 1
                elif c == '\t': w += 4
                else: break
            learn_lines.append((t, w))
        analyzer.learn(learn_lines)
        if analyzer.type_depth:
            depth_str = ", ".join(
                f"{t}→{analyzer.depth_to_kind(d)}"
                for t, d in sorted(analyzer.type_depth.items(), key=lambda x: x[1]))
            print(f"記号階層を学習: {depth_str}")
    else:
        # フォールバック（旧・簡易判定）
        maru_is_sub = detect_item_hierarchy(merged_texts)
        if maru_is_sub:
            print("記号階層: （１）＝号、①＝号の下位 と判定")

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
            node = classify_line(val, maru_is_sub=maru_is_sub, analyzer=analyzer)
            if node:
                # 'text'（＝どのパターンにも該当せず本文扱い）は判定不能候補
                if node.get('type') == 'text' and use_gemini:
                    gemini_candidates.append((len(flat), val, text_cursor))
                flat.append(node)
            text_cursor += 1

    # --- Gemini による判定不能段落の再判定 ---
    if use_gemini and gemini_candidates:
        try:
            from gemini_helper import classify_with_gemini, is_available
            if is_available():
                print(f"Gemini判定: 判定不能な{len(gemini_candidates)}段落を問い合わせます...")
                n_resolved = 0
                for flat_idx, target, tcur in gemini_candidates:
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
