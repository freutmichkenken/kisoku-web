# -*- coding: utf-8 -*-
"""
号の記号（①②、（１）、ア、一、a) 等）の種類と番号の判定

就業規則の項・号・その下位に使われる記号は文書ごとにまちまちなので、
行頭の記号の種類・番号・後ろの本文を判定する（identify_symbol）。
どの記号が号・号の細分かは、kisoku_parser.marker_depths が項ごとの入れ子で決める
（条ごとに記号の使い方が違う文書があるため、文書全体の表にはしない）。
"""

import re

# ------------------------------------------------------------
# 記号タイプの定義
#   各タイプは (名前, 正規表現, 番号抽出関数)
#   正規表現は「行頭の空白を除いた文字列」に対してマッチさせる
# ------------------------------------------------------------

_MARU_MAP = {chr(0x2460 + i): i + 1 for i in range(20)}   # ①-⑳
_MARU_MAP.update({chr(0x3251 + i): i + 21 for i in range(15)})  # ㉑-㉟
_MARU_MAP.update({chr(0x32b1 + i): i + 36 for i in range(15)})  # ㊱-㊿

_MARU_CHARS = ''.join(_MARU_MAP.keys())

_KANJI_DIGIT = {'〇':0,'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9}
_KANJI_UNIT = {'十':10,'百':100,'千':1000}


def kanji_to_int(s):
    """漢数字を int に（「十一」=11、「百十」=110、「千二百」=1200、「二〇」=20）。
    「〇」を含み「十百千万」を含まない漢数字は位取りの数字として読む。
    漢数字として読めない文字は無視する（何も読めなければ 0）。"""
    if s and all(c in _KANJI_DIGIT for c in s) and '〇' in s:
        return int(''.join(str(_KANJI_DIGIT[c]) for c in s))
    total = section = digit = 0
    for c in s:
        if c in _KANJI_DIGIT:
            digit = _KANJI_DIGIT[c]
        elif c in _KANJI_UNIT:
            section += (digit or 1) * _KANJI_UNIT[c]
            digit = 0
        elif c == '万':
            total += (section + digit or 1) * 10000
            section = digit = 0
    return total + section + digit

_KATAKANA_AIUEO = 'アイウエオカキクケコサシスセソタチツテトナニヌネノ'
_KATAKANA_IROHA = 'イロハニホヘトチリヌルヲワカヨタレソツネナラム'


def _z2i(s):
    """全角/半角数字を int に。"""
    s = s.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    try:
        return int(s)
    except ValueError:
        return 0


# 記号タイプ判定器。順序が重要（より限定的なものを先に）。
# 各要素: (type_name, compiled_regex, extractor)
#   extractor(match) -> (番号int, 本文str)
_SYMBOL_TYPES = [
    # 全角括弧数字 （１）
    ("paren_zen_num",
     re.compile(r'^（([０-９0-9]+)）[　\s\t]*(.*)$'),
     lambda m: (_z2i(m.group(1)), m.group(2))),
    # 半角括弧数字 (1)
    ("paren_han_num",
     re.compile(r'^\(([０-９0-9]+)\)[　\s\t]*(.*)$'),
     lambda m: (_z2i(m.group(1)), m.group(2))),
    # 全角括弧漢数字 （一）（十一）。「（一般職）」のように漢数字以外を含むものは除く
    ("paren_zen_kanji",
     re.compile(r'^（([一二三四五六七八九十]+)）[　\s\t]*(.*)$'),
     lambda m: (kanji_to_int(m.group(1)), m.group(2))),
    # 半角括弧漢数字 (一)
    ("paren_han_kanji",
     re.compile(r'^\(([一二三四五六七八九十]+)\)[　\s\t]*(.*)$'),
     lambda m: (kanji_to_int(m.group(1)), m.group(2))),
    # 丸数字 ①〜㊿
    ("maru",
     re.compile(r'^([①-⑳㉑-㉟㊱-㊿])[　\s\t]*(.*)$'),
     lambda m: (_MARU_MAP.get(m.group(1), 0), m.group(2))),
    # 数字ドット 1. １．（ドット必須。スペースのみは項本文と紛らわしいため除外）
    ("num_dot",
     re.compile(r'^([０-９0-9]+)[．.]\s*(.*)$'),
     lambda m: (_z2i(m.group(1)), m.group(2))),
    # 漢数字＋区切り 一、二、 一　（スペース区切りも許容）
    ("kanji",
     re.compile(r'^([一二三四五六七八九十]+)[、，)）.．]\s*(.*)$|'
                r'^([一二三四五六七八九十]+)[　\s]+(.*)$'),
     lambda m: (kanji_to_int(m.group(1) or m.group(3)),
                m.group(2) if m.group(2) is not None else m.group(4))),
    # 全角括弧カタカナ （ア）（イ）
    ("paren_zen_kana",
     re.compile(r'^（([ア-ン])）[　\s\t]*(.*)$'),
     lambda m: (_katakana_index(m.group(1)), m.group(2))),
    # 半角括弧カタカナ (ア)
    ("paren_han_kana",
     re.compile(r'^\(([ア-ン])\)[　\s\t]*(.*)$'),
     lambda m: (_katakana_index(m.group(1)), m.group(2))),
    # 全角括弧英小文字 （a）
    ("paren_zen_alpha",
     re.compile(r'^（([a-zａ-ｚ])）[　\s\t]*(.*)$'),
     lambda m: (ord(m.group(1).lower()[0]) - ord('a') + 1
                if m.group(1) else 0, m.group(2))),
    # 半角括弧英小文字 (a)
    ("paren_han_alpha",
     re.compile(r'^\(([a-z])\)[　\s\t]*(.*)$'),
     lambda m: (ord(m.group(1)) - ord('a') + 1, m.group(2))),
    # 括弧ローマ数字 (i)(ii)
    ("paren_roman",
     re.compile(r'^\(([ivx]+)\)[　\s\t]*(.*)$'),
     lambda m: (_roman_to_int(m.group(1)), m.group(2))),
    # カタカナ＋区切り ア、 ア) ア　（スペース区切りも許容、1文字のみ）
    ("katakana",
     re.compile(r'^([ア-ン])[、，)）.．。]\s*(.*)$|'
                r'^([ア-ン])[　\s]+(.*)$'),
     lambda m: (_katakana_index(m.group(1) or m.group(3)),
                m.group(2) if m.group(2) is not None else m.group(4))),
    # 英小文字＋区切り a) a.
    ("alpha",
     re.compile(r'^([a-z])[)）.．][　\s\t]*(.*)$'),
     lambda m: (ord(m.group(1)) - ord('a') + 1, m.group(2))),
]


def _roman_to_int(s):
    vals = {'i':1,'v':5,'x':10}
    total = 0; prev = 0
    for c in reversed(s.lower()):
        v = vals.get(c, 0)
        total += -v if v < prev else v
        prev = max(prev, v)
    return total


def _katakana_index(c):
    """カタカナ文字の順序番号。アイウエオ順・イロハ順の両方を試す。"""
    if c in _KATAKANA_AIUEO:
        return _KATAKANA_AIUEO.index(c) + 1
    if c in _KATAKANA_IROHA:
        return _KATAKANA_IROHA.index(c) + 1
    return 0


# 記号の後ろがこれらで始まり、かつ前の行が読点で終わっていれば、
# 号ではなく本文中の参照（kisoku_parser.merge_lines で使う）
#   例:「…にかかわらず、／（１）に掲げる場合」「…のうち、／（１）は」
# 平仮名で始まる号（「（５）より良い職場…」）もあるので、記号の判定そのものには使わない
REF_AFTER_MARKER = re.compile(
    r'^(?:の|に|を|は|が|で(?!き)|から|まで|より|及び|又は|並びに|若しくは|ないし|乃至|[、，])')


def identify_symbol(text):
    """
    行頭の記号タイプを返す。
    返り値: (type_name, number, body) または (None, 0, text)
    """
    s = text.lstrip('　 \t')
    for type_name, pat, extractor in _SYMBOL_TYPES:
        m = pat.match(s)
        if m:
            num, body = extractor(m)
            return type_name, num, body.strip()
    return None, 0, s
