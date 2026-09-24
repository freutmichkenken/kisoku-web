# -*- coding: utf-8 -*-
"""
記号階層の動的判定モジュール

就業規則の項・号・その下位に使われる記号（①②、（１）、ア、一、a) 等）は
文書ごとにまちまちだが、「同一文書内では記号と階層が一貫する」という前提のもと、
記号の初出順序と字下げの深さから、各記号がどの階層かを動的に決定する。

使い方:
    from marker_hierarchy import MarkerAnalyzer
    analyzer = MarkerAnalyzer()
    analyzer.learn(list_of_(text, indent_width))   # 文書全体を学習
    kind = analyzer.classify(text)  # 'item' / 'item_sub' / 'item_sub2' / None
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

_KANJI_NUM = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10}

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
     re.compile(r'^\(([0-9]+)\)[　\s\t]*(.*)$'),
     lambda m: (_z2i(m.group(1)), m.group(2))),
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
                r'^([一二三四五六七八九十])[　\s]+(.*)$'),
     lambda m: (_KANJI_NUM.get(m.group(1) or m.group(3), 0),
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


# ------------------------------------------------------------
# 階層アナライザ
# ------------------------------------------------------------

class MarkerAnalyzer:
    """
    文書全体を学習し、各記号タイプが階層のどの深さかを決定する。

    深さの決め方:
      1. 記号タイプごとの平均字下げ幅を集計
      2. 字下げが浅い順に「号(item) → 号の下位(item_sub) → item_sub2 ...」を割り当て
      3. 字下げが同じ/取れない場合は、初出順を使う
    """

    def __init__(self):
        self.type_depth = {}   # type_name -> 深さ(0=号, 1=号の下位, 2=...)
        self._learned = False

    def learn(self, lines):
        """
        lines: [(text, indent_width), ...]
               text が None の要素は「条の切れ目」を表し、
               そこで入れ子の状態をリセットする。
               （別々の条で ① と （１） が同じ号レベルに使われる場合に、
                 それらを親子と誤解しないため）

        階層の決め方（ネスト検出）:
          文書を上から読み、記号タイプの入れ子関係で深さを決める。
          - 番号が1で始まる新タイプ → 現在のスタックの一段深い子
          - 既にスタックにあるタイプに戻る → そのタイプまでスタックを巻き戻す
          - 番号2以上で初出のタイプ → 既存系列の途中とみなし、その時点の深さで登録
        字下げは補助的に使う（同深さの曖昧性解消）。
        """
        first_seen = {}
        seq = []
        for idx, (text, indent) in enumerate(lines):
            if text is None:          # 条の切れ目
                seq.append((None, 0, 0))
                continue
            tname, num, _ = identify_symbol(text)
            if tname is not None and tname not in first_seen:
                first_seen[tname] = idx
            seq.append((tname, num, indent))

        if not first_seen:
            self._learned = True
            return

        depth = {}
        active_stack = []   # 現在開いている記号タイプ（浅い→深い）

        for tname, num, indent in seq:
            if tname is None:
                # 条が変わったので入れ子状態をリセット
                active_stack = []
                continue

            if tname in active_stack:
                # 既出タイプに戻った → そのタイプより深い子を閉じる
                while active_stack and active_stack[-1] != tname:
                    active_stack.pop()
                # 深さは既に確定済み
                continue

            # 新しいタイプ
            if num <= 1:
                # 新系列の開始 → 現在の最深の一段下に積む
                active_stack.append(tname)
                depth[tname] = len(active_stack) - 1
            else:
                # 番号2以上で初出 → 既存系列の途中。現在のスタック末尾と同じ深さに置く
                # （新たな子ではなく、既存の深さレベルの別記法とみなさない方が安全なので
                #   スタックに積んで現在深さで登録する）
                active_stack.append(tname)
                depth[tname] = len(active_stack) - 1

        # 未登録タイプを初出順で補完
        nxt = (max(depth.values()) + 1) if depth else 0
        for tname in sorted(first_seen, key=lambda t: first_seen[t]):
            if tname not in depth:
                depth[tname] = nxt
                nxt += 1

        self.type_depth = depth
        self._learned = True

    def depth_to_kind(self, depth):
        """深さを種別名に変換。0=item, 1=item_sub, 2=item_sub2, ..."""
        if depth <= 0:
            return 'item'
        if depth == 1:
            return 'item_sub'
        return f'item_sub{depth}'

    def classify(self, text):
        """
        行を分類。
        返り値: (kind, number, body) または (None, 0, text)
          kind: 'item' / 'item_sub' / 'item_sub2' / ...
        """
        tname, num, body = identify_symbol(text)
        if tname is None:
            return None, 0, text
        depth = self.type_depth.get(tname, 0)
        return self.depth_to_kind(depth), num, body
