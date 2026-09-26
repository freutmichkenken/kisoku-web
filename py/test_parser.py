# -*- coding: utf-8 -*-
"""kisoku_parser の最小検証。 python3 py/test_parser.py で実行。"""

from kisoku_parser import classify_line, merge_lines

# 番号の前後に空白がある「第 24 条 (…)」も条見出しとして分けて認識する
_m = merge_lines([('text', '…できるものとする。'), ('text', '第 35 条 (休職期間)')])
assert len(_m) == 2, _m
_n = classify_line('第 24 条 (制裁の種類、程度)')
assert (_n['type'], _n['number_str'], _n['title']) == ('article', '第24条', '制裁の種類、程度'), _n
# 行頭の条文参照「第 9 条に定める…」は条見出しにしない
assert classify_line('第 9 条に定める手続による。')['type'] != 'article'
print("OK")
