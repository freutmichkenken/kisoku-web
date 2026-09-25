# -*- coding: utf-8 -*-
"""crossref._article_number_para の最小検証。 python3 py/test_crossref.py で実行。"""

from types import SimpleNamespace

from crossref import _article_number_para


def _run(levels, idx):
    paragraphs = [SimpleNamespace(_p=i) for i in range(len(levels))]
    nums = SimpleNamespace(get=lambda p: levels[p])
    return _article_number_para(paragraphs, idx, nums)


# テンプレ1・4型：前条の第1項 / 見出し（第N条）/ 第1項 → 見出しへ移す
assert _run([(2, '第%3条'), (3, '%4'), (2, '第%3条'), (3, '%4')], 3) == 2
# テンプレ2・3型：前条の第1項（第N条）/ 見出し / 第1項（第N条） → 前条ではなく第1項へ移す
assert _run([(3, '第%4条'), (2, '%3　'), (3, '第%4条')], 1) == 2
# 番号の無い段落は動かさない（次条の見出しへ移らない）
assert _run([(2, '第%3条'), None, (2, '第%3条')], 1) == 1
print("OK")
