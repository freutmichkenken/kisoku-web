# -*- coding: utf-8 -*-
"""crossref の最小検証。 python3 py/test_crossref.py で実行。"""

from types import SimpleNamespace

from crossref import _article_number_para, _convert_typed_refs


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

# 手打ちの「第2条」を囲んだ参照は \n 付きに作り替え、「第2条第1項」は作り替えない
_info = {"bookmarks": {"_R": {"para": "第2条　本文", "start": 0, "end": 3}},
         "fields": [{"instr": " REF _R \\h ", "result": "第2条", "bookmark": "_R", "para": "x"},
                    {"instr": " REF _R \\h ", "result": "第2条第1項", "bookmark": "_R", "para": "x"}]}
_fields, _skipped = _convert_typed_refs(_info)
assert _fields[0]["instr"] == " REF _R \\n \\h ", _fields[0]["instr"]
assert _fields[1]["instr"] == " REF _R \\h " and len(_skipped) == 1
print("OK")
