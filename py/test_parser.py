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


# --- 段落（行）の並びから docx を作って解析し、章・条・項・号の骨組みで比べる ---
import contextlib, io, os, tempfile
from docx import Document
import kisoku_parser


def _outline(lines):
    with tempfile.TemporaryDirectory() as w:
        path = os.path.join(w, 'x.docx')
        d = Document()
        for ln in lines:
            d.add_paragraph(ln)
        d.save(path)
        with contextlib.redirect_stdout(io.StringIO()):
            tree = kisoku_parser.parse_kisoku(path, use_gemini=False)
    out = []

    def items(lst, depth):
        for it in lst:
            out.append('  ' * depth + f"{it['number']}){it['body']}")
            items(it.get('sub_items', []), depth + 1)

    def article(a):
        note = f" 注釈={a['title_note']}" if a.get('title_note') else ''
        out.append(f"{a['number_str']}（{a['title']}）{note}")
        for p in a['paragraphs']:
            out.append(f"  {p['number']}:{p['body']}")
            items(p.get('items', []), 2)
    for n in tree:
        if n['type'] == 'chapter':
            out.append(f"{n['number_str']}|{n['title']}")
            out += [f"  前文:{t}" for t in n.get('preamble', [])]
            for a in n['articles']:
                article(a)
        elif n['type'] == 'article':
            article(n)
        elif n['type'] == 'back_matter':
            out.append(f"後付け:{n['texts'][0]}")
    return out


_CASES = [
    # 「第1条」だけの行の次の本文が、条ごと消えない（章の最初の条でも）
    (['第1章　総則', '第1条', '１　この規則は…定める。', '２　前項に…。', '第2条（定義）', '１．この規則で…いう。'],
     ['第1章|総則', '第1条（）', '  1:この規則は…定める。', '  2:前項に…。',
      '第2条（定義）', '  1:この規則で…いう。']),
    # 括弧書きの注記は次の条のタイトルにせず、消さずに直前の本文・号に戻す
    (['第1章　総則', '第1条（適用範囲）　正社員に適用する。', '（パートタイマーを除く）',
      '第2条　…による。', '（令和６年４月１日一部改定）', '第2章　採用',
      '第3条（書類）　次の書類を提出する。', '（１）履歴書', '（マイナンバーの記載がないもの）',
      '（試用期間）', '第4条　試用期間は3か月とする。'],
     ['第1章|総則', '第1条（適用範囲）', '  1:正社員に適用する。（パートタイマーを除く）',
      '第2条（）', '  1:…による。（令和６年４月１日一部改定）', '第2章|採用',
      '第3条（書類）', '  1:次の書類を提出する。', '    1)履歴書（マイナンバーの記載がないもの）',
      '第4条（試用期間）', '  1:試用期間は3か月とする。']),
    # 折り返し行頭の「第X章の…」「第X条（…）の…」「第 9 条 に…」「附則に…」は見出しにしない
    (['第1章　総則', '第1条（適用除外）　パートタイマーには、', '第3章の規定を適用しない。',
      '２　会社は、', '第10条（解雇）の規定にかかわらず、', '第 9 条 に定める手続による。',
      '３　嘱託は、', '附則に定める経過措置による。', '第2条（定義）　…。', '附則', 'この規則は…施行する。'],
     ['第1章|総則', '第1条（適用除外）', '  1:パートタイマーには、第3章の規定を適用しない。',
      '  2:会社は、第10条（解雇）の規定にかかわらず、第 9 条 に定める手続による。',
      '  3:嘱託は、附則に定める経過措置による。', '第2条（定義）', '  1:…。', '後付け:附則']),
    # 章タイトルのあとの説明文は章タイトルに混ぜない。注釈（※）は次の条に付けない
    (['第1章　総則', '［例］　規程例です。', '（目的）　※パターン1', '第1条　…。', '第2条（定義）　…。'],
     ['第1章|総則', '  前文:［例］　規程例です。', '第1条（目的） 注釈=※パターン1', '  1:…。', '第2条（定義）', '  1:…。']),
    # 項番号: タブ区切り、第10項以降、本文が数字で始まる「３．１ヶ月」
    (['第1条（a）　a。', '2\tb。', '３　c。', '４　d。', '５　e。', '６　f。', '７　g。', '８　h。', '９　i。',
      '１０　j。', '１１．１ヶ月k。'],
     ['第1条（a）', '  1:a。', '  2:b。', '  3:c。', '  4:d。', '  5:e。', '  6:f。', '  7:g。', '  8:h。',
      '  9:i。', '  10:j。', '  11:１ヶ月k。']),
    # 折り返し行頭の数量・小数・号の参照は、項・号にしない
    (['第1条（賃金）　割増賃金は通常の賃金の', '1.25倍とする。', '２　定年は満', '60 歳とする。',
      '３　契約期間は、', '2 年以内とする。', '４　次の事項を守る。', '（１）遅刻しない。',
      '（２）前号にかかわらず、', '（１）に掲げる場合は届け出る。'],
     ['第1条（賃金）', '  1:割増賃金は通常の賃金の1.25倍とする。', '  2:定年は満60 歳とする。',
      '  3:契約期間は、2 年以内とする。', '  4:次の事項を守る。', '    1)遅刻しない。',
      '    2)前号にかかわらず、（１）に掲げる場合は届け出る。']),
    # 号を「1.」「2.」で書いても2以降を項にしない。条見出しの同じ行の「1 週間」の1は消さない
    (['第1条（書類）　次の書類を提出する。', '1. 履歴書', '2. 住民票', '3. 誓約書', '２　前項の書類は…。',
      '第2条（週休）　1 週間の休日は2日とする。', '第3条（端数）　1.5 時間を超えるとき。', '第4条（目的）　１　この規則は…。'],
     ['第1条（書類）', '  1:次の書類を提出する。', '    1)履歴書', '    2)住民票', '    3)誓約書', '  2:前項の書類は…。',
      '第2条（週休）', '  1:1 週間の休日は2日とする。', '第3条（端数）', '  1:1.5 時間を超えるとき。',
      '第4条（目的）', '  1:この規則は…。']),
]
for _lines, _expected in _CASES:
    _got = _outline(_lines)
    assert _got == _expected, '\n'.join(['期待:'] + _expected + ['実際:'] + _got)


# 章の前文（条に属さない本文）が整形後の docx に残る
from apply_style import apply_style
import json
with tempfile.TemporaryDirectory() as _w:
    _src = os.path.join(_w, 'src.docx')
    _d = Document()
    for _ln in ['第1章　総則', '［例］　規程例です。', '第1条（目的）　…。']:
        _d.add_paragraph(_ln)
    _d.save(_src)
    with contextlib.redirect_stdout(io.StringIO()):
        _tree = kisoku_parser.parse_kisoku(_src, use_gemini=False)
    _jp = os.path.join(_w, 'x.json')
    with open(_jp, 'w', encoding='utf-8') as _f:
        json.dump(_tree, _f, ensure_ascii=False)
    _tpl = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates', '就業規則テンプレ1.docx')
    _dst = os.path.join(_w, 'out.docx')
    with contextlib.redirect_stdout(io.StringIO()):
        apply_style(json_path=_jp, template_path=_tpl, template_key='t1', output_path=_dst,
                    source_docx=_src, show_notice=False)
    _texts = [p.text for p in Document(_dst).paragraphs]
    assert '［例］　規程例です。' in _texts, _texts
print("OK (structure)")

# Gemini判定は本文扱いの段落で呼ばれ、章・条の判定は採用しない
import gemini_helper
_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_test_gemini.docx')
_d = Document()
_d.add_paragraph('第1条 (目的)')
_d.add_table(rows=1, cols=1)     # 表の直後の行は前の行と結合されない
_d.add_paragraph('【火曜日】')
_d.save(_path)
_calls = []
gemini_helper.is_available = lambda: True
gemini_helper.classify_with_gemini = lambda t, b, a: (_calls.append(t), 'article')[1]
kisoku_parser.apply_gemini_split = lambda *a, **k: None
try:
    with contextlib.redirect_stdout(io.StringIO()):
        _tree = kisoku_parser.parse_kisoku(_path, use_gemini=True)
finally:
    os.remove(_path)
assert _calls == ['【火曜日】'], _calls
assert sum(1 for n in _tree if n['type'] == 'article') == 1, _tree
print("OK (gemini)")
