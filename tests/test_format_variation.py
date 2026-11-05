# -*- coding: utf-8 -*-
"""
書式ゆらぎテスト：正しく解析できるサンプルの構造が、書式を崩しても変わらないかを見る。

  1. サンプルを解析し、その木から「標準形」の行を作る
  2. 標準形を解析した結果（S0）を基準にする
  3. 標準形に決まった規則でゆらぎを加えた行を解析した結果（S1）を S0 と比べる
     wrap        条・項・号の行を25〜40字で切って別の段落にする（PDF風）
     digit_space 数字の並びの前後に半角空白を入れる
     title_line  条タイトルを条の前の行に出す
     para_dot    第2項以降の番号を「２．」の形にする
     wrap_digit  wrap と digit_space の両方
  乱数は使わない。不一致は条ごとに数え、既知の不一致（tests/cases/format_variation_known.json）
  に無い条が不一致になったら終了コード1。

  python3 tests/test_format_variation.py                 実行
  python3 tests/test_format_variation.py --update-known  いまの不一致を既知として書き出す
"""
import json
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from outline_util import ROOT, make_docx, parse_path  # noqa: E402

SAMPLES = ['01_01_202601', '01_05_202206', 'kitei_01', 'shugyokisoku_sample']
VARIANTS = ['wrap', 'digit_space', 'title_line', 'para_dot', 'wrap_digit']
KNOWN = os.path.join(ROOT, 'tests', 'cases', 'format_variation_known.json')
WIDTHS = [25, 28, 31, 34, 37, 40]

H2F = str.maketrans('0123456789', '０１２３４５６７８９')
KATA = 'アイウエオカキクケコサシスセソタチツテトナニヌネノ'
DIGITS = re.compile(r'[0-9０-９]+(?:[.,．，][0-9０-９]+)*')


def fw(n):
    return str(n).translate(H2F)


def ws(s):
    return re.sub(r'\s+', '', s)        # \s は全角空白・タブも含む


# --- 木 → 標準形 ---------------------------------------------------------
# 標準形の行は (種類, 見出し部分, 本文, 条タイトル) の組。見出し部分は番号（とタイトル）。
def item_mark(depth, i):
    if depth == 0:
        return f'（{fw(i)}）'
    if depth == 1:
        assert i <= 20, '号の番号が20を超えた'
        return chr(0x2460 + i - 1)
    return KATA[i - 1] + '　'


def standard_lines(tree):
    out = []

    def items(lst, depth):
        assert depth <= 2, '号の深さが3を超えた'
        for i, it in enumerate(lst, 1):
            out.append(('item', item_mark(depth, i), it['body'], ''))
            items(it.get('sub_items', []), depth + 1)

    def article(a):
        title = a['title'].strip()
        head = a['number_str'] + (f'（{title}）' if title else '')
        paras = a['paragraphs']
        first = paras[0] if paras else None
        out.append(('article', head, first['body'] if first else '', title))
        if first:
            items(first.get('items', []), 0)
        for i, p in enumerate(paras[1:], 2):
            out.append(('para', fw(i), p['body'], ''))
            items(p.get('items', []), 0)

    for n in tree:
        if n['type'] == 'chapter':
            out.append(('chapter', n['number_str'], n['title'], ''))
            for a in n['articles']:
                article(a)
        elif n['type'] == 'article':
            article(n)
    return out


def render(kind, head, body):
    """標準形の1行の文字列。"""
    if kind == 'chapter':
        return f'{head}　{body}'
    if kind == 'article':
        return f'{head}　{body}' if body else head
    if kind == 'para':
        return f'{head}　{body}'
    return f'{head}{body}'          # 号


# --- ゆらぎ ------------------------------------------------------------
def spaced(s):
    return DIGITS.sub(lambda m: f' {m.group(0)} ', s)


def wrap_text(prefix, body, width):
    """prefix（見出し）の途中では切らず、width 字ごとに切る。最初の切れ目は prefix より後ろ。"""
    text = prefix + body
    if len(text) <= width:
        return [text]
    first = max(width, len(prefix))
    chunks = [text[:first]]
    rest = text[first:]
    while rest:
        chunks.append(rest[:width])
        rest = rest[width:]
    return chunks


def vary(lines, variant):
    out = []
    counter = 0
    for kind, head, body, title in lines:
        digit = variant in ('digit_space', 'wrap_digit')
        wrap = variant in ('wrap', 'wrap_digit')
        h, b = head, body
        if variant == 'para_dot' and kind == 'para':
            out.append(f'{head}．{body}')
            continue
        if digit:
            # 号の記号「（１）」の内側には空白を入れない（PDF由来でもこの崩れ方は少ない）
            h, b = (h if kind == 'item' else spaced(h)), spaced(b)
        if kind == 'chapter':
            out.append(render(kind, h, b))
            continue
        if variant == 'title_line' and kind == 'article' and title:
            out.append(f'（{title}）')
            h = head[:head.index('（')]
            out.append(render(kind, h, b))
            continue
        sep = '　' if kind in ('article', 'para') and b else ''
        if kind == 'para' and not b:
            sep = ''
        if wrap:
            width = WIDTHS[counter % len(WIDTHS)]
            counter += 1
            out += wrap_text(h + sep, b, width)
        else:
            out.append(render(kind, h, b))
    return out


# --- 比較 --------------------------------------------------------------
def structure(tree):
    """章（番号・タイトル）と、章の中の条（番号・タイトル・項・号の木）を比べる形にする。"""
    chapters = []
    arts = []       # [(ラベル, 内容)]
    seen = {}

    def flat(items, d, acc):
        for it in items:
            acc.append((d, ws(it['body'])))
            flat(it.get('sub_items', []), d + 1, acc)
        return acc

    for n in tree:
        if n['type'] != 'chapter':
            continue
        chapters.append((ws(n['number_str']), ws(n['title'])))
        for a in n['articles']:
            label = ws(a['number_str'])
            seen[label] = seen.get(label, 0) + 1
            key = label if seen[label] == 1 else f'{label}#{seen[label]}'
            paras = [(ws(p['body']), flat(p.get('items', []), 0, [])) for p in a['paragraphs']]
            arts.append((key, (label, ws(a['title']), paras)))
    return chapters, arts


def diff(base, other):
    """不一致の章・条のラベルの並び（基準の順、基準に無いものを後ろ）。"""
    bc, ba = structure(base)
    oc, oa = structure(other)
    bad = []
    for i, c in enumerate(bc):
        if i >= len(oc) or oc[i] != c:
            bad.append(c[0])
    for i in range(len(bc), len(oc)):
        bad.append(oc[i][0])
    od = dict(oa)
    for key, content in ba:
        if od.get(key) != content:
            bad.append(key)
    bd = dict(ba)
    bad += [key for key, _ in oa if key not in bd]
    return bad


def parse_lines_file(lines, w, name):
    path = os.path.join(w, name + '.docx')
    make_docx(lines, path)
    return parse_path(path)


def run():
    results = {}                # "サンプル/ゆらぎ" → 不一致のラベル
    warnings = []
    with tempfile.TemporaryDirectory() as w:
        for s in SAMPLES:
            orig = parse_path(os.path.join(ROOT, 'tests', 'samples', s + '.docx'))
            std = standard_lines(orig)
            base = parse_lines_file([render(*x[:3]) for x in std], w, s + '_s0')
            # 標準形の解析が元の解析と違うと、比べる意味が薄れるので知らせる
            miss = diff(orig, base)
            if miss:
                warnings.append(f'{s}: 標準形(S0)が元の解析結果と違う条: {miss}')
            for v in VARIANTS:
                other = parse_lines_file(vary(std, v), w, f'{s}_{v}')
                results[f'{s}/{v}'] = diff(base, other)
    return results, warnings


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    results, warnings = run()
    for m in warnings:
        print('[注意]', m)
    if '--update-known' in sys.argv:
        known = {k: v for k, v in results.items() if v}
        with open(KNOWN, 'w', encoding='utf-8', newline='\n') as f:
            json.dump(known, f, ensure_ascii=False, indent=1)
            f.write('\n')
        print(f'既知の不一致を書き出しました: {KNOWN}（{len(known)}組）')
        return 0
    known = json.load(open(KNOWN, encoding='utf-8')) if os.path.exists(KNOWN) else {}
    new_total = 0
    for key, bad in results.items():
        kn = known.get(key, [])
        new = [b for b in bad if b not in kn]
        fixed = [k for k in kn if k not in bad]
        if new:
            new_total += len(new)
            print(f'[新しい不一致] {key}: {new}')
        if fixed:
            print(f'[直った] {key}: {fixed}（format_variation_known.json から外してください）')
    print('組み合わせごとの不一致の条の数（既知を含む）:')
    total = 0
    for s in SAMPLES:
        print('  ' + s + '  ' + ' '.join(f'{v}={len(results[f"{s}/{v}"])}' for v in VARIANTS))
        total += sum(len(results[f'{s}/{v}']) for v in VARIANTS)
    print(f'合計 {total} / うち新しい不一致 {new_total}')
    return 1 if new_total else 0


if __name__ == '__main__':
    sys.exit(main())
