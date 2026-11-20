# -*- coding: utf-8 -*-
"""
回帰テスト用の文章（tests/regression_texts/*.txt）を docx にして解析し、
ファイルに書いた期待する骨組み（#> の行）と要確認の印（#hint> の行）に一致するか確かめる。

  python3 tests/test_regression_texts.py

#> の行が無いファイル（heading_like_references.txt など、別のテストが読むもの）は比べない。
期待を変えるときは、意図した変更かをユーザーに確かめてから、ファイルの #> / #hint> の行を直す。
"""
import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from outline_util import ROOT, outline, parse_lines  # noqa: E402

TEXTS = os.path.join(ROOT, 'tests', 'regression_texts')
NO_HINT = '（なし）'


def load(path):
    """(本文の行, 期待する骨組み or None, 期待する印 or None)。"""
    body, expect, hints = [], [], []
    with open(path, encoding='utf-8') as f:
        for ln in f:
            ln = ln.rstrip('\n')
            if ln.startswith('#> '):
                expect.append(ln[3:])
            elif ln.startswith('#hint> '):
                if ln[7:] != NO_HINT:
                    hints.append(ln[7:])
            elif ln.strip() and not ln.startswith('#'):
                body.append(ln)
    return body, (expect or None), (hints if expect else None)


def article_hints(tree):
    """木のどこかに付いた _article_hint を、出てくる順に集める。"""
    out = []

    def walk(n):
        if isinstance(n, dict):
            out.extend(n.get('_article_hint', []))
            for v in n.values():
                if isinstance(v, list):
                    for x in v:
                        walk(x)
    for n in tree:
        walk(n)
    return out


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    files = sorted(glob.glob(os.path.join(TEXTS, '*.txt')))
    ok = skipped = 0
    failed = []
    for path in files:
        name = os.path.basename(path)
        body, expect, hints = load(path)
        if expect is None:
            skipped += 1
            continue
        if not body:
            failed.append((name, '本文の行がありません', [], []))
            continue
        tree = parse_lines(body)
        got = outline(tree)
        got_hints = article_hints(tree)
        if got != expect:
            failed.append((name, '骨組みが違います', expect, got))
        elif got_hints != hints:
            failed.append((name, '要確認の印が違います', hints, got_hints))
        else:
            ok += 1
    for name, why, exp, got in failed:
        print(f'[失敗] {name}: {why}')
        print('  期待:')
        print('\n'.join('    ' + x for x in exp))
        print('  実際:')
        print('\n'.join('    ' + x for x in got))
    print(f'合格 {ok} / 失敗 {len(failed)} / 比べないファイル {skipped}')
    # ファイルが1つも読めなければ、テストが空回りしているので失敗にする
    if not files or ok + len(failed) == 0:
        print('比べるファイルがありません:', TEXTS)
        return 1
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
