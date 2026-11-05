# -*- coding: utf-8 -*-
"""
誤判定例（tests/cases/misdetection_cases.json）の解析結果を、期待する構造と比べる。

  python3 tests/test_misdetection.py              全例を比べる
  python3 tests/test_misdetection.py --show NAME  1例の実際の骨組みを表示する

  expected_outline  期待する正しい骨組み
  known_failure     まだ直っていない例（値は直す予定の段階）。不一致でもCIは落とさない。
                    一致したら「直った」と表示する（known_failure を外す）
  pending_spec      正しい構造が仕様の判断待ち。比べない
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from outline_util import ROOT, outline, parse_lines  # noqa: E402

CASES = os.path.join(ROOT, 'tests', 'cases', 'misdetection_cases.json')


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    cases = json.load(open(CASES, encoding='utf-8'))['cases']
    if len(sys.argv) >= 3 and sys.argv[1] == '--show':
        c = next((c for c in cases if c['name'] == sys.argv[2]), None)
        if c is None:
            print('例が見つかりません:', sys.argv[2])
            return 1
        print('\n'.join(outline(parse_lines(c['lines']))))
        return 0
    ok = known = pending = fixed = new = 0
    for c in cases:
        if 'pending_spec' in c:
            pending += 1
            continue
        exp = c.get('expected_outline')
        if exp is None:
            print(f"[新しい失敗] {c['name']}: expected_outline がありません")
            new += 1
            continue
        got = outline(parse_lines(c['lines']))
        if got == exp:
            if 'known_failure' in c:
                fixed += 1
                print(f"[直った] {c['name']}（known_failure を外してください）")
            else:
                ok += 1
        elif 'known_failure' in c:
            known += 1
        else:
            new += 1
            print(f"[新しい失敗] {c['name']}: {c.get('note', '')}")
            print('  期待:')
            print('\n'.join('    ' + x for x in exp))
            print('  実際:')
            print('\n'.join('    ' + x for x in got))
    print(f"合格 {ok} / 既知の失敗 {known} / 仕様未決 {pending} / 直った {fixed} / 新しい失敗 {new}")
    return 1 if new else 0


if __name__ == '__main__':
    sys.exit(main())
