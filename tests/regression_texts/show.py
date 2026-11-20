# -*- coding: utf-8 -*-
"""回帰テスト用の文章（tests/regression_texts/*.txt）を docx にして解析し、骨組みを表示する。

  python3 tests/regression_texts/show.py tests/regression_texts/<名前>.txt
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from outline_util import outline, parse_lines  # noqa: E402


def read_lines(path):
    """# で始まる説明の行を除いた、本文の行（docx の1段落ずつ）を返す。"""
    with open(path, encoding='utf-8') as f:
        return [ln.rstrip('\n') for ln in f if ln.strip() and not ln.startswith('#')]


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    print('\n'.join(outline(parse_lines(read_lines(sys.argv[1])))))
