# -*- coding: utf-8 -*-
"""テスト共通：行から一時docxを作って解析し、骨組み（outline）にする。"""
import contextlib
import io
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'py'))

from docx import Document  # noqa: E402
import kisoku_parser  # noqa: E402


def make_docx(lines, path):
    d = Document()
    for ln in lines:
        if ln.startswith('@') and '|' in ln:
            style, text = ln[1:].split('|', 1)
            d.add_paragraph(text, style=style)
        else:
            d.add_paragraph(ln)
    d.save(path)


def parse_path(path):
    with contextlib.redirect_stdout(io.StringIO()):
        return kisoku_parser.parse_kisoku(path, use_gemini=False)


def parse_lines(lines):
    with tempfile.TemporaryDirectory() as w:
        path = os.path.join(w, 'x.docx')
        make_docx(lines, path)
        return parse_path(path)


def outline(tree):
    out = []

    def items(lst, d):
        for it in lst:
            out.append('  ' * d + f"{it['number']}){it['body']}")
            items(it.get('sub_items', []), d + 1)

    def art(a):
        note = f" 注釈={a['title_note']}" if a.get('title_note') else ''
        out.append(f"{a['number_str']}（{a['title']}）{note}")
        for p in a['paragraphs']:
            out.append(f"  {p['number']}:{p['body']}")
            items(p.get('items', []), 2)
    for n in tree:
        t = n['type']
        if t == 'chapter':
            out.append(f"{n['number_str']}|{n['title']}")
            out += [f"  前文:{x}" for x in n.get('preamble', [])]
            for a in n['articles']:
                art(a)
        elif t == 'article':
            art(n)
        elif t == 'front_matter':
            out.append('前付け:' + ' / '.join(n['texts']))
        elif t == 'back_matter':
            out.append('後付け:' + ' / '.join(n['texts']))
        else:
            out.append(t)
    return out
