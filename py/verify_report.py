# -*- coding: utf-8 -*-
"""
原本との照合レポート

自動処理には限界があるので「必ず元ファイルと照合してください」と
案内していますが、その照合を機械にやらせるモジュールです。

やること
--------
元docxと整形後docxのテキストを正規化して突き合わせ、

  1. 原本にあって出力に無い文章（欠落）
  2. 出力にあって原本に無い文章（想定外の増加・AIの創作）
  3. 出力で回数が増えている文章（重複）

を一覧にします。判定は完全に機械的で、AIは使いません。

正規化の考え方
--------------
整形後は番号がスタイルの自動採番になるため、テキストとしては
番号が存在しません。一方、原本は「第１条　（目的）」のように
手打ちされています。そのまま比べると全部が差分になるので、

  - NFKC正規化（全角/半角・丸数字を吸収）
  - 行頭の番号（第X条、２、(1)、①、ア. など）を除去
  - 空白をすべて除去

してから比較します。

段落の分割・結合には強い作りです。全文を連結した1本の文字列に対して
部分一致を見るので、条が2段落に分かれても、逆に結合されても、
「文章が残っているか」だけを見ます。
"""

import re
import unicodedata
from collections import Counter

from docx.oxml.ns import qn

# 行頭の番号らしきもの（NFKC後の半角前提）
_NUM_PREFIX = re.compile(
    r'^(?:'
    r'第[0-9一二三四五六七八九十百千]+(?:条|章|節|款|編|項|号)(?:の[0-9]+)?'
    r'|[0-9]{1,3}(?:[.)、,:：]|(?=\s))'
    r'|\([0-9]{1,3}\)'
    r'|\([一二三四五六七八九十]{1,3}\)'
    r'|[ア-ヶ][.)、]'
    r'|\([ア-ヶ]\)'
    r'|[a-zA-Z][.)]'
    r'|\([a-zA-Z]\)'
    r'|[ivxIVX]{1,4}[.)]'
    r'|[一二三四五六七八九十]{1,3}[、.]'
    r')[\s]*'
)

_WS = re.compile(r'[\s\u3000\u200b]+')

# 丸数字などの記号は NFKC で普通の数字になってしまい、
# 「①以外の…」のように区切りが無いと番号と判別できなくなるため、
# NFKC より前に落としておく。
_MARKER_PREFIX = re.compile(
    r'^[\s\u3000]*[\u2460-\u2473\u3251-\u325f\u32b1-\u32bf'
    r'\u24b6-\u24e9\u2170-\u2179\u2160-\u2169]+[\s\u3000]*')


def normalize(text):
    """比較用にテキストを正規化する。"""
    if not text:
        return ""
    s = text
    # 丸数字・囲み文字の行頭記号を先に落とす
    for _ in range(3):
        s2 = _MARKER_PREFIX.sub('', s)
        if s2 == s:
            break
        s = s2
    s = unicodedata.normalize('NFKC', s)
    # 行頭の番号を（入れ子になっていることもあるので）繰り返し落とす
    for _ in range(3):
        s2 = _NUM_PREFIX.sub('', s.lstrip())
        if s2 == s:
            break
        s = s2
    s = _WS.sub('', s)
    # 括弧の種類差を吸収
    s = s.replace('｢', '「').replace('｣', '」')
    return s


# ============================================================
# テキストの取り出し
# ============================================================

def _p_text(p):
    """
    w:p の表示テキスト。段落内の改行（w:br / w:cr）は "\n"、
    タブは "\t" にする。1つの段落に複数の論理行が同居していることが
    あるため、あとで行に分けて比較できるようにしておく。
    """
    out = []
    for el in p.iter():
        if el.tag == qn('w:t'):
            out.append(el.text or '')
        elif el.tag in (qn('w:br'), qn('w:cr')):
            out.append('\n')
        elif el.tag == qn('w:tab'):
            out.append('\t')
    return ''.join(out)


def _is_toc_paragraph(p):
    """
    目次の段落か。スタイル名ではなく、目次特有のフィールドで判定する。
    （原本の目次はパーサが除外し、出力の目次はツールが挿入するため、
      どちらも比較対象から外す）
    """
    for el in p.iter(qn('w:instrText')):
        s = (el.text or '').upper()
        if 'PAGEREF' in s or ' TOC ' in f' {s} ':
            return True
    return False


def extract(doc, skip_toc=True, extra_skip=()):
    """
    docx から段落テキストを取り出す（表の中の段落も含む）。
    戻り値: [(元のテキスト, 正規化テキスト), ...]（正規化後が空のものは除く）
    """
    out = []
    for p in doc.element.body.iter(qn('w:p')):
        if skip_toc and _is_toc_paragraph(p):
            continue
        # 段落内に改行があると、1つの w:p に条見出しと本文が同居する。
        # 整形後は別々の段落になるので、行に分けてから比較する。
        for raw in _p_text(p).split('\n'):
            norm = normalize(raw)
            if not norm:
                continue
            if norm in extra_skip:
                continue
            out.append((raw, norm))
    return out


# ツールが挿入する目次まわりの定型文（比較対象から外す）
_TOOL_TEXTS = {
    normalize("目　次"),
    normalize("【目次】Wordで開いたときに自動作成されます"
              "（作成されない場合は Ctrl+A → F9）"),
    normalize("【目次】Ctrl+A → F9 を押すと目次が作成されます"),
}


# ============================================================
# 照合
# ============================================================

def compare(src_doc, out_doc):
    """原本と出力を突き合わせる。"""
    src = extract(src_doc)
    out = extract(out_doc, extra_skip=_TOOL_TEXTS)

    src_all = ''.join(n for _, n in src)
    out_all = ''.join(n for _, n in out)

    # --- 欠落: 原本の文章が出力のどこにも無い ---
    missing = [(raw, norm) for raw, norm in src if norm not in out_all]
    # --- 増加: 出力の文章が原本のどこにも無い ---
    added = [(raw, norm) for raw, norm in out if norm not in src_all]

    # --- 重複: 出力での出現回数が原本より多い ---
    # 原本にも同じ段落が存在する場合だけを見る。原本に0回のものは、
    # 段落の分割・結合でできた文字列（連結すれば原本にある）なので
    # 重複ではない。
    cs = Counter(n for _, n in src)
    co = Counter(n for _, n in out)
    dup = []
    for norm, c in co.items():
        c_src = cs.get(norm, 0)
        if c_src >= 1 and c > c_src and len(norm) >= 8:
            raw = next(r for r, n in out if n == norm)
            dup.append((raw, c_src, c))

    return {
        "src_paragraphs": len(src),
        "out_paragraphs": len(out),
        "src_chars": len(src_all),
        "out_chars": len(out_all),
        "missing": missing,
        "added": added,
        "duplicated": dup,
    }


_NUMS = r'[0-9〇一二三四五六七八九十百千]+'     # 全角の数字は normalize（NFKC）で半角になっている
_HEAD_LABEL = re.compile(rf'^第{_NUMS}[章条節](の{_NUMS})*')
_MIN_PARTIAL = 8


def tree_texts(tree):
    """解析済みの木から、本文の文章（条タイトル・項・号・その下）を正規化して集める。"""
    out = set()

    def walk(node):
        if not isinstance(node, dict):
            return
        for key in ("title", "body"):
            n = normalize(node.get(key) or "")
            if n:
                out.add(n)
        # 図の中の文字（テキストボックスなど）も原本の段落として照合されるので集める
        for dw in node.get("drawings") or []:
            n = normalize(dw.get("text") or "") if isinstance(dw, dict) else ""
            if n:
                out.add(n)
        for key in ("articles", "paragraphs", "items", "sub_items", "sub_items2"):
            for child in node.get(key) or []:
                walk(child)

    for node in tree or []:
        if node.get("type") in ("chapter", "article"):
            walk(node)
    return out


def split_edited(result, base_tree, edited_tree):
    """
    画面で編集した文章を、欠落・増加から「編集で消した文章」「編集で足した文章」へ移す。
    編集で変えた文章まで「整形で落ちた」と報告しないため。
    原本の1段落が木では行の結合で長くなることがあるので、含まれていれば編集の分とみなす。
    """
    base, edited = tree_texts(base_tree), tree_texts(edited_tree)
    removed, added = base - edited, edited - base

    def hit(norm, pool, other):
        """
        pool（編集で消えた／足した文章）に当たり、other（反対側の木の文章）には無いか。
        - 反対側の木にまだ含まれている文章は、編集で消した／足したものではない（本当の欠落・増加）
        - 一部が含まれるだけで編集の分とみなすのは、短い語句が偶然含まれるのを避けるため8文字以上に限る
        """
        # 原本の段落は「第1条この規則は…」のように番号と本文が続いていることがある（木の本文には番号が無い）
        norm = _HEAD_LABEL.sub("", norm)
        if not norm or any(norm in o for o in other):
            return False
        return norm in pool or (len(norm) >= _MIN_PARTIAL and any(norm in p for p in pool))

    def take(items, pool, other):
        ok = [it for it in items if hit(it[1], pool, other)]
        rest = [it for it in items if not hit(it[1], pool, other)]
        return ok, rest

    result["edit_removed"], result["missing"] = take(result["missing"], removed, edited)
    result["edit_added"], result["added"] = take(result["added"], added, base)
    return result


# ============================================================
# 出力
# ============================================================

def summary_lines(result, limit=5):
    """画面表示用の要約。"""
    lines = []
    d = result["out_chars"] - result["src_chars"]
    lines.append(f"  原本 {result['src_chars']:,}文字 / "
                 f"出力 {result['out_chars']:,}文字（差 {d:+,}）")
    n_m, n_a, n_d = (len(result["missing"]), len(result["added"]),
                     len(result["duplicated"]))
    edited = _edited_summary(result)
    if not (n_m or n_a or n_d):
        lines.append("  ★ 欠落・増加ともに検出されませんでした")
        if edited:
            lines.append(edited)
        return lines
    if edited:
        lines.append(edited)
    if n_m:
        lines.append(f"  ★ 原本にあって出力に無い文章: {n_m}件")
        for raw, _ in result["missing"][:limit]:
            lines.append(f"      - {raw.strip()[:40]}")
        if n_m > limit:
            lines.append(f"      …ほか {n_m - limit}件")
    if n_a:
        lines.append(f"  ★ 出力にあって原本に無い文章: {n_a}件")
        for raw, _ in result["added"][:limit]:
            lines.append(f"      + {raw.strip()[:40]}")
        if n_a > limit:
            lines.append(f"      …ほか {n_a - limit}件")
    if n_d:
        lines.append(f"  ★ 出力で回数が増えている文章: {n_d}件")
        for raw, c_src, c_out in result["duplicated"][:limit]:
            lines.append(f"      × {raw.strip()[:34]}（原本{c_src}回→出力{c_out}回）")
        if n_d > limit:
            lines.append(f"      …ほか {n_d - limit}件")
    lines.append("  → 詳細は照合レポートをご確認ください")
    return lines


def _edited_summary(result):
    n_r, n_a = len(result.get("edit_removed") or []), len(result.get("edit_added") or [])
    return f"  （画面での編集による変更: 消した文章 {n_r}件・足した文章 {n_a}件）" if n_r or n_a else ""


def write_report(result, path, src_name="", out_name=""):
    """照合レポートを Markdown で書き出す。"""
    L = []
    L.append("# 原本との照合レポート\n")
    if src_name or out_name:
        L.append(f"- 原本: `{src_name}`")
        L.append(f"- 出力: `{out_name}`\n")
    L.append("## 集計\n")
    L.append("| | 原本 | 出力 |")
    L.append("|---|---:|---:|")
    L.append(f"| 段落数 | {result['src_paragraphs']:,} | "
             f"{result['out_paragraphs']:,} |")
    L.append(f"| 文字数（正規化後） | {result['src_chars']:,} | "
             f"{result['out_chars']:,} |")
    L.append("")

    def section(title, items, fmt, note):
        L.append(f"## {title}（{len(items)}件）\n")
        if not items:
            L.append("検出されませんでした。\n")
            return
        L.append(note + "\n")
        for it in items:
            L.append(fmt(it))
        L.append("")

    section("原本にあって出力に無い文章", result["missing"],
            lambda it: f"- {it[0].strip()}",
            "**整形の過程で落ちた可能性があります。**"
            "原本の該当箇所を確認してください。")
    section("出力にあって原本に無い文章", result["added"],
            lambda it: f"- {it[0].strip()}",
            "原本に無い文章です。AIによる再構成が入った条や、"
            "行の結合位置がずれた箇所で出ることがあります。")
    section("出力で回数が増えている文章", result["duplicated"],
            lambda it: f"- {it[0].strip()}"
                       f"（原本 {it[1]}回 → 出力 {it[2]}回）",
            "同じ文章が原本より多く出ています。"
            "図表の複製が重複した可能性があります。")
    if "edit_removed" in result:
        section("画面の編集で消した文章", result["edit_removed"],
                lambda it: f"- {it[0].strip()}",
                "画面で編集・削除した文章です。上の「出力に無い文章」には数えていません。")
        section("画面の編集で足した文章", result["edit_added"],
                lambda it: f"- {it[0].strip()}",
                "画面で編集・追加した文章です。上の「原本に無い文章」には数えていません。")

    L.append("---\n")
    L.append("## この照合について\n")
    L.append("- 番号（第X条・①など）は自動採番になるため、"
             "比較前に行頭の番号を除去しています")
    L.append("- 全角/半角・空白・丸数字の違いは吸収しています")
    L.append("- 段落の分割・結合は差分になりません"
             "（文章が残っていれば「あり」と判定します）")
    L.append("- 目次は原本・出力とも比較対象外です")
    L.append("- **差分ゼロでも、順序の入れ替わりや階層の誤りは"
             "検出できません。最終確認は目視でお願いします**")

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return path
