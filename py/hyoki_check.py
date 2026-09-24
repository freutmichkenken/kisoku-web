# -*- coding: utf-8 -*-
"""
表記ゆれ検出モジュール（ルールベース）

kisoku_parser.py が生成した JSON の本文テキストを走査し、
同一文書内で表記がゆれている語を検出します。
apply_style.py がこの結果を使って、該当箇所を青字にし、
段落ごとにWordコメントを付けます。

設計方針
--------
- AIは使いません。判定は「辞書」と「機械的な正規化」の2本立てです。
  同じ文書を何度整形しても同じ結果になることを優先しています。
- 検出は2段構えです。
    1) 辞書グループ … 既知のゆれ（及び/および、従業員/社員 など）
    2) 自動グルーピング … 漢字骨格＋末尾送り仮名が一致する異表記
       （取扱い/取り扱い/取扱、申出/申し出 など）を機械的に発見
- 「多数派＝正しい」とは決めつけず、ゆれているグループの表記は
  多数派・少数派を問わず**すべて**マークします。
  どちらに統一するかは人間が決めるべきという考え方です。

使い方
------
    from hyoki_check import analyze
    report = analyze(data)          # data = JSON(list)
    spans = report.scan("〜本文〜")  # 描画する文字列に対して位置を取る
    print(report.summary_lines())

なぜ「JSONに位置を書き込む」方式にしないのか
--------------------------------------------
条見出しは出力時に「（　）」で囲まれるなど、JSONの文字列と
実際に出力される文字列が一致しません。位置を先に固定すると
ずれるため、**出力直前の確定文字列に対して再スキャンする**
方式にしています。
"""

import re
from collections import Counter, defaultdict

# ------------------------------------------------------------
# 文字クラス
# ------------------------------------------------------------
KANJI = r'\u4e00-\u9fff\u3005'      # 漢字（々を含む）
HIRA = r'\u3041-\u3096'             # ひらがな
KATA = r'\u30a1-\u30f6\u30fc'       # カタカナ（長音符を含む）

_RE_KANJI_OR_KATA = re.compile(rf'[{KANJI}{KATA}]')

# skip_prev / skip_next で前後を見る文字数。
# 文書全体をスライスすると O(文字数^2) になるため固定長に制限する。
# （skip_prev は末尾一致 `$`、skip_next は先頭一致 `^` の指定を前提）
_CTX_WINDOW = 16

# 複合動詞の連用形などに現れる「中間の送り仮名」として許可するもの。
# 「休職の期間」の「の」のような助詞を混ぜないための安全弁。
ALLOWED_MID = set("りしちきいみびぎじえけめげ")

# 「取扱」と「取扱い」のように、送り仮名の有無だけが違う場合に
# 同一グループとみなす末尾かな。活用語尾（う・る・す等）は入れない。
ALLOWED_TAIL_OPT = set("いりけめげえき")

# 語末の送り仮名として認めるかな（名詞化する i段・e段のみ）。
# 「を」「は」などの助詞を巻き込まないための制限。
ALLOWED_TAIL = "いりけめげえきしちみびぎじ"

# ------------------------------------------------------------
# 名詞用法／動詞用法の切り分け
#   公用文の慣行では、複合の語は名詞なら送り仮名を省き（申出・取扱・
#   支払・届出）、動詞なら送り仮名を付ける（申し出る・取り扱う）。
#   つまり「申出」と「申し出る」は表記ゆれではないので、
#   動詞として使われている箇所は検出対象から外す。
# ------------------------------------------------------------

# 直後がこのかなで始まれば動詞の活用とみなす
# （う音便・促音便・命令形・下一段の語尾など。助詞は入れない）
VERB_HEAD = set("うわっえおるれらろよくぐすつぬぶむずたて")

# 連用形（i段）のあとに続けば動詞とみなす助動詞の頭
RENYO_HEAD = set("またてだでしょ")

# 連用形になりうる送り仮名（i段・e段）
I_ROW = set("いりきぎしちみびじえけげめせね")

# 「な」で始まる場合、否定の助動詞（ない・なかった・なければ）なら動詞。
# 「支払など」「取扱なし」のような名詞用法と区別するため次の1文字を見る。
NAI_FOLLOW = set("いかけっ")


def _kana_run(text, pos, limit=8):
    """pos から続くひらがなの並びを返す（先頭 limit 文字まで見れば足りる）。"""
    m = re.match(rf'[{HIRA}]*', text[pos:pos + limit])
    return m.group(0) if m else ""


def _is_verb_usage(text, form, i):
    """
    text の位置 i にある form が、動詞として使われているか判定する。
      「申し出て」「支払います」「取り扱った」→ True（動詞）
      「申し出があった」「支払を行う」「取扱書」  → False（名詞）
    """
    # form の中で最後の漢字がどこかを求め、その直後からのかなを見る
    last_kanji = -1
    for idx, ch in enumerate(form):
        if re.match(rf'[{KANJI}]', ch):
            last_kanji = idx
    if last_kanji < 0:
        return False
    run = _kana_run(text, i + last_kanji + 1)
    if not run:
        return False                      # 漢字・記号・文末が続く → 名詞
    if run[0] in VERB_HEAD:
        return True                       # 出る／払う／扱った／出て
    if run[0] == "な" and len(run) > 1 and run[1] in NAI_FOLLOW:
        return True                       # 届け出ない／届け出なければ
    if run[0] in I_ROW and len(run) > 1 and run[1] in RENYO_HEAD:
        return True                       # 支払います／取り扱いました
    return False                          # 「い」＋助詞など → 名詞


# ============================================================
# 辞書グループ
#   forms      : 同義とみなす表記のリスト
#   category   : コメント・集計時の分類名
#   no_kanji_prefix : 直前が漢字/カタカナなら複合語とみなして除外
#                     （「正社員」の「社員」を拾わないため）
#   skip_next  : 直後がこの正規表現に一致したら除外
#   skip_prev  : 直前がこの正規表現に一致したら除外
# ============================================================

DICT_GROUPS = [
    # ---- 接続詞・副詞のかな漢字ゆれ ----
    {"category": "かな漢字", "forms": ["及び", "および"]},
    {"category": "かな漢字", "forms": ["並びに", "ならびに"]},
    {"category": "かな漢字", "forms": ["又は", "または"]},
    {"category": "かな漢字", "forms": ["若しくは", "もしくは"]},
    {"category": "かな漢字", "forms": ["但し", "ただし"]},
    {"category": "かな漢字", "forms": ["且つ", "かつ"],
     "skip_prev": r"(な|かつ)$", "skip_next": r"^(かつ|つ)"},
    {"category": "かな漢字", "forms": ["従って", "したがって"]},
    {"category": "かな漢字", "forms": ["尚", "なお"],
     "skip_next": r"^(かつ|さら|す|し|る)"},
    {"category": "かな漢字", "forms": ["全て", "すべて"]},
    {"category": "かな漢字", "forms": ["直ちに", "ただちに"]},
    {"category": "かな漢字", "forms": ["速やかに", "すみやかに"]},
    {"category": "かな漢字", "forms": ["予め", "あらかじめ"]},
    {"category": "かな漢字", "forms": ["更に", "さらに"]},
    {"category": "かな漢字", "forms": ["既に", "すでに"]},
    {"category": "かな漢字", "forms": ["概ね", "おおむね"]},
    {"category": "かな漢字", "forms": ["著しく", "いちじるしく"]},
    {"category": "かな漢字", "forms": ["極めて", "きわめて"]},
    {"category": "かな漢字", "forms": ["出来る", "できる"]},
    {"category": "かな漢字", "forms": ["行う", "行なう"]},
    {"category": "かな漢字", "forms": ["表す", "表わす"]},
    {"category": "かな漢字", "forms": ["合わせて", "併せて", "あわせて"]},

    # ---- 用語のゆれ ----
    {"category": "用語", "forms": ["従業員", "社員", "職員"],
     "no_kanji_prefix": True},
    {"category": "用語", "forms": ["会社", "当社"],
     "no_kanji_prefix": True},
    {"category": "用語", "forms": ["賃金", "給与", "給料"],
     "no_kanji_prefix": True, "skip_next": r"^(規程|規則|台帳|支払明細)"},
    {"category": "用語", "forms": ["事業場", "事業所"],
     "no_kanji_prefix": True},
    {"category": "用語", "forms": ["上長", "上司"],
     "no_kanji_prefix": True},
    {"category": "用語", "forms": ["年次有給休暇", "有給休暇", "年休"],
     "no_kanji_prefix": True},

    # ---- 単位・助数詞 ----
    {"category": "単位", "forms": ["か月", "ヵ月", "ヶ月", "カ月", "ケ月", "箇月"]},
    {"category": "単位", "forms": ["か所", "ヵ所", "ヶ所", "カ所", "ケ所", "箇所"]},
]


# ============================================================
# 規則／規程／規定 の誤用検出
#
#   この3語は表記ゆれではなく、本来それぞれ意味が違います。
#     規則 … 就業規則そのもの
#     規程 … 独立した別文書（賃金規程・育児介護休業規程）
#     規定 … 条文の定め（「前項の規定により」）
#   単純にゆれとして扱うと、正しく使い分けている大半が引っかかり、
#   本当の誤用が埋もれます。そこで「誤用のパターン」だけを狙います。
#
#   confidence:
#     "sure"   … 日本語として成立しないもの。ほぼ確実に誤用
#     "likely" … 慣行から外れるもの。正しいこともあるので要確認
# ============================================================

MISUSE_RULES = [
    # --- 規程は名詞。活用しないので「規程する」は必ず誤り ---
    {
        "id": "kitei_verb",
        "regex": r'規程(?=する|しな|して|した|され|すべき|せず)',
        "correct": "規定",
        "confidence": "sure",
        "message": "「規程」は独立した文書を指す名詞なので活用しません。"
                   "条文が定める意味であれば「規定」が適切です。",
    },
    # --- 「前項の規程」は、前項という条文の定めを指すので「規定」 ---
    {
        "id": "kitei_ref",
        "regex": (r'(?:前項|前条|同項|同条|本条|本項|次項|次条|前各号|各号|'
                  r'第[0-9０-９一二三四五六七八九十百]+[条項号])'
                  r'(?:の|に)規程'),
        "correct": "規定",
        "confidence": "sure",
        "message": "条・項・号が定める内容を指しているので「規定」が適切です。"
                   "「規程」は独立した別文書を指します。",
    },
    # --- 「〜規定に基づき」のように文書名として使っているもの ---
    #     慣行では独立文書は「規程」。ただし「〜規定」という名称の
    #     文書も実在するので要確認扱いにする。
    {
        "id": "bunsho_kitei",
        "regex": r'[一-龯ァ-ヶー]{2,8}規定(?=[にはがをでも、。」）\s]|$)',
        "correct": "規程",
        "confidence": "likely",
        # 「〜法の規定」「安全衛生管理の規定」などは条文の定めなので除外
        "skip_prev": r'(?:の|法|令|条例|規則|規程|条|項|号)$',
        "message": "独立した文書の名称であれば「規程」を使うのが慣行です"
                   "（文書名が「〜規定」であれば問題ありません）。",
    },
]

# 自称（この規則／この規程）の検出
_SELF_RE = re.compile(r'(?:この|本|当)(規則|規程)')


def _detect_doc_kind(data):
    """
    文書自身が「就業規則」なのか「〜規程」なのかを推定する。
    自称の誤用（就業規則なのに「この規程」と書く）を判定するために使う。
    判定できなければ None を返し、そのチェックは行わない。
    """
    texts = []
    for node in data:
        if node.get("type") == "front_matter":
            texts.extend(node.get("texts") or [])
    head = "".join(t for t in iter_body_texts(data) if t)[:400]
    joined = re.sub(r'[\s\u3000]', '', "".join(texts) + head)
    if re.search(r'この就業規則|就業規則[（(]以下', joined):
        return "規則"
    if re.search(r'この[一-龯ァ-ヶー]{2,10}規程|[一-龯ァ-ヶー]{2,10}規程[（(]以下',
                 joined):
        return "規程"
    return None


# ============================================================
# 本文テキストの収集
#   対象は本文（条見出し・項・号・号の下位）のみ。
#   前付け・後付け（附則等）・表は元XMLをそのまま複製する領域なので
#   触りません。
# ============================================================

def iter_body_texts(data):
    """JSON から本文テキストを順に取り出すジェネレータ。"""

    def walk_item(it):
        yield it.get("body", "")
        for su in (it.get("sub_items") or []):
            yield su.get("body", "")
            for su2 in (su.get("sub_items2") or []):
                yield su2.get("body", "")

    def walk_article(art):
        yield art.get("title", "")
        yield (art.get("title_note") or "")
        for para in (art.get("paragraphs") or []):
            yield para.get("body", "")
            for it in (para.get("items") or []):
                yield from walk_item(it)

    for node in data:
        t = node.get("type")
        if t == "chapter":
            yield node.get("title", "")
            for art in (node.get("articles") or []):
                yield from walk_article(art)
        elif t == "article":
            yield from walk_article(node)


# ============================================================
# 自動グルーピング（送り仮名ゆれ）
# ============================================================

_MID_CLASS = ''.join(sorted(ALLOWED_MID))
# 「賃金の取扱い」を1語として拾ってしまわないよう、語中のかなは
# 送り仮名らしいものだけに限定する（助詞では語が切れる）。
_TOKEN_RE = re.compile(
    rf'[{KANJI}]+(?:[{_MID_CLASS}][{KANJI}]+)*[{ALLOWED_TAIL}]?')
_KATA_TOKEN_RE = re.compile(rf'[{KATA}]{{3,}}')


def _iter_tokens(text):
    """
    漢字語トークン（名詞用法のみ）を取り出す。
    「支払った」「取り扱います」のような動詞の活用形は、
    送り仮名の付け方が名詞とは違うので候補にしない。
    """
    for m in _TOKEN_RE.finditer(text):
        tok = m.group(0)
        if _is_verb_usage(text, tok, m.start()):
            continue
        yield tok


def _decompose(tok):
    """
    トークンを (漢字骨格, 末尾かな, 中間かなのリスト) に分解する。
    条件を満たさない場合は None。
    """
    # 末尾のひらがなを切り離す
    m = re.match(rf'^(.*?)([{HIRA}]*)$', tok)
    head, tail = m.group(1), m.group(2)
    if not head:
        return None
    # head は 漢字(かな漢字)* の形のはず
    if not re.fullmatch(rf'[{KANJI}]+(?:[{HIRA}][{KANJI}]+)*', head):
        return None
    skeleton = re.sub(rf'[{HIRA}]', '', head)
    mids = re.findall(rf'[{HIRA}]', head)
    if len(skeleton) < 2 or len(skeleton) > 6:
        return None
    if len(tail) > 1:
        return None
    # 中間の送り仮名は許可リストのものだけ（助詞混入を防ぐ）
    for c in mids:
        if c not in ALLOWED_MID:
            return None
    return skeleton, tail, mids


def _tail_compatible(t1, t2):
    """末尾かなが同一グループとみなせるか。"""
    if t1 == t2:
        return True
    if t1 == "" and t2 in ALLOWED_TAIL_OPT and len(t2) == 1:
        return True
    if t2 == "" and t1 in ALLOWED_TAIL_OPT and len(t1) == 1:
        return True
    return False


def _auto_groups_okurigana(counter):
    """漢字骨格が一致する異表記をグルーピングする。"""
    buckets = defaultdict(dict)   # skeleton -> {surface: tail}
    for surf in counter:
        d = _decompose(surf)
        if not d:
            continue
        skeleton, tail, _ = d
        buckets[skeleton][surf] = tail

    groups = []
    for skeleton, forms in buckets.items():
        if len(forms) < 2:
            continue
        # 末尾かなの互換性で union-find
        names = list(forms)
        parent = {n: n for n in names}

        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x

        def union(a, b):
            ra, rb = find(a), find(b)
            if ra != rb:
                parent[ra] = rb

        for i in range(len(names)):
            for j in range(i + 1, len(names)):
                if _tail_compatible(forms[names[i]], forms[names[j]]):
                    union(names[i], names[j])

        clusters = defaultdict(list)
        for n in names:
            clusters[find(n)].append(n)
        for members in clusters.values():
            if len(members) >= 2:
                groups.append({"category": "送り仮名", "forms": sorted(members),
                               "noun_only": True})
    return groups


def _auto_groups_katakana(counter):
    """カタカナ語の長音符ゆれ（ユーザ/ユーザー など）。"""
    buckets = defaultdict(set)
    for surf in counter:
        if not re.fullmatch(rf'[{KATA}]{{3,}}', surf):
            continue
        key = surf.rstrip('ー')
        if len(key) < 3:
            continue
        buckets[key].add(surf)
    groups = []
    for key, forms in buckets.items():
        if len(forms) >= 2:
            groups.append({"category": "カタカナ", "forms": sorted(forms)})
    return groups


# ============================================================
# レポート
# ============================================================

class HyokiReport:
    """検出結果。出力直前のテキストに対して scan() で位置を返す。"""

    def __init__(self, groups, counts, width_stats=None):
        self.groups = groups          # [{category, forms, counts:{form:n}}]
        self.counts = counts          # Counter（全表層形の出現数）
        self.width_stats = width_stats or {}
        # スキャン用インデックス：長い表記を優先して当てる
        self._index = []
        self._regex_index = []      # 誤用ルール（正規表現）
        for gid, g in enumerate(groups):
            if g.get("regex"):
                self._regex_index.append((re.compile(g["regex"]), gid))
                continue
            for form in g["forms"]:
                self._index.append((form, gid))
        self._index.sort(key=lambda x: -len(x[0]))

    # --------------------------------------------------------
    def scan(self, text):
        """
        text 内の該当箇所を返す。
        戻り値: [{"start":int, "end":int, "surface":str, "gid":int}, ...]
        （位置順・重なりなし）
        """
        if not text:
            return []
        found = []
        # 誤用ルールを先に当てる（重なった場合はこちらを優先したいので
        # 長さを最大にして扱う）
        for rx, gid in self._regex_index:
            g = self.groups[gid]
            for m in rx.finditer(text):
                # 実際に着色するのは対象語だけ（前後の「前項の」等は含めない）
                w = g["forms"][0]
                i = text.find(w, m.start(), m.end())
                if i < 0:
                    continue
                sp = g.get("skip_prev")
                if sp and re.search(sp, text[max(0, m.start() - _CTX_WINDOW):
                                             m.start()]):
                    continue
                found.append({"start": i, "end": i + len(w),
                              "surface": w, "gid": gid, "misuse": True})
        for form, gid in self._index:
            g = self.groups[gid]
            start = 0
            while True:
                i = text.find(form, start)
                if i < 0:
                    break
                start = i + 1
                if not self._context_ok(text, i, form, g):
                    continue
                found.append({"start": i, "end": i + len(form),
                              "surface": form, "gid": gid})
        # 重なりを解消（長いものを優先）
        found.sort(key=lambda s: (s["start"],
                                  0 if s.get("misuse") else 1,
                                  -(s["end"] - s["start"])))
        result = []
        last_end = -1
        for s in found:
            if s["start"] >= last_end:
                result.append(s)
                last_end = s["end"]
        return result

    @staticmethod
    def _context_ok(text, i, form, group):
        if group.get("no_kanji_prefix") and i > 0:
            if _RE_KANJI_OR_KATA.match(text[i - 1]):
                return False
        # 送り仮名グループは名詞用法だけを対象にする
        # （「申出」と「申し出る」は表記ゆれではないため）
        if group.get("noun_only") and _is_verb_usage(text, form, i):
            return False
        sp = group.get("skip_prev")
        if sp and re.search(sp, text[max(0, i - _CTX_WINDOW):i]):
            return False
        sn = group.get("skip_next")
        if sn and re.match(sn, text[i + len(form):i + len(form) + _CTX_WINDOW]):
            return False
        return True

    # --------------------------------------------------------
    def group_label(self, gid):
        """コメント用の「行う 12件／行なう 2件」という文字列。"""
        g = self.groups[gid]
        parts = [f"{f} {g['counts'].get(f, 0)}件" for f in g["forms"]
                 if g["counts"].get(f, 0) > 0]
        return "／".join(parts)

    def comment_text(self, spans):
        """段落ごとのWordコメント本文を作る（1段落1コメント）。"""
        seen = []
        for s in spans:
            key = (s["gid"], s["surface"])
            if key not in seen:
                seen.append(key)
        # 誤用ルールは専用の文面にする
        if len(seen) == 1 and self.groups[seen[0][0]].get("misuse"):
            g = self.groups[seen[0][0]]
            head = ("【用語の誤用】" if g["confidence"] == "sure"
                    else "【用語の確認】")
            return "\n".join([
                f"{head}「{g['forms'][0]}」は「{g['correct']}」の誤りと"
                f"思われます。" if g["confidence"] == "sure"
                else f"{head}「{g['forms'][0]}」は「{g['correct']}」が"
                     f"適切かもしれません。",
                g["message"],
                "※ 自動修正はしていません。ご確認ください。",
            ])

        lines = ["【表記ゆれ】この語は文書内で表記がゆれています。"]
        for gid, surface in seen:
            g = self.groups[gid]
            if g.get("misuse"):
                lines.append(f"・「{surface}」… 「{g['correct']}」の可能性"
                             f"（{g['message']}）")
                continue
            others = [f for f in g["forms"]
                      if f != surface and g["counts"].get(f, 0) > 0]
            if others:
                lines.append(
                    f"・「{surface}」… 他に「{'」「'.join(others)}」"
                    f"（{self.group_label(gid)}）")
            else:
                lines.append(f"・「{surface}」（{self.group_label(gid)}）")
        if any(self.groups[gid]["category"] == "送り仮名" for gid, _ in seen):
            lines.append("※ 送り仮名は名詞用法だけを対象にしています"
                         "（「申し出る」などの動詞は除外）。")
        lines.append("※ このコメントはゆれ1件につき初出の1箇所だけに付けています。"
                     "該当箇所は文書全体で青字にしています（自動修正なし）。")
        return "\n".join(lines)

    # --------------------------------------------------------
    def summary_lines(self):
        """画面表示用のサマリ。"""
        lines = []
        if not self.groups:
            lines.append("  表記ゆれは検出されませんでした。")
        else:
            by_cat = defaultdict(list)
            for gid, g in enumerate(self.groups):
                if g.get("misuse"):
                    mark = "誤り" if g["confidence"] == "sure" else "要確認"
                    by_cat[g["category"]].append(
                        f"「{g['forms'][0]}」→「{g['correct']}」"
                        f"（{g['counts'][g['forms'][0]]}件・{mark}）")
                else:
                    by_cat[g["category"]].append(self.group_label(gid))
            for cat in ["用語の誤用", "用語", "かな漢字", "送り仮名",
                        "単位", "カタカナ"]:
                if cat not in by_cat:
                    continue
                lines.append(f"  [{cat}]")
                for lab in sorted(by_cat[cat]):
                    lines.append(f"    ・{lab}")
        ws = self.width_stats
        if ws.get("mixed"):
            lines.append(
                f"  [数字の全角・半角] 半角 {ws['han']}件 ／ 全角 {ws['zen']}件 が混在"
                "（着色対象外・参考表示）")
        return lines


# ============================================================
# 解析エントリポイント
# ============================================================

def analyze(data, use_dict=True, use_auto=True, extra_groups=None,
            exclude_forms=None, noun_only_okurigana=True,
            check_misuse=True):
    """
    JSON(list) を解析して HyokiReport を返す。

    use_dict      : 辞書グループを使うか
    use_auto      : 自動グルーピング（送り仮名・カタカナ）を使うか
    extra_groups  : 事務所独自の同義語グループを追加する場合
                    例: [{"category":"用語","forms":["役職者","管理職"]}]
    exclude_forms : 検出から除外したい表記のリスト（誤検出の打ち消し用）
    check_misuse  : 規則／規程／規定 の誤用を検出するか。
                    この3語は表記ゆれではなく意味が違うため、
                    ゆれとしてではなく「誤用パターン」として検出する。
    noun_only_okurigana :
                    送り仮名グループを名詞用法だけに限るか。
                    公用文の慣行では名詞は送り仮名を省き（申出・取扱）、
                    動詞は付ける（申し出る・取り扱う）ため、既定では
                    動詞用法を検出対象から外している。False にすると
                    動詞も含めてすべて拾う。
    """
    texts = [t for t in iter_body_texts(data) if t]
    joined = "\n".join(texts)

    # --- 全表層形の出現数を数える ---
    counter = Counter()
    for tok in _iter_tokens(joined):
        counter[tok] += 1
    for tok in _KATA_TOKEN_RE.findall(joined):
        counter[tok] += 1

    exclude = set(exclude_forms or [])
    candidates = []

    if use_dict:
        candidates.extend(DICT_GROUPS)
    if extra_groups:
        candidates.extend(extra_groups)
    if use_auto:
        auto_ok = _auto_groups_okurigana(counter)
        if not noun_only_okurigana:
            for g in auto_ok:
                g["noun_only"] = False
        candidates.extend(auto_ok)
        candidates.extend(_auto_groups_katakana(counter))

    # --- 実際に文書内に2表記以上あるグループだけ残す ---
    groups = []
    seen_sig = set()
    for g in candidates:
        g_eff = dict(g)
        g_eff["forms"] = [f for f in g["forms"] if f not in exclude]
        if len(g_eff["forms"]) < 2:
            continue
        raw = _count_group(joined, g_eff)
        counts = {f: n for f, n in raw.items() if n > 0}
        if len(counts) < 2:
            continue
        sig = tuple(sorted(counts))
        if sig in seen_sig:
            continue
        seen_sig.add(sig)
        ng = dict(g)
        ng["forms"] = sorted(counts, key=lambda f: -counts[f])
        ng["counts"] = counts
        groups.append(ng)

    # --- 規則／規程／規定 の誤用ルールを足す ---
    if check_misuse:
        doc_kind = _detect_doc_kind(data)
        for rule in MISUSE_RULES:
            if not re.search(rule["regex"], joined):
                continue
            wrong = "規程" if rule["correct"] == "規定" else "規定"
            groups.append({
                "category": "用語の誤用",
                "forms": [wrong],
                "counts": {wrong: len(re.findall(rule["regex"], joined))},
                "regex": rule["regex"],
                "skip_prev": rule.get("skip_prev"),
                "misuse": True,
                "correct": rule["correct"],
                "confidence": rule["confidence"],
                "message": rule["message"],
                "rule_id": rule["id"],
            })
        # 自称の誤り（就業規則なのに「この規程」と書いている等）
        if doc_kind:
            wrong = "規程" if doc_kind == "規則" else "規則"
            rx = r'(?:この|本|当)' + wrong
            n = len(re.findall(rx, joined))
            if n:
                groups.append({
                    "category": "用語の誤用",
                    "forms": [wrong],
                    "counts": {wrong: n},
                    "regex": rx,
                    "skip_prev": None,
                    "misuse": True,
                    "correct": doc_kind,
                    "confidence": "sure",
                    "message": f"この文書自身は「{doc_kind}」なので、"
                               f"自分を指す場合は「{doc_kind}」が適切です。",
                    "rule_id": "self_ref",
                })

    width_stats = _width_stats(joined)
    return HyokiReport(groups, counter, width_stats)


def _count_group(text, group):
    """
    グループ内の各表記の出現数を数える。
    「取扱い」を「取扱」として二重に数えないよう、長い表記を優先して
    重なりなしで割り当てる。
    """
    spans = []
    for form in group["forms"]:
        start = 0
        while True:
            i = text.find(form, start)
            if i < 0:
                break
            start = i + 1
            if HyokiReport._context_ok(text, i, form, group):
                spans.append((i, i + len(form), form))
    spans.sort(key=lambda s: (s[0], -(s[1] - s[0])))
    counts = Counter()
    last_end = -1
    for st, en, form in spans:
        if st >= last_end:
            counts[form] += 1
            last_end = en
    return counts


def _width_stats(text):
    """数字の全角・半角混在を数える（着色はしない／参考表示のみ）。"""
    han = len(re.findall(r'(?<![0-9])[0-9]+', text))
    zen = len(re.findall(r'(?<![０-９])[０-９]+', text))
    return {"han": han, "zen": zen, "mixed": bool(han and zen)}
