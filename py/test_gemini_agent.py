# -*- coding: utf-8 -*-
"""
Antigravity エージェントによるまとめて判定の検証。 python3 py/test_gemini_agent.py で実行。
ネットワークは使わない（_http と _sleep と時計を差し替える）。
"""

import io
import json
import os
import contextlib
from types import SimpleNamespace

os.environ["GEMINI_API_KEY"] = "dummy-key"

import gemini_helper as gh  # noqa: E402

_real_time = gh.time
_clock = [1000.0]
gh.time = SimpleNamespace(time=lambda: _clock[0], sleep=lambda s: None)
_sleeps = []


def _fake_sleep(sec):
    _sleeps.append(sec)
    _clock[0] += sec


gh._sleep = _fake_sleep


def _output(text, status="completed", env="environments/env1"):
    r = {"status": status,
         "steps": [{"type": "thought", "content": [{"type": "text", "text": "考え中"}]},
                   {"type": "model_output", "content": [{"type": "text", "text": text}]}]}
    if env:
        r["environment_id"] = env
    return r


def install(results):
    """results: interaction ごとの応答(dict)。None は終わらない（running のまま）。"""
    calls = []
    queue = list(results)
    started = []

    def http(method, url, key, body=None, timeout=120):
        calls.append((method, url, body))
        if method == "POST" and url.endswith("/interactions"):
            started.append(len(started) + 1)
            return 200, json.dumps({"id": f"i{len(started)}", "status": "running",
                                    "environment_id": "environments/env1"})
        if method == "GET" and "/interactions/" in url:
            r = queue[len(started) - 1] if len(started) - 1 < len(queue) else None
            if r is None:
                return 200, json.dumps({"id": "x", "status": "running"})
            return 200, json.dumps(r)
        return 200, "{}"                      # cancel / DELETE

    gh._http = http
    del _sleeps[:]
    _clock[0] = 1000.0
    return calls


def labels_json(d):
    return json.dumps({"labels": {str(k): v for k, v in d.items()}}, ensure_ascii=False)


def quiet(fn, *a, **k):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        r = fn(*a, **k)
    return r, buf.getvalue()


LINES = ["第1章　総則", "（目的）", "第1条　この規則は、…", "①　履歴書", "付記"]

# --- 正常：2つの target が返る。POST の中身、* 付きの行、環境の削除 ---
calls = install([_output(labels_json({1: "article_title", 3: "item"}))])
res, _ = quiet(gh.classify_document_with_agent, LINES, [1, 3])
assert res == {1: "article_title", 3: "item"}, res
post = [c for c in calls if c[0] == "POST" and c[1].endswith("/interactions")]
assert len(post) == 1
body = post[0][2]
assert body["agent"] == "antigravity-preview-09-2026"
assert body["background"] is True and body["environment"] == "remote"
assert body["agent_config"]["model"] == "gemini-3.8-flash"
assert body["agent_config"]["max_total_tokens"] <= 100000
assert "*1: （目的）" in body["input"] and "*3: ①　履歴書" in body["input"]
assert "\n0: 第1章　総則" in body["input"] and "*0:" not in body["input"]
assert ("DELETE", gh._API_BASE + "/environments/env1", None) in calls
assert gh._sleep and _sleeps[0] == gh._AGENT_POLL_SEC

# --- 開始フックに ID が渡る。フックが例外でも処理は続く ---
seen = []
gh.agent_started_hook = seen.append
install([_output(labels_json({1: "body"}))])
res, _ = quiet(gh.classify_document_with_agent, LINES, [1])
assert seen == ["i1"] and res == {1: "body"}
gh.agent_started_hook = lambda i: 1 / 0
install([_output(labels_json({1: "body"}))])
res, _ = quiet(gh.classify_document_with_agent, LINES, [1])
assert res == {1: "body"}
gh.agent_started_hook = None

# --- コードフェンス付きJSONは通る ---
install([_output("```json\n" + labels_json({1: "body"}) + "\n```")])
assert quiet(gh.classify_document_with_agent, LINES, [1])[0] == {1: "body"}

# --- JSONの前後に説明文が付いても通る ---
install([_output("判定結果です。\n" + labels_json({1: "body"}) + "\n以上")])
assert quiet(gh.classify_document_with_agent, LINES, [1])[0] == {1: "body"}
# 説明文に { } があっても、フェンスの前に説明文があっても通る
install([_output("{番号: kind} の形で返します。\n```json\n" + labels_json({1: "body"}) + "\n```\n※{注意}")])
assert quiet(gh.classify_document_with_agent, LINES, [1])[0] == {1: "body"}

# --- 開始の POST が 503 でも送り直さない（1回目の interaction を止められなくなるため）---
_posts = []
def _http_503(method, url, key, body=None, timeout=120):
    _posts.append(url)
    return 503, '{"error": {"message": "busy"}}'
gh._http = _http_503
res, out = quiet(gh.classify_document_with_agent, LINES, [1])
assert res is None and _posts == [gh._API_BASE + "/interactions"], _posts

# --- 検証に反したら捨てる（None）---
for bad in (labels_json({1: "body"}),                          # 足りない
            labels_json({1: "body", 3: "item", 4: "body"}),    # 余計
            labels_json({1: "body", 3: "heading"}),            # 不正な kind
            labels_json({1: "body", "x": "item"}),             # キーが数字でない
            "これはJSONではありません",                          # 壊れたJSON
            json.dumps({"labels": [1, 3]}),                    # labels が dict でない
            ""):                                               # 本文なし
    install([_output(bad)])
    assert quiet(gh.classify_document_with_agent, LINES, [1, 3])[0] is None, bad

# --- APIキーが無ければ None ---
os.environ.pop("GEMINI_API_KEY")
assert gh.classify_document_with_agent(LINES, [1]) is None
os.environ["GEMINI_API_KEY"] = "dummy-key"

# --- cancelled → None。残りのチャンクは送られない ---
long_lines = [f"{i} " + "あ" * 990 for i in range(80)]       # 約8万文字 → 3チャンク
calls = install([_output("", status="cancelled")])
res, out = quiet(gh.classify_document_with_agent, long_lines, [0, 40, 79])
assert res is None and "中止しました" in out
assert len([c for c in calls if c[0] == "POST" and c[1].endswith("/interactions")]) == 1
assert ("DELETE", gh._API_BASE + "/environments/env1", None) in calls

# --- 期限超過 → cancel が POST され、環境も消す ---
calls = install([None])
res, out = quiet(gh.classify_document_with_agent, LINES, [1])
assert res is None
assert ("POST", gh._API_BASE + "/interactions/i1/cancel", None) in calls
assert ("DELETE", gh._API_BASE + "/environments/env1", None) in calls
assert _clock[0] - 1000.0 >= gh._AGENT_DEADLINE_SEC

# --- 分割：target の無いチャンクは送られず、各プロンプトは40,000文字以下 ---
prompts = []
install([])


def _http_split(method, url, key, body=None, timeout=120):
    if method == "POST" and url.endswith("/interactions"):
        prompts.append(body["input"])
        return 200, json.dumps({"id": f"i{len(prompts)}", "status": "running"})
    if method == "GET":
        n = len(prompts)
        marked = [int(l[1:l.index(":")]) for l in prompts[-1].split("\n") if l.startswith("*")]
        return 200, json.dumps(_output(labels_json({i: "body" for i in marked}), env=None))
    return 200, "{}"


gh._http = _http_split
res, out = quiet(gh.classify_document_with_agent, long_lines, [0, 79])   # 真ん中のチャンクは target 無し
assert res == {0: "body", 79: "body"}, res
assert len(prompts) == 2, len(prompts)
assert all(len(p) <= gh._AGENT_PROMPT_MAX for p in prompts)
assert "*0:" in prompts[0] and "*79:" in prompts[1] and "*79:" not in prompts[0]
# 2回目の送信前に、直前チャンクの推定入力に応じて待っている
assert any(s > 1 for s in _sleeps)
# 1行で上限を超える場合は送らない
prompts[:] = []
res, out = quiet(gh.classify_document_with_agent, ["あ" * 45000, "b"], [0])
assert res is None and prompts == []

# --- 条の再構成（Antigravity）---
ART = {"paragraphs": [{"body": "次の書類を提出すること。履歴書誓約書", "items": []}]}
ART2 = {"paragraphs": [{"body": "賃金は次のとおりとする。", "items": [{"body": "基本給"}]}]}
fixed = [{"body": "次の書類を提出すること。", "items": [{"body": "履歴書"}, {"body": "誓約書"}]}]
ans = json.dumps({"articles": {"0": {"paragraphs": fixed},
                               "1": {"paragraphs": "壊れた形"}}}, ensure_ascii=False)
calls = install([_output("説明です\n" + ans)])
res, out = quiet(gh.restructure_articles_with_agent, [(ART, ["押し込み"]), (ART2, [])])
assert res == {0: fixed}, res                       # 形の悪い条1だけ落とす
post = [c for c in calls if c[0] == "POST" and c[1].endswith("/interactions")]
assert len(post) == 1                               # 2条を1回の依頼にまとめる
p = post[0][2]["input"]
assert "【条 0】" in p and "【条 1】" in p and "押し込み" in p and "入力に存在しない文言" in p
assert "依頼1/1を送信（2条" in out

# 回答が読めない・失敗・中止のときは空（元の構造のまま）。2条目以降の依頼は中止で送らない
install([_output("JSONではない")])
assert quiet(gh.restructure_articles_with_agent, [(ART, [])])[0] == {}
install([_output("", status="failed")])
assert quiet(gh.restructure_articles_with_agent, [(ART, [])])[0] == {}
big = {"paragraphs": [{"body": "あ" * 20000, "items": []}]}
calls = install([_output("", status="cancelled")])
assert quiet(gh.restructure_articles_with_agent, [(big, []), (big, [])])[0] == {}
assert len([c for c in calls if c[0] == "POST" and c[1].endswith("/interactions")]) == 1

# --- parse_kisoku 統合 ---
import kisoku_parser as kp  # noqa: E402

gh.time = _real_time
SAMPLE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                      "tests", "samples", "01_01_202601.docx")
gh.restructure_article_with_gemini = lambda *a, **k: None
gh.split_with_gemini = lambda *a, **k: None
gh.extract_subitems_with_gemini = lambda *a, **k: (None, None)
paragraph_calls = []
gh.classify_with_gemini = lambda t, b, a: (paragraph_calls.append(t), None)[1]


# このサンプルの判定不能段落は出力の木に現れない位置にあるため、
# 反映の有無は classify_auto の呼び出しと「再判定した件数」の表示で確かめる。
_orig_auto = kp.classify_auto
auto_calls = []


def _spy_auto(kind, text):
    auto_calls.append((kind, text))
    return _orig_auto(kind, text)


kp.classify_auto = _spy_auto
agent_calls = []


def fake_agent(lines, targets):
    agent_calls.append((list(lines), list(targets)))
    return {targets[0]: "paragraph", targets[1]: "chapter"}


gh.classify_document_with_agent = fake_agent
restructure_agent_calls = []
gh.restructure_articles_with_agent = lambda items: (restructure_agent_calls.append(items), {})[1]
base, _ = quiet(kp.parse_kisoku, SAMPLE, use_gemini=False)
del auto_calls[:]
tree, out = quiet(kp.parse_kisoku, SAMPLE, use_gemini=True, gemini_mode="agent")
assert len(agent_calls) == 1 and len(agent_calls[0][1]) == 2 and not paragraph_calls
lines, targets = agent_calls[0]
t0, t1 = lines[targets[0]], lines[targets[1]]
assert ("paragraph", t0) in auto_calls                                  # 項は反映される
assert not any(c[0] == "chapter" and c[1] == t1 for c in auto_calls)   # 章は反映されない
assert "1/2段落を再判定しました" in out, out
assert tree == base

# 条の再構成：agent では Antigravity にまとめて送り、1条ずつの Gemini は呼ばない
rest_calls = []
gh.restructure_article_with_gemini = lambda *a, **k: (rest_calls.append(1), None)[1]
_orig_collect = kp._collect_suspicious_articles
kp._collect_suspicious_articles = lambda t, n: [(a, ["テスト"]) for a in _articles(t)[:2]]


def _articles(t):
    return [a for n in t if n.get("type") == "chapter" for a in n.get("articles", [])]


quiet(kp.parse_kisoku, SAMPLE, use_gemini=True, gemini_mode="agent")
assert len(restructure_agent_calls[-1]) == 2 and rest_calls == []
quiet(kp.parse_kisoku, SAMPLE, use_gemini=True)
assert len(rest_calls) == 2
kp._collect_suspicious_articles = _orig_collect
gh.restructure_article_with_gemini = lambda *a, **k: None
paragraph_calls[:] = []

# 使えなかった（None）ときはルールの判定のまま。1段落ずつの方式には落とさない
gh.classify_document_with_agent = lambda l, t: None
tree2, out = quiet(kp.parse_kisoku, SAMPLE, use_gemini=True, gemini_mode="agent")
assert "ルールの判定のまま" in out and not paragraph_calls
assert tree2 == base

# 既定('paragraph')ではまとめて判定を呼ばず、従来の1段落ずつを使う
agent_calls[:] = []
gh.classify_document_with_agent = fake_agent
quiet(kp.parse_kisoku, SAMPLE, use_gemini=True)
assert agent_calls == [] and len(paragraph_calls) == 2

print("OK")
