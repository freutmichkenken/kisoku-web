# -*- coding: utf-8 -*-
"""
Gemini 判定補助モジュール（任意・無料枠対応）

正規表現・スタイルで判定できなかった「判定不能な段落」だけを、
前後数行の文脈とともに Gemini に送り、章/条/項/号/本文のどれかを
判定してもらうためのヘルパー。

【重要な注意】
  - 無料枠ではプロンプトが Google の学習に使われる可能性があります。
    会社名など機密を含む場合は注意してください。送信は判定不能段落の
    前後数行に限定していますが、完全な匿名化ではありません。
  - APIキーが未設定・ライブラリ未導入・通信失敗のいずれでも、
    呼び出し側は従来どおり（本文扱い）にフォールバックできるよう
    None を返します。

【使い方】
  import os
  os.environ["GEMINI_API_KEY"] = "＜あなたのキー＞"   # または Colab の userdata

  from gemini_helper import classify_with_gemini
  kind = classify_with_gemini(
      target_text="１　本規程における子とは…",
      before_lines=["（子および従業員の定義）"],
      after_lines=["①　実子", "②　養子"],
  )
  # kind は 'chapter'/'article'/'article_title'/'paragraph'/'item'/'item_sub'/'body' または None
"""

import os
import json
import time

# 返してよいラベル（想定外の値が返ってきたら None 扱いにする）
VALID_KINDS = {
    "chapter", "article", "article_title",
    "paragraph", "item", "item_sub", "body",
}

# 無料枠で使えるモデルの候補（上から順に、使えるものを自動選択）
# 2026年時点では Flash / Flash-Lite 系が無料枠。環境変数 GEMINI_MODEL で上書き可能。
# 新しめの軽量モデルを優先（レート制限内で高速・安価）。
_MODEL_CANDIDATES = [
    "gemini-3.5-flash-lite",        # 最優先（使えれば最新の軽量版）
    "gemini-3.1-flash-lite",        # 軽量安定版（無料枠向き）
    "gemini-2.5-flash-lite",        # 定番の軽量版
    "gemini-flash-lite-latest",     # エイリアス（将来に追従）
    "gemini-2.5-flash",             # Flash 通常版
    "gemini-flash-latest",          # Flash エイリアス
]

# レート制限対策：呼び出し間隔（秒）。無料枠は 5〜15 RPM なので
# 最低でも 4〜5秒あける（15 RPM でも 4秒間隔が安全）。
_MIN_INTERVAL_SEC = 4.5
_last_call_time = [0.0]

# 429（レート超過）時のリトライ回数と待機
_MAX_RETRIES = 2

_PROMPT_TEMPLATE = """あなたは日本の就業規則の文書構造を判定する専門家です。
以下の「対象行」が、就業規則の構造上どの要素かを判定してください。

判定候補（この中から必ず1つだけ選ぶ）:
- chapter        : 章見出し（例「第1章 総則」）
- article        : 条見出し（例「第5条（採用時の提出書類）」）
- article_title  : 条のタイトルだけの行（例「（目的）」）
- paragraph      : 項（条の中の段落。例「2 前項の…」）
- item           : 号（列挙項目。例「① 履歴書」「(1) 誓約書」）
- item_sub       : 号の下位項目（例「1. …」「イ …」）
- body           : 上記のどれでもない本文の続き

前後の文脈:
--- 前の行 ---
{before}
--- 対象行（これを判定）---
{target}
--- 後の行 ---
{after}

出力は必ず次のJSON形式のみ。説明や前置きは一切書かないこと:
{{"kind": "＜候補のいずれか＞"}}
"""


def _get_api_key():
    """APIキーを環境変数または Colab userdata から取得。無ければ None。"""
    key = os.environ.get("GEMINI_API_KEY")
    if key:
        return key
    # Colab の Secrets(userdata) も試す
    try:
        from google.colab import userdata
        return userdata.get("GEMINI_API_KEY")
    except Exception:
        return None


def _has_genai():
    """google-generativeai ライブラリが使えるか（Colab では pip 済み）。"""
    try:
        import google.generativeai  # noqa: F401
        return True
    except ImportError:
        return False


def is_available():
    """
    Gemini が使える状態か。
    APIキーさえあれば使える（ライブラリが無い環境＝ブラウザ版では
    REST API を直接呼ぶ方式に自動で切り替わる）。
    """
    return bool(_get_api_key())


# ============================================================
# REST API 直呼び（ブラウザ版 / google-generativeai 未導入の環境用）
#   - ブラウザ（Pyodide）: Web Worker 内の同期 XMLHttpRequest で送信
#   - 通常の Python      : urllib で送信
#   Colab で google-generativeai が入っていれば、従来どおりそちらを使う。
# ============================================================

_API_BASE = "https://generativelanguage.googleapis.com/v1beta"


def _in_browser():
    import sys
    return sys.platform == "emscripten"


def _http(method, url, key, body=None, timeout=120):
    """
    HTTP リクエストを送り (status, text) を返す。
    通信自体が失敗した場合は例外を送出する。
    """
    headers = {"x-goog-api-key": key}
    data = None
    if body is not None:
        data = json.dumps(body, ensure_ascii=False)
        headers["Content-Type"] = "application/json"

    if _in_browser():
        from js import XMLHttpRequest
        xhr = XMLHttpRequest.new()
        xhr.open(method, url, False)          # 同期（Web Worker 内なので可）
        for k, v in headers.items():
            xhr.setRequestHeader(k, v)
        xhr.send(data)
        return int(xhr.status), str(xhr.responseText or "")

    import urllib.request
    import urllib.error
    req = urllib.request.Request(
        url, method=method,
        data=data.encode("utf-8") if data is not None else None,
        headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace")


def _sleep(sec):
    """
    確実に sec 秒待つ。ブラウザ版では time.sleep が即座に戻る場合が
    あるため、経過時間を見て足りなければ待ち続ける。
    """
    end = time.time() + sec
    try:
        time.sleep(sec)
    except Exception:
        pass
    while time.time() < end:
        pass


class _RestResponse:
    def __init__(self, text):
        self.text = text


class _RestModel:
    """google.generativeai.GenerativeModel と同じ呼び方ができる最小実装。"""

    def __init__(self, name, key):
        self.name = name
        self._key = key

    def generate_content(self, prompt, generation_config=None):
        cfg = dict(generation_config or {})
        gen = {}
        if "temperature" in cfg:
            gen["temperature"] = cfg["temperature"]
        if "response_mime_type" in cfg:
            gen["responseMimeType"] = cfg["response_mime_type"]
        body = {"contents": [{"role": "user", "parts": [{"text": prompt}]}]}
        if gen:
            body["generationConfig"] = gen
        url = f"{_API_BASE}/models/{self.name}:generateContent"
        status, text = _http("POST", url, self._key, body)
        if status != 200:
            # 429 などを呼び出し側のリトライ判定に渡すため、コードを含める
            raise RuntimeError(f"HTTP {status}: {text[:300]}")
        data = json.loads(text)
        parts = (((data.get("candidates") or [{}])[0]
                  .get("content") or {}).get("parts") or [])
        return _RestResponse("".join(p.get("text", "") for p in parts))


def _rest_list_models(key):
    """generateContent に対応したモデル名の集合（取得失敗なら空集合）。"""
    names = set()
    url = f"{_API_BASE}/models?pageSize=1000"
    try:
        status, text = _http("GET", url, key)
        if status != 200:
            print(f"[gemini_helper] モデル一覧の取得に失敗: HTTP {status}")
            return names
        for m in json.loads(text).get("models", []):
            if "generateContent" in m.get("supportedGenerationMethods", []):
                names.add(m.get("name", "").split("/")[-1])
    except Exception as e:
        print(f"[gemini_helper] モデル一覧の取得に失敗: {e}")
    return names


# モデルインスタンスを使い回す（毎回作ると遅い）
_model_cache = {}


def _get_model():
    key = _get_api_key()
    if key is None:
        return None
    if "model" in _model_cache:
        return _model_cache["model"]
    if not _has_genai():
        return _get_rest_model(key)
    try:
        import google.generativeai as genai
        genai.configure(api_key=key)

        # 環境変数で明示指定されていればそれを最優先
        env_model = os.environ.get("GEMINI_MODEL")
        candidates = ([env_model] if env_model else []) + _MODEL_CANDIDATES

        # 実際に使えるモデル名の集合を取得（generateContent 対応のもの）
        available = set()
        try:
            for m in genai.list_models():
                if "generateContent" in m.supported_generation_methods:
                    # m.name は "models/gemini-2.5-flash-lite" の形式
                    available.add(m.name.split("/")[-1])
        except Exception:
            available = set()   # 取得できなければ候補をそのまま試す

        chosen = None
        for name in candidates:
            if not name:
                continue
            if not available or name in available:
                chosen = name
                break
        if chosen is None:
            # 候補が全滅なら、利用可能なFlash系を拾う
            for name in sorted(available):
                if "flash" in name:
                    chosen = name
                    break
        if chosen is None:
            print(f"[gemini_helper] 使えるモデルが見つかりません。利用可能: {sorted(available)}")
            return None

        model = genai.GenerativeModel(chosen)
        _model_cache["model"] = model
        _model_cache["name"] = chosen
        print(f"[gemini_helper] 使用モデル: {chosen}")
        return model
    except Exception as e:
        print(f"[gemini_helper] モデル初期化に失敗: {e}")
        return None


def _choose_model_name(available):
    """候補リストと利用可能モデルから使うモデル名を決める。"""
    env_model = os.environ.get("GEMINI_MODEL")
    candidates = ([env_model] if env_model else []) + _MODEL_CANDIDATES
    for name in candidates:
        if not name:
            continue
        if not available or name in available:
            return name
    for name in sorted(available):
        if "flash" in name:
            return name
    return None


def _get_rest_model(key):
    """REST 直呼び用のモデルを用意する（ブラウザ版）。"""
    available = _rest_list_models(key)
    chosen = _choose_model_name(available)
    if chosen is None:
        print(f"[gemini_helper] 使えるモデルが見つかりません。利用可能: {sorted(available)}")
        return None
    model = _RestModel(chosen, key)
    _model_cache["model"] = model
    _model_cache["name"] = chosen
    print(f"[gemini_helper] 使用モデル: {chosen}（REST）")
    return model


def _throttle():
    """呼び出し間隔を空けてレート制限（RPM）を避ける。"""
    elapsed = time.time() - _last_call_time[0]
    if elapsed < _MIN_INTERVAL_SEC:
        _sleep(_MIN_INTERVAL_SEC - elapsed)
    _last_call_time[0] = time.time()


def _generate_with_retry(model, prompt):
    """
    レート制限対策つきで生成する。
    429 が出たら待機してリトライ。最終的に失敗したら例外を送出。
    """
    last_err = None
    for attempt in range(_MAX_RETRIES + 1):
        _throttle()
        try:
            return model.generate_content(
                prompt,
                generation_config={
                    "temperature": 0,
                    "response_mime_type": "application/json",
                },
            )
        except Exception as e:
            last_err = e
            msg = str(e)
            if "429" in msg or "quota" in msg.lower() or "rate" in msg.lower():
                # レート超過 → 少し長めに待ってリトライ
                if attempt < _MAX_RETRIES:
                    wait = 30 * (attempt + 1)
                    print(f"[gemini_helper] レート制限。{wait}秒待機してリトライ...")
                    _sleep(wait)
                    continue
            raise
    if last_err:
        raise last_err


def classify_with_gemini(target_text, before_lines=None, after_lines=None):
    """
    判定不能な1段落を Gemini に判定させる。

    target_text : 判定したい行（対象）
    before_lines: 前の数行（リスト）
    after_lines : 後の数行（リスト）

    返り値: VALID_KINDS のいずれか、または判定できなければ None
    """
    model = _get_model()
    if model is None:
        return None

    before = "\n".join(before_lines or []) or "（なし）"
    after = "\n".join(after_lines or []) or "（なし）"
    prompt = _PROMPT_TEMPLATE.format(
        before=before, target=target_text, after=after
    )

    try:
        resp = _generate_with_retry(model, prompt)
        text = (resp.text or "").strip()
        data = json.loads(text)
        kind = data.get("kind")
        if kind in VALID_KINDS:
            return kind
        return None
    except Exception as e:
        # 429（レート制限）やパース失敗など、あらゆる失敗で None
        print(f"[gemini_helper] 判定失敗（従来処理にフォールバック）: {e}")
        return None


# ============================================================
# 分割補助：1つに押し込まれた項・号を分割する
# ============================================================

_SPLIT_PROMPT = """あなたは日本の就業規則の校正専門家です。
以下は就業規則の「{unit_label}」1件として抽出された本文ですが、
本来は複数の{unit_label}に分けるべきものが1つに押し込まれている可能性があります。

【本文】
{body}

この本文を、本来あるべき{unit_label}の単位に分割してください。
- 番号（「3」やタブ、「①」など）や、意味の切れ目で分割します。
- 各断片の先頭から、元の番号・タブ・記号は取り除いてください。
- 分割の必要がなければ、1件だけの配列で返してください。
- 本文の文言は変えず、順序も保ってください。

出力は必ず次のJSON形式のみ。説明は書かないこと:
{{"parts": ["断片1の本文", "断片2の本文", ...]}}
"""


def split_with_gemini(body, unit_label="項"):
    """
    1つに押し込まれた項・号を Gemini に分割させる。

    body       : 分割対象の本文
    unit_label : "項" または "号"（プロンプトの文言に使う）

    返り値: 分割後の本文リスト（例 ["…", "…"]）。
            分割不要・失敗時は None（呼び出し側は元のまま使う）。
    """
    model = _get_model()
    if model is None:
        return None

    prompt = _SPLIT_PROMPT.format(unit_label=unit_label, body=body)
    try:
        resp = _generate_with_retry(model, prompt)
        text = (resp.text or "").strip()
        data = json.loads(text)
        parts = data.get("parts")
        if not isinstance(parts, list) or not parts:
            return None
        # 全て文字列で、かつ2件以上に分かれた場合のみ意味がある
        parts = [str(p).strip() for p in parts if str(p).strip()]
        if len(parts) <= 1:
            return None
        return parts
    except Exception as e:
        print(f"[gemini_helper] 分割失敗（元のまま使用）: {e}")
        return None


# ============================================================
# 潰れた下位項目の抽出
#   「産前の場合 妊娠23週まで…4週に1回 妊娠24週から…」のように、
#   号の見出しと、その下にあるべき下位項目が1つに潰れているものを
#   「見出し」＋「下位項目リスト」に分離する。
# ============================================================

_EXTRACT_SUB_PROMPT = """あなたは日本の就業規則の校正専門家です。
以下は就業規則の「号」1件として抽出された本文ですが、本来は
「見出し」と、その下にぶら下がる複数の「下位項目」が1つに潰れている
可能性があります（元は字下げで階層を表現していたもの）。

【本文】
{body}

これを次の形に分離してください:
- heading: 見出し部分（例「産前の場合」）。見出しが無ければ空文字。
- subs: 下位項目の配列（例「妊娠23週まで……4週に1回」）。
        各項目の先頭の記号・空白は取り除く。
        潰れていない（単一内容の）場合は subs を空配列にする。

本文の文言は変えず、順序も保つこと。
出力は必ず次のJSON形式のみ。説明は書かないこと:
{{"heading": "見出し", "subs": ["下位1", "下位2", ...]}}
"""


def extract_subitems_with_gemini(body):
    """
    潰れた号を「見出し」＋「下位項目リスト」に分離する。

    返り値: (heading, subs)  subs は list
            分離不要・失敗時は (None, None)
    """
    model = _get_model()
    if model is None:
        return None, None
    prompt = _EXTRACT_SUB_PROMPT.format(body=body)
    try:
        resp = _generate_with_retry(model, prompt)
        text = (resp.text or "").strip()
        data = json.loads(text)
        heading = data.get("heading")
        subs = data.get("subs")
        if not isinstance(subs, list):
            return None, None
        subs = [str(s).strip() for s in subs if str(s).strip()]
        if not subs:
            return None, None   # 下位が無ければ分離不要
        return (heading or "").strip(), subs
    except Exception as e:
        print(f"[gemini_helper] 下位抽出失敗（元のまま使用）: {e}")
        return None, None


# ============================================================
# 条単位の構造再構成
#   条まるごとを見せて、項・号・号の下位の階層を正しく組み直させる。
#   単体の項・号だけ見せるより、条全体の一貫性から判断できるため精度が高い。
# ============================================================

_RESTRUCTURE_PROMPT = """あなたは日本の就業規則の文書構造を校正する専門家です。
以下は就業規則の1つの条を、機械的に構造化した結果です。
階層（項・号・号の下位）の判定に誤りがある可能性があります。

【現在の構造】
{article_json}
{reasons_block}
【就業規則の階層ルール】
- 項(paragraph): 条の中の段落。「1 …」「2 …」と番号が振られる単位。
- 号(item): 項の中の列挙項目。「(1)」「①」「一、」等で列挙される。
- 号の下位(sub_items): 号の中をさらに細分した項目。「1.」「2.」等。
- 号の下位のさらに下(sub_items2): その中をもう一段細分した項目。「(ア)」「(イ)」等。

【よくある誤り】
1. 本来複数の項に分かれるべきものが1つの項に押し込まれている
2. 本来複数の号に分かれるべきものが1つの号に押し込まれている
3. 号の見出しと、その下にぶら下がる下位項目が1つに潰れている
   （例:「産前の場合妊娠23週まで……4週に1回妊娠24週から……」
     → 号「産前の場合」＋下位「妊娠23週まで……4週に1回」「妊娠24週から……」）
4. 「次の各号」「以下のとおり」等の予告文があるのに、受け皿の号が無い
   （例:「次の書類を提出すること。履歴書誓約書身元保証書…」
     → 項「次の書類を提出すること。」＋号「履歴書」「誓約書」「身元保証書」…）
   ※ 区切り記号が無く名詞が連続しているだけの場合もあります。
      意味の切れ目で号に分けてください。
5. 全角スペースで列を揃えた一覧が1行に潰れている
   （例:「…出頭する場合　　その日…出頭する場合　　その期間」
     → 号「…出頭する場合　その日」「…出頭する場合　その期間」）

【厳守事項 — これに反する出力は破棄されます】
- **入力に存在しない文言を絶対に追加しないでください。**
  一般的な就業規則の知識から内容を補ってはいけません。
  例:「賃金の構成は、次のとおりとする。」しか無い条に対して、
     「基本給」「諸手当」などの号を勝手に作ってはいけません。
- 「次のとおり」「次の各号」という予告文があっても、
  入力に該当する内容が無ければ、**号を作らずそのまま返してください。**
  （元の文書では図や表で示されている場合があります）
- 既存の文言を削除しないでください。
- 出力に含めてよいのは、入力にある文字だけです。

【指示】
- 誤りがあれば、入力にある文言だけを並べ替えて正しい階層に組み直してください。
- 本文の文言は一切変えないでください（記号や余分な空白の除去のみ可）。
- 順序を保ってください。
- 誤りが無ければ、現在の構造をそのまま返してください。
- 条のタイトル(title)は変更しないでください。

出力は必ず次のJSON形式のみ。説明や前置きは書かないこと:
{{"paragraphs": [
  {{"body": "項の本文",
    "items": [
      {{"body": "号の本文",
        "sub_items": [
          {{"body": "下位の本文",
            "sub_items2": [{{"body": "さらに下の本文"}}]}}
        ]}}
    ]}}
]}}

※ sub_items2 は、入力に該当する階層がある場合のみ含めてください。
   入力にある sub_items2 を削除してはいけません。
"""


def restructure_article_with_gemini(article, reasons=None):
    """
    条まるごとを Gemini に見せて、階層を組み直させる。

    article: 条のdict（number, title, paragraphs を持つ）
             tables は送らず、呼び出し側で保持すること。
    reasons: 機械側が「疑わしい」と判定した理由のリスト。
             プロンプトに含めることで、どこを重点的に見るか示す。

    返り値: 組み直された paragraphs のリスト、または None（失敗・変更不要）
    """
    model = _get_model()
    if model is None:
        return None

    # 送信用に簡略化した構造を作る（表は含めない）
    simple = {"paragraphs": []}
    for p in article.get("paragraphs", []):
        sp = {"body": p.get("body", ""), "items": []}
        for it in p.get("items", []):
            si = {"body": it.get("body", "")}
            subs = it.get("sub_items", [])
            if subs:
                sub_list = []
                for su in subs:
                    d = {"body": su.get("body", "")}
                    subs2 = su.get("sub_items2", [])
                    if subs2:
                        d["sub_items2"] = [{"body": s2.get("body", "")}
                                           for s2 in subs2]
                    sub_list.append(d)
                si["sub_items"] = sub_list
            sp["items"].append(si)
        simple["paragraphs"].append(sp)

    article_json = json.dumps(simple, ensure_ascii=False, indent=1)

    # 機械側の検出理由を伝える（重点的に見てほしい箇所のヒント）
    reasons_block = ""
    if reasons:
        lines = "\n".join(f"- {r}" for r in reasons)
        reasons_block = (
            "\n【この条が疑わしいと判定された理由】\n"
            f"{lines}\n"
            "上記の箇所を重点的に見直してください。\n"
        )

    prompt = _RESTRUCTURE_PROMPT.format(
        article_json=article_json, reasons_block=reasons_block)

    try:
        resp = _generate_with_retry(model, prompt)
        text = (resp.text or "").strip()
        data = json.loads(text)
        paragraphs = data.get("paragraphs")
        if not isinstance(paragraphs, list) or not paragraphs:
            return None
        # 妥当性チェック: 各項に body があること
        for p in paragraphs:
            if not isinstance(p, dict) or "body" not in p:
                return None
        return paragraphs
    except Exception as e:
        print(f"[gemini_helper] 条の再構成に失敗（元のまま使用）: {e}")
        return None
