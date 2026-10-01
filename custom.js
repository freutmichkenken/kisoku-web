// 就業規則 整形アプリ — カスタムテンプレート（書式の設定画面とプレビュー）
// 設定値の形は py/custom_template.py の normalize() と同じ。
// 正式な検証は Python 側で行い、ここでは入力欄の範囲に収めるだけにする。
(() => {
  const VERSION = 1;
  const STORE_KEY = "kisoku.custom";
  const TEXT_WIDTH_PT = 439.35;  // A4・左右余白 30mm/25mm の本文幅（テンプレ1〜4共通）
  const DEFAULT_TAB_PT = 42;     // Word の既定のタブ間隔（840twips）

  const FONTS = [
    { name: "ＭＳ 明朝", css: '"ＭＳ 明朝", "MS Mincho"' },
    { name: "ＭＳ ゴシック", css: '"ＭＳ ゴシック", "MS Gothic"' },
    { name: "游明朝", css: '"游明朝", "Yu Mincho", YuMincho' },
    { name: "游ゴシック", css: '"游ゴシック", "Yu Gothic", YuGothic' },
    { name: "BIZ UD明朝 Medium", css: '"BIZ UD明朝 Medium", "BIZ UDMincho"' },
    { name: "BIZ UDゴシック", css: '"BIZ UDゴシック", "BIZ UDGothic"' },
  ];
  const FORMATS = [
    ["decimalFullWidth", "１, ２, ３"],
    ["decimal", "1, 2, 3"],
    ["japaneseCounting", "一, 二, 三"],
    ["decimalEnclosedCircle", "①, ②, ③"],
    ["aiueoFullWidth", "ア, イ, ウ"],
    ["irohaFullWidth", "イ, ロ, ハ"],
    ["none", "番号なし"],
  ];
  const SEPS = [["zen", "全角スペース"], ["space", "半角スペース"], ["tab", "タブ"], ["none", "なし"]];
  const ALIGNS = [["left", "両端揃え"], ["center", "中央揃え"]];
  const LEVELS = [
    ["chapter", "章"], ["article", "条（見出し）"], ["para1", "第1項"], ["paraN", "第2項以降"],
    ["item", "号"], ["sub", "号の下位"], ["sub2", "号の下位の下"],
  ];

  // ---------- テンプレ1〜4に近い設定 ----------
  const lv = (fmt, pre, suf, sep, first, left, size = 10.5, bold = false,
    align = "left", before = 0, after = 0) =>
    ({ fmt, pre, suf, sep, first, left, size, bold, font: "", align, before, after });
  const common = {
    item: lv("decimal", "(", ")", "tab", 3, 5),
    sub: lv("decimal", "", ".", "tab", 5.5, 8),
    sub2: lv("aiueoFullWidth", "(", ")", "tab", 7.5, 10),
  };
  const base = (b, layout, levels) => ({
    version: VERSION, base: b, layout, font: "ＭＳ 明朝", size: 10.5, line: 16, table_size: 10.5,
    levels: { ...structuredClone(common), ...levels },
  });
  const PRESETS = {
    t1: base("t1", "inline", {
      chapter: lv("decimalFullWidth", "第", "章", "zen", 0, 0, 16, false, "center", 27),
      article: lv("decimalFullWidth", "第", "条", "space", 0, 0, 12, false, "left", 18),
      para1: lv("decimal", "", "", "tab", 0, 2),
      paraN: lv("decimal", "", "", "tab", 0, 2),
    }),
    t2: base("t2", "separate", {
      chapter: lv("decimalFullWidth", "第", "章", "zen", 0, 0, 16, true, "center", 36),
      article: lv("decimal", "第", "条", "zen", 1, 1, 10.5, false, "left", 18),
      para1: lv("none", "", "", "tab", 0, 1),
      paraN: lv("decimal", "", "", "zen", 0, 1),
    }),
    t3: base("t3", "separate", {
      chapter: lv("decimalFullWidth", "第", "章", "zen", 0, 0, 16, true, "center", 36),
      article: lv("decimal", "第", "条", "zen", 1, 1, 10.5, false, "left", 18),
      para1: lv("none", "", "", "tab", 0, 1),
      paraN: lv("decimal", "", "", "zen", 1, 1),
    }),
    t4: base("t4", "inline", {
      chapter: lv("decimalFullWidth", "第", "章", "zen", 0, 0, 16, false, "center", 27),
      article: lv("decimalFullWidth", "第", "条", "space", 0, 0, 12, false, "left", 18),
      para1: lv("none", "", "", "tab", 2, 1),
      paraN: lv("decimal", "", ".", "tab", 0, 1),
      sub: lv("decimal", "", "．", "tab", 5.5, 8),
    }),
  };

  // ---------- 見本の例文（プレビューと見本 docx で共通） ----------
  const P = (body, items = []) => ({ type: "paragraph", body, items });
  const SAMPLE = [
    { type: "chapter", title: "総則", articles: [
      { type: "article", title: "目的", paragraphs: [
        P("この規則は、株式会社〇〇（以下「会社」という。）の従業員の労働条件、服務規律その他の就業に関する事項を定めるものである。"),
        P("この規則に定めのない事項については、労働基準法その他の法令の定めるところによる。"),
      ] },
      { type: "article", title: "適用範囲", paragraphs: [
        P("この規則は、会社に雇用されるすべての従業員に適用する。ただし、パートタイマーの就業に関する事項については、別に定めるところによる。"),
      ] },
    ] },
    { type: "chapter", title: "人事", articles: [
      { type: "article", title: "採用時の提出書類", paragraphs: [
        P("従業員として採用された者は、採用された日から２週間以内に、次の書類を提出しなければならない。", [
          { body: "履歴書" },
          { body: "住民票記載事項証明書" },
          { body: "その他会社が指定する書類", sub_items: [
            { body: "扶養家族がいる場合は、次の書類", sub_items2: [
              { body: "健康保険被扶養者（異動）届" },
              { body: "給与所得者の扶養控除等（異動）申告書" },
            ] },
            { body: "通勤手当の支給を受ける場合は、通勤経路の届出書" },
          ] },
        ]),
        P("前項の書類の記載事項に変更があったときは、速やかに会社に届け出なければならない。"),
      ] },
    ] },
  ];
  // apply_style は項の number で第1項かどうかを見分ける
  SAMPLE.forEach((ch) => ch.articles.forEach((a) =>
    a.paragraphs.forEach((p, i) => { p.number = i + 1; })));

  // ---------- 番号の文字 ----------
  const KANJI = "〇一二三四五六七八九";
  function kanji(n) {
    if (n <= 0 || n >= 10000) return String(n);
    let out = "";
    [[1000, "千"], [100, "百"], [10, "十"]].forEach(([u, c]) => {
      const d = Math.floor(n / u) % 10;
      if (d) out += (d > 1 ? KANJI[d] : "") + c;
    });
    return out + (n % 10 ? KANJI[n % 10] : "");
  }
  const AIUEO = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン";
  const IROHA = "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス";
  function formatNumber(fmt, n) {
    switch (fmt) {
      case "decimalFullWidth": return String(n).replace(/\d/g, (d) => String.fromCharCode(d.charCodeAt(0) + 0xFEE0));
      case "japaneseCounting": return kanji(n);
      case "decimalEnclosedCircle":
        if (n >= 1 && n <= 20) return String.fromCharCode(0x2460 + n - 1);
        if (n >= 21 && n <= 35) return String.fromCharCode(0x3251 + n - 21);
        if (n >= 36 && n <= 50) return String.fromCharCode(0x32B1 + n - 36);
        return String(n);
      case "aiueoFullWidth": return AIUEO[(n - 1) % AIUEO.length];
      case "irohaFullWidth": return IROHA[(n - 1) % IROHA.length];
      case "none": return "";
      default: return String(n);
    }
  }
  const numberText = (l, n) => (l.fmt === "none" ? "" : l.pre + formatNumber(l.fmt, n) + l.suf);
  // py/custom_template.py の line_height() と同じ式
  const lineHeight = (s, l) => Math.max(s.line, l.size * 1.25);

  // ---------- 設定値を入力欄の範囲に収める ----------
  const RANGES = { size: [6, 36], line: [6, 72], table_size: [6, 36] };
  const LV_RANGES = { first: [0, 30], left: [0, 30], size: [6, 36], before: [0, 72], after: [0, 72] };
  const clamp = (v, [lo, hi], fallback) => {
    const n = Number(v);
    return Number.isFinite(n) ? Math.min(hi, Math.max(lo, n)) : fallback;
  };
  const pick = (v, choices, fallback) => (choices.some(([k]) => k === v) ? v : fallback);
  const str = (v, fallback, max = 8) => (typeof v === "string" ? v.replace(/[\x00-\x1f\x7f%]/g, "").slice(0, max) : fallback);

  function sanitize(s) {
    if (!s || typeof s !== "object" || s.version !== VERSION || !s.levels) return null;
    const d = PRESETS[["t1", "t2", "t3", "t4"].includes(s.base) ? s.base : "t1"];
    const out = {
      version: VERSION,
      base: d.base,
      layout: s.layout === "separate" ? "separate" : "inline",
      font: str(s.font, d.font, 40).trim() || d.font,
      levels: {},
    };
    for (const k of Object.keys(RANGES)) out[k] = clamp(s[k], RANGES[k], d[k]);
    for (const [k] of LEVELS) {
      const src = s.levels[k] || {}, dl = d.levels[k];
      const l = {
        fmt: pick(src.fmt, FORMATS, dl.fmt), pre: str(src.pre, dl.pre), suf: str(src.suf, dl.suf),
        sep: pick(src.sep, SEPS, dl.sep), bold: typeof src.bold === "boolean" ? src.bold : dl.bold,
        font: str(src.font, "", 40).trim(), align: pick(src.align, ALIGNS, dl.align),
      };
      for (const r of Object.keys(LV_RANGES)) l[r] = clamp(src[r], LV_RANGES[r], dl[r]);
      out.levels[k] = l;
    }
    return out;
  }

  // ---------- 状態 ----------
  let settings = null;
  let onChange = () => {};
  const store = {
    get() { try { return localStorage.getItem(STORE_KEY); } catch { return null; } },
    set(v) { try { localStorage.setItem(STORE_KEY, v); } catch {} },
  };
  function load() {
    try { settings = sanitize(JSON.parse(store.get() || "null")); } catch { settings = null; }
    if (!settings) settings = structuredClone(PRESETS.t1);
  }
  const save = () => store.set(JSON.stringify(settings));

  // ---------- フォーム ----------
  const h = (tag, attrs = {}, ...kids) => {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") el.className = v;
      else if (k === "text") el.textContent = v;
      else el.setAttribute(k, v);
    }
    kids.flat().forEach((c) => el.append(c));
    return el;
  };
  let uid = 0;
  const field = (label, input, cls = "") => {
    const id = "cf" + (++uid);
    input.id = id;
    return h("div", { class: "cf " + cls }, h("label", { for: id, text: label }), input);
  };
  const select = (opts, value) => {
    const el = h("select");
    opts.forEach(([v, t]) => el.append(h("option", { value: v, text: t })));
    el.value = value;
    return el;
  };
  const number = (value, [lo, hi], step) =>
    h("input", { type: "number", min: lo, max: hi, step, inputmode: "decimal", value });
  const fontOptions = (value, withSame) => {
    const opts = (withSame ? [["", "本文と同じ"]] : []).concat(FONTS.map((f) => [f.name, f.name]));
    // 読み込んだ設定に一覧にないフォントがあれば、選べるように足す
    if (value && !FONTS.some((f) => f.name === value)) opts.push([value, value]);
    return select(opts, value);
  };

  let root = null;

  function bindNumber(input, obj, key, range) {
    // 打っている途中は範囲内の値だけ反映し、確定したら範囲に収めて書き戻す
    input.addEventListener("input", () => {
      const n = Number(input.value);
      if (input.value !== "" && Number.isFinite(n) && n >= range[0] && n <= range[1]) {
        obj[key] = n;
        changed();
      }
    });
    input.addEventListener("change", () => {
      obj[key] = clamp(input.value, range, obj[key]);
      input.value = obj[key];
      changed();
    });
  }
  function bindValue(input, obj, key, prop = "value") {
    input.addEventListener(prop === "checked" ? "change" : "input", () => {
      obj[key] = prop === "checked" ? input.checked : input.value;
      if (input.tagName === "INPUT" && input.type === "text") {
        // 「%」は Word の番号の書式と衝突するので入れられないようにする
        obj[key] = str(obj[key], "");
        if (input.value !== obj[key]) input.value = obj[key];
      }
      changed();
    });
  }

  function buildForm() {
    root.querySelector("#cfGlobal").replaceChildren(...globalFields());
    root.querySelector("#cfLevels").replaceChildren(...LEVELS.map(([k, label]) => levelBox(k, label)));
    updateLayoutNotes();
  }

  function globalFields() {
    const s = settings;
    const layout = h("fieldset", { class: "cf-layout" }, h("legend", { text: "条の組み方" }));
    [["inline", "第1条（目的）を1行にする"], ["separate", "（目的）の次の行を「第1条」から始める"]].forEach(([v, t]) => {
      const id = "cfLayout_" + v;
      const r = h("input", { type: "radio", name: "cfLayout", id, value: v });
      r.checked = s.layout === v;
      r.addEventListener("change", () => { s.layout = v; updateLayoutNotes(); changed(); });
      layout.append(h("div", { class: "cf-radio" }, r, h("label", { for: id, text: t })));
    });
    const font = fontOptions(s.font, false);
    bindValue(font, s, "font");
    const size = number(s.size, RANGES.size, 0.5); bindNumber(size, s, "size", RANGES.size);
    const line = number(s.line, RANGES.line, 0.5); bindNumber(line, s, "line", RANGES.line);
    const tsize = number(s.table_size, RANGES.table_size, 0.5); bindNumber(tsize, s, "table_size", RANGES.table_size);
    return [layout, h("div", { class: "cf-grid" },
      field("本文のフォント", font, "wide"),
      field("本文の文字サイズ（pt）", size),
      field("行間（固定値・pt）", line),
      field("表の文字サイズ（pt）", tsize))];
  }

  function levelBox(key, label) {
    const l = settings.levels[key];
    const num = [];
    const fmt = select(FORMATS, l.fmt); bindValue(fmt, l, "fmt");
    const pre = h("input", { type: "text", maxlength: 8, value: l.pre, autocomplete: "off" }); bindValue(pre, l, "pre");
    const suf = h("input", { type: "text", maxlength: 8, value: l.suf, autocomplete: "off" }); bindValue(suf, l, "suf");
    const sep = select(SEPS, l.sep); bindValue(sep, l, "sep");
    num.push(field("番号の形式", fmt), field("番号と本文の間", sep),
      field("番号の前の文字", pre), field("番号の後の文字", suf));
    const first = number(l.first, LV_RANGES.first, 0.5); bindNumber(first, l, "first", LV_RANGES.first);
    const left = number(l.left, LV_RANGES.left, 0.5); bindNumber(left, l, "left", LV_RANGES.left);
    const size = number(l.size, LV_RANGES.size, 0.5); bindNumber(size, l, "size", LV_RANGES.size);
    const font = fontOptions(l.font, true); bindValue(font, l, "font");
    const align = select(ALIGNS, l.align); bindValue(align, l, "align");
    const before = number(l.before, LV_RANGES.before, 1); bindNumber(before, l, "before", LV_RANGES.before);
    const after = number(l.after, LV_RANGES.after, 1); bindNumber(after, l, "after", LV_RANGES.after);
    const bold = h("input", { type: "checkbox" }); bold.checked = l.bold; bindValue(bold, l, "bold", "checked");
    const boldId = "cf" + (++uid); bold.id = boldId;

    return h("details", { class: "cf-level", "data-level": key },
      h("summary", {}, h("span", { class: "cf-lname", text: label }), h("span", { class: "cf-ex" })),
      h("p", { class: "cf-note", "data-note": key }),
      h("div", { class: "cf-grid cf-num" }, num),
      h("div", { class: "cf-grid" },
        field("1行目の開始位置（字）", first), field("2行目以降の開始位置（字）", left),
        field("文字サイズ（pt）", size), field("フォント", font),
        field("配置", align),
        h("div", { class: "cf cf-check" }, bold, h("label", { for: boldId, text: "太字" })),
        field("段落の前の間隔（pt）", before), field("段落の後の間隔（pt）", after)));
  }

  function updateLayoutNotes() {
    const sep = settings.layout === "separate";
    const notes = {
      article: sep ? "見出しの行には番号を表示しません。この番号は第1項の行頭に表示されます。" : "",
      para1: sep ? "番号は「条（見出し）」の設定を使います。" : "",
    };
    root.querySelectorAll("[data-note]").forEach((p) => {
      p.textContent = notes[p.dataset.note] || "";
      p.hidden = !p.textContent;
    });
    const p1 = root.querySelector('[data-level="para1"] .cf-num');
    if (p1) p1.hidden = sep;
  }

  function updateExamples() {
    root.querySelectorAll(".cf-level").forEach((d) => {
      const k = d.dataset.level;
      const src = k === "para1" && settings.layout === "separate" ? settings.levels.article : settings.levels[k];
      const n = k === "paraN" ? 2 : 1;
      d.querySelector(".cf-ex").textContent = numberText(src, n) || "番号なし";
    });
  }

  // ---------- プレビュー ----------
  function fontCss(name) {
    const f = FONTS.find((x) => x.name === name);
    const gothic = /ゴシック|Gothic/i.test(name);
    const q = f ? f.css : `"${name.replace(/["\\]/g, "")}"`;
    return gothic ? `${q}, "BIZ UDGothic", sans-serif` : `${q}, "BIZ UDMincho", serif`;
  }

  // 例文を、段（スタイル）・番号・本文の並びに展開する
  function lines(s) {
    const L = s.levels, sep = s.layout === "separate";
    const out = [];
    let ch = 0, art = 0;
    const push = (lvKey, numLv, n, text) => {
      const num = numLv ? numberText(numLv, n) : "";
      out.push({ lv: L[lvKey], num, sep: num ? numLv.sep : "none", text });
    };
    SAMPLE.forEach((c) => {
      push("chapter", L.chapter, ++ch, c.title);
      c.articles.forEach((a) => {
        art++;
        push("article", sep ? null : L.article, art, `（${a.title}）`);
        a.paragraphs.forEach((p, i) => {
          if (i === 0) push("para1", sep ? L.article : L.para1, sep ? art : 1, p.body);
          else push("paraN", L.paraN, i + 1, p.body);
          (p.items || []).forEach((it, m) => {
            push("item", L.item, m + 1, it.body);
            const subs = it.sub_items || [];
            subs.forEach((su, n) => {
              // 号の下位が1つだけのときは番号を付けない（Python 側と同じ）
              push("sub", subs.length >= 2 ? L.sub : null, n + 1, su.body);
              (su.sub_items2 || []).forEach((s2, o) => push("sub2", L.sub2, o + 1, s2.body));
            });
          });
        });
      });
    });
    return out;
  }

  function renderPreview() {
    const s = settings;
    const page = root.querySelector("#cfPage");
    const unit = s.size;   // 1字 = 本文の文字サイズ（pt）
    page.style.width = TEXT_WIDTH_PT + "pt";
    page.replaceChildren(...lines(s).map((ln) => {
      const l = ln.lv;
      const p = h("p");
      // フォント名は読み込んだ docx から来ることもあるので、1項目ずつ代入する
      // （cssText に連結すると「;」を含む名前で別の指定が混ざる）
      Object.assign(p.style, {
        fontFamily: fontCss(l.font || s.font), fontSize: `${l.size}pt`,
        fontWeight: l.bold ? "700" : "400", lineHeight: `${lineHeight(s, l)}pt`,
        margin: `${l.before}pt 0 ${l.after}pt`, paddingLeft: `${l.left * unit}pt`,
        textIndent: `${(l.first - l.left) * unit}pt`,
        textAlign: l.align === "center" ? "center" : "justify",
      });
      if (ln.num) {
        const sepText = { zen: "\u3000", space: " " }[ln.sep] || "";
        const span = h("span", { class: "pv-num", text: ln.num + sepText });
        if (ln.sep === "tab") { span.dataset.tab = ""; span.dataset.first = l.first * unit; span.dataset.left = l.left * unit; }
        p.append(span);
      }
      p.append(ln.text);
      return p;
    }));
    // タブ：番号が2行目以降の開始位置より手前で終われば、そこまで進む。
    // 越える場合は Word と同じく次の既定のタブ位置まで進む。
    // ponytail: 中央揃えの段のタブは位置を厳密に再現しない（章に使う想定がないため）
    page.querySelectorAll(".pv-num[data-tab]").forEach((sp) => {
      const first = Number(sp.dataset.first), left = Number(sp.dataset.left);
      const w = sp.getBoundingClientRect().width / currentScale * 0.75; // px → pt
      const end = first + w;
      const target = end < left - 0.1 ? left : (Math.floor(end / DEFAULT_TAB_PT) + 1) * DEFAULT_TAB_PT;
      sp.style.width = (target - first) + "pt";
    });
    fit();
  }

  let currentScale = 1;
  function fit() {
    const wrap = root.querySelector("#cfFit");
    const page = root.querySelector("#cfPage");
    page.style.transform = "none";
    wrap.style.width = "";
    const natural = page.offsetWidth;
    // 「実際の大きさで表示」では縮小せず、横にスクロールして見る
    const actual = root.querySelector("#pvActual").checked;
    currentScale = actual ? 1 : Math.min(1, wrap.clientWidth / natural);
    if (actual) wrap.style.width = natural + "px";
    page.style.transform = `scale(${currentScale})`;
    wrap.style.height = page.offsetHeight * currentScale + "px";
  }

  function changed() {
    save();
    updateExamples();
    renderPreview();
    onChange(settings);
  }

  // ---------- 公開する操作 ----------
  window.KisokuCustom = {
    PRESETS, SAMPLE,
    init(el, opts = {}) {
      root = el;
      onChange = opts.onChange || onChange;
      load();
      root.querySelector("#pvActual").addEventListener("change", () => renderPreview());
      buildForm();
      updateExamples();
      // 非表示のうちは幅が0なので、表示されてから描き直す
      // 枠の幅が変わったとき（表示された・画面を回転した）だけ描き直す
      let lastWidth = -1;
      const box = root.querySelector(".pv-scroll");
      new ResizeObserver(() => {
        if (box.clientWidth === lastWidth || !root.offsetParent) return;
        lastWidth = box.clientWidth;
        renderPreview();
      }).observe(box);
      if (document.fonts) document.fonts.ready.then(() => { if (root.offsetParent) renderPreview(); });
    },
    get: () => structuredClone(settings),
    set(s) {
      const v = sanitize(s);
      if (!v) return false;
      settings = v;
      save();
      buildForm();
      changed();
      return true;
    },
    usePreset(key) { return this.set(structuredClone(PRESETS[key])); },
    refresh() { if (root.offsetParent) renderPreview(); },
  };
})();
