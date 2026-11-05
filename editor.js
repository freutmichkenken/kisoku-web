// 就業規則 整形アプリ — 改正案の編集欄（新旧対照表の右側で、条・章を1つずつ編集する）
// 木（py/kisoku_parser.py の出力）の条・章を複製して編集し、確定したら呼び出し元に返す。
// 文字はすべて textContent / value で扱い、innerHTML に本文を入れない。
(() => {
  const ZEN = (n) => String(n).replace(/[0-9]/g, (d) => "０１２３４５６７８９"[d]);
  const KANA = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン";
  const kana = (n) => KANA[n - 1] || ZEN(n);
  const clone = (v) => JSON.parse(JSON.stringify(v));

  // 子の一覧のキーと、新しく作る子の形
  const CHILD = { paragraph: "items", item: "sub_items", sub: "sub_items2" };
  const NEW = {
    paragraph: () => ({ type: "paragraph", number: 0, body: "", items: [] }),
    item: () => ({ type: "item", number: 0, body: "" }),
    sub: () => ({ type: "item_sub", number: 0, body: "" }),
    sub2: () => ({ type: "item_sub2", number: 0, body: "" }),
  };
  const CHILD_KIND = { paragraph: "item", item: "sub", sub: "sub2" };
  const ADD_LABEL = { paragraph: "号を追加", item: "細分を追加" };   // sub の下（さらに下）は追加しない
  const INSIDE = { paragraph: "号・表・図", item: "細分・表・図", sub: "細分・表・図", sub2: "表・図" };

  function el(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }
  function button(text, onClick, cls = "btn small", aria) {
    const b = el("button", cls, text);
    b.type = "button";
    if (aria) b.setAttribute("aria-label", aria);
    b.addEventListener("click", onClick);
    return b;
  }
  // 中に号・表・図があるか（削除の確認に使う）
  const hasInside = (node) => {
    const key = CHILD[kindOf(node)];
    return (key && (node[key] || []).length > 0) || (node.tables || []).length > 0 || (node.drawings || []).length > 0;
  };
  function kindOf(node) {
    if (node.type === "paragraph") return "paragraph";
    if (node.type === "item_sub") return "sub";
    if (node.type === "item_sub2") return "sub2";
    return "item";
  }
  function marks(node) {
    const out = [];
    (node.tables || []).forEach(() => out.push("［表］"));
    (node.drawings || []).forEach((d) => out.push(d.type === "image" ? "［画像］" : "［図］"));
    return out;
  }

  // ---------- 表・図・画像 ----------
  const IMAGE_MAX = 5 * 1024 * 1024;     // 1枚の上限（py/web_entry.py の IMAGE_MAX_BYTES と同じ）
  const IMAGE_TYPES = ["image/png", "image/jpeg"];
  function figureText(key, f) {
    if (key === "tables") {
      const cells = (f.rows || []).flat().filter(Boolean).join("／");
      return "［表］" + (cells ? "　" + cells.slice(0, 40) : "");
    }
    if (f.type === "image") return "［画像］" + (f.name ? "　" + f.name.slice(0, 40) : "");
    return "［図］" + (f.text ? "　" + f.text.slice(0, 40) : "");
  }
  // 表・図・画像の1行。↑↓は条の中の前・次の項・号へ移す（整形ではその項・号の直後に出る）
  // ponytail: 同じ項・号の中での並べ替えはできない（表が先、図・画像が後の順で出る）。要るなら一覧内の入れ替えを足す。
  function figureRow(node, key, idx, ctx) {
    const f = node[key][idx];
    const row = el("div", "ed-fig");
    if (f.type === "image" && f.data) {
      const img = el("img", "ed-thumb");
      img.alt = "";
      img.src = `data:${f.mime};base64,${f.data}`;
      row.appendChild(img);
    }
    row.appendChild(el("span", "ed-mark", figureText(key, f)));
    const tools = el("span", "ed-tools");
    const at = ctx.units.indexOf(node);
    const what = key === "tables" ? "表" : f.type === "image" ? "画像" : "図";
    tools.appendChild(button("↑", () => ctx.move(node, key, idx, -1), "btn small", "前の項・号へ移動")).disabled = at <= 0;
    tools.appendChild(button("↓", () => ctx.move(node, key, idx, 1), "btn small", "次の項・号へ移動")).disabled = at < 0 || at >= ctx.units.length - 1;
    tools.appendChild(button("削除", () => {
      if (!confirm(`${what}を削除しますか？`)) return;
      node[key].splice(idx, 1);
      ctx.redraw();
    }, "btn small"));
    row.appendChild(tools);
    return row;
  }
  function figureRows(node, ctx) {
    const out = [];
    ["tables", "drawings"].forEach((key) => (node[key] || []).forEach((_, i) => out.push(figureRow(node, key, i, ctx))));
    return out;
  }
  // 画像を選ばせて node.drawings の末尾に足す
  function pickImage(node, ctx) {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = IMAGE_TYPES.join(",");
    input.addEventListener("change", () => {
      const file = input.files && input.files[0];
      if (!file) return;
      if (!IMAGE_TYPES.includes(file.type) || file.size === 0 || file.size > IMAGE_MAX) {
        alert("画像は PNG か JPEG（1枚 5MB まで）にしてください。");
        return;
      }
      const r = new FileReader();
      r.onload = () => {
        const s = String(r.result);
        const data = s.slice(s.indexOf(",") + 1);
        node.drawings = node.drawings || [];
        node.drawings.push({ type: "image", name: file.name, mime: file.type, data });
        ctx.redraw();
      };
      r.onerror = () => alert("画像を読み込めませんでした。");
      r.readAsDataURL(file);
    });
    input.click();
  }
  // 条の中の項・号・細分…を、文書の順（整形で出る順）に並べる
  function unitsOf(art) {
    const out = [];
    (art.paragraphs || []).forEach((p) => {
      out.push(p);
      (p.items || []).forEach((it) => {
        out.push(it);
        (it.sub_items || []).forEach((s) => {
          out.push(s);
          (s.sub_items2 || []).forEach((s2) => out.push(s2));
        });
      });
    });
    return out;
  }
  function label(kind, i) {
    if (kind === "paragraph") return `第${i}項`;
    if (kind === "item") return `（${ZEN(i)}）`;
    if (kind === "sub") return kana(i);
    return `（${kana(i)}）`;
  }

  // 番号を位置で振り直す（整形の出力は位置で番号を付ける。第1項の書式は number===1 で決まる）
  function renumber(art) {
    (art.paragraphs || []).forEach((p, i) => {
      p.number = i + 1;
      (p.items || []).forEach((it, j) => {
        it.number = j + 1;
        (it.sub_items || []).forEach((s, k) => {
          s.number = k + 1;
          (s.sub_items2 || []).forEach((s2, m) => { s2.number = m + 1; });
        });
      });
    });
    return art;
  }

  function newArticle() {
    return { type: "article", number: 0, branch: 0, number_str: "", title: "", title_note: "",
             paragraphs: [NEW.paragraph()] };
  }
  function newChapter() {
    return { type: "chapter", number: 0, number_str: "", title: "", articles: [] };
  }

  // 一覧（項・号・細分…）の1つ分の編集欄。list は親の配列、i はその中の位置。
  function unitEditor(list, i, kind, ctx) {
    const redraw = ctx.redraw;
    const node = list[i];
    const box = el("div", "ed-unit ed-" + kind);
    const head = el("div", "ed-unit-head");
    head.appendChild(el("span", "ed-label", label(kind, i + 1)));
    const tools = el("span", "ed-tools");
    tools.appendChild(button("↑", () => { [list[i - 1], list[i]] = [list[i], list[i - 1]]; redraw(); },
      "btn small", "上へ")).disabled = i === 0;
    tools.appendChild(button("↓", () => { [list[i + 1], list[i]] = [list[i], list[i + 1]]; redraw(); },
      "btn small", "下へ")).disabled = i === list.length - 1;
    tools.appendChild(button("削除", () => {
      const what = label(kind, i + 1);
      const msg = hasInside(node) ? `${what}を削除しますか？中の${INSIDE[kind]}も消えます。` : `${what}を削除しますか？`;
      if (!confirm(msg)) return;
      list.splice(i, 1);
      redraw();
    }, "btn small"));
    const childKey = CHILD[kind];
    if (ADD_LABEL[kind]) {
      tools.appendChild(button(ADD_LABEL[kind], () => {
        node[childKey] = node[childKey] || [];
        node[childKey].push(NEW[CHILD_KIND[kind]]());
        redraw();
      }, "btn small"));
    }
    tools.appendChild(button("画像を追加", () => pickImage(node, ctx), "btn small"));
    head.appendChild(tools);
    box.appendChild(head);
    const ta = el("textarea", "ed-text");
    ta.value = node.body || "";
    ta.rows = Math.min(8, Math.max(2, Math.ceil((ta.value.length || 1) / 40)));
    ta.setAttribute("aria-label", label(kind, i + 1));
    ta.addEventListener("input", () => { node.body = ta.value; });
    box.appendChild(ta);
    figureRows(node, ctx).forEach((r) => box.appendChild(r));
    if (childKey) {
      (node[childKey] || []).forEach((_, j) => box.appendChild(unitEditor(node[childKey], j, CHILD_KIND[kind], ctx)));
    }
    return box;
  }

  function footer(items) {
    const f = el("div", "ed-foot");
    items.forEach((b) => f.appendChild(b));
    return f;
  }

  // 条の編集欄。h: { commit(art), cancel(), insertAfter(art), remove() }
  function articleForm(article, h) {
    const art = clone(article);
    const form = el("div", "ed-form");
    const redraw = () => {
      form.textContent = "";
      const t = el("label", "ed-field");
      t.appendChild(el("span", "ed-label", "条タイトル"));
      const inp = el("input", "ed-title");
      inp.type = "text"; inp.value = art.title || "";
      inp.addEventListener("input", () => { art.title = inp.value; });
      t.appendChild(inp);
      form.appendChild(t);
      const ctx = {
        redraw,
        units: unitsOf(art),
        // 表・図・画像を前（dir=-1、その項・号の末尾へ）か次（dir=1、先頭へ）の項・号へ移す
        move(node, key, idx, dir) {
          const target = this.units[this.units.indexOf(node) + dir];
          if (!target) return;
          const [f] = node[key].splice(idx, 1);
          target[key] = target[key] || [];
          if (dir < 0) target[key].push(f); else target[key].unshift(f);
          redraw();
        },
      };
      (art.paragraphs || []).forEach((_, i) => form.appendChild(unitEditor(art.paragraphs, i, "paragraph", ctx)));
      marks(art).forEach((m) => form.appendChild(el("div", "ed-mark", m)));
      form.appendChild(button("項を追加", () => {
        art.paragraphs = art.paragraphs || [];
        art.paragraphs.push(NEW.paragraph());
        redraw();
      }));
      form.appendChild(footer([
        button("確定", () => h.commit(renumber(art)), "btn small primary"),
        button("取消", () => h.cancel()),
        button("下に条を追加", () => h.insertAfter(renumber(art))),
        button("条を削除", () => { if (confirm("この条を削除しますか？")) h.remove(); }),
      ]));
    };
    redraw();
    return form;
  }

  // 章の編集欄。h: { commit(ch), cancel(), insertArticle(ch), insertAfter(ch), remove() }
  function chapterForm(chapter, h) {
    const ch = clone(chapter);
    const form = el("div", "ed-form");
    const t = el("label", "ed-field");
    t.appendChild(el("span", "ed-label", "章タイトル"));
    const inp = el("input", "ed-title");
    inp.type = "text"; inp.value = ch.title || "";
    inp.addEventListener("input", () => { ch.title = inp.value; });
    t.appendChild(inp);
    form.appendChild(t);
    form.appendChild(footer([
      button("確定", () => h.commit(ch), "btn small primary"),
      button("取消", () => h.cancel()),
      button("先頭に条を追加", () => h.insertArticle(ch)),
      button("下に章を追加", () => h.insertAfter(ch)),
      button("章を削除", () => {
        const n = (ch.articles || []).length;
        if (confirm(n ? `この章と章内の${n}条を削除しますか？` : "この章を削除しますか？")) h.remove();
      }),
    ]));
    return form;
  }

  window.KisokuEditor = { articleForm, chapterForm, newArticle, newChapter, renumber, clone };
})();
