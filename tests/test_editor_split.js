// editor.js の「ここから新しい条にする」「前の条にまとめる」のテスト。実行: node tests/test_editor_split.js
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

// editor.js は読み込み時にメニューを閉じるイベントを登録するので、document は形だけ用意する
const ctx = { window: {}, document: { addEventListener() {} } };
vm.runInNewContext(fs.readFileSync(path.join(__dirname, "..", "editor.js"), "utf8"), ctx);
const E = ctx.window.KisokuEditor;
const plain = (x) => JSON.parse(JSON.stringify(x));
const para = (body, extra = {}) => Object.assign({ type: "paragraph", number: 0, body, items: [] }, extra);
const art = (title, bodies) => ({ type: "article", number: 1, branch: 0, number_str: "", title, title_note: "",
  paragraphs: bodies.map((b) => (typeof b === "string" ? para(b) : b)) });
// 解析で「条見出しの疑い」の印が付いた項
const marked = (body, hint, extra = {}) => para(body, Object.assign({ _article_hint: [hint] }, extra));
const second = (p) => plain(E.splitArticle(art("", ["一", p]), 1))[1];

// 第5項から分ける：第1〜4項が残り、第5〜8項が新しい条の第1〜4項になる。元の条は変えない
{
  const a = art("試用期間", ["一", "二", "三", "四", "五", "六", "七", "八"]);
  const [x, y] = plain(E.splitArticle(a, 4));
  assert.deepEqual(x.paragraphs.map((p) => [p.number, p.body]), [[1, "一"], [2, "二"], [3, "三"], [4, "四"]]);
  assert.deepEqual(y.paragraphs.map((p) => [p.number, p.body]), [[1, "五"], [2, "六"], [3, "七"], [4, "八"]]);
  assert.equal(x.title, "試用期間");
  assert.equal(y.title, "");
  assert.equal(a.paragraphs.length, 8);
}

// 印のある項：括弧の中をタイトルにし、見出しの部分だけ除く。号・表も一緒に移り、印は消える
{
  const p = marked("第６条＃２（本採用拒否）次の各号の…", "第６条＃２（本採用拒否）",
    { items: [{ type: "item", number: 1, body: "号" }], tables: [{ rows: [] }] });
  const [x, y] = plain(E.splitArticle(art("試用期間", ["一", p, "二"]), 1));
  assert.equal(x.paragraphs.length, 1);
  assert.deepEqual([y.title, y.paragraphs[0].body], ["本採用拒否", "次の各号の…"]);
  assert.equal(y.paragraphs[0].items[0].body, "号");
  assert.equal(y.paragraphs[0].tables.length, 1);
  assert.equal(y.paragraphs[0]._article_hint, undefined);
  assert.equal(y.paragraphs[1].number, 2);
}

// 印のある項の、いろいろな書き方
{
  // 入れ子の括弧のタイトル
  let y = second(marked("第6条（賃金（基本給等））支給額を定める。", "第6条（賃金（基本給等））"));
  assert.deepEqual([y.title, y.paragraphs[0].body], ["賃金（基本給等）", "支給額を定める。"]);
  // 平仮名で始まる本文
  y = second(marked("第６条－２（解雇）やむを得ない事由", "第６条－２（解雇）"));
  assert.deepEqual([y.title, y.paragraphs[0].body], ["解雇", "やむを得ない事由"]);
  // 参照の書き換えで番号が変わった・参照先なしの印が付いた
  y = second(marked("第７条＃２（本採用拒否）次の", "第６条＃２（本採用拒否）"));
  assert.deepEqual([y.title, y.paragraphs[0].body], ["本採用拒否", "次の"]);
  y = second(marked("第６条［参照先なし］＃２（本採用拒否）次の", "第６条＃２（本採用拒否）"));
  assert.deepEqual([y.title, y.paragraphs[0].body], ["本採用拒否", "次の"]);
}

// 見出しだけの項は、中に号・表・図が無ければ除き、次の項を第1項にする
{
  const [, y] = plain(E.splitArticle(art("", ["一", marked("第6条の2(兼業)", "第6条の2(兼業)"), "許可制とする。"]), 1));
  assert.deepEqual([y.title, y.paragraphs.map((p) => p.body)], ["兼業", ["許可制とする。"]]);
  const [, z] = plain(E.splitArticle(art("", ["一",
    marked("第6条の2【兼業】", "第6条の2【兼業】", { items: [{ type: "item", number: 1, body: "号" }] }), "二"]), 1));
  assert.equal(z.paragraphs.length, 2);
}

// 印の無い項は、見出しの形でも参照でも本文に手を付けない（参照の文言を消さないため）
for (const body of ["第６条－２（本採用拒否）次の各号の…", "第6条（安全への配慮）　できる限り",
  "第25条（休職）と第26条（復職）を適用する。", "第25条（休職）　と第26条", "第6条の2（兼業）違反を理由として",
  "第10条（休職）の規定により", "（ただし、管理職を除く。）本文", "第25条の遵守事項に違反したとき"]) {
  const y = second(body);
  assert.deepEqual([y.title, y.paragraphs[0].body], ["", body], body);
}
// 印があっても、タイトルが違う見出しで始まる本文（参照）には手を付けない
{
  const y = second(marked("第25条（休職）と第26条", "第６条＃２（本採用拒否）"));
  assert.deepEqual([y.title, y.paragraphs[0].body], ["", "第25条（休職）と第26条"]);
}

// 第1項・範囲外では分けない
assert.equal(E.splitArticle(art("", ["一", "二"]), 0), null);
assert.equal(E.splitArticle(art("", ["一", "二"]), 2), null);

// 前の条にまとめる：タイトルは本文の先頭に残し、番号は続きになる
{
  const m = plain(E.mergeArticles(art("試用期間", ["一", "二"]), art("本採用拒否", ["三", "四"])));
  assert.deepEqual(m.paragraphs.map((p) => [p.number, p.body]), [[1, "一"], [2, "二"], [3, "（本採用拒否）三"], [4, "四"]]);
  assert.equal(m.title, "試用期間");
  // タイトルが無ければ本文はそのまま。項が無い条はタイトルだけの項になる
  assert.equal(plain(E.mergeArticles(art("a", ["一"]), art("", ["二"]))).paragraphs[1].body, "二");
  assert.equal(plain(E.mergeArticles(art("a", ["一"]), art("t", []))).paragraphs[1].body, "（t）");
  // 条の前の節の行（lead）も消さない
  const b = Object.assign(art("t", ["二"]), { lead: ["第2節　服務"] });
  assert.deepEqual(plain(E.mergeArticles(art("a", ["一"]), b)).paragraphs.map((p) => p.body), ["一", "第2節　服務", "（t）二"]);
}

// 分けてまとめ直すと、本文の文言は（タイトルの括弧書きとして）残る
{
  const a = art("試用期間", ["一", marked("第６条－２（本採用拒否）次の各号の…", "第６条－２（本採用拒否）"), "三"]);
  const [x, y] = E.splitArticle(a, 1);
  assert.deepEqual(plain(E.mergeArticles(x, y)).paragraphs.map((p) => p.body), ["一", "（本採用拒否）次の各号の…", "三"]);
}

console.log("OK");
