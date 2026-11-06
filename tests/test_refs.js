// refs.js（本文中の「第◯条」の書き換え）のテスト。実行: node tests/test_refs.js
"use strict";
const assert = require("node:assert/strict");
const R = require("../refs.js");

const art = (number, body, branch = 0) => ({ type: "article", number, branch, number_str: "", title: "",
  paragraphs: [{ type: "paragraph", number: 1, body, items: [] }] });
const bodies = (tree) => tree.flatMap((n) => (n.type === "chapter" ? n.articles : [n])).map((a) => a.paragraphs[0].body);
const clone = (x) => JSON.parse(JSON.stringify(x));

// 数の読み書き
assert.equal(R.parseNum("１２"), 12);
assert.equal(R.parseNum("十二"), 12);
assert.equal(R.parseNum("百五"), 105);
assert.equal(R.formatNum(13, "１２"), "１３");
assert.equal(R.formatNum(13, "12"), "13");
assert.equal(R.formatNum(20, "十九"), "二十");

// 条を間に足すと、後ろの条への参照が1つずつずれる。法令の参照・連なった参照はそのまま
{
  const before = R.assignIds([
    art(1, "第２条及び第3条による。労基法第２条、第３条並びにこの規則第3条を参照。"),
    art(2, "第十条は無い。高年齢者雇用安定法第9条第2項に基づく。"),
    art(3, "第１条から前条まで。賃金規程第2条による。"),
  ]);
  const after = clone(before);
  after.splice(1, 0, art(0, "新しい条"));
  R.assignIds(after);
  const missing = R.rewriteTree(before, after, false);
  assert.deepEqual(bodies(after), [
    "第３条及び第4条による。労基法第２条、第３条並びにこの規則第4条を参照。",
    "新しい条",
    "第十条［参照先なし］は無い。高年齢者雇用安定法第9条第2項に基づく。",
    "第１条から前条まで。賃金規程第2条による。",
  ]);
  assert.deepEqual(missing, ["第十条"].map((ref) => ({ at: "第3条", ref })));
}

// 法令名の直後の括弧、労働協約・定款、項がいくつも挟まる並び
{
  const before = R.assignIds([art(1, "労働基準法（第2条）、労働協約第2条、定款第2条、労基法第2条第1項、第2項及び第3条、本規則第2条"),
    art(2, "x"), art(3, "y")]);
  const after = clone(before);
  after.unshift(art(0, "new"));
  R.assignIds(after);
  R.rewriteTree(before, after, false);
  assert.equal(bodies(after)[1], "労働基準法（第2条）、労働協約第2条、定款第2条、労基法第2条第1項、第2項及び第3条、本規則第3条");
}

// 名前の後の開き括弧だけでは外部扱いにしない（「懲戒の基準（第2条）」は自社の条）
{
  const before = R.assignIds([art(1, "懲戒の基準（第2条）による。"), art(2, "x")]);
  const after = clone(before);
  after.unshift(art(0, "new"));
  R.assignIds(after);
  R.rewriteTree(before, after, false);
  assert.equal(bodies(after)[1], "懲戒の基準（第3条）による。");
}

// 括弧書き・「の規定（」を挟んだ法令の参照、項・号や「まで及び」でつながる参照、ほかの規程名
{
  const before = R.assignIds([
    art(1, "労働基準法（昭和22年法律第49号）第2条、労基法の規定（第3条）、労基法第2条第1項第1号、第3条、" +
      "労基法第2条から第2条の5まで及び第3条、出張旅費細則第2条、本規程第2条、第2条"),
    art(2, "x"), art(3, "y"),
  ]);
  const after = clone(before);
  after.unshift(art(0, "new"));
  R.assignIds(after);
  R.rewriteTree(before, after, false);
  assert.equal(bodies(after)[1], "労働基準法（昭和22年法律第49号）第2条、労基法の規定（第3条）、労基法第2条第1項第1号、第3条、" +
    "労基法第2条から第2条の5まで及び第3条、出張旅費細則第2条、本規程第3条、第3条");
}

// 条を消すと、その条を指す参照に印を付けて知らせる。印の付いた参照は、後の編集で書き換えない
{
  const before = R.assignIds([art(1, "第２条と第３条"), art(2, "x"), art(3, "y")]);
  const after = clone(before);
  after.splice(1, 1);
  const missing = R.rewriteTree(before, after, false);
  assert.deepEqual(bodies(after), ["第２条［参照先なし］と第２条", "y"]);
  assert.deepEqual(missing, [{ at: "第1条", ref: "第２条" }]);
  assert.deepEqual(R.deadRefs(after), [{ at: "第1条", ref: "第２条" }]);
  const next = clone(after);
  next.unshift(art(0, "new"));
  R.assignIds(next);
  R.rewriteTree(after, next, false);
  assert.deepEqual(bodies(next).slice(1), ["第２条［参照先なし］と第３条", "y"]);
}

// 元の docx の番号（欠番・枝番号）から通し番号へ
{
  const base = R.assignIds([art(1, "第7条の2及び第8条"), art(7, "a"), art(7, "b", 2), art(8, "c")]);
  const tree = clone(base);
  R.rewriteTree(base, tree, true);
  assert.equal(bodies(tree)[0], "第3条及び第4条");
}

// 章の中の条・号の本文も対象。漢数字は漢数字のまま
{
  const before = R.assignIds([{ type: "chapter", title: "", articles: [
    art(1, "a"),
    { type: "article", number: 2, title: "", paragraphs: [{ type: "paragraph", body: "",
      items: [{ type: "item", body: "第一条に違反したとき", sub_items: [] }] }] },
  ] }]);
  const after = clone(before);
  after[0].articles.unshift(art(0, "new"));
  R.assignIds(after);
  R.rewriteTree(before, after, false);
  assert.equal(after[0].articles[2].paragraphs[0].items[0].body, "第二条に違反したとき");
}

// 参照のチェック：無い条・無い項・無い号を拾い、法令への参照とある項・号は拾わない
{
  const two = { type: "article", number: 2, title: "", paragraphs: [
    { type: "paragraph", body: "x", items: [{ type: "item", body: "a" }, { type: "item", body: "b" }] },
    { type: "paragraph", body: "y", items: [] },
  ] };
  const tree = [{ type: "chapter", title: "", articles: [
    art(1, `第2条第1項第2号、第2条第3項、第2条第1項第3号、第2条第2項第1号、労働基準法第2条第5項による。第9条${R.DEAD}`),
    two,
  ] }];
  const r = R.checkRefs(tree);
  assert.deepEqual(r.map((x) => [x.kind, x.ref]), [
    ["para", "第2条第3項"], ["item", "第2条第1項第3号"], ["item", "第2条第2項第1号"], ["dead", "第9条"],
  ]);
  assert.deepEqual(r[0].path, [0, 0]);
  assert.equal(r[0].at, "第1条");
  // 元の番号で書かれた木（欠番あり）：第5条は2番目の条
  const orig = [art(1, "第5条第2項、第5条第3項"), { ...two, number: 5 }];
  assert.deepEqual(R.checkRefs(orig, true).map((x) => x.ref), ["第5条第3項"]);
}

console.log("test_refs: OK");
