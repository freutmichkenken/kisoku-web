// 本文中の「第◯条」の参照を、条の追加・削除・並べ替えに合わせて書き換える。
// 画面（index.html）とテスト（tests/test_refs.js）の両方から使うので、DOM には触らない。
(function (root) {
  "use strict";

  const ZEN = "０１２３４５６７８９";
  const KAN = "〇一二三四五六七八九";
  const NUM = "[0-9０-９]+|[〇一二三四五六七八九十百]+";
  const REF = new RegExp(`第(${NUM})条(?:の(${NUM}))?`, "g");
  // 直前がこれで終わる「第◯条」は、ほかの法令・規程などへの参照なので書き換えない
  const EXTERNAL = /(法律?|令|規則|規程|条例|附則|協定書?|協約|定款|告示|通達|指針|細則|要領|内規|基準|約款|契約書?|同法|同令)$/;
  // この語の直後に開き括弧があれば、括弧の中の「第◯条」もその法令などを指す（「労働基準法（第89条）」）。
  // 「基準」「規則」のように一般の語と紛らわしいもの（「懲戒の基準（第10条）」）は含めない
  const LAW_PAREN = /(法律?|令|条例|協約|定款|同法|同令)[（(]$/;
  // ただし「この規則」「本規程」などは、この規則そのものを指す
  const SELF = /(この|本|当)(就業)?(規則|規程)$/;
  // 前の参照と、この区切りだけでつながっていれば、前の参照と同じ扱いにする
  // （「労基法第32条、第33条」「第32条第1項第1号、第33条」「第32条から第32条の5まで及び第40条」など）
  //  （「第32条第1項、第2項及び第33条」のように項・号がいくつ挟まってもよい。つなぎの語は1つ以上要る）
  const D = "[0-9０-９〇一二三四五六七八九十百]+";
  const CONN = "、|，|,|及び|および|又は|または|若しくは|もしくは|並びに|ならびに|から|ないし|乃至|～|〜|・";
  // 参照先の無い参照（削除した条・元から無い番号）に付ける印。印の付いた参照は、以後書き換えない
  // （印が無いと、後の編集で同じ番号になった別の条への参照とみなしてしまう）
  const DEAD = "［参照先なし］";
  const CHAIN = new RegExp(`^(?:第${D}項|第${D}号|の規定|まで|前段|後段|ただし書|${DEAD}|${CONN})*(?:${CONN})` +
    `(?:第${D}項|第${D}号)?$`);
  // 参照の直前の文字列から、名前のすぐ後の括弧書き（「（昭和22年法律第49号）」「（以下「労基法」という。）」）と
  // 「の規定（」「の定め（」を取り除き、その前の名前で判定できるようにする。
  // 開き括弧だけ（「懲戒の基準（第10条）」）は外さない（「基準」で外部と誤判定しないように）
  function nameBefore(before) {
    let b = before.replace(/[\s　]+$/, "");
    for (let i = 0; i < 5; i++) {
      const prev = b;
      b = b.replace(/（[^（）]*）$/, "").replace(/\([^()]*\)$/, "")
        .replace(/の(?:規定|定め)[（(]$/, "").replace(/[\s　]+$/, "");
      if (LAW_PAREN.test(b)) b = b.slice(0, -1);
      if (b === prev) break;
    }
    return b;
  }

  function parseNum(s) {
    if (/^[0-9０-９]+$/.test(s)) return Number(s.replace(/[０-９]/g, (c) => ZEN.indexOf(c)));
    // 漢数字（百の位まで）
    let n = 0, cur = 0;
    for (const c of s) {
      if (c === "百") { n += (cur || 1) * 100; cur = 0; }
      else if (c === "十") { n += (cur || 1) * 10; cur = 0; }
      else cur = cur * 10 + KAN.indexOf(c);
    }
    return n + cur;
  }
  // 元の書き方（半角・全角・漢数字）で数を書く
  function formatNum(n, like) {
    if (/^[0-9]+$/.test(like)) return String(n);
    if (/^[０-９]+$/.test(like)) return String(n).replace(/[0-9]/g, (d) => ZEN[d]);
    if (n >= 1000) return String(n).replace(/[0-9]/g, (d) => ZEN[d]);
    const h = Math.floor(n / 100), t = Math.floor((n % 100) / 10), o = n % 10;
    return (h ? (h > 1 ? KAN[h] : "") + "百" : "") + (t ? (t > 1 ? KAN[t] : "") + "十" : "") + (o ? KAN[o] : "");
  }

  // 条の一覧（文書の順）。章の中の条も含む
  function articles(tree) {
    const out = [];
    (tree || []).forEach((n) => {
      if (!n || typeof n !== "object") return;
      if (n.type === "chapter") (n.articles || []).forEach((a) => { if (a && typeof a === "object") out.push(a); });
      else if (n.type === "article") out.push(n);
    });
    return out;
  }

  // 条に見分けるための印（_rid）を付ける。付いていない条（新しく足した条）にだけ付ける
  function assignIds(tree) {
    const all = articles(tree);
    let max = 0;
    all.forEach((a) => { if (Number.isInteger(a._rid)) max = Math.max(max, a._rid); });
    const seen = new Set();
    all.forEach((a) => {
      // 複製した条（コピー）で同じ印が重なったら、後ろの方を新しい条として扱う
      if (!Number.isInteger(a._rid) || seen.has(a._rid)) a._rid = ++max;
      seen.add(a._rid);
    });
    return tree;
  }

  // 番号の付け方 → { 番号のキー: _rid }
  //   seq: 通し番号（改正案の表示と同じ）。original: 元の docx の番号（number・branch）
  function numbering(tree, original) {
    const map = new Map();
    articles(tree).forEach((a, i) => {
      if (!Number.isInteger(a._rid)) return;
      const key = original && Number.isInteger(a.number) && a.number > 0
        ? (a.branch ? `${a.number}の${a.branch}` : String(a.number))
        : String(i + 1);
      if (!map.has(key)) map.set(key, a._rid);
    });
    return map;
  }

  // 1つの文字列の参照を書き換える。from: キー→_rid、to: _rid→新しい番号
  // 戻り値 { text, missing: [消えた条を指していた参照の文字列] }
  function rewriteText(text, from, to) {
    if (typeof text !== "string" || text.indexOf("条") < 0) return { text, missing: [] };
    const missing = [];
    let out = "", last = 0, prevEnd = -1, prevExternal = false;
    REF.lastIndex = 0;
    let m;
    while ((m = REF.exec(text))) {
      const before = text.slice(0, m.index);
      const between = prevEnd >= 0 ? text.slice(prevEnd, m.index) : null;
      let external;
      if (between !== null && CHAIN.test(between)) external = prevExternal;
      else {
        const b = nameBefore(before);
        external = EXTERNAL.test(b) && !SELF.test(b);
      }
      prevEnd = m.index + m[0].length; prevExternal = external;
      if (external) continue;
      if (text.startsWith(DEAD, m.index + m[0].length)) continue;   // 削除した条への参照（印付き）
      const key = m[2] ? `${parseNum(m[1])}の${parseNum(m[2])}` : String(parseNum(m[1]));
      // 参照先の無い参照（消えた条・元から無い番号）は、番号を残して印を付ける
      const n = from.has(key) ? to.get(from.get(key)) : undefined;
      const repl = n === undefined ? m[0] + DEAD : `第${formatNum(n, m[1])}条`;
      if (n === undefined) missing.push(m[0]);
      if (repl === m[0]) continue;
      out += text.slice(last, m.index) + repl;
      last = m.index + m[0].length;
    }
    return { text: out + text.slice(last), missing };
  }

  // 条の本文（前文・項・号・号の下位）を順にたどる
  function eachText(art, fn) {
    if (Array.isArray(art.lead)) art.lead = art.lead.map((t) => (typeof t === "string" ? fn(t) : t));
    const walk = (list, kids) => (Array.isArray(list) ? list : []).forEach((u) => {
      if (!u || typeof u !== "object") return;
      if (typeof u.body === "string") u.body = fn(u.body);
      kids.forEach((k) => walk(u[k], k === "items" ? ["sub_items"] : k === "sub_items" ? ["sub_items2"] : []));
    });
    walk(art.paragraphs, ["items"]);
  }

  // before の番号の付け方（original なら元の番号）で書かれた参照を、after の通し番号に書き換える。
  // after の条には _rid が付いていること。戻り値は消えた条を指している参照 [{ at: "第５条", ref: "第３条" }]
  function rewriteTree(before, after, original) {
    const from = numbering(before, original);
    const to = new Map();
    articles(after).forEach((a, i) => { if (Number.isInteger(a._rid)) to.set(a._rid, i + 1); });
    const missing = [];
    articles(after).forEach((a, i) => {
      eachText(a, (t) => {
        const r = rewriteText(t, from, to);
        r.missing.forEach((ref) => missing.push({ at: `第${i + 1}条`, ref }));
        return r.text;
      });
    });
    return missing;
  }

  // 参照先の無い参照（印付き）の一覧 [{ at: "第５条", ref: "第３条" }]。直すまで画面で知らせ続けるために使う
  function deadRefs(tree) {
    const out = [];
    const re = new RegExp(`(第(?:${NUM})条(?:の(?:${NUM}))?)${DEAD}`, "g");
    articles(tree).forEach((a, i) => {
      eachText(a, (t) => {
        for (const m of t.matchAll(re)) out.push({ at: `第${i + 1}条`, ref: m[1] });
        return t;
      });
    });
    return out;
  }

  const api = { assignIds, rewriteTree, rewriteText, deadRefs, numbering, parseNum, formatNum, DEAD };
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.KisokuRefs = api;
})(typeof self !== "undefined" ? self : this);
