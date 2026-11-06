// 就業規則 整形アプリ（ブラウザ版）— 処理担当の Web Worker
// Python（Pyodide）をここで動かすので、重い処理中も画面は固まらない。
// 就業規則ファイルはこの端末のブラウザ内だけで処理され、外部に送られない。
// （Gemini を有効にした場合のみ、判定対象の文章が Google に送られる）

const PYODIDE_VERSION = "0.29.5";
const PYODIDE_BASE = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`;

// このフォルダ構成と一致させること（ファイルを増やしたらここにも追加）
const PY_FILES = [
  "web_entry.py",
  "kisoku_parser.py",
  "apply_style.py",
  "custom_template.py",
  "settings_reader.py",
  "marker_hierarchy.py",
  "hyoki_check.py",
  "crossref.py",
  "verify_report.py",
  "gemini_helper.py",
  "comparison.py",
];
const TEMPLATE_FILES = [
  "就業規則テンプレ1.docx",
  "就業規則テンプレ2.docx",
  "就業規則テンプレ3.docx",
  "就業規則テンプレ4.docx",
];
const DOCX_WHEEL = "wheels/python_docx-1.2.0-py3-none-any.whl";

importScripts(PYODIDE_BASE + "pyodide.js");

let py = null;
let entry = null;
let runNo = 0;
let current = null; // { srcPath, jsonPath }

const post = (msg, transfer) => self.postMessage(msg, transfer || []);
const status = (text) => post({ type: "status", text });

async function fetchOk(path, as) {
  // no-cache: 更新したファイルが古いキャッシュのまま使われるのを防ぐ
  const res = await fetch(path, { cache: "no-cache" });
  if (!res.ok) throw new Error(`${path} を読み込めません（HTTP ${res.status}）`);
  return as === "text" ? res.text() : new Uint8Array(await res.arrayBuffer());
}

async function init() {
  status("（1/3）");
  py = await loadPyodide({ indexURL: PYODIDE_BASE });
  py.setStdout({ batched: (s) => post({ type: "log", text: s }) });
  py.setStderr({ batched: (s) => post({ type: "log", text: s }) });

  status("（2/3）");
  await py.loadPackage(["lxml", "typing-extensions"]);
  py.FS.mkdirTree("/app/lib");
  py.FS.mkdirTree("/app/templates");
  py.FS.mkdirTree("/work");
  py.unpackArchive(await fetchOk(DOCX_WHEEL), "zip", { extractDir: "/app/lib" });

  status("（3/3）");
  await Promise.all([
    ...PY_FILES.map(async (f) =>
      py.FS.writeFile(`/app/${f}`, await fetchOk(`py/${f}`, "text"))),
    ...TEMPLATE_FILES.map(async (f) =>
      py.FS.writeFile(`/app/templates/${f}`,
        await fetchOk(`templates/${encodeURIComponent(f)}`))),
  ]);

  py.runPython(`
import sys
for p in ("/app", "/app/lib"):
    if p not in sys.path:
        sys.path.insert(0, p)
`);
  entry = py.pyimport("web_entry");
  post({ type: "ready" });
}

function safeName(name) {
  return name.replace(/[\\/:*?"<>|]/g, "_").trim() || "input.docx";
}

function readOut(path) {
  const bytes = py.FS.readFile(path);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function handleParse(msg) {
  runNo += 1;
  const dir = `/work/run_${runNo}`;
  py.FS.mkdirTree(dir);
  const srcPath = `${dir}/${safeName(msg.name)}`;
  py.FS.writeFile(srcPath, new Uint8Array(msg.bytes));

  const res = JSON.parse(entry.parse(srcPath, !!msg.useGemini,
    !!msg.layoutAsFigure, msg.apiKey || "", msg.geminiMode || "paragraph"));
  // jsonPath: 現行（比べる元）の木。parsedPath: この解析の結果。editedPath: 画面で編集した木
  current = { srcPath, jsonPath: res.json_path, parsedPath: res.json_path, editedPath: null };
  post({ type: "parsed", id: msg.id, chapters: res.chapters, articles: res.articles });
}

async function handleFormat(msg) {
  if (!current) throw new Error("先に docx を読み込んでください");
  const res = JSON.parse(entry.format_docx(current.srcPath, current.jsonPath,
    msg.templateKey, !!msg.checkHyoki, !!msg.insertToc,
    !!msg.keepCrossref, msg.custom || "", current.editedPath || ""));

  const docx = readOut(res.output_path);
  const out = {
    type: "formatted", id: msg.id,
    docxName: res.output_path.split("/").pop(),
    docx,
  };
  const transfer = [docx];
  if (res.report_path) {
    out.reportName = res.report_path.split("/").pop();
    out.reportText = new TextDecoder().decode(py.FS.readFile(res.report_path));
  }
  post(out, transfer);
}

async function handleCompare(msg) {
  if (!current) throw new Error("先に docx を読み込んでください");
  // 改正案の記号は、③で選んだテンプレートの書き方で表示する
  const v = JSON.parse(entry.compare_rows(current.jsonPath, current.editedPath || "",
    msg.templateKey || "", msg.custom || ""));
  post({ type: "compared", id: msg.id, rows: v.rows, blocks: v.blocks, marks: v.marks });
}

async function handleComparisonDocx(msg) {
  if (!current) throw new Error("先に docx を読み込んでください");
  const path = entry.make_comparison(current.srcPath, current.jsonPath, msg.notes || "{}",
    current.editedPath || "", msg.templateKey || "", msg.custom || "",
    JSON.stringify(msg.labels || []));
  const docx = readOut(path);
  post({ type: "comparisonDocx", id: msg.id, docxName: path.split("/").pop(), docx }, [docx]);
}

// 編集を始めるときの木（現行）を返す。parsed なら、比較の表示中でも解析した木そのもの
async function handleGetTree(msg) {
  if (!current) throw new Error("先に docx を読み込んでください");
  post({ type: "tree", id: msg.id, tree: entry.read_tree(msg.parsed ? current.parsedPath : current.jsonPath) });
}

// 比べる相手の docx を解析して木だけ返す（今の作業の解析結果は変えない）
async function handleParseOther(msg) {
  runNo += 1;
  const dir = `/work/other_${runNo}`;
  py.FS.mkdirTree(dir);
  const srcPath = `${dir}/${safeName(msg.name)}`;
  py.FS.writeFile(srcPath, new Uint8Array(msg.bytes));
  try {
    const res = JSON.parse(entry.parse(srcPath, false, !!msg.layoutAsFigure, "", "paragraph"));
    post({ type: "otherTree", id: msg.id, tree: entry.read_tree(res.json_path),
      chapters: res.chapters, articles: res.articles });
  } finally {
    // 比べるたびにメモリ上のファイルが増えないよう、使い終わったら消す
    // （消せなくても比較はできるので、失敗は無視する）
    try {
      for (const f of py.FS.readdir(dir)) if (f !== "." && f !== "..") py.FS.unlink(`${dir}/${f}`);
      py.FS.rmdir(dir);
    } catch {}
  }
}

// 画面の編集内容を受け取る。base: 下書きの現行の木（null ならこの解析の結果）、edited: 編集した木（null なら編集なし）
async function handleUseTrees(msg) {
  if (!current) throw new Error("先に docx を読み込んでください");
  const dir = current.srcPath.slice(0, current.srcPath.lastIndexOf("/"));
  current.jsonPath = msg.base ? entry.write_tree(`${dir}/_base.json`, msg.base) : current.parsedPath;
  current.editedPath = msg.edited ? entry.write_tree(`${dir}/_edited.json`, msg.edited) : null;
  post({ type: "treesUsed", id: msg.id });
}

async function handleSample(msg) {
  runNo += 1;
  const path = entry.make_sample(msg.custom || "", msg.tree || "[]", `/work/sample_${runNo}`);
  const docx = readOut(path);
  post({ type: "sample", id: msg.id, docxName: path.split("/").pop(), docx }, [docx]);
}

async function handleReadCustom(msg) {
  runNo += 1;
  const path = `/work/import_${runNo}.docx`;
  py.FS.writeFile(path, new Uint8Array(msg.bytes));
  try {
    post({ type: "customRead", id: msg.id, settings: entry.read_custom(path) });
  } finally {
    py.FS.unlink(path);
  }
}

const ready = init().catch((e) => {
  post({ type: "fatal", text: String(e && e.message ? e.message : e) });
  throw e;
});

self.onmessage = async (ev) => {
  const msg = ev.data;
  try {
    await ready;
    if (msg.type === "parse") await handleParse(msg);
    else if (msg.type === "format") await handleFormat(msg);
    else if (msg.type === "sample") await handleSample(msg);
    else if (msg.type === "compare") await handleCompare(msg);
    else if (msg.type === "getTree") await handleGetTree(msg);
    else if (msg.type === "parseOther") await handleParseOther(msg);
    else if (msg.type === "useTrees") await handleUseTrees(msg);
    else if (msg.type === "comparisonDocx") await handleComparisonDocx(msg);
    else if (msg.type === "readCustom") await handleReadCustom(msg);
  } catch (e) {
    let text = String(e && e.message ? e.message : e);
    // Python の例外はトレースバック全体が来るので、最後の行を要約に使う
    const lines = text.trim().split("\n");
    post({ type: "error", id: msg.id, text: lines[lines.length - 1], detail: text });
  }
};
