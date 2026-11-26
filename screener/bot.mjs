// Серверний бот скрінера: запускає той самий код зі screener/index.html у Node,
// робить скан, дописує сетапи з «Сетапів» у спільний журнал і дораховує результати.
// Використання: node screener/bot.mjs <шлях-до-journal.json> [<шлях-до-bot-status.json>]
import fs from "node:fs";
import vm from "node:vm";
import path from "node:path";
import {fileURLToPath} from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const MERGE = process.argv[2] === "--merge";   // node bot.mjs --merge ours.json theirs.json
const JPATH = (MERGE ? process.argv[3] : process.argv[2]) || path.join(here, "journal.json");
const SPATH = process.argv[3] || path.join(path.dirname(JPATH), "bot-status.json");
const html = fs.readFileSync(path.join(here, "index.html"), "utf8");
const code = html.match(/<script>([\s\S]*)<\/script>/)[1];

// налаштування бота: стандартні + те, що лежить у screener/bot-config.json (якщо є)
let botCfg = {};
try { botCfg = JSON.parse(fs.readFileSync(path.join(here, "bot-config.json"), "utf8")); } catch (e) {}

// попередній журнал
let prev = {cleared: 0, trades: []};
try {
  const o = JSON.parse(fs.readFileSync(JPATH, "utf8"));
  prev = Array.isArray(o) ? {cleared: 0, trades: o} : {cleared: +o.cleared || 0, trades: o.trades || []};
} catch (e) {}

// ---- мінімальна «сторінка»
const els = {};
const mkEl = () => ({
  style: {}, dataset: {}, value: "", textContent: "", innerHTML: "", hidden: false, disabled: false,
  classList: {toggle() {}, add() {}, remove() {}, contains() { return false; }},
  addEventListener() {}, querySelector() { return mkEl(); }, querySelectorAll() { return []; },
  focus() {}, click() {}, scrollIntoView() {}, setSelectionRange() {}, appendChild() {}, remove() {},
  setAttribute() {}, getAttribute() { return null; }, closest() { return null; }, select() {},
});
const store = new Map();
const localStorage = {
  getItem: k => (store.has(k) ? store.get(k) : null),
  setItem: (k, v) => store.set(k, String(v)),
  removeItem: k => store.delete(k),
};
store.set("smc-cfg", JSON.stringify(Object.assign({loop: 0, sound: "0", vibro: "0", notify: "0", lang: "uk"}, botCfg, {ghtoken: ""})));
store.set("smc-journal", JSON.stringify(prev.trades));
store.set("smc-journal-cleared", String(prev.cleared));

// GitHub (API і raw) з бота не чіпаємо: журнал пишемо через git у workflow
const realFetch = globalThis.fetch;
const fetchW = (url, opt) => (/api\.github\.com|raw\.githubusercontent\.com/.test(String(url))
  ? Promise.resolve(new Response("{}", {status: 503}))
  : realFetch(url, opt));

const document = {
  getElementById: id => (els[id] ||= mkEl()),
  querySelector: () => mkEl(), querySelectorAll: () => [],
  addEventListener() {}, createElement: () => mkEl(), createTreeWalker: () => ({nextNode: () => null}),
  documentElement: {lang: "uk"}, body: mkEl(), title: "", execCommand: () => false,
};
const ctx = {
  document, localStorage, fetch: fetchW, AbortController, Response, Blob, URL,
  console, setTimeout, clearTimeout, setInterval: () => 0, clearInterval() {},
  navigator: {}, btoa, atob,
  NodeFilter: {SHOW_TEXT: 4}, location: {reload() {}}, isSecureContext: true, addEventListener() {},
};
ctx.window = ctx; ctx.self = ctx; ctx.top = MERGE ? {} : ctx; ctx.BOT = true;
vm.createContext(ctx);
// код сторінки сам запускає scan() при завантаженні (window.self===window.top)
vm.runInContext(code, ctx, {filename: "index.html"});
const ev = s => vm.runInContext(s, ctx);
if (MERGE) {
  let other = {cleared: 0, trades: []};
  try { const o = JSON.parse(fs.readFileSync(process.argv[4], "utf8")); other = Array.isArray(o) ? {cleared: 0, trades: o} : {cleared: +o.cleared || 0, trades: o.trades || []}; } catch (e) {}
  const cl = Math.max(prev.cleared, other.cleared);
  const J = ev("jMerge")(prev.trades, other.trades, cl);
  fs.writeFileSync(JPATH, ev("jText")(cl, J));
  console.log(`merged: ${J.length} trades`);
  process.exit(0);
}

const t0 = Date.now();
const sleep = ms => new Promise(r => setTimeout(r, ms));
while (Date.now() - t0 < 12 * 60000) {
  await sleep(1000);
  if (!ev("busy") && (ev("last") || /Помилка|Error/.test(els.status?.textContent || ""))) break;
}
await sleep(500);
await ev("updJournal()");

const J = ev("J"), cl = ev("jCleared");
const statusText = els.status?.textContent || "";
const last = ev("last");
const shown = last ? last.sigs.filter(s => s.state === "ENTER").length : 0;
fs.mkdirSync(path.dirname(JPATH), {recursive: true});
fs.writeFileSync(JPATH, ev("jText")(cl, J));
fs.writeFileSync(SPATH, JSON.stringify({
  at: new Date().toISOString(), status: statusText, ok: !!last,
  exchange: (statusText.split(" · ")[0] || ""), signals: last ? last.sigs.length : 0, enterNow: shown,
  diag: ev("DIAG"),
  trades: J.length, open: J.filter(j => j.st === "wait" || j.st === "open").length,
}, null, 1));
console.log(statusText);
console.log(`journal: ${J.length} trades (було ${prev.trades.length})`);
if (!last) process.exitCode = 2;
