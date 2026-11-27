// Серверний бот скрінера: запускає той самий код зі screener/index.html у Node,
// робить скан, дописує сетапи з «Сетапів» у спільний журнал і дораховує результати.
// Використання: node screener/bot.mjs <шлях-до-journal.json> [<шлях-до-bot-status.json>]
import fs from "node:fs";
import vm from "node:vm";
import path from "node:path";
import {fileURLToPath} from "node:url";
import crypto from "node:crypto";
import {loadKey, openSub, send} from "./push.mjs";

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

// ---- push-сповіщення на телефони (підписки: screener/subs.json у гілці screener-data, зашифровані)
let prevStatus = {};
try { prevStatus = JSON.parse(fs.readFileSync(SPATH, "utf8")); } catch (e) {}
let push = null;
try { push = await pushAll(prevStatus.push || {}); } catch (e) { push = {err: String(e && e.message || e)}; }

async function pushAll(pp) {
  const now = Date.now();
  const info = {at: new Date(now).toISOString(), subs: 0, sent: 0, fail: 0, err: "",
    notified: (pp.notified || []).slice(-400), tested: Object.assign({}, pp.tested), dead: Object.assign({}, pp.dead),
    seen: Object.assign({}, pp.seen), last: pp.last || null};
  const phase = ev("phase"), dupOf = ev("dupOf"), winEnd = ev("winEnd"), nice = ev("nice"), PX = ev("PX"), rrAt = ev("rrAt");
  // нові сетапи, які ще можна торгувати (ліміт ще стоїть або по ринку ще RR ≥ 2), не дублікати
  const done = new Set(info.notified);
  const fresh = J.filter(j => !done.has(j.id) && j.t0 > cl && now - j.t0 < 45 * 60000 && !dupOf(j, J) && ["go", "mkt"].includes(phase(j, now)));
  fresh.forEach(j => info.notified.push(j.id));
  const KEY = process.env.SCREENER_PUSH_KEY || "";
  if (!KEY) { info.err = "no SCREENER_PUSH_KEY secret"; return info; }
  const key = loadKey(KEY);
  let subs = [];
  try { subs = JSON.parse(fs.readFileSync(path.join(path.dirname(JPATH), "subs.json"), "utf8")).subs || []; } catch (e) {}
  info.subs = subs.length;
  const ids = new Set(subs.map(x => x.id));
  for (const m of [info.tested, info.dead, info.seen]) for (const k of Object.keys(m)) if (!ids.has(k)) delete m[k];
  const hh = t => { try { return new Intl.DateTimeFormat("uk-UA", {timeZone: "Europe/Kyiv", hour: "2-digit", minute: "2-digit", hour12: false}).format(new Date(t)); }
    catch (e) { const d = new Date(t + 2 * 3600e3); return String(d.getUTCHours()).padStart(2, "0") + ":" + String(d.getUTCMinutes()).padStart(2, "0"); } };
  const topic = id => crypto.createHash("sha256").update(id).digest("base64url").slice(0, 22);
  const one = (j, en) => {
    const F = v => PX(v, j.dec), end = winEnd(j), ph = phase(j, now), em = j.dir === "LONG" ? "🟢" : "🔴";
    const title = `${em} ${nice(j.symbol)} ${j.dir} · ${en ? "enter now" : "вхід зараз"}`;
    let body = `${en ? "Entry" : "Вхід"} ${F(j.entry)} · SL ${F(j.sl)} · TP ${F(j.tp)} · RR ${(+j.rr).toFixed(1)}\n`;
    body += ph === "mkt"
      ? (en ? `Limit already filled: market entry still OK until ${hh(end)} (RR ${rrAt(j, j.lp).toFixed(1)} from price)` : `Ліміт уже налився: по ринку ще можна до ${hh(end)} (RR ${rrAt(j, j.lp).toFixed(1)} від ціни)`)
      : (en ? `Valid until ${hh(end)}` : `Діє до ${hh(end)}`);
    body += ` · ${en ? "grade" : "клас"} ${j.tier}${j.src ? " · " + j.src : ""}`;
    return {title, body, tag: j.id, url: "./", ttl: Math.max(120, Math.round((end - now) / 1000))};
  };
  for (const rec of subs) {
    if (info.dead[rec.id] && info.dead[rec.id] > (rec.at || 0)) continue;   // ще не перепідписався
    let o;
    try { o = openSub(key, rec); } catch (e) { info.fail++; continue; }
    const en = o.lang === "en", pr = o.prefs || {};
    const msgs = [];
    if (rec.test && rec.test > (info.tested[rec.id] || 0)) {
      info.tested[rec.id] = rec.test;
      msgs.push({title: en ? "✅ Test from the bot" : "✅ Тест від бота",
        body: en ? "Notifications reach this device. New setups will arrive like this." : "Сповіщення на цей пристрій доходять. Нові сетапи приходитимуть так само.",
        tag: "test", url: "./", ttl: 3600});
    }
    const mine = fresh.filter(j => (pr.dir === "LONG" || pr.dir === "SHORT" ? j.dir === pr.dir : true) && (pr.tier === "A" ? j.tier === "A" : true));
    if (mine.length > 3) {
      const end = Math.min(...mine.map(j => winEnd(j)));
      msgs.push({title: en ? `🟢 ${mine.length} new setups` : `🟢 Нових сетапів: ${mine.length}`,
        body: mine.map(j => `${nice(j.symbol)} ${j.dir} · RR ${(+j.rr).toFixed(1)} · ${j.tier}`).join("\n"),
        tag: "many-" + now, url: "./", ttl: Math.max(120, Math.round((end - now) / 1000))});
    } else mine.forEach(j => msgs.push(one(j, en)));
    for (const m of msgs) {
      let st = 0;
      try { st = await send(key, o.sub, m, {ttl: m.ttl, topic: topic(m.tag), fetchFn: realFetch}); } catch (e) { st = 0; }
      if (st >= 200 && st < 300) { info.sent++; info.seen[rec.id] = now; info.last = now; }
      else { info.fail++; if (st === 404 || st === 410) { info.dead[rec.id] = now; break; } }
    }
  }
  return info;
}
const statusText = els.status?.textContent || "";
const last = ev("last");
const shown = last ? last.sigs.filter(s => s.state === "ENTER").length : 0;
fs.mkdirSync(path.dirname(JPATH), {recursive: true});
fs.writeFileSync(JPATH, ev("jText")(cl, J));
fs.writeFileSync(SPATH, JSON.stringify({
  at: new Date().toISOString(), status: statusText, ok: !!last,
  exchange: (statusText.split(" · ")[0] || ""), signals: last ? last.sigs.length : 0, enterNow: shown,
  diag: ev("DIAG"), push,
  trades: J.length, open: J.filter(j => j.st === "wait" || j.st === "open").length,
}, null, 1));
console.log(statusText);
console.log(`journal: ${J.length} trades (було ${prev.trades.length})`);
if (push) console.log(`push: subs ${push.subs || 0}, sent ${push.sent || 0}, fail ${push.fail || 0}${push.err ? ", " + push.err : ""}`);
if (!last) process.exitCode = 2;
