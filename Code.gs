/**
 * ゲットチャレンジ記録（全員の記録を集める用）
 * リクエスト箱・配布ルーレットの Apps Script とは別のものです。
 * ゲットチャレンジ用の Apps Script（無題のプロジェクト）にこれをまるごと貼り付けてください。
 * 記録は、カードリクエスト箱のスプレッドシートの「ゲットチャレンジ記録」「ゲットチャレンジ見た目」シートに入ります
 * （シートはなければ自動で作られます。ほかのシートには書きこみません）。
 */

// 記録を入れるスプレッドシート（カードリクエスト箱と同じもの）
const SHEET_ID = "1i5JR5t0Szowl4zxlLopsNgRT1PayENAUH1HY-4JeXac";

const SHEET = "ゲットチャレンジ記録";
const HEAD = ["id", "日時", "持ち主", "かぎ", "パック", "シャッフル前の5枚", "欲しい（前の場所）", "GET（後の場所）", "シャッフル後（後の場所ごとの前の場所）", "当たり", "消した"];
const MAX_PUSH = 50;          // 1回に送れる記録の数
const STATS_CACHE_SEC = 600;  // 全員の統計を覚えておく時間（秒）
const LOOK_SHEET = "ゲットチャレンジ見た目";   // スクショ読み取り用の、カードの見た目（小さな色の並び）。画像そのものは保存しない
const LOOK_HEAD = ["パック", "カード名", "レア度", "見た目", "送った人", "日時"];
const LOOK_PER_CARD = 8;      // 1種類のカードにつき覚えておく数
const RARITY_OK = ["d1", "d2", "d3", "d4", "s1", "s2", "s3", "sh1", "sh2", "cr", "item"];

function doGet(e) {
  return json({ ok: true, app: "getchallenge" });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents || "{}"); } catch (err) { return json({ ok: false, error: "bad_json" }); }
  try {
    if (b.action === "gc_push") return json(push(b));
    if (b.action === "gc_mine") return json(mine(b));
    if (b.action === "gc_stats") return json({ ok: true, stats: stats(String(b.pack || "")) });
    if (b.action === "gc_looks_add") return json(looksAdd(b));
    if (b.action === "gc_looks") return json({ ok: true, looks: looks(String(b.pack || "")) });
    return json({ ok: false, error: "bad_action" });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function book() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// 最初に1回だけ：この関数を選んで「実行」すると、許可の確認が出て、シートが作られます
function setup() {
  sheet(); lookSheet();
  Logger.log("できました：" + book().getUrl());
}

function sheet() {
  const ss = book();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

// かぎはそのまま保存せず、ハッシュにして保存する
function hash(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "gc:" + String(s || ""), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(raw).slice(0, 22);
}

function cleanUid(u) {
  u = String(u || "");
  return /^(p:\d{8,20}|d:[\w-]{8,64})$/.test(u) ? u : "";
}

function validRecord(r) {
  if (!r || typeof r.id !== "string" || !/^[a-z0-9]{6,20}$/.test(r.id)) return false;
  if (!Array.isArray(r.cards) || r.cards.length !== 5) return false;
  if (!Array.isArray(r.want) || !r.want.length || r.want.length > 5) return false;
  if (!r.want.every((w) => Number.isInteger(w) && w >= 0 && w < 5)) return false;
  if (!Number.isInteger(r.get) || r.get < 0 || r.get > 4) return false;
  // after[後の場所] = 前の場所。0〜4がひとつずつ入っていること
  if (!Array.isArray(r.after) || r.after.length !== 5) return false;
  if ([0, 1, 2, 3, 4].some((x) => r.after.indexOf(x) < 0)) return false;
  return true;
}

function push(b) {
  const uid = cleanUid(b.uid);
  if (!uid) return { ok: false, error: "bad_uid" };
  const key = hash(b.key);
  const recs = (Array.isArray(b.records) ? b.records : []).slice(0, MAX_PUSH);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet();
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, HEAD.length).getValues() : [];
    const byId = {};
    rows.forEach((r, i) => { byId[r[0]] = i; });
    // 同じ持ち主がほかのかぎで送ってきたら受け付けない（なりすまし防止）
    const owned = rows.find((r) => r[2] === uid);
    if (owned && owned[3] !== key) return { ok: false, error: "bad_key" };
    const saved = [], add = [];
    let moved = false;
    recs.forEach((r) => {
      if (!validRecord(r)) return;
      const i = byId[r.id];
      if (i !== undefined) {
        if (rows[i][2] !== uid) {
          // ログイン前にこのスマホでつけた記録を、ログインした持ち主にまとめる（そのスマホのかぎが合うときだけ）
          const from = cleanUid(r.moveFrom);
          if (!from || rows[i][2] !== from || from.indexOf("d:") !== 0 || rows[i][3] !== hash(from.slice(2))) return;
          sh.getRange(i + 2, 3, 1, 2).setValues([[uid, key]]); rows[i][2] = uid; rows[i][3] = key; moved = true;
        }
        if (r.deleted && !rows[i][10]) sh.getRange(i + 2, 11).setValue(true);
        saved.push(r.id);
        return;
      }
      if (r.deleted) { saved.push(r.id); return; }
      const cards = r.cards.map((c) => c ? [Number(c.no) || 0, String(c.name || "").slice(0, 30), String(c.r || "").slice(0, 4)] : null);
      const hit = r.want.indexOf(r.after[r.get]) >= 0;
      add.push([r.id, new Date(Number(r.at) || Date.now()), uid, key, String(r.pack || "").slice(0, 10), JSON.stringify(cards), JSON.stringify(r.want), r.get, JSON.stringify(r.after), hit, false]);
      saved.push(r.id);
    });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, HEAD.length).setValues(add);
    if (add.length || moved || recs.some((r) => r && r.deleted)) CacheService.getScriptCache().removeAll(cacheKeys());
    return { ok: true, saved: saved };
  } finally {
    lock.releaseLock();
  }
}

function mine(b) {
  const uid = cleanUid(b.uid);
  if (!uid) return { ok: false, error: "bad_uid" };
  const key = hash(b.key);
  const sh = sheet(), last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, HEAD.length).getValues() : [];
  const list = rows.filter((r) => r[2] === uid && r[3] === key && !r[10]).map(toRecord);
  return { ok: true, records: list };
}

function toRecord(r) {
  const cards = JSON.parse(r[5] || "[]").map((c) => c ? { no: c[0], name: c[1], r: c[2] } : null);
  return { id: r[0], at: new Date(r[1]).getTime(), pack: r[4], cards: cards, want: JSON.parse(r[6] || "[]"), get: Number(r[7]), after: JSON.parse(r[8] || "[]") };
}

function cacheKeys() {
  const packs = ["_"];
  const sh = sheet(), last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 5, last - 1, 1).getValues().forEach((v) => { if (packs.indexOf(v[0]) < 0) packs.push(v[0]); });
  return packs.map((p) => "stats:" + p);
}

// 全員の統計（ページの「自分の記録」と同じ形で返す）
function stats(pack) {
  const cache = CacheService.getScriptCache();
  const ck = "stats:" + (pack || "_");
  const hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  const sh = sheet(), last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, HEAD.length).getValues() : [];
  const RG = { d1: "◆1〜2", d2: "◆1〜2", d3: "◆3", d4: "◆4", s1: "☆1", s2: "☆2", s3: "☆3以上", sh1: "色違い", sh2: "色違い", cr: "☆3以上" };
  const z = () => [0, 1, 2, 3, 4].map(() => [0, 0, 0, 0, 0]);
  const S = { n: 0, hit: 0, expect: 0, getN: [0, 0, 0, 0, 0], getHit: [0, 0, 0, 0, 0], wantAt: [0, 0, 0, 0, 0], wantN: 0,
    moveAll: z(), moveWant: z(), rar: {}, packs: {}, users: 0, updated: Date.now() };
  const users = {};
  rows.forEach((row) => {
    if (row[10]) return;
    if (pack && row[4] !== pack) return;
    let r; try { r = toRecord(row); } catch (e) { return; }
    const A = r.after;
    if (!Array.isArray(A) || A.length !== 5) return;
    users[row[2]] = 1;
    const got = A[r.get], h = r.want.indexOf(got) >= 0;
    S.n++; if (h) S.hit++;
    S.expect += r.want.length / 5;
    S.getN[r.get]++; if (h) S.getHit[r.get]++;
    for (let q = 0; q < 5; q++) { const d = A.indexOf(q); if (d >= 0) S.moveAll[q][d]++; }
    r.want.forEach((w) => {
      const d = A.indexOf(w); if (d < 0) return;
      S.wantAt[d]++; S.wantN++; S.moveWant[w][d]++;
      const g = RG[(r.cards[w] || {}).r] || "?"; S.rar[g] = S.rar[g] || [0, 0]; S.rar[g][0]++; if (got === w) S.rar[g][1]++;
    });
    S.packs[r.pack] = S.packs[r.pack] || [0, 0]; S.packs[r.pack][0]++; if (h) S.packs[r.pack][1]++;
  });
  S.users = Object.keys(users).length;
  try { cache.put(ck, JSON.stringify(S), STATS_CACHE_SEC); } catch (e) {}
  return S;
}

// ---- カードの見た目（みんなで覚える） ----
function lookSheet() {
  const ss = book();
  let sh = ss.getSheetByName(LOOK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOOK_SHEET);
    sh.getRange(1, 1, 1, LOOK_HEAD.length).setValues([LOOK_HEAD]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function looksAdd(b) {
  const uid = cleanUid(b.uid);
  if (!uid) return { ok: false, error: "bad_uid" };
  const list = (Array.isArray(b.looks) ? b.looks : []).slice(0, 20).filter((x) =>
    x && /^[A-Za-z0-9-]{1,10}$/.test(String(x.pack || "")) && typeof x.name === "string" && x.name.length > 0 && x.name.length <= 30 &&
    RARITY_OK.indexOf(x.r) >= 0 && typeof x.t === "string" && /^[A-Za-z0-9+/=]{380,390}$/.test(x.t));
  if (!list.length) return { ok: true, saved: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = lookSheet(), last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, LOOK_HEAD.length).getValues() : [];
    const add = [], packs = {};
    list.forEach((x) => {
      const same = rows.map((r, i) => ({ r, i })).filter((o) => o.r[0] === x.pack && o.r[1] === x.name && o.r[2] === x.r);
      const mine = same.find((o) => o.r[4] === uid);
      if (mine) { sh.getRange(mine.i + 2, 4, 1, 3).setValues([[x.t, uid, new Date()]]); }       // 同じ人は1つだけ（新しいものにする）
      else if (same.length < LOOK_PER_CARD) { add.push([x.pack, x.name, x.r, x.t, uid, new Date()]); rows.push([x.pack, x.name, x.r, x.t, uid]); }
      packs[x.pack] = 1;
    });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, LOOK_HEAD.length).setValues(add);
    CacheService.getScriptCache().removeAll(Object.keys(packs).map((p) => "looks:" + p));
    return { ok: true, saved: list.length };
  } finally {
    lock.releaseLock();
  }
}

function looks(pack) {
  if (!pack) return {};
  const cache = CacheService.getScriptCache(), ck = "looks:" + pack, hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  const sh = lookSheet(), last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
  const out = {};
  rows.forEach((r) => { if (r[0] !== pack) return; const k = r[1] + "|" + r[2]; (out[k] = out[k] || []).push(r[3]); });
  try { cache.put(ck, JSON.stringify(out), 600); } catch (e) {}
  return out;
}
