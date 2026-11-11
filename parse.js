// リプ欄の読み取り結果（行ごと）から、名前と@IDを取り出す（ページとテストの両方で使う）
(function (root) {
  "use strict";
  const CJK = "\\u3000-\\u30ff\\u3400-\\u9fff\\uf900-\\ufaff\\uff00-\\uffef";
  const SPACE_IN_CJK = new RegExp("([" + CJK + "])\\s+(?=[" + CJK + "])", "g");
  // @IDのあとに続く「・17分」などの時間。読み取りで「-17%」「・.16 分」「.12」「 16 分」「1759」のように崩れるので、区切りはゆるく見る
  const isTimeRest = (r) => /^[・·\-–—.:\s]*\d/.test(r);
  const isTruncRest = (r) => /^(…|\.\.|--|-\+|—|–|-$|\.$)/.test(r);
  // 読み取りで出やすい文字のゆれをそろえる：丸数字（①→1）・全角→半角、@に似た記号→@
  const normText = (s) => String(s || "").normalize("NFKC").replace(/[④®©＠]/g, "@");
  function cleanName(s) {
    s = String(s || "").replace(SPACE_IN_CJK, "$1").replace(/\s+/g, " ").trim();
    // 前後のごみ（読み取りで出やすい記号）を取る
    s = s.replace(/^[\s~〜()（）\[\]{}|_.,、。'"`*+\-=?!:;<>^—–・·]+/, "").replace(/[\s~〜()（）\[\]{}|_,、'"`*+\-=?:;<>^—–・·]+$/, "");
    return s.slice(0, 40);
  }
  // line: { text, words: [{ text, confidence }] }
  function parseLine(line) {
    const raw = String(line.text || "").replace(/[\r\n]+/g, " ");
    const t = normText(raw).trim();
    if (/^返信先/.test(t.replace(/\s/g, ""))) return null;
    let id = "", pre = "", rest = "";
    const m = t.match(/@\s?([A-Za-z0-9_]{2,15})/);
    if (m) { id = m[1]; pre = t.slice(0, m.index); rest = t.slice(m.index + m[0].length); }
    else {
      // 「@」が読めなかったとき（「ok ok11053.7 分」「kyo osaiakuuue・11 分」など）：英数字のあとに時間が続くものを@IDとみなす
      const m2 = t.match(/(^|\s)([A-Za-z0-9_]{3,15})([・·\-–—.:\s]*\d.*)$/);
      if (!m2) return null;
      id = m2[2]; pre = t.slice(0, m2.index + m2[1].length); rest = m2[3];
    }
    rest = rest.trim();
    const time = isTimeRest(rest), trunc = !time && isTruncRest(rest);
    if (!time && !trunc) return null;                       // 本文の中の@（名前の行ではない）
    // 名前：@IDより前の単語のうち、読み取りの自信があるものだけ。@IDの単語の自信も見ておく
    let name = "", conf = 0, n = 0, idConf = 100;
    const words = Array.isArray(line.words) ? line.words : [];
    if (words.length) {
      const parts = []; let seenId = false;
      for (const w of words) {
        const wt = normText(w.text);
        if (!seenId && (wt.indexOf("@") >= 0 || wt.indexOf(id) >= 0)) { seenId = true; idConf = Number(w.confidence) || 0; continue; }
        if (seenId) continue;
        n++; conf += Number(w.confidence) || 0;
        // 記号だけの単語（絵文字の読みまちがいなど）は入れない
        if ((Number(w.confidence) || 0) >= 45 && /[0-9A-Za-z぀-ヿ㐀-鿿０-ｚ]/.test(wt)) parts.push(wt);
      }
      name = cleanName(parts.join(" "));
      if (name && /…\s*$/.test(pre) && !/…$/.test(name)) name += "…";
      name = name.replace(/\s+…$/, "…");
      conf = n ? conf / n : 0;
    } else { name = cleanName(pre); conf = 60; }
    if (!cleanName(pre)) return null;                       // 名前のない行（本文の先頭が@など）
    // IDが途中で切れているときも、確かめてもらう（発表の文に入るため）。@IDの読み取りの自信がとても低いときも
    return { name, id, trunc, idConf, sure: conf >= 70 && !!name && !trunc && idConf >= 20 };
  }
  // 行の一覧 → 参加者の一覧（同じ人は1人にまとめる・はぶくIDは入れない）
  function collect(lines, list, skip) {
    const skipSet = new Set((skip || []).map((s) => String(s).replace(/^@/, "").toLowerCase()).filter(Boolean));
    const out = list ? list.slice() : [];
    lines.forEach((l) => {
      const e = parseLine(l);
      if (!e || skipSet.has(e.id.toLowerCase())) return;
      out.push(e);
    });
    return dedupe(out);
  }
  const keyOf = (e) => (e.trunc ? "n:" + e.name : "i:" + e.id.toLowerCase());
  // 2つの文字のちがい（1文字ちがいなら、読みまちがいの同じ人とみなす）
  function dist(a, b) {
    a = a.toLowerCase(); b = b.toLowerCase();
    if (Math.abs(a.length - b.length) > 1) return 9;
    const d = []; for (let i = 0; i <= a.length; i++) { d[i] = [i]; }
    for (let j = 1; j <= b.length; j++) d[0][j] = j;
    for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return d[a.length][b.length];
  }
  // 名前の比べ方：「ー」「一」「-」、空白、末尾の「…」のちがいは無視。どちらかが空・片方がもう片方を含む・1文字ちがいも「同じ」とみなす
  const normName = (s) => cleanName(s).replace(/…$/, "").replace(/[ー一\-－–—]/g, "ー").replace(/\s+/g, "").toLowerCase();
  function sameName(x, y) {
    const a = normName(x.name), b = normName(y.name);
    if (!a || !b || a === b) return true;
    if (a.indexOf(b) >= 0 || b.indexOf(a) >= 0) return true;
    return dist(a, b) <= 1;
  }
  function dedupe(list) {
    const map = new Map();
    list.forEach((e) => {
      const k = keyOf(e), old = map.get(k);
      if (!old) { map.set(k, Object.assign({ seen: 1 }, e)); return; }
      old.seen = (old.seen || 1) + 1;
      // 名前は、読み取りの自信がある方・長い方を残す
      if ((!old.sure && e.sure) || (!old.name && e.name)) old.name = e.name, old.sure = e.sure;
      if (e.idConf !== undefined && (old.idConf === undefined || e.idConf > old.idConf)) old.idConf = e.idConf;
      if (e.manual) old.manual = true;
    });
    // 途中で切れた@IDは、最後まで読めた同じ人にまとめる
    const all = Array.from(map.values()), full = all.filter((e) => !e.trunc);
    const kept = all.filter((e) => !(e.trunc && full.some((f) => f.id.toLowerCase().indexOf(e.id.toLowerCase()) === 0)));
    // @IDが1文字だけちがい、名前も同じ（または似ている・どちらかの読み取りに自信がない）ものは、読みまちがいとして1人にまとめる
    // （何枚ものスクショで多く読めたほう・読み取りの自信が高いほうの@IDを残し、確かめてもらう印をつける）
    const res = [];
    kept.forEach((e) => {
      const i = e.manual ? -1 : res.findIndex((r) => !r.manual && !r.trunc && !e.trunc && r.id.length >= 4 && dist(r.id, e.id) === 1 && (sameName(r, e) || !r.sure || !e.sure));
      if (i < 0) { res.push(e); return; }
      const r = res[i];
      // 「@」が「o」「0」「Q」などに読まれて@IDの頭にくっついたもの（osaiakuuue と saiakuuue）は、短いほうが正しい
      const glued = (a, b) => a.length === b.length + 1 && /^[oO0Qq]/.test(a) && a.slice(1).toLowerCase() === b.toLowerCase();
      const better = glued(r.id, e.id) ? true : glued(e.id, r.id) ? false : (e.seen || 1) > (r.seen || 1) || ((e.seen || 1) === (r.seen || 1) && (e.idConf || 0) > (r.idConf || 0));
      const keep = better ? e : r, drop = better ? r : e;
      keep.alt = (keep.alt || []).concat(drop.alt || [], drop.id).filter((x, j, a) => x.toLowerCase() !== keep.id.toLowerCase() && a.indexOf(x) === j);
      // 名前は、読み取りの自信があるほう・長いほうを残す
      if (drop.name && (!keep.name || (drop.sure && !keep.sure) || (drop.sure === keep.sure && normName(drop.name).length > normName(keep.name).length))) keep.name = drop.name;
      if (drop.manual) keep.manual = true;
      if (drop.on !== undefined && keep.on === undefined) keep.on = drop.on;
      keep.sure = false;
      res[i] = keep;
    });
    return res;
  }
  const api = { parseLine, collect, dedupe, cleanName, keyOf, normText };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.RR_PARSE = api;
})(typeof window !== "undefined" ? window : globalThis);
