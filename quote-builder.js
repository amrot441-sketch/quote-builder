/* Dragon Electronics quotation builder \u2013 core (no DOM).
 * Fills Yousouf's "Quotation Template 2026 05 13" (version of 23 Nov 2026, no Project line) (sheets "With Disc Column" + "NO Disc") by editing the workbook XML,
 * so every style, image and print setting is kept. Output: one sheet named "Quotation".
 *  - Terms, letterhead and signature come from the "NO Disc" sheet (it has the complete, highlighted warranty lines).
 *  - With a discount, the column headings and price cells follow the "With Disc Column" sheet (Normal + Discounted price).
 *  - Unit prices are entered INCLUDING VAT (website price); only "Total Amount with VAT" is kept.
 * Runs in Chrome and in Node 18+ (CompressionStream / DecompressionStream "deflate-raw"). */
(function (root) {
  "use strict";

  // ------------------------------------------------------------------ zip
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(buf) { let c = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  async function streamBytes(bytes, stream) {
    const out = new Response(new Blob([bytes]).stream().pipeThrough(stream));
    return new Uint8Array(await out.arrayBuffer());
  }
  const inflate = b => streamBytes(b, new DecompressionStream("deflate-raw"));
  const deflate = b => streamBytes(b, new CompressionStream("deflate-raw"));
  const enc = new TextEncoder(), dec = new TextDecoder("utf-8");

  async function unzip(ab) {
    const u8 = new Uint8Array(ab), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error("this is not a valid Excel file");
    const count = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
    const files = new Map();
    for (let i = 0; i < count; i++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      p += 46 + nlen + xlen + clen;
      const start = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
      const raw = u8.subarray(start, start + csize);
      files.set(name, method === 0 ? raw.slice() : await inflate(raw));
    }
    return files;
  }

  async function zip(files) {
    const names = [...files.keys()].sort((a, b) => (a === "[Content_Types].xml" ? -1 : b === "[Content_Types].xml" ? 1 : 0));
    const parts = [], central = []; let offset = 0;
    for (const name of names) {
      const data = files.get(name), nameB = enc.encode(name), crc = crc32(data), comp = await deflate(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 8, true);
      lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true); lh.setUint32(14, crc, true);
      lh.setUint32(18, comp.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, nameB.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), nameB, comp);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, 8, true);
      ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true); ch.setUint32(16, crc, true); ch.setUint32(20, comp.length, true);
      ch.setUint32(24, data.length, true); ch.setUint16(28, nameB.length, true); ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), nameB);
      offset += 30 + nameB.length + comp.length;
    }
    const cdSize = central.reduce((s, b) => s + b.length, 0);
    const eocd = new DataView(new ArrayBuffer(22));
    eocd.setUint32(0, 0x06054b50, true); eocd.setUint16(8, names.length, true); eocd.setUint16(10, names.length, true);
    eocd.setUint32(12, cdSize, true); eocd.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(eocd.buffer)];
    const out = new Uint8Array(all.reduce((s, b) => s + b.length, 0)); let o = 0;
    for (const b of all) { out.set(b, o); o += b.length; }
    return out;
  }

  // ------------------------------------------------------------------ helpers
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const unesc = s => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
  const pad2 = n => String(n).padStart(2, "0");
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  function isoToDate(iso) { const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); }
  const range = (a, b) => { const r = []; for (let i = a; i <= b; i++) r.push(i); return r; };   // inclusive
  const colNum = c => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const VAT = 1.15;

  // Rows of "Quotation Template 2026 05 13" (both sheets share the row numbers)
  // ("NO Disc" sheet; the "With Disc Column" sheet is only used for its two heading rows)
  const R = {
    headings: [19, 20], header: 20, example: range(21, 36), blankBeforeTotals: 37, totals: [38, 39, 40], terms: 42,
    attention: 14, tel: 15, mobile: 16, email: 17,
    pay: { full: 45, days: 46 }, delivery: 47, transport: 48, validUntil: 49,
    warranty: range(52, 77), install: 83, stainless: 85, closing: 87, last: 97
  };
  const LANDMARKS = [["B9", /^To:/], ["I9", /^Date/], ["I10", /^Quotation No/], ["B14", /^Attention/], ["B15", /^Tel/], ["B19", /^Model/],
    ["B42", /^Terms & Conditions/], ["B45", /Payment/], ["B46", /Payment/], ["B47", /Delivery/], ["B48", /Delivery/], ["B49", /Validity/],
    ["B52", /Warranty/], ["B79", /Exchange/], ["B83", /Installation/], ["B85", /^Note/], ["I38", /Total Amount without VAT/]];
  const LANDMARKS_DISC = [["B9", /^To:/], ["B19", /^Model/], ["H19", /^Normal Unit/], ["I19", /^Discounted Unit/], ["I38", /Total Amount without VAT/]];

  function sheetPaths(files) {
    const wb = dec.decode(files.get("xl/workbook.xml")), rels = dec.decode(files.get("xl/_rels/workbook.xml.rels"));
    const out = {};
    for (const m of wb.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/>/g)) {
      const t = rels.match(new RegExp(`<Relationship [^>]*Id="${m[2]}"[^>]*Target="([^"]+)"`)) || rels.match(new RegExp(`<Relationship [^>]*Target="([^"]+)"[^>]*Id="${m[2]}"`));
      if (t) out[unesc(m[1])] = { rid: m[2], path: "xl/" + t[1].replace(/^\/?xl\//, "").replace(/^\.\//, ""), tag: m[0] };
    }
    return out;
  }

  // ------------------------------------------------------------------ template check
  function checkTemplate(files) {
    if (!files.get("xl/workbook.xml") || !files.get("xl/sharedStrings.xml")) return ["this is not an Excel quotation file"];
    const sh = sheetPaths(files);
    if (!sh["NO Disc"] || !sh["With Disc Column"]) return ["it needs the two sheets 'With Disc Column' and 'NO Disc'"];
    const si = dec.decode(files.get("xl/sharedStrings.xml")).match(/<si>[\s\S]*?<\/si>/g) || [];
    const problems = [];
    for (const [name, marks] of [["NO Disc", LANDMARKS], ["With Disc Column", LANDMARKS_DISC]]) {
      const s = dec.decode(files.get(sh[name].path) || new Uint8Array());
      for (const [ref, re] of marks) {
        const m = s.match(new RegExp(`<c r="${ref}"[^>]*t="s"[^>]*><v>(\\d+)</v></c>`));
        const txt = m ? unesc((si[+m[1]] || "").replace(/<[^>]+>/g, "")).trim() : "";
        if (!re.test(txt)) problems.push(`${name} ${ref} should start with "${re.source.replace(/[\^\\]/g, "")}" but has "${txt || "nothing"}"`);
      }
    }
    return problems;
  }

  // ------------------------------------------------------------------ build
  async function buildQuote(templateBytes, spec, images) {
    images = images || {};
    const files = await unzip(templateBytes);
    const bad = checkTemplate(files);
    if (bad.length) throw new Error("This template's layout is different from 'Quotation Template 2026 05 13' (" + bad.slice(0, 2).join("; ") + ").");
    const get = n => dec.decode(files.get(n));
    const put = (n, s) => files.set(n, enc.encode(s));
    const sheets = sheetPaths(files), base = sheets["NO Disc"], disc = sheets["With Disc Column"];
    const SHEET = base.path, SHEET_RELS = SHEET.replace(/worksheets\/(sheet\d+\.xml)$/, "worksheets/_rels/$1.rels");
    let s = get(SHEET), sst = get("xl/sharedStrings.xml");
    const sDisc = get(disc.path);
    const discount = spec.layout === "discount";
    const warnings = [];

    const siList = sst.match(/<si>[\s\S]*?<\/si>/g) || [];
    let nSi = siList.length; const newSi = [];
    const sidx = text => {
      text = String(text);
      const sp = text !== text.trim() || text.includes("  ") ? ' xml:space="preserve"' : "";
      newSi.push(`<si><t${sp}>${esc(text).replace(/&quot;/g, '"')}</t></si>`); return nSi++;
    };
    const rowOf = (xml, r) => { const m = xml.match(new RegExp(`<row r="${r}"[^>]*?(?:/>|>[\\s\\S]*?</row>)`)); return m ? m[0] : ""; };
    const cellText = (r, col) => {
      const m = rowOf(s, r).match(new RegExp(`<c r="${col}${r}"[^>]*t="s"[^>]*><v>(\\d+)</v></c>`));
      return m ? unesc(siList[+m[1]].replace(/<[^>]+>/g, "")) : null;
    };
    // replace / insert one cell inside a row's XML
    const setCellInRow = (rx, ref, inner, typeAttr, styleOverride) => {
      const re = new RegExp(`<c r="${ref}"((?:\\s+[a-z]+="[^"]*")*)\\s*(?:/>|>[\\s\\S]*?</c>)`);
      const m = rx.match(re);
      const style = styleOverride != null ? String(styleOverride) : m && (m[1].match(/\ss="(\d+)"/) || [])[1];
      const xml = inner ? `<c r="${ref}"${style ? ` s="${style}"` : ""}${typeAttr || ""}>${inner}</c>` : `<c r="${ref}"${style ? ` s="${style}"` : ""}/>`;
      if (m) return rx.replace(re, xml);
      const col = ref.match(/[A-Z]+/)[0];
      const body = rx.match(/^(<row [^>]*?)(\/>|>([\s\S]*?)<\/row>)$/);
      const cells = (body[3] || "").match(/<c r="[A-Z]+\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || [];
      const at = cells.findIndex(c => colNum(c.match(/<c r="([A-Z]+)/)[1]) > colNum(col));
      if (at < 0) cells.push(xml); else cells.splice(at, 0, xml);
      return `${body[1]}>${cells.join("")}</row>`;
    };

    // ---- what to keep / show / change (original row numbers)
    const t = spec.terms || {}, cust = spec.customer || {};
    const isCompany = !!cust.company;
    const DELETE = new Set(R.example), SHOW = new Set(), CLEAR_B = new Set();
    const EDIT = new Map();                                  // row -> [[col, text|null, opts]]
    const edit = (r, col, text, opts) => { if (!EDIT.has(r)) EDIT.set(r, []); EDIT.get(r).push([col, text, opts || {}]); };
    const showOrDel = (cond, rows) => rows.forEach(r => (cond ? SHOW : DELETE).add(r));

    // customer block
    let addr = cust.address || []; if (typeof addr === "string") addr = addr.split(/\n/).map(x => x.trim()).filter(Boolean);
    const to = isCompany ? String(cust.to || "").toUpperCase() : cust.to;
    edit(9, "C", to || null); edit(10, "C", addr[0] || null); edit(11, "C", addr[1] || null);
    edit(12, "C", addr[2] || null); edit(13, "C", addr.slice(3).join(", ") || null);
    if (isCompany) edit(R.attention, "C", cust.attention || null);
    else DELETE.add(R.attention);
    edit(R.tel, "C", cust.tel || null); edit(R.mobile, "C", cust.mobile || null);
    if (cust.email) edit(R.email, "C", cust.email); else DELETE.add(R.email);
    edit(10, "J", spec.quote_no || null);
    if (spec.date) edit(9, "J", null, { num: Math.round((isoToDate(spec.date) - Date.UTC(1899, 11, 30)) / 86400000) });

    // payment
    const pay = t.payment || { type: "full" };
    if (pay.type === "days") { showOrDel(false, [R.pay.full]); showOrDel(true, [R.pay.days]); edit(R.pay.days, "C", `Strictly within ${pay.days} days as from invoice date`); }
    else {
      showOrDel(true, [R.pay.full]); showOrDel(false, [R.pay.days]);
      const txt = pay.type === "down_delivery" ? `${pay.percent}% down payment upon order confirmation, balance before delivery`
        : pay.type === "down_week" ? `${pay.percent}% down payment upon order confirmation, payable within 1 week as from invoice date`
        : pay.type === "custom" ? pay.text : null;
      if (txt) edit(R.pay.full, "C", txt);
    }
    // delivery
    SHOW.add(R.delivery);
    if (t.delivery && t.delivery !== "Ex-stock, subject unsold") edit(R.delivery, "C", t.delivery);
    if (t.transport_fee) { SHOW.add(R.transport); CLEAR_B.add(R.transport); } else DELETE.add(R.transport);
    // validity
    const val = t.validity || { days: 14 };
    let vd;
    if (val.until) vd = isoToDate(val.until);
    else {   // the template only has "Strictly valid until": turn N days into a date (weekend -> Friday before)
      vd = new Date((spec.date ? isoToDate(spec.date) : new Date()).getTime() + (val.days || 14) * 86400000);
      if (vd.getUTCDay() === 6) vd = new Date(vd.getTime() - 86400000); else if (vd.getUTCDay() === 0) vd = new Date(vd.getTime() - 2 * 86400000);
    }
    SHOW.add(R.validUntil);
    edit(R.validUntil, "C", `Strictly valid until ${pad2(vd.getUTCDate())}-${MONTHS[vd.getUTCMonth()]}-${vd.getUTCFullYear()}`);
    // totals: only "Total Amount with VAT" (prices are entered including VAT)
    DELETE.add(R.totals[0]); DELETE.add(R.totals[1]);
    // installation, stainless note
    if (t.installation === false) DELETE.add(R.install); else SHOW.add(R.install);
    if (!t.stainless_note) DELETE.add(R.stainless);
    // warranty lines
    const norm = x => (x || "").replace(/:/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
    const wanted = new Set((spec.warranties || []).map(norm));
    let firstW = true, prevKept = false; const known = new Set();
    for (const r of R.warranty) {
      const c = cellText(r, "C"), sub = cellText(r, "D") || "";
      const label = norm(c), full = sub.trim().endsWith(":") ? norm((c || "") + " " + sub) : label;
      known.add(label); known.add(full);
      let keep;
      if (label.startsWith("* except led bar")) keep = prevKept;
      else keep = wanted.has(label) || wanted.has(full);
      if (keep) { SHOW.add(r); if (!label.startsWith("* except")) { if (firstW) firstW = false; else CLEAR_B.add(r); } }
      else DELETE.add(r);
      if (!label.startsWith("* except")) prevKept = keep;
    }
    for (const w of wanted) if (!known.has(w)) warnings.push(`Warranty line not found in the template: ${w}`);

    // ---- product block plan (segments are kept on one page; the heading bar is repeated on every page with products)
    // one segment per product (category bars and gaps are added while placing them)
    const segs = [], cats = [];
    for (const p of spec.products) if (!cats.includes(p.category)) cats.push(p.category);
    cats.forEach(cat => spec.products.filter(p => p.category === cat).forEach(p => {
      const items = [["product", p]];
      (p.features || []).forEach(f => items.push(["line", f]));
      if (p.capacity) items.push(["line", "Capacity:", p.capacity]);
      (p.dims || []).forEach(([a, b], i) => items.push(["line", i === 0 && /^(product|size)$/i.test(a) ? "Dimension:" : a.replace(/:?$/, ":"), b]));
      if (p.url) items.push(["link", p.url]);
      for (const f of p.notes || []) {
        if (f.length > 105) {
          let cut = f.lastIndexOf(". ", 104); cut = cut >= 0 ? cut + 1 : f.lastIndexOf(" ", 104);
          items.push(["note", f.slice(0, cut).trimEnd()]); items.push(["note", "    " + f.slice(cut).trim()]);
        } else items.push(["note", f]);
      }
      segs.push({ cat, items });
    }));
    // row heights of the template (after deletions) to plan the pages
    const tplRowH = new Map();
    for (const m of s.matchAll(/<row r="(\d+)"([^>]*)>/g)) { const h = m[2].match(/\sht="([\d.]+)"/); tplRowH.set(+m[1], h ? +h[1] : 14.4); }
    const tH = r => (DELETE.has(r) ? 0 : (tplRowH.has(r) ? tplRowH.get(r) : 14.4));
    const itemH = it => (it[0] === "banner" ? 18 : it[0] === "heading" ? tH(it[1]) : 14.4);
    const PAGE_PT = 841.89 - (0.275590551181102 + 0.393700787401575) * 72;
    const CAP = PAGE_PT / 0.767 * 0.96;
    let TITLE = 0; for (let r = 1; r <= 6; r++) TITLE += tH(r);
    const HEAD = R.headings.reduce((a, r) => a + tH(r), 0);
    let used = 0; for (let r = 1; r <= R.header; r++) used += tH(r);
    // Products are never split over two pages. When the next product does not fit in the space left on a page,
    // a later product that does fit is moved up, so pages are not left half empty.
    const plan = [], forced = [];       // forced: plan indexes where a new page starts (heading rows)
    let prevCat = null, pageStart = true;
    const withLead = (seg, atStart) => {
      const lead = [];
      if (seg.cat !== prevCat) { if (prevCat !== null && !atStart) lead.push(["blank"]); lead.push(["banner", seg.cat]); }
      else if (!atStart) lead.push(["blank"]);
      return lead.concat(seg.items);
    };
    const hOf = items => items.reduce((a, it) => a + itemH(it), 0);
    const queue = segs.slice();
    while (queue.length) {
      let idx = 0, items = withLead(queue[0], pageStart);
      if (used + hOf(items) > CAP) {
        const j = queue.findIndex((sg, k) => k > 0 && used + hOf(withLead(sg, pageStart)) <= CAP);
        if (j > 0) { idx = j; items = withLead(queue[j], pageStart); }
        else if (hOf(withLead(queue[0], true)) <= CAP - TITLE - HEAD) {
          forced.push(plan.length);
          R.headings.forEach(r => plan.push(["heading", r]));
          used = TITLE + HEAD; pageStart = true;
          items = withLead(queue[0], true);
        }
      }
      for (const it of items) { const h = itemH(it); if (used + h > CAP) used = TITLE; used += h; plan.push(it); }
      prevCat = queue[idx].cat; pageStart = false;
      queue.splice(idx, 1);
    }
    const N = plan.length;
    const delSorted = [...DELETE].sort((a, b) => a - b);
    const delBefore = r => { let k = 0; for (const d of delSorted) { if (d < r) k++; else break; } return k; };
    const newrow = r => {
      if (DELETE.has(r)) throw new Error("internal: deleted row referenced " + r);
      return r - delBefore(r) + (r > R.header ? N : 0);
    };
    const firstNew = newrow(R.header) + 1;

    // cell builders per layout (styles taken from the template's own example rows)
    const cell = (ref, style, o = {}) => {
      if (o.text != null) return `<c r="${ref}" s="${style}" t="s"><v>${sidx(o.text)}</v></c>`;
      if (o.formula != null) return `<c r="${ref}" s="${style}"><f>${o.formula}</f><v>${o.num}</v></c>`;
      if (o.num != null) return `<c r="${ref}" s="${style}"><v>${o.num}</v></c>`;
      return `<c r="${ref}" s="${style}"/>`;
    };
    const styleAt = (xml, r, ref, dflt) => { const m = rowOf(xml, r).match(new RegExp(`<c r="${ref}"[^>]*\\ss="(\\d+)"`)); return m ? +m[1] : dflt; };
    const ST = { discProdL: styleAt(sDisc, 22, "L22", 84), discLine: styleAt(sDisc, 24, "H24", 83) };
    const ex = p => ({ unit: p.price, normal: p.normal || p.price });   // unit prices INCLUDING VAT
    const newRows = [], merges = [], productRows = [], links = [];
    const headingXml = r => (discount ? rowOf(sDisc, r) : "") || rowOf(s, r);
    let amountSum = 0;
    plan.forEach((item, k) => {
      const r = firstNew + k, kind = item[0];
      if (kind === "heading") {
        let rx = headingXml(item[1]).replace(/<c r="[M-Z]\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, "");
        rx = rx.replace(/<row r="\d+"/, `<row r="${r}"`).replace(/(<c r=")([A-Z]+)\d+(")/g, (m, a, col, b) => a + col + r + b);
        newRows.push(rx);
      } else if (kind === "banner") {
        const cells = cell(`B${r}`, 18, { text: item[1] }) + [..."CDEFGHIJ"].map(x => cell(`${x}${r}`, 18)).join("") + cell(`L${r}`, 55) + cell(`M${r}`, 56);
        newRows.push(`<row r="${r}" s="2" customFormat="1" ht="18" spans="2:13">${cells}</row>`); merges.push(`B${r}:J${r}`);
      } else if (kind === "product") {
        const p = item[1], qty = p.qty || 1, e = ex(p), amount = qty * e.unit; amountSum += amount;
        let cells;
        if (discount) cells = cell(`B${r}`, 19, { text: p.model }) + cell(`C${r}`, 20, { text: p.name }) + cell(`D${r}`, 21) + cell(`E${r}`, 22) + cell(`F${r}`, 22)
          + cell(`G${r}`, 25, { num: qty }) + cell(`H${r}`, 45, { num: e.normal }) + cell(`I${r}`, 45, { num: e.unit }) + cell(`J${r}`, 45, { num: amount, formula: `I${r}*G${r}` }) + cell(`L${r}`, ST.discProdL);
        else cells = cell(`B${r}`, 19, { text: p.model }) + cell(`C${r}`, 20, { text: p.name }) + cell(`D${r}`, 21) + cell(`E${r}`, 22) + cell(`F${r}`, 23) + cell(`G${r}`, 24)
          + cell(`H${r}`, 25, { num: qty }) + cell(`I${r}`, 45, { num: e.unit }) + cell(`J${r}`, 45, { num: amount, formula: `H${r}*I${r}` }) + cell(`L${r}`, 57);
        newRows.push(`<row r="${r}" s="${discount ? 4 : 3}" customFormat="1" spans="2:12">${cells}</row>`);
        const nBody = (p.features || []).length + (p.capacity ? 1 : 0) + (p.dims || []).length;
        productRows.push([r, nBody, p]);
      } else {
        const [ctext, dtext] = [item[1], item[2]];
        let cStyle = 7, dStyle = 7, dOpt = dtext != null ? { text: dtext } : {}, cOpt = {};
        if (kind === "line") cOpt = { text: ctext };
        if (kind === "note") { cStyle = 35; cOpt = { text: ctext }; }
        if (kind === "link") {
          cOpt = { text: "More info:" }; dStyle = 34;
          const url = item[1];
          const disp = url.length <= 95 ? url : (() => { const last = url.replace(/\/$/, "").split("/").pop(); const sh = `www.dragonelectronics.mu/\u2026/${last}`; return sh.length <= 95 ? sh : "View this product on www.dragonelectronics.mu"; })();
          dOpt = { text: disp }; links.push([`D${r}`, url, disp]);
        }
        const tailCells = discount ? cell(`E${r}`, 22) + cell(`F${r}`, 22) + cell(`G${r}`, 25) + cell(`H${r}`, ST.discLine) + cell(`I${r}`, ST.discLine) + cell(`J${r}`, ST.discLine) + cell(`L${r}`, 60)
          : cell(`E${r}`, 22) + cell(`F${r}`, 31) + cell(`G${r}`, 32) + cell(`H${r}`, 28) + cell(`I${r}`, 28) + cell(`J${r}`, 28) + cell(`L${r}`, 57);
        const cells = kind === "blank" ? cell(`B${r}`, 29) + cell(`C${r}`, 7) + cell(`D${r}`, 7) + tailCells
          : cell(`B${r}`, 29) + cell(`C${r}`, cStyle, cOpt) + cell(`D${r}`, dStyle, dOpt) + tailCells;
        newRows.push(`<row r="${r}" s="${discount ? 4 : 3}" customFormat="1" spans="2:12">${cells}</row>`);
      }
    });

    // ---- rebuild sheetData
    const i1 = s.indexOf("<sheetData>"), i2 = s.indexOf("</sheetData>");
    let head = s.slice(0, i1), data = s.slice(i1 + 11, i2), tail = s.slice(i2 + 12);
    const rows = data.match(/<row r="\d+"[^>]*\/>|<row r="\d+"[\s\S]*?<\/row>/g) || [];
    if (rows.join("") !== data) throw new Error("internal: unexpected sheet layout");
    const shiftRefs = str => str.replace(/(?<![A-Z$])(\$?)([A-Z]{1,2})(\$?)(\d+)(?!\d)/g, (m, d1, col, d2, row) => d1 + col + d2 + (DELETE.has(+row) ? row : newrow(+row)));
    const totalsVal = { [R.totals[2]]: amountSum };
    const out = [];
    for (let rx of rows) {
      const r = +rx.match(/<row r="(\d+)"/)[1];
      if (DELETE.has(r)) continue;
      if (discount && R.headings.includes(r)) rx = rowOf(sDisc, r) || rx;          // headings with Normal + Discounted price
      if (SHOW.has(r)) rx = rx.replace(' hidden="1"', "");
      if (r === R.transport) rx = rx.replace(new RegExp(`(<c r="[C-G]${R.transport}" s=")44(")`, "g"), (m, a, b) => a + "7" + b);   // drop the yellow reminder highlight
      if (CLEAR_B.has(r)) rx = setCellInRow(rx, `B${r}`, "", "");
      for (const [col, txt, o] of EDIT.get(r) || []) {
        if (o.num != null) rx = setCellInRow(rx, `${col}${r}`, `<v>${o.num}</v>`, "");
        else if (txt) rx = setCellInRow(rx, `${col}${r}`, `<v>${sidx(txt)}</v>`, ' t="s"');
        else rx = setCellInRow(rx, `${col}${r}`, "", "");
      }
      if (r === R.totals[2]) rx = rx.replace(/<f>[^<]*<\/f>/, `<f>SUM(J${R.headings[0]}:J${R.blankBeforeTotals})</f>`);
      if (totalsVal[r] != null) rx = rx.replace(new RegExp(`(<c r="J${r}"[^>]*><f>[^<]*</f><v>)[^<]*(</v>)`), (m, a, b) => a + totalsVal[r] + b);
      const nr = newrow(r);
      if (nr !== r || r > R.header) {
        rx = rx.replace(/<row r="\d+"/, `<row r="${nr}"`);
        rx = rx.replace(/(<c r=")([A-Z]+)\d+(")/g, (m, a, col, b) => a + col + nr + b);
      }
      rx = rx.replace(/<f>([^<]*)<\/f>/g, (m, f) => `<f>${shiftRefs(f)}</f>`);
      out.push(rx);
      if (r === R.header) out.push(...newRows);
    }
    data = out.join("");

    // ---- merges, hyperlinks, page breaks
    const oldM = [...tail.matchAll(/<mergeCell ref="([^"]+)"\/>/g)].map(m => m[1]);
    const keepM = oldM.filter(m => !(m.match(/\d+/g) || []).some(x => DELETE.has(+x))).map(shiftRefs);
    const allM = keepM.concat(merges);
    tail = tail.replace(/<mergeCells count="\d+">[\s\S]*?<\/mergeCells>|<mergeCells count="\d+"\/>/, `<mergeCells count="${allM.length}">` + allM.map(m => `<mergeCell ref="${m}"/>`).join("") + "</mergeCells>");
    let srels = files.has(SHEET_RELS) ? get(SHEET_RELS) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    for (const m of tail.matchAll(/<hyperlink [^>]*r:id="([^"]+)"[^>]*\/>/g)) srels = srels.replace(new RegExp(`<Relationship [^>]*Id="${m[1]}"[^>]*/>`), "");
    tail = tail.replace(/<hyperlinks>[\s\S]*?<\/hyperlinks>/, "");
    if (links.length) {
      const hl = "<hyperlinks>" + links.map(([ref, url, disp], i) => `<hyperlink ref="${ref}" r:id="rIdQL${i + 1}" display="${esc(disp)}"/>`).join("") + "</hyperlinks>";
      tail = tail.replace(/<printOptions|<pageMargins/, m => hl + m);
      srels = srels.replace("</Relationships>", links.map(([ref, url], i) => `<Relationship Id="rIdQL${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(url)}" TargetMode="External"/>`).join("") + "</Relationships>");
    }
    tail = tail.replace(/<rowBreaks[\s\S]*?<\/rowBreaks>/, "");

    const heights = new Map();
    for (const m of data.matchAll(/<row r="(\d+)"([^>]*)>/g)) {
      const attrs = m[2];
      if (attrs.includes('hidden="1"')) heights.set(+m[1], 0);
      else { const h = attrs.match(/\sht="([\d.]+)"/); heights.set(+m[1], h ? +h[1] : 14.4); }
    }
    const H = r => (heights.has(r) ? heights.get(r) : 14.4);
    const sumH = (a, b) => { let x = 0; for (let r = a; r <= b; r++) x += H(r); return x; };
    const lastRow = newrow(R.last);
    const forcedRows = new Set(forced.map(k => firstNew + k));
    const keep = [];
    keep.push([newrow(R.blankBeforeTotals), newrow(R.totals[2])]);
    const terms = newrow(R.terms);
    keep.push(sumH(terms, lastRow) <= CAP - TITLE ? [terms, lastRow] : [newrow(R.closing), lastRow]);
    const starts = new Map(keep);
    const breaks = [];
    used = 0;
    for (let r = 1; r <= lastRow; r++) {
      if (forcedRows.has(r)) { breaks.push(r - 1); used = TITLE; }
      else if (starts.has(r)) {
        const bh = sumH(r, starts.get(r));
        if (used + bh > CAP && used > TITLE + 1 && bh <= CAP - TITLE) { breaks.push(r - 1); used = TITLE; }
      }
      if (used + H(r) > CAP) used = TITLE;
      used += H(r);
    }
    if (breaks.length) {
      const rb = `<rowBreaks count="${breaks.length}" manualBreakCount="${breaks.length}">` + breaks.map(b => `<brk id="${b}" max="16383" man="1"/>`).join("") + "</rowBreaks>";
      tail = tail.includes("<drawing ") ? tail.replace("<drawing ", rb + "<drawing ") : tail.replace("</worksheet>", rb + "</worksheet>");
    }
    head = head.replace(/<dimension ref="A1:([A-Z]+)(\d+)"\/>/, (m, c, n) => `<dimension ref="A1:${c}${lastRow + 1}"/>`);
    head = head.replace(/<sheetView /, m => (head.includes("tabSelected=") ? m : '<sheetView tabSelected="1" '));
    head = head.replace(/<selection [^>]*\/>/, '<selection activeCell="A1" sqref="A1"/>').replace(/ topLeftCell="[^"]*"/, "");
    put(SHEET, head + "<sheetData>" + data + "</sheetData>" + tail);
    put(SHEET_RELS, srels);

    // ---- drawing: keep logo + signature (shifted), drop the example photo, add product photos
    const drawRel = (srels.match(/<Relationship [^>]*Target="\.\.\/drawings\/(drawing\d+\.xml)"/) || [])[1];
    const DRAW = "xl/drawings/" + drawRel, DRAW_RELS = "xl/drawings/_rels/" + drawRel + ".rels";
    let d = get(DRAW), rels = get(DRAW_RELS);
    const anchors = d.match(/<xdr:(?:twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:(?:twoCellAnchor|oneCellAnchor)>/g) || [];
    const body = [];
    for (let a of anchors) {
      const r0 = +a.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/)[1];
      if (r0 + 1 >= R.example[0] && r0 + 1 <= R.example[R.example.length - 1]) continue;
      a = a.replace(/<xdr:row>(\d+)<\/xdr:row>/g, (m, n) => { const r1 = +n + 1; let nr = r1; while (DELETE.has(nr)) nr++; return `<xdr:row>${newrow(nr) - 1}</xdr:row>`; });
      body.push(a);
    }
    const EMU_PX = 9525, ROW_EMU = Math.floor(14.4 * 12700);
    const colEmu = { 6: 72 * EMU_PX, 7: 114 * EMU_PX, 8: 114 * EMU_PX, 9: 113 * EMU_PX };
    const regionW = colEmu[6] + colEmu[7] + colEmu[8] + colEmu[9];
    const xAnchor = x => { let col = 6; while (col < 9 && x >= colEmu[col]) { x -= colEmu[col]; col++; } return [col, Math.floor(x)]; };
    productRows.forEach(([prow, nBody, p], idx) => {
      const img = images[p.model];
      if (!img) { warnings.push(`No photo for ${p.model}`); return; }
      const i = idx + 1, aspect = img.width / img.height;
      let h = Math.min(2400000, nBody * ROW_EMU - 90000), w = Math.floor(h * aspect);
      if (w > 3000000) { w = 3000000; h = Math.floor(w / aspect); }
      if (h <= 0) return;
      const x = Math.floor((regionW - w) / 2), [fc, fo] = xAnchor(x), [tc, to] = xAnchor(x + w);
      const rowsH = Math.floor((h + 20000) / ROW_EMU), rem = h + 20000 - rowsH * ROW_EMU;
      const name = `quote_photo_${pad2(i)}.jpeg`, rid = `rIdQ${i}`;
      files.set(`xl/media/${name}`, img.bytes);
      rels = rels.replace("</Relationships>", `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${name}"/></Relationships>`);
      body.push(`<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${fc}</xdr:col><xdr:colOff>${fo}</xdr:colOff><xdr:row>${prow}</xdr:row><xdr:rowOff>20000</xdr:rowOff></xdr:from>`
        + `<xdr:to><xdr:col>${tc}</xdr:col><xdr:colOff>${to}</xdr:colOff><xdr:row>${prow + rowsH}</xdr:row><xdr:rowOff>${rem}</xdr:rowOff></xdr:to>`
        + `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${300 + i}" name="Photo ${esc(p.model)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>`
        + `<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`
        + `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`);
    });
    const pre = anchors.length ? d.slice(0, d.indexOf(anchors[0])) : d.replace(/<\/xdr:wsDr>[\s\S]*$/, "");
    const post = anchors.length ? d.slice(d.lastIndexOf(anchors[anchors.length - 1]) + anchors[anchors.length - 1].length) : "</xdr:wsDr>";
    put(DRAW, pre + body.join("") + post);
    const usedR = new Set([...body.join("").matchAll(/r:embed="([^"]+)"/g)].map(m => m[1]));
    for (const m of [...rels.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*\/>/g)]) if (!usedR.has(m[1])) rels = rels.replace(m[0], "");
    put(DRAW_RELS, rels);

    // ---- remove the "With Disc Column" sheet, rename "NO Disc" to "Quotation"
    let w = get("xl/workbook.xml");
    const sheetTags = [...w.matchAll(/<sheet [^>]*\/>/g)].map(m => m[0]);
    const discIdx = sheetTags.indexOf(disc.tag), baseIdx = sheetTags.indexOf(base.tag);
    w = w.replace(disc.tag, "").replace(base.tag, base.tag.replace(/name="[^"]+"/, 'name="Quotation"'));
    w = w.replace(/<definedName ([^>]*)>([^<]*)<\/definedName>/g, (m, attrs, val) => {
      const lm = attrs.match(/localSheetId="(\d+)"/);
      if (lm && +lm[1] === discIdx) return "";
      if (/'With Disc Column'!/.test(val)) return "";
      let a2 = attrs; if (lm && +lm[1] > discIdx) a2 = attrs.replace(/localSheetId="\d+"/, `localSheetId="${+lm[1] - 1}"`);
      let v2 = val.replace(/'NO Disc'!/g, "'Quotation'!");
      if (/_xlnm\.Print_Area/.test(attrs)) v2 = v2.replace(/(\$[A-Z]+\$)(\d+)$/, (mm, c) => c + lastRow);
      return `<definedName ${a2}>${v2}</definedName>`;
    });
    w = w.replace(/<definedNames>\s*<\/definedNames>/, "");
    w = w.replace(/activeTab="\d+"/, 'activeTab="0"').replace(/ firstSheet="\d+"/, "");
    w = w.replace(/<calcPr([^>]*?)\s*\/>/, (m, a) => "<calcPr" + a + (a.includes("fullCalcOnLoad") ? "" : ' fullCalcOnLoad="1"') + "/>");
    put("xl/workbook.xml", w);
    let wr = get("xl/_rels/workbook.xml.rels");
    wr = wr.replace(new RegExp(`<Relationship [^>]*Id="${disc.rid}"[^>]*/>`), "");
    put("xl/_rels/workbook.xml.rels", wr);
    const discRels = disc.path.replace(/worksheets\/(sheet\d+\.xml)$/, "worksheets/_rels/$1.rels");
    if (files.has(discRels)) {
      const dr = (get(discRels).match(/Target="\.\.\/drawings\/(drawing\d+\.xml)"/) || [])[1];
      if (dr && dr !== drawRel) { files.delete("xl/drawings/" + dr); files.delete("xl/drawings/_rels/" + dr + ".rels"); }
      files.delete(discRels);
    }
    files.delete(disc.path);
    let ct = get("[Content_Types].xml");
    for (const part of [disc.path]) ct = ct.replace(new RegExp(`<Override PartName="/${part.replace(/[.]/g, "\\.")}"[^>]*/>`), "");
    ct = ct.replace(/<Override PartName="\/xl\/drawings\/(drawing\d+\.xml)"[^>]*\/>/g, (m, f) => (files.has("xl/drawings/" + f) ? m : ""));
    if (!/Extension="jpeg"/i.test(ct)) ct = ct.replace("<Default ", '<Default Extension="jpeg" ContentType="image/jpeg"/><Default ');
    put("[Content_Types].xml", ct);
    if (files.has("docProps/app.xml")) {
      let a = get("docProps/app.xml");
      a = a.replace(/<vt:lpstr>With Disc Column<\/vt:lpstr>/, "").replace(/<vt:lpstr>NO Disc<\/vt:lpstr>/, "<vt:lpstr>Quotation</vt:lpstr>");
      a = a.replace(/(<TitlesOfParts><vt:vector size=")\d+/, "$11").replace(/(<vt:i4>)\d+(<\/vt:i4>)/, "$11$2");
      put("docProps/app.xml", a);
    }
    const allRels = [...files.keys()].filter(k => k.endsWith(".rels")).map(k => get(k)).join("");
    for (const k of [...files.keys()]) if (k.startsWith("xl/media/") && !allRels.includes(k.slice(3)) && !allRels.includes("media/" + k.split("/").pop())) files.delete(k);

    // ---- shared strings
    sst = sst.replace("</sst>", newSi.join("") + "</sst>").replace(/uniqueCount="\d+"/, `uniqueCount="${nSi}"`);
    sst = sst.replace(/ count="(\d+)"/, (m, n) => ` count="${+n + newSi.length}"`);
    put("xl/sharedStrings.xml", sst);

    return { bytes: await zip(files), warnings, total: amountSum, totalExVat: amountSum / VAT };
  }

  // ------------------------------------------------------------------ business rules
  function warrantyFor(p) {
    const n = (p.name + " " + p.category).toLowerCase();
    if (n.includes("washtower")) return "LG WashTower";
    if (n.includes("dishwasher")) return "LG Dishwasher";
    if (n.includes("refrigerator") || n.includes("freezer")) return "LG Refrigerator";
    if (n.includes("dryer") && !n.includes("washer") && !n.includes("washing")) return "LG Dryer";
    if (n.includes("washer") || n.includes("washing")) return "LG Washer";
    if (n.includes("microwave")) return n.includes("neochef") ? "LG Microwave NeoChef" : n.includes("solardom") ? "LG Microwave SolarDOM" : "LG Microwave Non Inverter";
    if (n.includes("sound bar") || n.includes("soundbar")) return "LG Sound Bar";
    if (n.includes("xboom") || n.includes("x-boom")) return "LG X-BOOM";
    if (n.includes("dehumidifier")) return "LG Dehumidifier";
    if (n.includes("signage")) return "LG Digital Signage";
    if (n.includes("iptv") || n.includes("hotel tv")) return "LG IPTV";
    if (/\btv\b|oled|qned|nanocell|uhd/.test(n)) return "LG TV";
    if (n.includes("multi split")) return "LG Multi Split AC";
    if (n.includes("commercial") && (n.includes("air") || n.includes("ac"))) return "LG Commercial AC";
    if (n.includes("air con") || n.includes("air condition") || /\bac\b/.test(n)) return "LG Room Air Con";
    if (n.includes("vacuum")) return /a9t/.test(n) ? "LG Vacuum Cleaner A9T series" : /a9k/.test(n) ? "LG Vacuum Cleaner A9K series" : "LG Vacuum Cleaner VC5420NHT";
    if (n.includes("ariston")) return "ARISTON";
    return null;
  }

  const COMPANY_WORDS = /^(ltd|ltee|lt\u00e9e|limited|co|company|plc|inc|sarl|sa|cie|group|groupe|corporation|corp|enterprise|enterprises|trading|holdings|hotel|resort|school|college|ministry|bank|society|association|club|clinic|hospital)\.?$/i;
  function looksLikeCompany(name) {
    const words = String(name || "").replace(/[()]/g, " ").split(/\s+/).filter(Boolean);
    if (/^(mr|mrs|ms|miss|dr|me|mme|mlle)\.?$/i.test(words[0] || "")) return false;
    return words.some(w => COMPANY_WORDS.test(w)) || words.length > 3;
  }
  function initials(name) {
    const titles = /^(mr|mrs|ms|miss|dr|me|m|mme|mlle)\.?$/i, skip = /^(ltd|ltee|lt\u00e9e|limited|co|company|plc|inc|sarl|sa|and|&|the|of)\.?$/i;
    const words = String(name || "").replace(/[()]/g, " ").split(/\s+/).filter(w => w && !titles.test(w));
    if (looksLikeCompany(name)) return words.filter(w => !skip.test(w)).map(w => w[0]).join("").toUpperCase().slice(0, 4) || "XX";
    if (words.length >= 2) return (words[0][0] + words[words.length - 1][0]).toUpperCase();
    return (words[0] || "XX").slice(0, 2).toUpperCase();
  }

  const API = { buildQuote, checkTemplate, warrantyFor, initials, looksLikeCompany, unzip, zip };
  if (typeof module !== "undefined" && module.exports) module.exports = API;
  else root.DQCore = API;
})(typeof window !== "undefined" ? window : globalThis);

/* Dragon Electronics Quote Builder \u2013 in-page panel. Loaded by the bookmark on www.dragonelectronics.mu. */
(function () {
  "use strict";
  const BASE = (document.currentScript && document.currentScript.src || "").replace(/[^/]*(\?.*)?$/, "");
  if (!/(^|\.)dragonelectronics\.mu$/.test(location.hostname)) {
    window.alert("Open www.dragonelectronics.mu first, then click the Quote Builder bookmark again.");
    return;
  }
  const existing = document.getElementById("dq-panel");
  if (existing) { existing.style.display = "flex"; return; }

  const store = {
    get(k, d) { try { const v = localStorage.getItem("dq_" + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem("dq_" + k, JSON.stringify(v)); } catch (e) { /* private mode */ } }
  };
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const fmt = n => "Rs " + Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const pad2 = n => String(n).padStart(2, "0");
  const isoLocal = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  const h = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  // ------------------------------------------------------------------ UI
  const css = `
  #dq-panel{position:fixed;top:0;right:0;width:440px;max-width:100vw;height:100vh;z-index:2147483646;background:#fff;color:#222;
    box-shadow:-4px 0 24px rgba(0,0,0,.25);display:flex;flex-direction:column;font:13px/1.4 Arial,Helvetica,sans-serif}
  #dq-panel *{box-sizing:border-box;font-family:inherit}
  #dq-panel header{background:#c8102e;color:#fff;padding:12px 16px;display:flex;align-items:center;justify-content:space-between}
  #dq-panel header b{font-size:15px}
  #dq-panel header button{background:none;border:0;color:#fff;font-size:22px;cursor:pointer;line-height:1}
  #dq-body{overflow:auto;padding:12px 16px 16px;flex:1}
  #dq-panel fieldset{border:1px solid #ddd;border-radius:6px;margin:0 0 12px;padding:8px 10px 10px}
  #dq-panel legend{font-weight:bold;padding:0 4px;color:#c8102e}
  #dq-panel label{display:block;margin:6px 0 2px;color:#555;font-size:12px}
  #dq-panel input[type=text],#dq-panel input[type=number],#dq-panel input[type=date],#dq-panel select,#dq-panel textarea{
    width:100%;padding:6px 8px;border:1px solid #bbb;border-radius:4px;font-size:13px;background:#fff;color:#222}
  #dq-panel textarea{min-height:90px;font-family:Consolas,monospace}
  #dq-panel .row{display:flex;gap:8px}#dq-panel .row>*{flex:1}
  #dq-panel .chk{display:flex;align-items:center;gap:6px;margin-top:8px;color:#333;font-size:13px}
  #dq-panel .hint{color:#888;font-size:11px;margin-top:2px}
  #dq-go{width:100%;padding:10px;border:0;border-radius:6px;background:#c8102e;color:#fff;font-size:15px;font-weight:bold;cursor:pointer}
  #dq-go:disabled{background:#999;cursor:wait}
  #dq-log{margin-top:12px;font-size:12px}
  #dq-log .ok{color:#1a7f37}#dq-log .warn{color:#b35900}#dq-log .err{color:#c8102e;font-weight:bold}
  #dq-panel table{width:100%;border-collapse:collapse;margin-top:8px;font-size:12px}
  #dq-panel th,#dq-panel td{border-bottom:1px solid #eee;padding:4px;text-align:left;vertical-align:top}
  #dq-panel td.n,#dq-panel th.n{text-align:right;white-space:nowrap}
  #dq-panel .dl{display:inline-block;margin-top:10px;padding:8px 12px;background:#1a7f37;color:#fff;border-radius:5px;text-decoration:none;font-weight:bold}
  #dq-panel .prow{display:flex;gap:6px;align-items:center;margin-top:6px;flex-wrap:wrap}
  #dq-panel .prow .pm{flex:1 1 150px;text-transform:uppercase}#dq-panel .prow .pq{flex:0 0 58px;padding:6px 4px;text-align:center}
  #dq-panel .prow .pd{display:flex;align-items:center;gap:4px;margin:0;font-size:12px;color:#333;white-space:nowrap;cursor:pointer}
  #dq-panel .prow .pp{order:5;flex:1 1 100%;margin-left:0;border-color:#c8102e}
  #dq-panel .prow .px{background:none;border:0;color:#999;font-size:18px;cursor:pointer;padding:0 2px;line-height:1}
  #dq-panel .prow .px:hover{color:#c8102e}
  #dq-addrow{margin-top:8px;padding:5px 10px;border:1px dashed #c8102e;background:#fff;color:#c8102e;border-radius:4px;cursor:pointer;font-size:12px}
  #dq-panel details{margin-top:12px}#dq-panel summary{cursor:pointer;color:#c8102e}`;
  const st = document.createElement("style"); st.textContent = css; document.head.appendChild(st);

  const today = new Date();
  const until = new Date(today); until.setDate(until.getDate() + 14);
  const panel = document.createElement("div"); panel.id = "dq-panel";
  panel.innerHTML = `
  <header><b>Quote Builder \u00b7 Dragon Electronics</b><button id="dq-close" title="Close">\u00d7</button></header>
  <div id="dq-body">
    <fieldset><legend>Customer</legend>
      <label>To (name or company) *</label><input type="text" id="dq-to" placeholder="Mr. Jean DOE  or  ABC TRADING LTD">
      <div class="chk"><input type="checkbox" id="dq-company"> Customer is a company <span class="hint">(adds the Attention line, name in CAPITALS)</span></div>
      <div id="dq-cobox" style="display:none">
        <label>Attention (Title First name SURNAME)</label><input type="text" id="dq-att" placeholder="Mr. Jean DOE, Purchasing Manager">
      </div>
      <label>Address</label><input type="text" id="dq-addr1" placeholder="Street"><input type="text" id="dq-addr2" placeholder="Town" style="margin-top:4px">
      <div class="row"><div><label>Tel</label><input type="text" id="dq-tel"></div><div><label>Mobile</label><input type="text" id="dq-mob"></div></div>
      <div class="chk"><input type="checkbox" id="dq-hasemail"> Add email</div>
      <input type="text" id="dq-email" placeholder="name@example.com" style="display:none">
    </fieldset>
    <fieldset><legend>Products</legend>
      <div class="hint">Model, quantity (empty = 1). Tick <b>Discount</b> to type your own price, incl. VAT.</div>
      <div id="dq-rows"></div>
      <button type="button" id="dq-addrow">+ Add product</button>
      <label>Price</label>
      <select id="dq-price"><option value="retail">Normal retail price</option><option value="promo">Promo price where available</option></select>
      <div class="chk"><input type="checkbox" id="dq-disc"> Show Normal + Discounted price columns</div>
      <div class="hint">Unit prices include VAT; the quote shows the Total Amount with VAT only.</div>
    </fieldset>
    <fieldset><legend>Terms</legend>
      <label>Payment</label>
      <div class="row"><select id="dq-pay">
        <option value="full">Payment in full upon order confirmation</option>
        <option value="days">Within \u2026 days of invoice</option>
        <option value="down_delivery">\u2026% down payment, balance before delivery</option>
        <option value="down_week">\u2026% down payment, within 1 week (Kitchen Maker)</option>
        <option value="custom">Custom text</option></select>
        <input type="number" id="dq-payn" value="30" style="max-width:80px;display:none"></div>
      <input type="text" id="dq-paytxt" placeholder="Custom payment text" style="display:none;margin-top:4px">
      <label>Delivery</label>
      <div class="row"><select id="dq-deliv"><option value="stock">Ex-stock, subject unsold</option><option value="days">Within \u2026 days as from invoice date</option></select>
        <input type="number" id="dq-delivn" min="1" max="30" value="30" style="max-width:80px;display:none"></div>
      <div class="hint" id="dq-delivhint" style="display:none">Maximum 30 days.</div>
      <div class="chk"><input type="checkbox" id="dq-transport"> Add the transport fee line</div>
      <label>Validity</label>
      <input type="date" id="dq-until" value="${isoLocal(until)}">
      <div class="hint">Written as "Strictly valid until dd-mmm-yyyy". A Saturday or Sunday is moved back to the Friday before. Check public holidays yourself.</div>
      <div class="chk"><input type="checkbox" id="dq-install" checked> Include the installation clause</div>
    </fieldset>
    <fieldset><legend>Quotation</legend>
      <div class="row"><div><label>Quotation No</label><input type="text" id="dq-no"></div><div><label>Date</label><input type="date" id="dq-date" value="${isoLocal(today)}"></div></div>
      <div class="hint">DP | customer initials | running number. Last number used: <input type="number" id="dq-counter" style="width:64px;padding:2px 4px"></div>
      <label>Quotation template</label>
      <div id="dq-tplstat" class="hint"></div>
      <input type="file" id="dq-tpl" accept=".xlsx" style="font-size:12px;margin-top:4px">
    </fieldset>
    <button id="dq-go">Generate quotation</button>
    <div id="dq-log"></div>
    <details><summary>Quote history on this computer</summary><div id="dq-hist"></div></details>
  </div>`;
  document.body.appendChild(panel);
  const $ = id => panel.querySelector("#" + id);
  $("dq-close").onclick = () => { panel.style.display = "none"; };
  $("dq-counter").value = store.get("counter", 0);

  const upd = () => {
    $("dq-paytxt").style.display = $("dq-pay").value === "custom" ? "block" : "none";
    $("dq-payn").style.display = ["days", "down_delivery", "down_week"].includes($("dq-pay").value) ? "block" : "none";
    $("dq-delivn").style.display = $("dq-delivhint").style.display = $("dq-deliv").value === "days" ? "block" : "none";
    $("dq-email").style.display = $("dq-hasemail").checked ? "block" : "none";
    $("dq-cobox").style.display = $("dq-company").checked ? "block" : "none";
  };
  $("dq-deliv").addEventListener("change", () => upd());
  $("dq-delivn").addEventListener("change", () => { const n = Math.round(+$("dq-delivn").value); $("dq-delivn").value = Math.min(30, Math.max(1, n || 30)); });
  ["dq-pay"].forEach(id => $(id).addEventListener("change", () => {
    if (id === "dq-pay") $("dq-payn").value = $("dq-pay").value === "days" ? 30 : 50;
    upd();
  }));
  let noEdited = false, coEdited = false;
  $("dq-company").addEventListener("change", () => { coEdited = true; upd(); });
  $("dq-hasemail").addEventListener("change", upd);
  $("dq-to").addEventListener("input", () => { if (!coEdited) { $("dq-company").checked = DQCore.looksLikeCompany($("dq-to").value); upd(); } });
  $("dq-price").addEventListener("change", () => { if ($("dq-price").value === "promo") $("dq-disc").checked = true; });
  // product lines: model | qty | [ ] Discount -> price
  function addRow(model, qty) {
    const row = document.createElement("div"); row.className = "prow";
    row.innerHTML = `<input type="text" class="pm" placeholder="Model, e.g. GC-X31FFCRB"><input type="number" class="pq" min="1" placeholder="Qty">`
      + `<label class="pd"><input type="checkbox" class="pdc"> Discount</label><input type="text" class="pp" placeholder="Discounted price Rs (incl. VAT), e.g. 180000" style="display:none">`
      + `<button type="button" class="px" title="Remove this product">\u00d7</button>`;
    const q = s => row.querySelector(s);
    q(".pm").value = model || ""; q(".pq").value = qty || "";
    q(".pdc").addEventListener("change", () => {
      q(".pp").style.display = q(".pdc").checked ? "block" : "none";
      if (q(".pdc").checked) { $("dq-disc").checked = true; q(".pp").focus(); }
    });
    q(".px").addEventListener("click", () => { row.remove(); if (!$("dq-rows").children.length) addRow(); });
    q(".pm").addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); const r = addRow(); r.querySelector(".pm").focus(); } });
    q(".pm").addEventListener("paste", e => {        // pasting a list of models makes one line per model
      const txt = (e.clipboardData || window.clipboardData).getData("text");
      const lines = txt.split(/[\r\n;,]+/).map(x => x.trim()).filter(Boolean);
      if (lines.length < 2) return;
      e.preventDefault();
      lines.forEach((l, i) => {
        const m = l.match(/^([A-Za-z0-9.\-\/]+)\s*(?:[x\u00d7*]\s*|\s)?(\d+)?$/i) || [null, l, ""];
        if (i === 0) { q(".pm").value = m[1]; q(".pq").value = m[2] || ""; } else addRow(m[1], m[2]);
      });
    });
    $("dq-rows").appendChild(row);
    return row;
  }
  addRow(); addRow(); addRow();
  $("dq-addrow").addEventListener("click", () => addRow().querySelector(".pm").focus());
  const suggestNo = () => { if (!noEdited) $("dq-no").value = `DP | ${DQCore.initials($("dq-to").value)} | ${pad2((+$("dq-counter").value || 0) + 1)}`; };
  $("dq-to").addEventListener("input", suggestNo); $("dq-counter").addEventListener("input", suggestNo);
  $("dq-no").addEventListener("input", () => { noEdited = true; });
  suggestNo(); upd();
  const renderHist = () => {
    const hist = store.get("history", []);
    $("dq-hist").innerHTML = hist.length ? `<table><tr><th>Date</th><th>No</th><th>Customer</th><th class="n">Total</th></tr>${hist.slice().reverse().map(q => `<tr><td>${h(q.date)}</td><td>${h(q.no)}</td><td>${h(q.customer)}</td><td class="n">${fmt(q.total)}</td></tr>`).join("")}</table>` : "<div class='hint'>No quotes yet.</div>";
  };
  renderHist();

  const log = (msg, cls) => { const d = document.createElement("div"); if (cls) d.className = cls; d.innerHTML = msg; $("dq-log").appendChild(d); d.scrollIntoView({ block: "nearest" }); return d; };

  // ------------------------------------------------------------------ product lookup
  let sitemap = null;
  async function productUrl(model) {
    if (!sitemap) {
      try { const t = await (await fetch("/sitemap.xml", { cache: "no-store" })).text(); sitemap = [...t.matchAll(/<loc>([^<]+)<\/loc>/g)].map(m => m[1]); }
      catch (e) { sitemap = []; }
    }
    const m = model.toLowerCase();
    const hit = sitemap.find(u => u.toLowerCase().endsWith("-" + m) || u.toLowerCase().endsWith("/" + m));
    return hit || location.origin + "/" + model.replace(/-/g, "");
  }

  function loadFrame(url) {
    return new Promise((resolve, reject) => {
      const f = document.createElement("iframe");
      f.style.cssText = "position:fixed;left:-20000px;top:0;width:1300px;height:900px;border:0;opacity:0;pointer-events:none";
      const timer = setTimeout(() => { f.remove(); reject(new Error("page took too long to load")); }, 45000);
      f.onload = () => { clearTimeout(timer); resolve(f); };
      f.src = url; document.body.appendChild(f);
    });
  }

  async function readProduct(model) {
    let url = await productUrl(model);
    const probe = await fetch(url, { cache: "no-store" });
    if (!probe.ok) throw new Error("not on the website \u2013 it may have been removed, or check the model code");
    url = probe.url;
    const f = await loadFrame(url);
    try {
      const doc = f.contentDocument, W = ms => sleep(ms);
      if (!doc) throw new Error("the product page could not be opened");
      const btns = () => [...doc.querySelectorAll("button")];
      let kfB = null;
      for (let i = 0; i < 40 && !kfB; i++) { kfB = btns().find(x => x.innerText.trim() === "Key Features"); if (!kfB) await W(250); }
      if (!kfB) throw new Error("product page not found on the website");
      await W(400);
      const btn = t => btns().find(x => x.innerText.trim().startsWith(t));
      let box = kfB; for (let i = 0; i < 4 && box; i++) box = box.parentElement;
      btn("Show More")?.click(); await W(600);
      const skip = new Set(["Key Features", "Dimensions & Specs", "Show Less", "Show More"]);
      const lines = t => t.split("Warranty Period")[0].split("\n").map(s => s.trim()).filter(s => s && !skip.has(s) && !/^Key Features\s*Dimensions/.test(s));
      const kf = lines(box.innerText);
      btn("Dimensions & Specs")?.click(); await W(900);
      const ds = lines(box.innerText);
      let capacity = null; const ci = ds.indexOf("Capacity");
      if (ci >= 0 && ds[ci + 1] && ds[ci + 2]) { const hd = ds[ci + 1].split("\t"), v = ds[ci + 2].split("\t"); capacity = hd.map((x, k) => x.trim() + " " + (v[k] || "").trim()).join("  |  "); }
      const dims = []; const di = ds.findIndex(s => s === "Dimensions & Weight" || s === "Dimensions");
      const rest = di >= 0 ? ds.slice(di + 1) : [];
      for (let k = 0; k < rest.length; k++) {
        let lab, val; const p = rest[k].split("\t").map(s => s.trim()).filter(Boolean);
        if (p.length >= 2) { lab = p[0]; val = p.slice(1).join(" "); }
        else if (k + 1 < rest.length && /^(W\s?[\d.]|[\d.]+\s*Kg)/i.test(rest[k + 1])) { lab = rest[k]; val = rest[k + 1]; k++; }
        else continue;
        const ex = dims.find(d => d[0] === lab);
        if (ex && /^[\d.]+\s*Kg$/i.test(val)) ex[1] += " | " + val; else dims.push([lab, val]);
      }
      const body = doc.body.innerText;
      const mi = body.search(/MODEL\s/), ki = body.indexOf("Key Features", mi);
      const head = body.slice(mi, ki > mi ? ki : mi + 250).replace(/\s+/g, " ").trim();
      const num = re => { const m = head.match(re); return m ? Number(m[1].replace(/,/g, "")) : null; };
      const crumb = (body.match(/Home\s*\/\s*([^\n]+)/) || [])[1] || "";
      const nm = doc.title.replace(/\s*-\s*Dragon Electronics\s*$/, "");
      const pageModel = (head.match(/MODEL\s*(\S+)/) || [])[1] || model;
      const data = {
        model: pageModel, name: nm, url,
        retail: num(/Retail\s*Rs\s*([\d,]+)/) || num(/Rs\s*([\d,]+)/), promo: num(/PROMO\s*Rs\s*([\d,]+)/),
        promoValid: (head.match(/Valid Until (.+?)(?: or While Stock Lasts|$)/) || [])[1] || null,
        noStock: /No Stock|Out of Stock/i.test(head),
        category: crumb.split("/").map(s => s.trim()).slice(0, -1)[0] || "Products",
        capacity, dims,
        features: kf.filter(s => !s.startsWith("*") && !/^MODEL\s/.test(s) && s !== nm),
        notes: kf.filter(s => s.startsWith("*")),
        image: doc.querySelector('meta[property="og:image"]')?.content || null
      };
      if (pageModel.replace(/-/g, "").toUpperCase() !== model.replace(/-/g, "").toUpperCase()) throw new Error(`the page found is for ${pageModel}, not ${model}`);
      if (!data.retail) throw new Error("price not found on the page");
      return data;
    } finally { f.remove(); }
  }

  async function photo(src) {
    const blob = await (await fetch(src)).blob();
    const bmp = await createImageBitmap(blob);
    const sc = Math.min(1, 1000 / Math.max(bmp.width, bmp.height));
    const c = document.createElement("canvas"); c.width = Math.round(bmp.width * sc); c.height = Math.round(bmp.height * sc);
    const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(bmp, 0, 0, c.width, c.height);
    const d = g.getImageData(0, 0, c.width, c.height).data; let x0 = c.width, y0 = c.height, x1 = 0, y1 = 0;
    for (let y = 0; y < c.height; y += 2) for (let x = 0; x < c.width; x += 2) {
      const i = (y * c.width + x) * 4;
      if (d[i] < 240 || d[i + 1] < 240 || d[i + 2] < 240) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    }
    if (x1 <= x0) { x0 = 0; y0 = 0; x1 = c.width - 1; y1 = c.height - 1; }
    const p = 8; x0 = Math.max(0, x0 - p); y0 = Math.max(0, y0 - p); x1 = Math.min(c.width, x1 + p); y1 = Math.min(c.height, y1 + p);
    const w = x1 - x0, hh = y1 - y0, c2 = document.createElement("canvas"); c2.width = w; c2.height = hh;
    c2.getContext("2d").drawImage(c, x0, y0, w, hh, 0, 0, w, hh);
    const jb = await new Promise(r => c2.toBlob(r, "image/jpeg", 0.92));
    return { bytes: new Uint8Array(await jb.arrayBuffer()), width: w, height: hh };
  }

  // ------------------------------------------------------------------ template (kept in this browser only, never online)
  const toB64 = buf => { const u = new Uint8Array(buf); let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const fromB64 = b64 => { const s = atob(b64), u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u.buffer; };
  let template = null;
  const tplStatus = () => {
    const saved = store.get("template_info", null);
    $("dq-tplstat").innerHTML = saved ? `\u2713 Saved on this computer: ${h(saved.name)} (${h(saved.date)}). Choose a file only to replace it.` : "<span style='color:#c8102e'>One-time step: choose your quotation template file (Quotation Template 2026 05 13.xlsx).</span>";
  };
  tplStatus();
  $("dq-tpl").addEventListener("change", async e => {
    const file = e.target.files[0]; if (!file) return;
    try {
      const buf = await file.arrayBuffer(), files = await DQCore.unzip(buf);
      const bad = DQCore.checkTemplate(files);
      if (bad.length) throw new Error("its layout is different from 'Quotation Template 2026 05 13' (" + bad.slice(0, 2).join("; ") + ").");
      store.set("template", toB64(buf)); store.set("template_info", { name: file.name, date: isoLocal(new Date()) });
      if (!store.get("template", null)) throw new Error("the browser refused to save it (private window?)");
      template = buf; tplStatus(); log("\u2713 Template saved on this computer.", "ok");
    } catch (err) { log("\u2717 Template not saved: " + h(err.message), "err"); }
    e.target.value = "";
  });
  async function getTemplate() {
    if (template) return template;
    const b64 = store.get("template", null);
    if (b64) return (template = fromB64(b64));
    if (BASE) { const r = await fetch(BASE + "template.xlsx", { cache: "no-store" }).catch(() => null); if (r && r.ok) return (template = await r.arrayBuffer()); }
    throw new Error("Choose your quotation template file first (Quotation section, bottom of the form).");
  }

  // ------------------------------------------------------------------ generate
  $("dq-go").onclick = async () => {
    const go = $("dq-go"); $("dq-log").innerHTML = "";
    const to = $("dq-to").value.trim();
    const items = [...$("dq-rows").querySelectorAll(".prow")].map(row => {
      const model = row.querySelector(".pm").value.trim().toUpperCase(), qtxt = row.querySelector(".pq").value.trim();
      const disc = row.querySelector(".pdc").checked, ptxt = row.querySelector(".pp").value.trim();
      if (!model) return null;
      if (!/^[A-Z0-9.\-\/]+$/.test(model)) return { bad: `${model} (model code)` };
      const qty = qtxt === "" ? 1 : Math.round(+qtxt);
      if (!(qty >= 1)) return { bad: `${model}: quantity "${qtxt}"` };
      const it = { model, qty };
      if (disc) {
        const pc = ptxt.match(/^-?\s*(\d+(?:\.\d+)?)\s*%$/), num = ptxt.replace(/^rs\.?\s*/i, "").replace(/[,\s]/g, "");
        if (pc && +pc[1] > 0 && +pc[1] < 100) it.percent = +pc[1];
        else if (/^\d+(\.\d+)?$/.test(num) && +num > 0) it.custom = +num;
        else return { bad: `${model}: enter the discounted price${ptxt ? ` ("${ptxt}" is not a price)` : ""}` };
      }
      return it;
    }).filter(Boolean);
    if (!to) return log("Enter the customer name.", "err");
    if (!items.length) return log("Enter at least one model.", "err");
    const bad = items.filter(i => i.bad); if (bad.length) return log("Could not read: " + bad.map(b => h(b.bad)).join(", "), "err");
    go.disabled = true; go.textContent = "Working\u2026";
    try {
      await getTemplate();
      const products = [], images = {}, problems = [];
      for (const it of items) {
        const line = log(`Looking up <b>${h(it.model)}</b>\u2026`);
        try {
          const d = await readProduct(it.model);
          const own = it.custom != null ? it.custom : it.percent != null ? Math.round(d.retail * (1 - it.percent / 100) * 100) / 100 : null;
          const usePromo = own == null && $("dq-price").value === "promo" && d.promo;
          const price = own != null ? own : usePromo ? d.promo : d.retail;
          if (own != null && own >= d.retail) problems.push(`${d.model}: your price ${fmt(own)} is not below the website price ${fmt(d.retail)}`);
          products.push({ category: d.category, model: d.model, name: d.name, qty: it.qty, price, normal: d.retail, url: d.url, features: d.features, capacity: d.capacity, dims: d.dims, notes: d.notes, _d: d, _promo: usePromo, _own: own != null });
          try { if (d.image) images[d.model] = await photo(d.image); } catch (e) { problems.push(`${d.model}: photo could not be read`); }
          line.className = "ok"; line.innerHTML = `\u2713 ${h(d.model)} \u2013 ${h(d.name)} \u2013 ${fmt(price)}${own != null ? ` <span class='warn'>(your price; website ${fmt(d.retail)})</span>` : ""}${d.noStock ? " <span class='warn'>(No Stock on website)</span>" : ""}`;
          if (!d.dims.length) problems.push(`${d.model}: no dimensions on the website`);
        } catch (e) {
          line.className = "err"; line.innerHTML = `\u2717 ${h(it.model)}: ${h(e.message)}`; problems.push(`${it.model}: ${e.message}`);
        }
      }
      if (!products.length) throw new Error("No product could be read, nothing to build.");
      const t = {};
      const pv = $("dq-pay").value, pn = +$("dq-payn").value || 0;
      t.payment = pv === "custom" ? { type: "custom", text: $("dq-paytxt").value || "As agreed" } : pv === "days" ? { type: "days", days: pn } : pv === "full" ? { type: "full" } : { type: pv, percent: pn };
      const dn = Math.min(30, Math.max(1, Math.round(+$("dq-delivn").value) || 30));
      t.delivery = $("dq-deliv").value === "days" ? `Within ${dn} days as from invoice date` : "Ex-stock, subject unsold"; t.transport_fee = $("dq-transport").checked;
      {
        const d = new Date($("dq-until").value + "T12:00:00");
        if (d.getDay() === 6) d.setDate(d.getDate() - 1); if (d.getDay() === 0) d.setDate(d.getDate() - 2);
        if (isoLocal(d) !== $("dq-until").value) log(`Validity moved to Friday ${isoLocal(d)} (the date fell on a weekend).`, "warn");
        t.validity = { until: isoLocal(d) };
      }
      t.installation = $("dq-install").checked;
      const spec = {
        customer: { to, company: $("dq-company").checked, address: [$("dq-addr1").value.trim(), $("dq-addr2").value.trim()].filter(Boolean), attention: $("dq-att").value.trim(), tel: $("dq-tel").value.trim(), mobile: $("dq-mob").value.trim(), email: $("dq-hasemail").checked ? $("dq-email").value.trim() : "" },
        quote_no: $("dq-no").value.trim(), date: $("dq-date").value, terms: t, layout: $("dq-disc").checked ? "discount" : "plain",
        warranties: [...new Set(products.map(DQCore.warrantyFor).filter(Boolean))], products
      };
      products.forEach(p => { if (!DQCore.warrantyFor(p)) problems.push(`${p.model}: no matching warranty line \u2013 check the Terms`); });
      log("Building the Excel file\u2026");
      const res = await DQCore.buildQuote(template, spec, images);
      const fname = `Quotation_${spec.quote_no.replace(/\s*\|\s*/g, "-").replace(/[^\w\-]+/g, "_")}_${to.replace(/^(mr|mrs|ms|miss|dr)\.?\s+/i, "").replace(/[^\w]+/g, "_")}.xlsx`;
      const url = URL.createObjectURL(new Blob([res.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const a = document.createElement("a"); a.href = url; a.download = fname; document.body.appendChild(a); a.click(); a.remove();
      const counterMatch = spec.quote_no.match(/(\d+)\s*$/);
      if (counterMatch && +counterMatch[1] > (+$("dq-counter").value || 0)) { store.set("counter", +counterMatch[1]); $("dq-counter").value = +counterMatch[1]; }
      const hist = store.get("history", []); hist.push({ date: spec.date, no: spec.quote_no, customer: to, total: res.total }); store.set("history", hist.slice(-200)); renderHist();
      const promoRows = products.filter(p => !p._promo && !p._own && p._d.promo);
      let html = `<div class="hint">Prices incl. VAT</div><table><tr><th>Model</th><th class="n">Qty</th><th class="n">Price</th><th class="n">Amount</th></tr>${products.map(p => `<tr><td>${h(p.model)}</td><td class="n">${p.qty}</td><td class="n">${fmt(p.price)}</td><td class="n">${fmt(p.qty * p.price)}</td></tr>`).join("")}<tr><th colspan="3">Total incl. VAT</th><th class="n">${fmt(res.total)}</th></tr></table>`;
      if (promoRows.length) html += `<div class="warn" style="margin-top:6px">Promo prices available: ${promoRows.map(p => `${h(p.model)} ${fmt(p._d.promo)}${p._d.promoValid ? " until " + h(p._d.promoValid) : ""}`).join("; ")}.</div>`;
      [...problems, ...res.warnings].forEach(w => { html += `<div class="warn">\u26a0 ${h(w)}</div>`; });
      html += `<a class="dl" href="${url}" download="${h(fname)}">\u2b07 Download ${h(fname)}</a><div class="hint">For the PDF: open the file in Excel \u2192 File \u2192 Save as \u2192 PDF.</div>`;
      log(html);
      noEdited = false; suggestNo();
    } catch (e) {
      log("\u2717 " + h(e.message || e), "err");
    } finally { go.disabled = false; go.textContent = "Generate quotation"; }
  };
})();
