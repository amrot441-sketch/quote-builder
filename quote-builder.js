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
    attention: 14, tel: 15, mobile: 16, email: 17, sigRows: [92, 93, 94], sigName: 95, sigLine: 96,
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
    // extra cell styles derived from the template's own ones (alignment / number format / font)
    let stylesXml = get("xl/styles.xml");
    const xfList = () => [...stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/)[1].matchAll(/<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(m => m[0]);
    const derived = new Map();
    const deriveStyle = (base, o) => {
      const key = base + JSON.stringify(o);
      if (derived.has(key)) return derived.get(key);
      const list = xfList(); let x = list[base];
      if (!x) return base;
      if (o.fontId != null) x = x.replace(/fontId="\d+"/, `fontId="${o.fontId}"`);
      if (o.fillId != null) x = x.replace(/fillId="\d+"/, `fillId="${o.fillId}"`).replace(/^<xf (?![^>]*applyFill)/, '<xf applyFill="1" ');
      if (o.numFmtId != null) x = x.replace(/numFmtId="\d+"/, `numFmtId="${o.numFmtId}"`).replace(/^<xf (?![^>]*applyNumberFormat)/, '<xf applyNumberFormat="1" ');
      if (o.horizontal) {
        if (/<alignment /.test(x)) x = x.replace(/<alignment ([^>]*?)\/>/, (m, at) => `<alignment horizontal="${o.horizontal}" ${at.replace(/horizontal="[^"]*"\s*/, "")}/>`);
        else x = x.replace(/\/>$/, `><alignment horizontal="${o.horizontal}"/></xf>`);
      }
      const idx = list.length;
      stylesXml = stylesXml.replace(/<\/cellXfs>/, x + "</cellXfs>").replace(/<cellXfs count="\d+"/, `<cellXfs count="${idx + 1}"`);
      derived.set(key, idx); return idx;
    };
    const solidFill = rgb => {
      const fills = [...stylesXml.match(/<fills[^>]*>([\s\S]*?)<\/fills>/)[1].matchAll(/<fill>[\s\S]*?<\/fill>|<fill\/>/g)].map(m => m[0]);
      const ex = fills.findIndex(f => /patternType="solid"/.test(f) && f.includes(`rgb="${rgb}"`));
      if (ex >= 0) return ex;
      stylesXml = stylesXml.replace(/<\/fills>/, `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill></fills>`).replace(/<fills count="\d+"/, `<fills count="${fills.length + 1}"`);
      return fills.length;
    };
    const fontIdOf = base => +((xfList()[base] || "").match(/fontId="(\d+)"/) || [0, 0])[1];
    // a bold version of a style's font (re-uses an identical font when the template has one)
    const boldFont = base => {
      const fid = fontIdOf(base);
      const fonts = [...stylesXml.match(/<fonts[^>]*>([\s\S]*?)<\/fonts>/)[1].matchAll(/<font>[\s\S]*?<\/font>|<font\/>/g)].map(m => m[0]);
      const f = fonts[fid] || ""; if (/<b\/>/.test(f)) return fid;
      const fb = f.replace("<font>", "<font><b/>"), ex = fonts.indexOf(fb);
      if (ex >= 0) return ex;
      stylesXml = stylesXml.replace(/<\/fonts>/, fb + "</fonts>").replace(/<fonts count="\d+"/, `<fonts count="${fonts.length + 1}"`);
      return fonts.length;
    };
    // a bold red version of a style's font (promo / discount prices)
    const boldRedFont = base => {
      const fid = fontIdOf(base);
      const fonts = [...stylesXml.match(/<fonts[^>]*>([\s\S]*?)<\/fonts>/)[1].matchAll(/<font>[\s\S]*?<\/font>|<font\/>/g)].map(m => m[0]);
      let f = (fonts[fid] || "<font></font>").replace(/<b\/>/g, "").replace(/<color [^>]*\/>/g, "");
      f = f.replace("<font>", '<font><b/>').replace(/(<sz [^>]*\/>)/, '$1<color rgb="FFFF0000"/>');
      if (!/FFFF0000/.test(f)) f = f.replace("<font><b/>", '<font><b/><color rgb="FFFF0000"/>');
      const ex = fonts.indexOf(f); if (ex >= 0) return ex;
      stylesXml = stylesXml.replace(/<\/fonts>/, f + "</fonts>").replace(/<fonts count="\d+"/, `<fonts count="${fonts.length + 1}"`);
      return fonts.length;
    };
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
    // company VAT / BRN go on the two free lines under the address (the panel has two address lines)
    const ids = isCompany ? [["VAT:", cust.vat], ["BRN:", cust.brn]].filter(x => x[1] && String(x[1]).trim()) : [];
    if (ids.length && addr.length <= 2) {
      [12, 13].forEach((r, i) => { if (ids[i]) { edit(r, "B", ids[i][0]); edit(r, "C", String(ids[i][1]).trim()); } else edit(r, "C", null); });
    } else { edit(12, "C", addr[2] || null); edit(13, "C", addr.slice(3).join(", ") || null); }
    if (isCompany) edit(R.attention, "C", cust.attention || null);
    else DELETE.add(R.attention);
    if (cust.tel) edit(R.tel, "C", cust.tel); else DELETE.add(R.tel);
    if (cust.mobile) edit(R.mobile, "C", cust.mobile); else DELETE.add(R.mobile);      // empty contact lines are left out
    if (cust.email) edit(R.email, "C", cust.email); else DELETE.add(R.email);
    edit(10, "J", spec.quote_no || null);
    // signature: name + "Title | branch phone | branch email"; the template's signature image only for its own person
    const tplCell = (r, col) => {
      const c = (rowOf(s, r).match(new RegExp(`<c r="${col}${r}"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`)) || [""])[0];
      const v = (c.match(/<v>([^<]*)<\/v>/) || [])[1];
      if (/t="s"/.test(c) && v != null) return ((sst.match(/<si>[\s\S]*?<\/si>/g) || [])[+v] || "").replace(/<[^>]+>/g, "").replace(/&amp;/g, "&").trim();
      return ((c.match(/<t[^>]*>([^<]*)<\/t>/) || [])[1] || "").trim();
    };
    const tplSigName = tplCell(R.sigName, "B"), tplSigLine = tplCell(R.sigLine, "B");
    const signer = spec.signer && spec.signer.name ? spec.signer : null;
    const nm = x => String(x || "").toUpperCase().replace(/\s+/g, " ").trim();
    const hasSigImage = !signer || nm(signer.name) === nm(tplSigName);
    if (signer) {
      const suffix = tplSigLine.includes("|") ? tplSigLine.slice(tplSigLine.indexOf("|")) : "";
      edit(R.sigName, "B", signer.name);
      edit(R.sigLine, "B", signer.line || ((signer.title || "") + (suffix ? " " + suffix : "")).trim());
    }
    if (!hasSigImage) R.sigRows.forEach(r => DELETE.add(r));      // no image: name follows "Yours sincerely," directly
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
    edit(R.validUntil, "C", `Strictly valid until ${pad2(vd.getUTCDate())}-${MONTHS[vd.getUTCMonth()]}-${String(vd.getUTCFullYear()).slice(-2)}`);   // DD-MMM-YY like the quote date
    // totals: only "Total Amount with VAT" (prices are entered including VAT)
    DELETE.add(R.totals[0]); DELETE.add(R.totals[1]);
    // installation, stainless note
    if (t.installation === false) DELETE.add(R.install); else SHOW.add(R.install);
    // stainless steel note: only when the quote has an ARISTON (incl. HOTPOINT ARISTON) or TURBOAIR product
    const hasSS = t.stainless_note != null ? !!t.stainless_note
      : spec.products.some(p => /ariston|turbo\s?air/i.test(`${p.brand || ""} ${p.name || ""}`));
    if (!hasSS) DELETE.add(R.stainless);
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
    // every photo gets the same visual size (same area, about 13 cm2), sized from its own shape:
    // wide products come out shorter, tall ones narrower. Not linked to the length of the description.
    const CM = 360000, PHOTO_AREA = 13 * CM * CM, PHOTO_MAX_H = 5.0 * CM, PHOTO_MAX_W = 5.6 * CM, ROW_EMU_P = Math.floor(14.4 * 12700);
    const PHOTO_GAP_ROWS = 1;      // one empty line between the prices and the top of the photo
    const photoBox = img => {
      const aspect = img.width / img.height;
      let w = Math.sqrt(PHOTO_AREA * aspect), h = w / aspect;
      if (h > PHOTO_MAX_H) { h = PHOTO_MAX_H; w = h * aspect; }
      if (w > PHOTO_MAX_W) { w = PHOTO_MAX_W; h = w / aspect; }
      return { w: Math.floor(w), h: Math.floor(h), aspect };
    };
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
      // a short description gets empty lines so the photo never runs into the next product
      if (images[p.model]) { const need = PHOTO_GAP_ROWS + Math.ceil((photoBox(images[p.model]).h + 20000) / ROW_EMU_P); while (items.length - 1 < need) items.push(["pad"]); }
      segs.push({ cat, bar: p.bar || cat, items });
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
      if (prevCat !== null && !atStart) lead.push(["blank"]);
      lead.push(["banner", seg.bar]);            // every product has its own pink bar
      return lead.concat(seg.items);
    };
    const hOf = items => items.reduce((a, it) => a + itemH(it), 0);
    const queue = segs.slice();
    // the total (blank row + "Total Amount with VAT") must sit on the page of the last product
    let TOT = 0; for (let r = R.blankBeforeTotals; r <= R.totals[2]; r++) TOT += tH(r);
    const need = items => hOf(items) + (queue.length === 1 ? TOT : 0);
    while (queue.length) {
      let idx = 0, items = withLead(queue[0], pageStart);
      if (used + need(items) > CAP) {
        const j = queue.length > 1 ? queue.findIndex((sg, k) => k > 0 && used + hOf(withLead(sg, pageStart)) <= CAP) : -1;
        if (j > 0) { idx = j; items = withLead(queue[j], pageStart); }
        else if (need(withLead(queue[0], true)) <= CAP - TITLE - HEAD) {
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
    // quotation format (Nov 2026): category bar bold; model + description left; qty, prices, amounts and total centred
    ST.banner = deriveStyle(18, { fontId: boldFont(18) });
    ST.model = deriveStyle(19, { horizontal: "left" });
    ST.desc = deriveStyle(20, { horizontal: "left" });
    ST.qty = deriveStyle(25, { horizontal: "center" });
    ST.money = deriveStyle(45, { numFmtId: 4, horizontal: "center" });
    ST.moneyRed = deriveStyle(ST.money, { fontId: boldRedFont(ST.money) });     // promo / discount price and its amount
    // "No Stock" label: yellow background, bold black, between the features and the photo
    ST.stock = deriveStyle(7, { fillId: solidFill("FFFFFF00"), fontId: boldFont(7), horizontal: "center" });
    const STOCK_TEXT = spec.no_stock_text || "NO STOCK";
    let stockPending = [];
    const quoteDay = spec.date ? isoToDate(spec.date) : new Date();
    const stockLines = p => {
      if (!p.noStock) return [];
      const L = [STOCK_TEXT];
      if (p.noStockUntil) { const d = isoToDate(String(p.noStockUntil).slice(0, 10)); if (!isNaN(d) && d >= quoteDay) L.push("Arrival on or about", `${pad2(d.getUTCDate())}-${MONTHS[d.getUTCMonth()]}-${String(d.getUTCFullYear()).slice(-2)}`); }
      return L;
    };
    const ex = p => ({ unit: p.price, normal: p.normal || p.price });   // unit prices INCLUDING VAT
    const newRows = [], merges = [], productRows = [], links = [];
    // second price column: "Promo Unit" when any product is at its website promo price, else "Discounted Unit"
    const anyPromo = spec.products.some(p => p.isPromo);
    const sstText = i => ((sst.match(/<si>[\s\S]*?<\/si>/g) || [])[i] || "").replace(/<[^>]+>/g, "");
    const promoHead = rx => anyPromo ? rx.replace(/<c r="I(\d+)"([^>]*?)\st="s"([^>]*)>\s*<v>(\d+)<\/v>\s*<\/c>/, (m, n, a1, a2, v) =>
      /^Discounted Unit/.test(sstText(+v)) ? `<c r="I${n}"${a1} t="inlineStr"${a2}><is><t>${sstText(+v).replace(/^Discounted/, "Promo")}</t></is></c>` : m) : rx;
    const headingXml = r => (discount ? promoHead(rowOf(sDisc, r)) : "") || rowOf(s, r);
    let amountSum = 0;
    plan.forEach((item, k) => {
      const r = firstNew + k, kind = item[0];
      if (kind === "heading") {
        let rx = headingXml(item[1]).replace(/<c r="[M-Z]\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, "");
        rx = rx.replace(/<row r="\d+"/, `<row r="${r}"`).replace(/(<c r=")([A-Z]+)\d+(")/g, (m, a, col, b) => a + col + r + b);
        newRows.push(rx);
      } else if (kind === "banner") {
        const cells = cell(`B${r}`, ST.banner, { text: item[1] }) + [..."CDEFGHIJ"].map(x => cell(`${x}${r}`, ST.banner)).join("") + cell(`L${r}`, 55) + cell(`M${r}`, 56);
        newRows.push(`<row r="${r}" s="2" customFormat="1" ht="18" spans="2:13">${cells}</row>`); merges.push(`B${r}:J${r}`);
      } else if (kind === "product") {
        const p = item[1], qty = p.qty || 1, e = ex(p), amount = qty * e.unit; amountSum += amount;
        const red = !!(p.isPromo || p.isDisc || (p.normal && e.unit < e.normal - 0.001));   // promo or discount price -> bold red
        let cells;
        if (discount) cells = cell(`B${r}`, ST.model, { text: p.model }) + cell(`C${r}`, ST.desc, { text: p.name }) + cell(`D${r}`, 21) + cell(`E${r}`, 22) + cell(`F${r}`, 22)
          + cell(`G${r}`, ST.qty, { num: qty }) + cell(`H${r}`, ST.money, { num: e.normal }) + cell(`I${r}`, red ? ST.moneyRed : ST.money, { num: e.unit }) + cell(`J${r}`, red ? ST.moneyRed : ST.money, { num: amount, formula: `I${r}*G${r}` }) + cell(`L${r}`, ST.discProdL);
        else cells = cell(`B${r}`, ST.model, { text: p.model }) + cell(`C${r}`, ST.desc, { text: p.name }) + cell(`D${r}`, 21) + cell(`E${r}`, 22) + cell(`F${r}`, 23) + cell(`G${r}`, 24)
          + cell(`H${r}`, ST.qty, { num: qty }) + cell(`I${r}`, red ? ST.moneyRed : ST.money, { num: e.unit }) + cell(`J${r}`, red ? ST.moneyRed : ST.money, { num: amount, formula: `H${r}*I${r}` }) + cell(`L${r}`, 57);
        newRows.push(`<row r="${r}" s="${discount ? 4 : 3}" customFormat="1" spans="2:12">${cells}</row>`);
        stockPending = stockLines(p);
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
        let cells = (kind === "blank" || kind === "pad") ? cell(`B${r}`, 29) + cell(`C${r}`, 7) + cell(`D${r}`, 7) + tailCells
          : cell(`B${r}`, 29) + cell(`C${r}`, cStyle, cOpt) + cell(`D${r}`, dStyle, dOpt) + tailCells;
        if (stockPending.length && kind !== "blank") {      // lines under a No Stock product: label in F:G ("NO STOCK", "until \u2026")
          cells = cells.replace(new RegExp(`<c r="F${r}"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`), cell(`F${r}`, ST.stock, { text: stockPending.shift() }))
                       .replace(new RegExp(`<c r="G${r}"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`), cell(`G${r}`, ST.stock));
          merges.push(`F${r}:G${r}`);
        }
        if (kind === "blank") stockPending = [];
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
      if (discount && R.headings.includes(r)) rx = promoHead(rowOf(sDisc, r)) || rx;          // headings with Normal + Discounted price
      if (SHOW.has(r)) rx = rx.replace(' hidden="1"', "");
      if (r === R.transport) rx = rx.replace(new RegExp(`(<c r="[C-G]${R.transport}" s=")44(")`, "g"), (m, a, b) => a + "7" + b);   // drop the yellow reminder highlight
      if (CLEAR_B.has(r)) rx = setCellInRow(rx, `B${r}`, "", "");
      for (const [col, txt, o] of EDIT.get(r) || []) {
        if (o.num != null) rx = setCellInRow(rx, `${col}${r}`, `<v>${o.num}</v>`, "");
        else if (txt) rx = setCellInRow(rx, `${col}${r}`, `<v>${sidx(txt)}</v>`, ' t="s"');
        else rx = setCellInRow(rx, `${col}${r}`, "", "");
      }
      if (r === R.totals[2]) {
        rx = rx.replace(/<f>[^<]*<\/f>/, `<f>SUM(J${R.headings[0]}:J${R.blankBeforeTotals})</f>`);
        rx = rx.replace(new RegExp(`(<c r="J${r}"[^>]*?\\ss=")(\\d+)(")`), (m, a, st, b) => a + deriveStyle(+st, { numFmtId: 4, horizontal: "center" }) + b);
      }
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
      if (!hasSigImage && r0 + 1 >= R.sigRows[0] && r0 + 1 <= R.sigRows[R.sigRows.length - 1]) continue;   // signature image
      a = a.replace(/<xdr:row>(\d+)<\/xdr:row>/g, (m, n) => { const r1 = +n + 1; let nr = r1; while (DELETE.has(nr)) nr++; return `<xdr:row>${newrow(nr) - 1}</xdr:row>`; });
      body.push(a);
    }
    const EMU_PX = 9525, ROW_EMU = Math.floor(14.4 * 12700);
    const colEmu = { 6: 72 * EMU_PX, 7: 114 * EMU_PX, 8: 114 * EMU_PX, 9: 113 * EMU_PX };
    const regionW = colEmu[6] + colEmu[7] + colEmu[8] + colEmu[9];
    const xAnchor = (x, col = 6) => { while (col < 9 && x >= colEmu[col]) { x -= colEmu[col]; col++; } return [col, Math.floor(x)]; };
    productRows.forEach(([prow, nBody, p], idx) => {
      const img = images[p.model];
      if (!img) { warnings.push(`No photo for ${p.model}`); return; }
      const i = idx + 1, aspect = img.width / img.height;
      let { w, h } = photoBox(img);      // same box for every product
      if (h <= 0) return;
      const c0 = 7, rw = regionW - colEmu[6];   // every photo centred on the same line (columns H..J), clear of the No Stock label
      if (w > rw - 20000) { w = rw - 20000; h = Math.floor(w / aspect); }
      const x = Math.floor((rw - w) / 2), [fc, fo] = xAnchor(x, c0), [tc, to] = xAnchor(x + w, c0);
      const rowsH = Math.floor((h + 20000) / ROW_EMU), rem = h + 20000 - rowsH * ROW_EMU;
      const name = `quote_photo_${pad2(i)}.jpeg`, rid = `rIdQ${i}`;
      files.set(`xl/media/${name}`, img.bytes);
      rels = rels.replace("</Relationships>", `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${name}"/></Relationships>`);
      body.push(`<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${fc}</xdr:col><xdr:colOff>${fo}</xdr:colOff><xdr:row>${prow + PHOTO_GAP_ROWS}</xdr:row><xdr:rowOff>20000</xdr:rowOff></xdr:from>`
        + `<xdr:to><xdr:col>${tc}</xdr:col><xdr:colOff>${to}</xdr:colOff><xdr:row>${prow + PHOTO_GAP_ROWS + rowsH}</xdr:row><xdr:rowOff>${rem}</xdr:rowOff></xdr:to>`
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

    put("xl/styles.xml", stylesXml);
    return { bytes: await zip(files), warnings, total: amountSum, totalExVat: amountSum / VAT };
  }

  // ------------------------------------------------------------------ business rules
  function warrantyFor(p) {
    const n = (p.name + " " + p.category).toLowerCase(), brand = String(p.brand || "").toLowerCase(), nb = brand + " " + n;
    // other brands first (an ARISTON dishwasher must not get the LG line)
    if (nb.includes("ariston")) return "ARISTON";
    if (/turbo\s?air/.test(nb)) return "TURBOAIR";
    if (nb.includes("ufesa")) return "UFESA";
    if (/\bde fan\b|dragon electronics/.test(nb) || (brand && brand !== "lg" && n.includes("fan"))) return "DE Fan";
    if (brand && brand !== "lg") return null;
    if (/bracket|wall mount/.test(n)) return null;
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

  const fmtDate = d => `${pad2(d.getDate())}-${MONTHS[d.getMonth()]}-${d.getFullYear()}`;
  const API = { buildQuote, checkTemplate, warrantyFor, initials, looksLikeCompany, fmtDate, unzip, zip };
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
  #dq-panel .prow .pm{flex:1 1 150px;text-transform:uppercase}#dq-panel .prow .pm::placeholder{text-transform:none}#dq-panel .prow .pq{flex:0 0 58px;padding:6px 4px;text-align:center}
  #dq-panel .prow .pd{display:flex;align-items:center;gap:4px;margin:0;font-size:12px;color:#333;white-space:nowrap;cursor:pointer}
  #dq-panel .prow .pp{order:5;flex:1 1 100%;margin-left:0;border-color:#c8102e}
  #dq-panel .prow .px{background:none;border:0;color:#999;font-size:18px;cursor:pointer;padding:0 2px;line-height:1}
  #dq-panel .prow .px:hover{color:#c8102e}
  #dq-addrow{margin-top:8px;padding:5px 10px;border:1px dashed #c8102e;background:#fff;color:#c8102e;border-radius:4px;cursor:pointer;font-size:12px}
  #dq-panel select.need{border-color:#c8102e;box-shadow:0 0 0 2px rgba(200,16,46,.25)}
  #dq-panel details{margin-top:12px}#dq-panel summary{cursor:pointer;color:#c8102e}
  #dq-panel .sg{order:4;flex:1 1 100%;border:1px solid #c8102e;border-radius:4px;background:#fff;max-height:280px;overflow:auto;display:none;box-shadow:0 4px 12px rgba(0,0,0,.12)}
  #dq-panel .sgi{display:flex;align-items:center;gap:8px;padding:5px 8px;cursor:pointer;border-bottom:1px solid #f0f0f0;font-size:12px}
  #dq-panel .sgi:last-child{border-bottom:0}
  #dq-panel .sgi.act,#dq-panel .sgi:hover{background:#fdecef}
  #dq-panel .sgi img,#dq-panel .sgi .ph{width:34px;height:34px;object-fit:contain;flex:0 0 34px;background:#f6f6f6;border-radius:3px}
  #dq-panel .sgi .tx{flex:1;min-width:0}#dq-panel .sgi .tx b{display:block;color:#222}
  #dq-panel .sgi .tx span{display:block;color:#666;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  #dq-panel .sgi .pr{white-space:nowrap;color:#222;text-align:right}#dq-panel .sgi .pr i{display:block;color:#c8102e;font-style:normal}
  #dq-panel .sgnone{padding:6px 8px;color:#888;font-size:12px}
  #dq-cat{display:flex;align-items:center;gap:8px;margin:0 0 8px;padding:6px 8px;background:#f7f7f7;border-radius:4px;font-size:12px;color:#444}
  #dq-cat .ct{flex:1}#dq-cat .ct.need{color:#c8102e}
  #dq-catupd{padding:4px 10px;border:1px solid #c8102e;background:#fff;color:#c8102e;border-radius:4px;cursor:pointer;font-size:12px;white-space:nowrap}
  #dq-catupd:disabled{color:#999;border-color:#bbb;cursor:wait}
  #dq-panel .sgi .tag{display:inline-block;margin-left:4px;padding:0 4px;border-radius:3px;background:#eee;color:#555;font-size:10px;font-weight:normal}
  #dq-catbar{height:4px;background:#eee;border-radius:2px;margin:-4px 0 8px;display:none}#dq-catbar div{height:100%;width:0;background:#c8102e;border-radius:2px;transition:width .2s}`;
  const st = document.createElement("style"); st.textContent = css; document.head.appendChild(st);

  const today = new Date();
  const until = new Date(today); until.setDate(until.getDate() + 5);          // default validity: 5 days
  if (until.getDay() === 6) until.setDate(until.getDate() - 1); else if (until.getDay() === 0) until.setDate(until.getDate() - 2);
  const panel = document.createElement("div"); panel.id = "dq-panel";
  panel.innerHTML = `
  <header><b>Quote Builder \u00b7 Dragon Electronics</b><button id="dq-close" title="Close">\u00d7</button></header>
  <div id="dq-body">
    <fieldset><legend>Customer</legend>
      <label>To (name or company) *</label><input type="text" id="dq-to">
      <label>Customer type</label>
      <select id="dq-ctype"><option value="individual">Individual</option><option value="company">Company</option></select>
      <div class="hint">Company adds the Attention line and writes the name in CAPITALS.</div>
      <div id="dq-cobox" style="display:none">
        <label>Attention (Title First name SURNAME)</label><input type="text" id="dq-att">
        <div class="row"><div><label>VAT No</label><input type="text" id="dq-vat"></div><div><label>BRN</label><input type="text" id="dq-brn"></div></div>
        <div class="hint">Leave VAT or BRN empty if the company has none; that line is then left out.</div>
      </div>
      <label>Address</label><input type="text" id="dq-addr1"><input type="text" id="dq-addr2" style="margin-top:4px">
      <div class="row"><div><label>Tel</label><input type="text" id="dq-tel"></div><div><label>Mobile</label><input type="text" id="dq-mob"></div></div>
      <label>Email</label><input type="text" id="dq-email">
      <div class="hint">Leave Tel, Mobile or Email empty if the customer has none; that line is then left out.</div>
    </fieldset>
    <fieldset><legend>Products</legend>
      <div id="dq-cat"><div class="ct" id="dq-catstat"></div><button type="button" id="dq-catupd">Update product list</button></div>
      <div id="dq-catbar"><div></div></div>
      <div id="dq-pos" class="hint" style="margin:-2px 0 8px">Full range incl. products not on the website: <a href="#" id="dq-poslink" style="color:#c8102e">load the POS product list (FEATURES.TXT)</a><span id="dq-posstat"></span>
        <input type="file" id="dq-posfile" accept=".txt,.TXT,text/plain" style="display:none"></div>
      <div class="hint">Type part of a model or name (e.g. <b>X31</b> or <b>850L</b>) and pick from the list. Quantity empty = 1. Tick <b>Discount</b> to type your own price, incl. VAT.</div>
      <div id="dq-rows"></div>
      <div style="display:flex;align-items:center;gap:10px;margin-top:8px"><button type="button" id="dq-addrow" style="margin-top:0">+ Add product</button>
        <span id="dq-pcount" style="font-size:12px;color:#555"></span></div>
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
      <div class="hint">Written as "Strictly valid until dd-mmm-yy". A Saturday or Sunday is moved back to the Friday before. Check public holidays yourself.</div>
      <div class="chk"><input type="checkbox" id="dq-install" checked> Include the installation clause</div>
    </fieldset>
    <fieldset><legend>Quotation</legend>
      <div class="row"><div><label>Quotation No</label><input type="text" id="dq-no"></div><div><label>Date</label><input type="date" id="dq-date" value="${isoLocal(today)}"></div></div>
      <label>Signed by</label><select id="dq-signer"></select>

      <details id="dq-signedit"><summary>Edit the list of names</summary>
        <div class="hint">One person per line: <b>Name | Job title</b>. The branch phone and email are added after the title. Only the person whose name matches the template's signature gets the signature image.</div>
        <textarea id="dq-signlist" style="min-height:96px"></textarea>
        <button type="button" id="dq-signsave" style="margin-top:6px;padding:4px 10px;border:1px solid #c8102e;background:#fff;color:#c8102e;border-radius:4px;cursor:pointer;font-size:12px">Save list</button></details>
      <label>Save quotes in folder</label>
      <div style="display:flex;gap:8px;align-items:center"><span id="dq-dirstat" class="hint" style="flex:1;margin:0">Quotes go to your Downloads folder.</span>
        <button type="button" id="dq-dirbtn" style="padding:4px 10px;border:1px solid #c8102e;background:#fff;color:#c8102e;border-radius:4px;cursor:pointer;font-size:12px;white-space:nowrap">Choose folder</button></div>
      <div class="hint">DP | customer initials | running number. Last number used: <input type="number" id="dq-counter" style="width:64px;padding:2px 4px"></div>
      <label>Quotation template</label>
      <div id="dq-tplstat" class="hint"></div>
      <input type="file" id="dq-tpl" accept=".xlsx" style="font-size:12px;margin-top:4px">
    </fieldset>
    <button id="dq-go">Generate quotation</button>
    <div id="dq-log"></div>
    <details><summary>Quote history on this computer</summary><div id="dq-hist"></div></details>
    <div id="dq-ver" style="margin-top:14px;padding-top:8px;border-top:1px solid #eee;text-align:center;font-size:11px;color:#999">Quote Builder version 1.2 \u00b7 25-Nov-2026</div>
  </div>`;
  document.body.appendChild(panel);
  const $ = id => panel.querySelector("#" + id);
  $("dq-close").onclick = () => { panel.style.display = "none"; };
  $("dq-counter").value = store.get("counter", 0);

  const upd = () => {
    $("dq-paytxt").style.display = $("dq-pay").value === "custom" ? "block" : "none";
    $("dq-payn").style.display = ["days", "down_delivery", "down_week"].includes($("dq-pay").value) ? "block" : "none";
    $("dq-delivn").style.display = $("dq-delivhint").style.display = $("dq-deliv").value === "days" ? "block" : "none";
    $("dq-cobox").style.display = $("dq-ctype").value === "company" ? "block" : "none";
  };
  $("dq-deliv").addEventListener("change", () => upd());
  $("dq-delivn").addEventListener("change", () => { const n = Math.round(+$("dq-delivn").value); $("dq-delivn").value = Math.min(30, Math.max(1, n || 30)); });
  ["dq-pay"].forEach(id => $(id).addEventListener("change", () => {
    if (id === "dq-pay") $("dq-payn").value = $("dq-pay").value === "days" ? 30 : 50;
    upd();
  }));
  let noEdited = false, coEdited = false;
  $("dq-ctype").addEventListener("change", upd);

  $("dq-price").addEventListener("change", () => { if ($("dq-price").value === "promo") $("dq-disc").checked = true; });
  // product lines: model | qty | [ ] Discount -> price
  // how many products (and units) are in the quote so far
  function updCount() {
    const rows = [...$("dq-rows").querySelectorAll(".prow")].filter(r => r.querySelector(".pm").value.trim());
    const units = rows.reduce((s, r) => s + (Math.max(1, Math.round(+r.querySelector(".pq").value) || 1)), 0);
    $("dq-pcount").innerHTML = rows.length ? `<b>${rows.length}</b> product${rows.length === 1 ? "" : "s"} in this quote${units !== rows.length ? ` \u00b7 ${units} units` : ""}` : "No product yet";
  }
  function addRow(model, qty) {
    const row = document.createElement("div"); row.className = "prow";
    row.innerHTML = `<input type="text" class="pm" placeholder="Model, e.g. GC-X31FFCRB"><input type="number" class="pq" min="1" placeholder="Qty">`
      + `<label class="pd"><input type="checkbox" class="pdc"> Discount</label><input type="text" class="pp" placeholder="Discounted price Rs (incl. VAT), e.g. 180000" style="display:none">`
      + `<button type="button" class="px" title="Remove this product">\u00d7</button><div class="sg"></div>`;
    const q = s => row.querySelector(s);
    q(".pm").placeholder = "Model or name, e.g. X31 or 850L";
    // suggestions from the saved product list
    let sgItems = [], sgAct = -1;
    const sgBox = q(".sg");
    const sgHide = () => { sgBox.style.display = "none"; sgItems = []; sgAct = -1; };
    const sgPick = i => {
      const p = sgItems[i]; if (!p) return;
      q(".pm").value = p.model; sgHide(); q(".pq").focus(); updCount();
    };
    const sgShow = () => {
      const val = q(".pm").value.trim();
      if (val.length < 2 || !catalog.size) return sgHide();
      sgItems = searchCatalog(val, 8); sgAct = sgItems.length ? 0 : -1;
      if (!sgItems.length) { sgBox.innerHTML = `<div class="sgnone">No saved product matches "${h(val)}". Check the model, or update the product list.</div>`; sgBox.style.display = "block"; return; }
      sgBox.innerHTML = sgItems.map((p, i) => `<div class="sgi${i === 0 ? " act" : ""}" data-i="${i}"><span class="ph" data-id="${h(p.id)}"></span>`
        + `<div class="tx"><b>${h(p.model)}${p.offWeb ? `<span class="tag">not on website</span>` : ""}</b><span>${h(p.name)}</span></div>`
        + `<div class="pr">${fmt(p.retail)}${p.promo ? `<i>Promo ${fmt(p.promo)}</i>` : ""}</div></div>`).join("");
      sgBox.style.display = "block";
      sgBox.querySelectorAll(".sgi").forEach(el => el.addEventListener("mousedown", e => { e.preventDefault(); sgPick(+el.dataset.i); }));
      sgBox.querySelectorAll(".ph").forEach(async el => { const u = await thumbUrl(el.dataset.id); if (u && el.isConnected) { const im = document.createElement("img"); im.src = u; el.replaceWith(im); } });
    };
    const sgMove = d => {
      if (!sgItems.length) return;
      sgAct = (sgAct + d + sgItems.length) % sgItems.length;
      sgBox.querySelectorAll(".sgi").forEach((el, i) => el.classList.toggle("act", i === sgAct));
      sgBox.querySelectorAll(".sgi")[sgAct].scrollIntoView({ block: "nearest" });
    };
    q(".pm").addEventListener("input", sgShow);
    q(".pm").addEventListener("focus", sgShow);
    q(".pm").addEventListener("blur", () => setTimeout(sgHide, 150));
    q(".pm").value = model || ""; q(".pq").value = qty || "";
    q(".pdc").addEventListener("change", () => {
      q(".pp").style.display = q(".pdc").checked ? "block" : "none";
      if (q(".pdc").checked) { $("dq-disc").checked = true; q(".pp").focus(); }
    });
    q(".px").addEventListener("click", () => { row.remove(); if (!$("dq-rows").children.length) addRow(); updCount(); });
    q(".pm").addEventListener("keydown", e => {
      const open = sgBox.style.display === "block" && sgItems.length;
      if (e.key === "ArrowDown" && open) { e.preventDefault(); sgMove(1); }
      else if (e.key === "ArrowUp" && open) { e.preventDefault(); sgMove(-1); }
      else if (e.key === "Escape") sgHide();
      else if (e.key === "Enter") {
        e.preventDefault();
        if (open && sgAct >= 0) sgPick(sgAct);
        else { sgHide(); const r = addRow(); r.querySelector(".pm").focus(); }
      }
    });
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
    if ($("dq-pcount")) updCount();
    return row;
  }
  addRow(); addRow(); addRow();
  $("dq-rows").addEventListener("input", updCount); updCount();
  $("dq-addrow").addEventListener("click", () => addRow().querySelector(".pm").focus());
  // ---- people who sign the quote (list kept in Chrome on this computer)
  const DEFAULT_SIGNERS = [["Youvraj AUBEELUCK", "Shop Manager"], ["Farhaan ELLAHEE", "Sales Demonstrator"], ["Emillio MARUTHAMUTEE", "Sales Demonstrator"],
    ["Jayslay SANTUCK", "Sales Demonstrator"], ["Ourmila BALIRAM", "Sales Demonstrator"]];
  const getSigners = () => { const l = store.get("signers", null); return Array.isArray(l) && l.length ? l : DEFAULT_SIGNERS; };
  function fillSigners() {
    const list = getSigners(), cur = $("dq-signer").value;      // shared PC: no one is pre-selected, choose for each quote
    $("dq-signer").innerHTML = `<option value="">\u2013 Choose who signs this quote \u2013</option>` + list.map(([n, t]) => `<option value="${h(n)}"${n === cur ? " selected" : ""}>${h(n)} \u2013 ${h(t)}</option>`).join("");
    $("dq-signlist").value = list.map(([n, t]) => `${n} | ${t}`).join("\n");
  }
  fillSigners();
  $("dq-signer").addEventListener("change", () => $("dq-signer").classList.remove("need"));
  $("dq-signsave").addEventListener("click", () => {
    const list = $("dq-signlist").value.split(/\n/).map(l => l.split("|").map(x => x.trim())).filter(x => x[0]).map(([n, t]) => [n, t || ""]);
    if (!list.length) return;
    store.set("signers", list); fillSigners(); $("dq-signedit").open = false;
  });
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

  // ------------------------------------------------------------------ pink bar text per model (from Yousouf's sheet, 24 Nov 2026)
  // Models not listed here use the website's category (or the POS category).
  const PINK_BARS = {"100QNED85B8":"TV","50NU840B6LA":"TV","50NU850B6LA":"TV","55QNED70B6T":"TV","55QNED72B6T":"TV","55QNED85B6T":"TV","65QNED70B6T":"TV","65QNED72B6T":"TV","65QNED85B6T":"TV","75QNED70B6T":"TV","75QNED72B6T":"TV","75QNED85B6T":"TV","85QNED72B6T":"TV","86QNED85B6T":"TV","OLED55B66LA":"TV","OLED65B66LA":"TV","OLED65C66LA":"TV","OLED77C66LA":"TV","OLED83C66LA":"TV","GCB257SLWL":"Refrigerator","GCB411ELCM":"Refrigerator Only","GCB414ELCM":"Freezer Only","GCB459FMFW":"Refrigerator","GCJ257CQFW":"Refrigerator","GCX22FFQWS":"Refrigerator","GCX31FFCRB":"Refrigerator","GLB452PFBK":"Refrigerator","GML960PYFE":"Refrigerator","GNB262PFSF":"Refrigerator","GNB312PFGB":"Refrigerator","GNB389FLDM":"Refrigerator","GNH702HLHL":"Refrigerator","UA8F1CX":"Freezer Only","F0P2CYV2E":"Washing Machine","F2V5FYPYJ":"Washing Machine","F2V5GRPYJ":"Washer & Dryer","F2Y1HYP6J":"Washing Machine","F4V5RYP2T":"Washing Machine","F4X5EYP23":"Washing Machine","F4X5LYP26":"Washing Machine","F4Y7ERPYZ":"Washer & Dryer","F4Y7RRPYP":"Washer & Dryer","F4Y9LWP2ZB":"Washing Machine","RC90V9EV2W":"Dryer","RH10X56V2NP":"Dryer","RH90Y1MPTNP":"Dryer","T13V8TEHTX":"Washing Machine","T16V1NDHT2C":"Washing Machine","T18V1NDHT2":"Washing Machine","T23X2EHHST4":"Washing Machine","T65S1TDHT2":"Washing Machine","W4W8BVPKZHM":"Wash Tower","DFB512FP":"Dishwasher","DFC335HD":"Dishwasher","DFC435FP":"Dishwasher","MA3884VC":"Microwave","MA3885SPB":"Microwave","MA3885SPS":"Microwave","MH8265CIS":"Microwave","MH8265DIS":"Microwave","MJ3965ACS":"Microwave","MS3032JAS":"Microwave","MS4295CIS":"Microwave","MS4295DIS":"Microwave","MS5696HIT":"Microwave","AC12BK":"Air Conditioner","AC18BK":"Air Conditioner","AC24BK":"Air Conditioner","EZ12CYN":"Air Conditioner","EZ18CYN":"Air Conditioner","EZ24CYN":"Air Conditioner","S12EQ":"Air Conditioner","VM129CT":"Air Conditioner","VM189CT":"Air Conditioner","VM249CT":"Air Conditioner","BOUNCE":"XBOOM","POWER7000":"XBOOM","POWER9000":"XBOOM","RNC9":"XBOOM","S30A":"Sound Bar","S65TR":"Sound Bar","S75TR":"Sound Bar","S95AR":"Sound Bar","STAGE301":"XBOOM","STAGE501":"XBOOM","XL9T":"XBOOM","AGS72F/BK":"Built In Hob","AGS92S/BK":"Built In Hob","AMD54MXA":"Built In Microwave","AO235HX":"Built In Oven","AO435HB":"Built In Oven","AO458H1X":"Built In Oven","AO458HSUFB":"Built In Oven","AO458HSUFX":"Built In Oven","AOI5S8HM1XA":"Built In Oven","AOI78PT1SXA":"Built In Oven","AS1577CNE":"Built In Hob","DBC435TSL":"Built In Dishwasher","FTGHG751D/HA":"Built In Hob","H2IHD526A":"Built In Dishwasher","HQ2960SNE":"Built In Hob","HR602BA":"Built In Hob","IMPULSOBL/A/60":"Built In Hob","IMPULSOBL/A/83":"Built In Hob","MS5744IXA":"Built In Oven","PCN641T/IX/A":"Built In Hob","PCN751T/IX/A":"Built In Hob","PK951TGH":"Built In Hob","WS7D7632WB":"Built In Oven","WSED7612S":"Built In Oven","WSED7667M":"Built In Oven","AHBS6.7FLLX":"Kitchen Hood","AHBS9.7FLTIX":"Kitchen Hood","AHPN9.4FLMX1":"Kitchen Hood","CERTOSAIX/A/60":"Kitchen Hood","HHVP8.7FLTK":"Kitchen Hood","IXGL50":"Kitchen Hood","MORAN60":"Kitchen Hood","MORAN90":"Kitchen Hood","NORMAH6IX":"Kitchen Hood","SCARLETTXA90":"Kitchen Hood","SL161LPIX":"Kitchen Hood","SL191LPIX":"Kitchen Hood","SOFIAIX90":"Kitchen Hood","THUNDERBL/A/90":"Kitchen Hood","TT1460":"Kitchen Hood","TT1490":"Kitchen Hood","ZOEISLIX/A/90":"Kitchen Hood","HS68G8PHX/E/1":"Cooker","AR4010":"Rice Cooker","AR4080BASMATI":"Rice Cooker","CLASSICPINUP":"Toaster","CRUSTYFRY3000":"Deep Fryer","CV4000":"Steamer","DUODELUX":"Toaster","FUSIONPRO":"Hair Dryer","GENIUS":"Mixer","JUPITER1500M":"Hand Blender","KILIMANJARO":"Press Grill","LC5750":"Juicer","MERCURY2800":"Steam Iron","MOONSTONE":"Table Blender","NOVAESSENCE":"Steam Iron","PHOENIX":"Air Fryer","PPXTREME":"Hair Straightener","PULSAR600":"Hand Blender","PV2600GLIDE":"Steam Iron","RESIA":"Kettle","SC8460":"Hair Dryer","SUNRISEJOIE":"Toaster","SUNRISEVIE":"Kettle","VENEZIA":"Espresso Maker","BUTAN":"Ceiling Fan","FB3501T3":"Box Fan","FS4050T25":"Stand Fan","FW4027":"Wall Fan","NEPALNROBLE":"Ceiling Fan","A9KCORE":"Vacuum Cleaner","A9KULTRA":"Vacuum Cleaner","VC5420NHT":"Vacuum Cleaner","MD16GQSA1":"Dehumidifier","S3MFC":"Styler","A09":"Wall Bracket","S006":"Wall Bracket","ST113":"Wall Bracket","WM2615":"Wall Bracket"};
  const pinkBar = d => PINK_BARS[String(d.model || "").toUpperCase().replace(/[\s\-()]/g, "")] || d.category;

  // ------------------------------------------------------------------ product lookup
  const norm = s => String(s || "").toUpperCase().replace(/[\s\-()]/g, "");
  let sitemap = null;
  async function getSitemap(fresh) {
    if (!sitemap || fresh) {
      try { const t = await (await fetch("/sitemap.xml", { cache: "no-store" })).text(); sitemap = [...t.matchAll(/<loc>([^<]+)<\/loc>/g)].map(m => m[1].trim()); }
      catch (e) { sitemap = sitemap || []; }
    }
    return sitemap;
  }
  async function productUrl(model) {
    const sm = await getSitemap(), m = model.toLowerCase().replace(/\s+/g, ""), n = norm(model).toLowerCase();
    const hit = sm.find(u => { const l = u.toLowerCase(); return l.endsWith("-" + m) || l.endsWith("/" + m) || l.endsWith("-" + n) || l.endsWith("/" + n); });
    return hit || location.origin + "/" + model.replace(/[\s\-]/g, "");
  }

  // The product pages carry their data (Nuxt payload). Read it without opening the page.
  function reviveNuxt(arr) {
    const cache = new Map(), wrap = new Set(["Reactive", "ShallowReactive", "Ref", "ShallowRef", "EmptyRef", "EmptyShallowRef", "NuxtError", "Island"]);
    const hy = i => {
      if (typeof i !== "number") return i;
      if (i < 0) return undefined;
      if (cache.has(i)) return cache.get(i);
      const v = arr[i];
      if (v === null || typeof v !== "object") { cache.set(i, v); return v; }
      if (Array.isArray(v)) {
        if (typeof v[0] === "string" && (wrap.has(v[0]) || v[0] === "Date")) { const r = hy(v[1]); cache.set(i, r); return r; }
        if (typeof v[0] === "string" && (v[0] === "Set" || v[0] === "Map")) { const r = v.slice(1).map(hy); cache.set(i, r); return r; }
        const out = []; cache.set(i, out); v.forEach(x => out.push(hy(x))); return out;
      }
      const o = {}; cache.set(i, o); for (const k in v) o[k] = hy(v[k]); return o;
    };
    return hy(0);
  }
  const clean = s => String(s || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  const frag = html => new DOMParser().parseFromString(html || "", "text/html").body;
  const tableRows = html => [...frag(html).querySelectorAll("tr")].map(tr => [...tr.children].map(c => clean(c.textContent)).filter(Boolean)).filter(r => r.length);
  function parseProductPage(html, url) {
    const m = html.match(/<script[^>]*id="__NUXT_DATA__"[^>]*>([\s\S]*?)<\/script>/);
    if (!m) return null;
    const data = reviveNuxt(JSON.parse(m[1]));
    const ent = Object.entries((data && data.data) || {}).find(([k, v]) => k.startsWith("gql") && v && v.product && v.product.sku);
    if (!ent) return null;
    const p = ent[1].product, meta = {};
    (p.metaData || []).forEach(x => { if (x && x.key) meta[x.key] = x.value; });
    const doc = new DOMParser().parseFromString(html, "text/html");
    let crumbs = [];
    doc.querySelectorAll('script[type="application/ld+json"]').forEach(s => {
      try { const j = JSON.parse(s.textContent); (j["@graph"] || [j]).forEach(n => { if (n["@type"] === "BreadcrumbList") crumbs = (n.itemListElement || []).map(e => clean(e.name)); }); } catch (e) { /* ignore */ }
    });
    const cats = ((p.productCategories && p.productCategories.nodes) || []);
    const topCat = (cats.find(c => !c.parent) || {}).name || (((cats[0] || {}).parent || {}).node || {}).name;
    // features and * notes, in page order
    const sd = frag(p.shortDescription); sd.querySelectorAll(".warranty, table").forEach(e => e.remove());
    const features = [], notes = [];
    sd.querySelectorAll("p, li").forEach(el => {
      let t;
      if (el.tagName === "LI") t = clean([...el.childNodes].filter(n => !(n.nodeType === 1 && /^(UL|OL)$/.test(n.tagName))).map(n => n.textContent).join(" "));
      else t = clean(el.textContent);
      if (!t) return;
      (t.startsWith("*") ? notes : features).push(t);
    });
    // dimensions and weights (weight added to the matching dimension line)
    const dims = [];
    for (const r of tableRows(meta.dimension)) if (r.length >= 2) dims.push([r[0], r.slice(1).join(" ")]);
    for (const r of tableRows(meta.weight_specs)) {
      if (r.length < 2) continue;
      const ex = dims.find(d => d[0] === r[0]);
      if (ex) ex[1] += " | " + r.slice(1).join(" "); else dims.push([r[0], r.slice(1).join(" ")]);
    }
    let capacity = null;
    for (const [k, v] of Object.entries(meta)) {
      if (!/_specs$/.test(k) || k === "weight_specs" || !v) continue;
      const rows = tableRows(v);
      if (rows.length >= 2) { capacity = rows[0].map((x, i) => x + " " + (rows[1][i] || "")).join("  |  "); break; }
    }
    const retail = Number(p.rawRegularPrice || p.rawPrice) || null, sale = Number(p.rawSalePrice) || null;
    let promoValid = null;
    if (p.dateOnSaleTo) { const d = new Date(p.dateOnSaleTo); if (!isNaN(d)) promoValid = `${pad2(d.getDate())} ${["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.getMonth()]} ${d.getFullYear()}`; }
    const model = clean(p.sku);
    return {
      id: norm(model), keys: [...new Set([norm(model), norm(meta.product_model)].filter(Boolean))],
      model, name: clean(doc.title.replace(/\s*-\s*Dragon Electronics\s*$/, "")) || clean(p.name), url, brand: clean(meta.brand) || null,
      category: crumbs.length >= 3 ? crumbs[1] : (topCat || "Products"),
      retail, promo: sale && retail && sale < retail ? sale : null, promoValid,
      promoEnd: p.dateOnSaleTo && /^\d{4}-\d{2}-\d{2}/.test(p.dateOnSaleTo) ? p.dateOnSaleTo.slice(0, 10) : null,
      noStock: meta._no_stock === "yes" || p.stockStatus === "OUT_OF_STOCK",
      noStockUntil: meta._no_stock_until === "yes" && /^\d{4}-\d{2}-\d{2}/.test(meta._no_stock_until_date || "") ? meta._no_stock_until_date.slice(0, 10) : null,
      features, notes, dims, capacity,
      image: (doc.querySelector('meta[property="og:image"]') || {}).content || null,
      checked: new Date().toISOString()
    };
  }

  // ------------------------------------------------------------------ saved product list (IndexedDB, this computer only)
  const IDB = (() => {
    let dbp = null;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const r = indexedDB.open("dq-products", 3);
      r.onupgradeneeded = () => { const db = r.result; for (const s of ["products", "photos", "pos", "settings"]) if (!db.objectStoreNames.contains(s)) db.createObjectStore(s, { keyPath: "id" }); };
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    }));
    const run = async (store, mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => { const t = db.transaction(store, mode), req = fn(t.objectStore(store)); t.oncomplete = () => res(req && req.result); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error); });
    };
    return {
      get: (s, k) => run(s, "readonly", o => o.get(k)), put: (s, v) => run(s, "readwrite", o => o.put(v)),
      all: s => run(s, "readonly", o => o.getAll()), keys: s => run(s, "readonly", o => o.getAllKeys()),
      replaceAll: (s, list) => run(s, "readwrite", o => { o.clear(); list.forEach(v => o.put(v)); return null; })
    };
  })();
  const catalog = new Map(), alias = new Map();      // website products (saved copies)
  const posMap = new Map();                             // POS product list = the full range
  const onWebsite = id => { const p = catalog.get(id); return !!(p && !p.removed); };
  // a product that is not on the website now, as it will be quoted: POS price, saved website description/photo if any
  function offWebView(P) {
    const w = catalog.get(P.id);
    if (w) return Object.assign({}, w, { retail: P.retail, promo: null, promoValid: null, url: null, noStock: false, stale: false, pos: true, offWeb: true, posAsAt: P.posAsAt, descFrom: "website" });
    return Object.assign({}, P, { offWeb: true, descFrom: "pos" });
  }
  const remember = p => { catalog.set(p.id, p); (p.keys || [p.id]).forEach(k => alias.set(k, p.id)); };
  const findSaved = model => { const id = alias.get(norm(model)); return id ? catalog.get(id) : null; };
  const catInfo = () => store.get("catalog_info", null);
  function catStatus() {
    const info = catInfo(), n = [...catalog.values()].filter(p => !p.removed).length;
    const el = $("dq-catstat");
    if (!n && posMap.size) { el.className = "ct need"; el.textContent = `${posMap.size} products from the POS list. Load the website list too, for photos, promo prices and the full descriptions.`; $("dq-catupd").textContent = "Load product list"; return; }
    if (!n) { el.className = "ct need"; el.textContent = "No product list yet. Load it once (about 2 minutes) to get model suggestions and faster quotes."; $("dq-catupd").textContent = "Load product list"; return; }
    const d = info && info.date ? new Date(info.date) : null, old = d && (Date.now() - d) > 7 * 86400000;
    el.className = "ct" + (old ? " need" : "");
    const off = [...posMap.keys()].filter(id => !onWebsite(id)).length;
    el.textContent = `${n} on the website${off ? ` + ${off} more from the POS list` : ""}${d ? " \u00b7 updated " + DQCore.fmtDate(d) : ""}${old ? " \u00b7 update recommended" : ""}`;
    const pi = store.get("pos_info", null);
    $("dq-posstat").textContent = pi ? ` \u00b7 loaded: list of ${pi.asAt} (${pi.count} products)` : "";
    $("dq-catupd").textContent = "Update product list";
  }
  (async () => {
    try { (await IDB.all("products")).forEach(remember); (await IDB.all("pos")).forEach(p => posMap.set(p.id, p)); } catch (e) { /* no storage */ }
    catStatus();
  })();
  function searchCatalog(q, max) {
    const toks = q.toLowerCase().split(/\s+/).filter(Boolean), qn = norm(q).toLowerCase();
    const res = [];
    const pool = [...catalog.values()].filter(p => !p.removed);
    for (const P of posMap.values()) if (!onWebsite(P.id)) pool.push(offWebView(P));
    for (const p of pool) {
      const mk = norm(p.model).toLowerCase(), hay = `${p.model} ${p.name} ${p.category}`.toLowerCase(), hayN = hay.replace(/[\s\-]/g, "");
      let score = -1;
      if (mk === qn) score = 0;
      else if (mk.startsWith(qn)) score = 1;
      else if (qn.length >= 2 && mk.includes(qn)) score = 2;
      else if (toks.every(t => hay.includes(t) || hayN.includes(t.replace(/-/g, "")))) score = 3;
      if (score >= 0) res.push([score, p]);
    }
    res.sort((x, y) => x[0] - y[0] || x[1].model.localeCompare(y[1].model));
    return res.slice(0, max).map(r => r[1]);
  }
  const thumbs = new Map();
  async function thumbUrl(id) {
    if (thumbs.has(id)) return thumbs.get(id);
    let u = null;
    try { const ph = await IDB.get("photos", id); if (ph) u = URL.createObjectURL(new Blob([ph.bytes], { type: "image/jpeg" })); } catch (e) { /* none */ }
    thumbs.set(id, u); return u;
  }
  async function savedPhoto(p) {
    if (!p.image) return null;
    try { const ph = await IDB.get("photos", p.id); if (ph && ph.src === p.image) return { bytes: ph.bytes, width: ph.width, height: ph.height }; } catch (e) { /* none */ }
    const im = await photo(p.image);
    try { await IDB.put("photos", { id: p.id, src: p.image, bytes: im.bytes, width: im.width, height: im.height }); thumbs.delete(p.id); } catch (e) { /* storage full */ }
    return im;
  }
  async function pool(list, n, fn) { let i = 0; await Promise.all(Array.from({ length: Math.min(n, list.length) }, async () => { while (i < list.length) { const k = i++; await fn(list[k], k); } })); }
  $("dq-catupd").addEventListener("click", async () => {
    const btn = $("dq-catupd"), bar = $("dq-catbar"), fill = bar.firstElementChild, el = $("dq-catstat");
    btn.disabled = true; bar.style.display = "block"; fill.style.width = "0";
    const prog = (txt, f) => { el.className = "ct"; el.textContent = txt; fill.style.width = Math.round(f * 100) + "%"; };
    try {
      const urls = (await getSitemap(true)).filter(u => { const path = u.replace(/^https?:\/\/[^/]+\//, "").replace(/\/$/, ""); return path.split("/").length === 2 && !/^brand\//.test(path); });
      if (!urls.length) throw new Error("the website's product list (sitemap) could not be read");
      const before = new Map(catalog), seen = new Set(); let added = 0, changed = 0, done = 0, failed = 0;
      await pool(urls, 4, async u => {
        try {
          const r = await fetch(u, { cache: "no-store" });
          if (r.ok) {
            const p = parseProductPage(await r.text(), r.url);
            if (p && p.retail) {
              const o = before.get(p.id);
              if (!o || o.removed) added++; else if (o.retail !== p.retail || o.promo !== p.promo) changed++;
              p.removed = false; seen.add(p.id); remember(p); await IDB.put("products", p);
            }
          }
        } catch (e) { failed++; }
        prog(`Reading products\u2026 ${++done} / ${urls.length}`, 0.6 * done / urls.length);
      });
      if (!seen.size) throw new Error("no product could be read \u2013 the website may have changed");
      let removed = 0;
      for (const p of catalog.values()) if (!seen.has(p.id) && !p.removed) { p.removed = true; removed++; await IDB.put("products", p); }
      const have = new Map((await IDB.all("photos")).map(ph => [ph.id, ph.src]));
      const need = [...catalog.values()].filter(p => !p.removed && p.image && have.get(p.id) !== p.image);
      let pd = 0, pfail = 0;
      await pool(need, 3, async p => {
        try { await savedPhoto(p); } catch (e) { pfail++; }
        prog(`Saving photos\u2026 ${++pd} / ${need.length}`, 0.6 + 0.4 * pd / need.length);
      });
      store.set("catalog_info", { date: new Date().toISOString(), count: seen.size });
      catStatus();
      const parts = [`${seen.size} products saved`];
      if (before.size) parts.push(`${added} new`, `${changed} price change${changed === 1 ? "" : "s"}`, `${removed} removed from the website`);
      log(`\u2713 Product list updated: ${parts.join(", ")}.${failed || pfail ? ` ${failed + pfail} item(s) could not be read; they are tried again at the next update.` : ""}`, "ok");
    } catch (e) {
      catStatus(); log("\u2717 Product list not updated: " + h(e.message || e), "err");
    } finally { btn.disabled = false; bar.style.display = "none"; }
  });

  // ------------------------------------------------------------------ POS product list (FEATURES.TXT) \u2013 only for products NOT on the website
  const CP437 = "\u00c7\u00fc\u00e9\u00e2\u00e4\u00e0\u00e5\u00e7\u00ea\u00eb\u00e8\u00ef\u00ee\u00ec\u00c4\u00c5\u00c9\u00e6\u00c6\u00f4\u00f6\u00f2\u00fb\u00f9\u00ff\u00d6\u00dc\u00a2\u00a3\u00a5\u20a7\u0192\u00e1\u00ed\u00f3\u00fa\u00f1\u00d1\u00aa\u00ba\u00bf\u2310\u00ac\u00bd\u00bc\u00a1\u00ab\u00bb\u2591\u2592\u2593\u2502\u2524\u2561\u2562\u2556\u2555\u2563\u2551\u2557\u255d\u255c\u255b\u2510\u2514\u2534\u252c\u251c\u2500\u253c\u255e\u255f\u255a\u2554\u2569\u2566\u2560\u2550\u256c\u2567\u2568\u2564\u2565\u2559\u2558\u2552\u2553\u256b\u256a\u2518\u250c\u2588\u2584\u258c\u2590\u2580\u03b1\u00df\u0393\u03c0\u03a3\u03c3\u00b5\u03c4\u03a6\u0398\u03a9\u03b4\u221e\u03c6\u03b5\u2229\u2261\u00b1\u2265\u2264\u2320\u2321\u00f7\u2248\u00b0\u2219\u00b7\u221a\u207f\u00b2\u25a0\u00a0";
  function posText(buf) {
    const u8 = new Uint8Array(buf);
    try {   // saved as UTF-8 from a DOS (code page 437) print: letters like \u00e0/\u00f8 stand for \u03b1/\u00b0
      return new TextDecoder("utf-8", { fatal: true }).decode(u8).replace(/[\u0080-\u00ff]/g, c => CP437[c.charCodeAt(0) - 0x80]);
    } catch (e) { let s = ""; for (const b of u8) s += b < 0x80 ? String.fromCharCode(b) : CP437[b - 0x80]; return s; }
  }
  const titleCase = s => s.toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase()).replace(/\bTv\b/g, "TV").replace(/\bXboom\b/g, "XBOOM");
  const POS_CAT = { "TV": "TV", "WASHING MACHINE": "Washer & Dryer", "DRYER": "Washer & Dryer", "WASHER DRYER TOWER": "Washer & Dryer", "REFRIGERATOR": "Refrigerator & Freezer",
    "FREEZER ONLY": "Refrigerator & Freezer", "MICROWAVE OVEN": "Microwave", "AIR CONDITIONER": "Air Conditioner", "DISHWASHER": "Dishwasher", "XBOOM 1-BODY": "Audio",
    "XBOOM PARTY SPEAKER": "Audio", "HOME THEATER SYSTEM": "Audio", "KITCHEN HOOD": "Kitchen Hood", "BUILT-IN APPLIANCE": "Built-In Appliance", "VACUUM CLEANER": "Vacuum Cleaner",
    "ELECTRIC FAN": "Electric Fan", "DEHUMIDIFIER": "Dehumidifier", "COOKER": "Cooker", "STYLER": "Steam Clothing Care", "TV WALL BRACKET": "TV Wall Bracket",
    "ELECTRIC FRYER": "Small Domestic Appliance", "HAIR DRYER": "Small Domestic Appliance", "RICE COOKER": "Small Domestic Appliance" };
  function posToProduct(r, asAt) {
    const brand = r.brand === "H ARISTON" ? "ARISTON" : r.brand === "DE" ? "DE Fan" : r.brand === "OTHER" ? "" : r.brand;
    const L = r.lines.slice();
    // name: first line when it already starts with the brand, else "Brand Category"
    let name;
    if (brand && L[0] && L[0].toUpperCase().startsWith(brand.split(" ")[0].toUpperCase())) name = L.shift();
    else name = [brand === "DE Fan" ? "DE" : brand, titleCase(r.category)].filter(Boolean).join(" ");
    // notes: last block when it starts with "*" (continuation lines joined)
    const blocks = []; let cur = [];
    for (const l of L) { if (!l.trim()) { if (cur.length) blocks.push(cur); cur = []; } else cur.push(l.trim()); }
    if (cur.length) blocks.push(cur);
    const notes = [];
    if (blocks.length && blocks[blocks.length - 1][0].startsWith("*")) {
      for (const l of blocks.pop()) { if (l.startsWith("*") || !notes.length) notes.push(l); else notes[notes.length - 1] += " " + l; }
    }
    const features = [], dims = [], weights = [];
    let mode = null, capacity = [];
    const DIM = /W\s?[\d.]+\s*[xX]\s*D\s?[\d.]+\s*[xX]\s*H\s?[\d.]+\s*cm/;
    const lab = s => { s = s.replace(/[:\s]+$/, "").trim(); return !s || /^(dimensions?|weight|net weight|product)$/i.test(s) ? "Product" : s.replace(/^with box$/i, "Product with Box").replace(/^weight\s+/i, ""); };
    for (const b of blocks) for (const l of b) {
      if (/^(key features|specs):?$/i.test(l)) continue;
      if (/^dimensions?:?$/i.test(l)) { mode = "dim"; continue; }
      if (/^weight:?$/i.test(l)) { mode = "wt"; continue; }
      if (/^net capacity:?$/i.test(l)) { mode = "cap"; continue; }
      const dm = l.match(DIM), wm = l.match(/^(.*?)\s*:?\s*([\d.]+\s*Kg)$/i);
      if (dm) { dims.push([lab(l.slice(0, dm.index)), dm[0].replace(/\s*x\s*/g, " X ")]); continue; }
      if (wm && (mode === "wt" || /weight/i.test(l))) { weights.push([lab(wm[1]), wm[2].replace(/\s+/g, "")]); continue; }
      if (mode === "cap" && /\d\s*L$/i.test(l)) { capacity.push(l.replace(/\s+/g, " ")); continue; }
      features.push(l.replace(/\s{2,}/g, " "));
    }
    for (const [lb, w] of weights) {
      const d = dims.find(x => x[0].toLowerCase() === lb.toLowerCase()) || (lb === "Product" && dims.find(x => /^with stand$/i.test(x[0])));
      if (d) d[1] += " | " + w; else dims.push([lb, w]);
    }
    if (dims.length && /^with stand$/i.test(dims[0][0])) dims[0][0] = "Product";
    return {
      id: norm(r.model), keys: [norm(r.model)], model: r.model, name, brand: brand || null, url: null,
      category: POS_CAT[r.category] || titleCase(r.category), retail: r.price, promo: null, promoValid: null, noStock: false,
      features, notes, dims, capacity: capacity.length ? capacity.join("  |  ") : null, image: null,
      checked: new Date().toISOString(), pos: true, posAsAt: asAt, posCode: r.code
    };
  }
  function parsePosList(text) {
    const lines = text.split(/\r?\n/);
    const hdr = lines.find(l => l.startsWith("|P.CODE"));
    if (!hdr) throw new Error("this is not the POS product list (no P.CODE header)");
    const asAt = ((text.match(/PRODUCT LIST AS AT\s+([\d\/]+)/) || [])[1]) || "";
    const bars = []; for (let i = 0; i < hdr.length; i++) if (hdr[i] === "|") bars.push(i);
    const cut = l => bars.slice(0, -1).map((b, i) => (l.slice(b + 1, bars[i + 1]) || "").trim());
    const recs = []; let cur = null;
    for (const l of lines) {
      if (!l.startsWith("|") || l.startsWith("|P.CODE") || /^\|\s*\|BRAND\s/.test(l)) continue;
      const [code, c2, c3, feat, price] = cut(l);
      if (code) { cur = { code, category: c2, type: c3, lines: feat ? [feat] : [], price: Number(String(price).replace(/,/g, "")) || null, brand: "", model: "", n: 1 }; recs.push(cur); }
      else if (cur) { cur.n++; if (cur.n === 2) { cur.brand = c2; cur.model = c3; } cur.lines.push(feat); }
    }
    return { asAt, recs: recs.filter(r => r.model && r.price) };
  }
  $("dq-poslink").addEventListener("click", e => { e.preventDefault(); $("dq-posfile").click(); });
  $("dq-posfile").addEventListener("change", async e => {
    const file = e.target.files[0]; e.target.value = ""; if (!file) return;
    try {
      const { asAt, recs } = parsePosList(posText(await file.arrayBuffer()));
      if (!recs.length) throw new Error("no product found in the file");
      const list = recs.map(r => posToProduct(r, asAt));
      await IDB.replaceAll("pos", list);
      posMap.clear(); list.forEach(p => posMap.set(p.id, p));
      store.set("pos_info", { asAt, count: list.length, date: new Date().toISOString() });
      catStatus();
      const off = list.filter(p => !onWebsite(p.id));
      const diff = list.filter(p => onWebsite(p.id) && catalog.get(p.id).retail !== p.retail);
      log(`\u2713 POS list of ${h(asAt)} loaded: ${list.length} products${catalog.size ? `, ${list.length - off.length} also on the website, <b>${off.length} not on the website</b> (quoted with the POS price)` : ""}.`
        + (diff.length ? ` <span class="warn">Normal price different on the website: ${diff.slice(0, 15).map(p => `${h(p.model)} POS ${fmt(p.retail)} / web ${fmt(catalog.get(p.id).retail)}`).join("; ")}${diff.length > 15 ? "; \u2026" : ""}.</span>` : ""), "ok");
    } catch (err) { log("\u2717 POS list not added: " + h(err.message || err), "err"); }
  });

  // Live lookup at quote time: always reads the current page (prices change); the saved list gives the
  // page address, the photo, and a fallback when the product has been removed from the website.
  // Order: 1) the website now  2) not on the website -> POS list price (description/photo from the saved
  // website copy, else the POS text)  3) only a saved website copy -> used with a warning.
  async function readProduct(model) {
    const saved = findSaved(model), P = posMap.get(norm(model)) || (saved && posMap.get(saved.id));
    let url = saved && saved.url ? saved.url : await productUrl(model);
    let r = await fetch(url, { cache: "no-store" }).catch(() => null);
    if ((!r || !r.ok) && saved) { const u2 = await productUrl(model); if (u2 !== url) { url = u2; r = await fetch(url, { cache: "no-store" }).catch(() => null); } }
    if (!r || !r.ok) {
      if (P) return offWebView(P);
      if (saved) return Object.assign({}, saved, { stale: true });
      throw new Error("not on the website" + (posMap.size ? " and not in the POS list" : "") + " \u2013 check the model code");
    }
    let d = null;
    try { d = parseProductPage(await r.text(), r.url); } catch (e) { d = null; }
    if (!d) return readProductFrame(r.url, model);
    if (!d.keys.includes(norm(model))) throw new Error(`the page found is for ${d.model}, not ${model}`);
    if (!d.retail) throw new Error("price not found on the page");
    d.removed = false; remember(d); IDB.put("products", d).catch(() => {});
    return d;
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

  // backup method: open the page in a hidden frame and read what is shown on screen
  async function readProductFrame(url, model) {
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
      data.id = norm(pageModel); data.keys = [data.id];
      if (norm(pageModel) !== norm(model)) throw new Error(`the page found is for ${pageModel}, not ${model}`);
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

  // ------------------------------------------------------------------ save quotes into a folder on this computer (year / month subfolders)
  const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const canPickFolder = typeof window.showDirectoryPicker === "function";
  let quoteDir = null;
  async function dirStatus() {
    const st = $("dq-dirstat");
    if (!canPickFolder) { st.textContent = "This browser cannot save into a folder; quotes go to Downloads."; $("dq-dirbtn").style.display = "none"; return; }
    st.innerHTML = quoteDir ? `Saving to <b>${h(quoteDir.name)}</b> \u203a year \u203a month` : "Quotes go to your Downloads folder.";
    $("dq-dirbtn").textContent = quoteDir ? "Change folder" : "Choose folder";
  }
  (async () => { try { const s = await IDB.get("settings", "quoteDir"); if (s && s.handle) quoteDir = s.handle; } catch (e) { /* none */ } dirStatus(); })();
  $("dq-dirbtn").addEventListener("click", async () => {
    try {
      const d = await window.showDirectoryPicker({ id: "dq-quotes", mode: "readwrite", startIn: "documents" });
      quoteDir = d; await IDB.put("settings", { id: "quoteDir", handle: d }); dirStatus();
    } catch (e) { if (e && e.name !== "AbortError") log("\u2717 Folder not set: " + h(e.message || e), "err"); }
  });
  // must run straight after the click (Chrome only asks for permission during a click)
  async function folderReady() {
    const d = window.__dqTestDir || quoteDir;
    if (!d) return null;
    try {
      const opts = { mode: "readwrite" };
      if (d.queryPermission && (await d.queryPermission(opts)) === "granted") return d;
      if (d.requestPermission && (await d.requestPermission(opts)) === "granted") return d;
    } catch (e) { /* not allowed */ }
    return null;
  }
  async function saveToFolder(d, isoDate, fname, bytes) {
    const dt = isoDate ? new Date(isoDate + "T12:00:00") : new Date();
    const y = await d.getDirectoryHandle(String(dt.getFullYear()), { create: true });
    const m = await y.getDirectoryHandle(`${pad2(dt.getMonth() + 1)}-${MONTH_NAMES[dt.getMonth()]}`, { create: true });
    let name = fname, k = 2;
    const exists = async n => { try { await m.getFileHandle(n); return true; } catch (e) { return false; } };
    while (await exists(name)) name = fname.replace(/\.xlsx$/, ` (${k++}).xlsx`);
    const fh = await m.getFileHandle(name, { create: true });
    const w = await fh.createWritable(); await w.write(new Blob([bytes])); await w.close();
    return `${d.name}\\${dt.getFullYear()}\\${pad2(dt.getMonth() + 1)}-${MONTH_NAMES[dt.getMonth()]}\\${name}`;
  }

  // ---- "Quote Log.xlsx" in the quotes folder: one row per quote (kept outside Chrome)
  const LOG_NAME = "Quote Log.xlsx", LOG_HEAD = ["Date", "Quotation No", "Customer", "Signed by", "Products", "Total incl. VAT (Rs)", "File"];
  const xmlEsc = v => String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const colL = i => String.fromCharCode(65 + i);
  async function readLog(dir) {
    let fh; try { fh = await dir.getFileHandle(LOG_NAME); } catch (e) { return []; }
    const files = await DQCore.unzip(await (await fh.getFile()).arrayBuffer());
    const xml = new TextDecoder().decode(files.get("xl/worksheets/sheet1.xml") || new Uint8Array());
    const rows = [];
    for (const m of xml.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
      if (+m[1] === 1) continue;
      const r = [];
      for (const c of m[2].matchAll(/<c r="([A-Z]+)\d+"[^>]*?(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const t = (c[2] || "").match(/<t[^>]*>([\s\S]*?)<\/t>/), v = (c[2] || "").match(/<v>([^<]*)<\/v>/);
        const val = t ? t[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&") : v ? Number(v[1]) : "";
        r[c[1].charCodeAt(0) - 65] = val;
      }
      rows.push(LOG_HEAD.map((_, i) => r[i] == null ? "" : r[i]));
    }
    return rows;
  }
  async function writeLog(dir, rows) {
    const enc = new TextEncoder(), f = new Map();
    const cell = (v, r, i, st) => typeof v === "number" ? `<c r="${colL(i)}${r}" s="${st || 2}"><v>${v}</v></c>` : `<c r="${colL(i)}${r}" t="inlineStr"${st ? ` s="${st}"` : ""}><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
    const all = [LOG_HEAD, ...rows];
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`
      + `<cols><col min="1" max="1" width="13" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/><col min="3" max="3" width="32" customWidth="1"/><col min="4" max="4" width="24" customWidth="1"/><col min="5" max="5" width="50" customWidth="1"/><col min="6" max="6" width="20" customWidth="1"/><col min="7" max="7" width="48" customWidth="1"/></cols><sheetData>`
      + all.map((r, k) => `<row r="${k + 1}">${r.map((v, i) => cell(v, k + 1, i, k === 0 ? 1 : 0)).join("")}</row>`).join("")
      + `</sheetData><autoFilter ref="A1:G${all.length}"/></worksheet>`;
    f.set("[Content_Types].xml", enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`));
    f.set("_rels/.rels", enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`));
    f.set("xl/workbook.xml", enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Quotes" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Quotes!$A$1:$G$${all.length}</definedName></definedNames></workbook>`));
    f.set("xl/_rels/workbook.xml.rels", enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`));
    f.set("xl/styles.xml", enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC8102E"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`));
    f.set("xl/worksheets/sheet1.xml", enc.encode(sheet));
    const fh = await dir.getFileHandle(LOG_NAME, { create: true });
    const w = await fh.createWritable(); await w.write(new Blob([await DQCore.zip(f)])); await w.close();
  }
  // adds the quote (and any that could not be added before, e.g. while the log was open in Excel)
  async function addToLog(dir, row) {
    const pending = store.get("log_pending", []); pending.push(row);
    try {
      const rows = await readLog(dir);
      await writeLog(dir, rows.concat(pending));
      store.set("log_pending", []);
      return null;
    } catch (e) {
      store.set("log_pending", pending);
      return `${LOG_NAME} could not be updated (is it open in Excel? close it). This quote will be added to it with the next one.`;
    }
  }

  // ------------------------------------------------------------------ generate
  $("dq-go").onclick = async () => {
    const go = $("dq-go"); $("dq-log").innerHTML = "";
    const folderP = folderReady();          // ask for the folder permission now, while the click counts
    const to = $("dq-to").value.trim();
    const items = [...$("dq-rows").querySelectorAll(".prow")].map(row => {
      const model = row.querySelector(".pm").value.trim().toUpperCase(), qtxt = row.querySelector(".pq").value.trim();
      const disc = row.querySelector(".pdc").checked, ptxt = row.querySelector(".pp").value.trim();
      if (!model) return null;
      if (!/^[A-Z0-9.\-\/ ()]+$/.test(model)) return { bad: `${model} (model code)` };
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
    if (!$("dq-signer").value) { $("dq-signer").classList.add("need"); $("dq-signer").focus(); return log("Choose who signs this quote (Signed by, in the Quotation section).", "err"); }
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
          products.push({ bar: pinkBar(d), isPromo: !!usePromo, isDisc: own != null, noStock: !!d.noStock && !d.offWeb, noStockUntil: d.noStockUntil || null, category: d.category, brand: d.brand, model: d.model, name: d.name, qty: it.qty, price, normal: d.retail, url: d.url, features: d.features, capacity: d.capacity, dims: d.dims, notes: d.notes, _d: d, _promo: usePromo, _own: own != null });
          try { const im = await savedPhoto(d); if (im) images[d.model] = im; } catch (e) { problems.push(`${d.model}: photo could not be read`); }
          if (d.stale) problems.push(`${d.model}: not on the website any more \u2013 details and price taken from the saved product list (${DQCore.fmtDate(new Date(d.checked))}). Check them.`);
          line.className = "ok"; line.innerHTML = `\u2713 ${h(d.model)} \u2013 ${h(d.name)} \u2013 ${fmt(price)}${own != null ? ` <span class='warn'>(your price; website ${fmt(d.retail)})</span>` : ""}${d.noStock && !d.offWeb ? ` <span class='warn'>(No Stock on website${d.noStockUntil ? " \u2013 arrival on or about " + h(DQCore.fmtDate(new Date(d.noStockUntil + "T12:00:00"))) : ""})</span>` : ""}${d.stale ? " <span class='warn'>(saved copy \u2013 removed from website)</span>" : ""}${d.offWeb ? " <span class='warn'>(not on the website \u2013 POS price)</span>" : ""}`;
          if (!d.dims.length) problems.push(`${d.model}: no dimensions ${d.descFrom === "pos" ? "in the POS list" : "on the website"}`);
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
        let untilIso = isoLocal(d);
        // a quote must not promise a promo price after the promo ends
        const promoEnds = products.filter(p => p._promo && p._d.promoEnd).map(p => [p._d.promoEnd, p.model]).sort();
        if (promoEnds.length && promoEnds[0][0] < untilIso) {
          untilIso = promoEnds[0][0];
          problems.push(`Validity shortened to ${DQCore.fmtDate(new Date(untilIso + "T12:00:00"))}: the promo price of ${promoEnds.filter(x => x[0] === untilIso).map(x => x[1]).join(", ")} ends that day.`);
        }
        t.validity = { until: untilIso };
      }
      t.installation = $("dq-install").checked;
      const spec = {
        customer: { to, company: $("dq-ctype").value === "company", address: [$("dq-addr1").value.trim(), $("dq-addr2").value.trim()].filter(Boolean), attention: $("dq-att").value.trim(),
          vat: $("dq-ctype").value === "company" ? $("dq-vat").value.trim() : "", brn: $("dq-ctype").value === "company" ? $("dq-brn").value.trim() : "", tel: $("dq-tel").value.trim(), mobile: $("dq-mob").value.trim(), email: $("dq-email").value.trim() },
        signer: (() => { const n = $("dq-signer").value, e = getSigners().find(x => x[0] === n); return e ? { name: e[0], title: e[1] } : null; })(),
        quote_no: $("dq-no").value.trim(), date: $("dq-date").value, terms: t, layout: $("dq-disc").checked ? "discount" : "plain",
        warranties: [...new Set(products.map(DQCore.warrantyFor).filter(Boolean))], products
      };
      products.forEach(p => { if (!DQCore.warrantyFor(p)) problems.push(`${p.model}: no matching warranty line \u2013 check the Terms`); });
      const offOnes = products.filter(p => p._d.offWeb);
      if (offOnes.length) {
        const wd = offOnes.filter(p => p._d.descFrom === "website").map(p => p.model), pd = offOnes.filter(p => p._d.descFrom === "pos").map(p => p.model);
        problems.unshift(`Not on the website now \u2013 normal price from the POS list of ${offOnes[0]._d.posAsAt}, no promo, no "More info" link.`
          + (wd.length ? ` Description and photo from the last saved website copy: ${wd.join(", ")}.` : "")
          + (pd.length ? ` Description from the POS list (no photo): ${pd.join(", ")} \u2013 check the wording.` : ""));
      }
      log("Building the Excel file\u2026");
      const res = await DQCore.buildQuote(template, spec, images);
      const fname = `Quotation_${spec.quote_no.replace(/\s*\|\s*/g, "-").replace(/[^\w\-]+/g, "_")}_${to.replace(/^(mr|mrs|ms|miss|dr)\.?\s+/i, "").replace(/[^\w]+/g, "_")}.xlsx`;
      const url = URL.createObjectURL(new Blob([res.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      let savedAs = null, folderNote = "";
      const dir = await folderP;
      if (dir) { try { savedAs = await saveToFolder(dir, spec.date, fname, res.bytes); } catch (e) { folderNote = "Could not save into your folder (" + (e.message || e) + "), so the file was downloaded instead."; } }
      else if (quoteDir) folderNote = "Chrome did not allow saving into your folder this time, so the file was downloaded instead. Click Allow when Chrome asks.";
      if (!savedAs) { const a = document.createElement("a"); a.href = url; a.download = fname; document.body.appendChild(a); a.click(); a.remove(); }
      const counterMatch = spec.quote_no.match(/(\d+)\s*$/);
      if (counterMatch && +counterMatch[1] > (+$("dq-counter").value || 0)) { store.set("counter", +counterMatch[1]); $("dq-counter").value = +counterMatch[1]; }
      const hist = store.get("history", []); hist.push({ date: spec.date, no: spec.quote_no, customer: to, total: res.total, signedBy: spec.signer && spec.signer.name }); store.set("history", hist.slice(-200)); renderHist();
      $("dq-signer").value = "";            // shared PC: the next quote asks again who signs it
      const promoRows = products.filter(p => !p._promo && !p._own && p._d.promo);
      let html = `<div class="hint">Prices incl. VAT</div><table><tr><th>Model</th><th class="n">Qty</th><th class="n">Price</th><th class="n">Amount</th></tr>${products.map(p => `<tr><td>${h(p.model)}</td><td class="n">${p.qty}</td><td class="n">${fmt(p.price)}</td><td class="n">${fmt(p.qty * p.price)}</td></tr>`).join("")}<tr><th colspan="3">Total incl. VAT</th><th class="n">${fmt(res.total)}</th></tr></table>`;
      if (promoRows.length) html += `<div class="warn" style="margin-top:6px">Promo prices available: ${promoRows.map(p => `${h(p.model)} ${fmt(p._d.promo)}${p._d.promoValid ? " until " + h(p._d.promoValid) : ""}`).join("; ")}.</div>`;
      const posSet = new Set(products.filter(p => p._d.descFrom === "pos").map(p => "No photo for " + p.model));
      [...problems, ...res.warnings.filter(w => !posSet.has(w))].forEach(w => { html += `<div class="warn">\u26a0 ${h(w)}</div>`; });
      if (savedAs) {
        html += `<div class="ok" style="margin-top:6px">\u2713 Saved in <b>${h(savedAs)}</b></div>`;
        const logRow = [DQCore.fmtDate(new Date((spec.date || isoLocal(new Date())) + "T12:00:00")), spec.quote_no, to, (spec.signer && spec.signer.name) || "",
          products.map(p => `${p.model}${p.qty > 1 ? " x" + p.qty : ""}`).join(", "), Math.round(res.total * 100) / 100, savedAs];
        const logErr = await addToLog(dir, logRow);
        html += logErr ? `<div class="warn">\u26a0 ${h(logErr)}</div>` : `<div class="hint">Added to ${h(LOG_NAME)} in your quotes folder.</div>`;
      }
      if (folderNote) html += `<div class="warn">\u26a0 ${h(folderNote)}</div>`;
      html += `<a class="dl" href="${url}" download="${h(fname)}">\u2b07 ${savedAs ? "Download a copy" : "Download"} ${h(fname)}</a><div class="hint">For the PDF: open the file in Excel \u2192 File \u2192 Save as \u2192 PDF.</div>`;
      log(html);
      noEdited = false; suggestNo();
    } catch (e) {
      log("\u2717 " + h(e.message || e), "err");
    } finally { go.disabled = false; go.textContent = "Generate quotation"; }
  };
})();
