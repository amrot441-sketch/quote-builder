/* Dragon Electronics quotation builder \u2013 core (no DOM).
 * Fills Yousouf's DP_ISL_100 slim template (template.xlsx) by editing the workbook XML,
 * exactly like build_quote.py, so every style, image and print setting is kept.
 * Runs in Chrome and in Node 18+ (uses CompressionStream / DecompressionStream "deflate-raw"). */
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
    if (eocd < 0) throw new Error("template.xlsx is not a valid file");
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
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const reEsc = s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pad2 = n => String(n).padStart(2, "0");
  function isoToDate(iso) { const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); }

  // Template map (original row numbers of template.xlsx) \u2013 identical to build_quote.py
  const HEADER_ROW = 18;
  const EXAMPLE_ROWS = range(19, 205);
  const PAY = { days: [213], down_delivery: [214], down_week: [215, 216] };
  const DELIV = { stock: [217], "3-4 months": [218], "6 months": [219, 220] };
  const TRANSPORT_ROW = 221;
  const VALID = { days: 222, until: 223, closing_days: 224 };
  const POLICY = { standard: [255, 256, 257, 258], project: [259, 260, 261] };
  const INSTALL_ROW = 262, STAINLESS_ROW = 264;
  const WARRANTY_ROWS = range(227, 254);
  function range(a, b) { const r = []; for (let i = a; i < b; i++) r.push(i); return r; }

  // ------------------------------------------------------------------ build
  async function buildQuote(templateBytes, spec, images) {
    images = images || {};
    const files = await unzip(templateBytes);
    const get = n => dec.decode(files.get(n));
    const put = (n, s) => files.set(n, enc.encode(s));
    let s = get("xl/worksheets/sheet1.xml"), sst = get("xl/sharedStrings.xml");
    const warnings = [];

    const siList = sst.match(/<si>[\s\S]*?<\/si>/g) || [];
    let nSi = siList.length; const newSi = [];
    const sidx = text => {
      text = String(text);
      const sp = text !== text.trim() || text.includes("  ") ? ' xml:space="preserve"' : "";
      newSi.push(`<si><t${sp}>${esc(text)}</t></si>`); return nSi++;
    };
    const rowXml = r => { const m = s.match(new RegExp(`<row r="${r}"[^>]*?(?:/>|>[\\s\\S]*?</row>)`)); return m ? m[0] : ""; };
    const cellText = (r, col) => {
      const m = rowXml(r).match(new RegExp(`<c r="${col}${r}"[^>]*t="s"[^>]*><v>(\\d+)</v></c>`));
      return m ? siList[+m[1]].replace(/<[^>]+>/g, "") : null;
    };

    const t = spec.terms || {};
    const DELETE = new Set(EXAMPLE_ROWS), SHOW = new Set(), TEXT = new Map(), CLEAR_B = new Set();
    const showOrDel = (cond, rows) => rows.forEach(r => (cond ? SHOW : DELETE).add(r));

    const pay = t.payment || { type: "days", days: 30 };
    const ptype = pay.type === "custom" ? "days" : pay.type;
    for (const [k, rows] of Object.entries(PAY)) showOrDel(k === ptype, rows);
    if (pay.type === "days") TEXT.set("213|C", `Strictly within ${pay.days} days as from invoice date`);
    else if (pay.type === "custom") TEXT.set("213|C", pay.text);
    else if (pay.type === "down_delivery") TEXT.set("214|C", `${pay.percent}% down payment upon order confirmation, balance before delivery`);
    else if (pay.type === "down_week") TEXT.set("215|C", `${pay.percent}% down payment upon order confirmation, payable within 1 week as from invoice date`);

    const deliv = t.delivery || "stock";
    const dkey = typeof deliv === "object" ? "stock" : deliv;
    for (const [k, rows] of Object.entries(DELIV)) showOrDel(k === dkey, rows);
    if (typeof deliv === "object") TEXT.set("217|C", deliv.custom);
    if (t.transport_fee) { SHOW.add(TRANSPORT_ROW); CLEAR_B.add(TRANSPORT_ROW); } else DELETE.add(TRANSPORT_ROW);

    const val = t.validity || { days: 14 };
    const vkey = Object.keys(val)[0];
    for (const [k, r] of Object.entries(VALID)) showOrDel(k === vkey, [r]);
    if (vkey === "until") { const d = isoToDate(val.until); TEXT.set("223|C", `Strictly valid until ${pad2(d.getUTCDate())}/${pad2(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`); }
    else if (vkey === "days") TEXT.set("222|C", `${val.days} days as from quotation date`);
    else TEXT.set("224|C", `Strictly valid ${val.closing_days} days as from closing date for submission`);

    const pol = t.policy || "standard";
    for (const [k, rows] of Object.entries(POLICY)) showOrDel(k === pol, rows);
    if (t.installation === false) DELETE.add(INSTALL_ROW);
    DELETE.add(STAINLESS_ROW);

    const norm = x => (x || "").replace(/:/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
    const wanted = new Set((spec.warranties || []).map(norm));
    let firstW = true, prevKept = false; const known = new Set();
    for (const r of WARRANTY_ROWS) {
      const c = cellText(r, "C"), sub = cellText(r, "D") || "";
      const label = norm(c), full = sub.trim().endsWith(":") ? norm((c || "") + " " + sub) : label;
      known.add(label); known.add(full);
      let keep;
      if (r === 227) keep = wanted.has("generic");
      else if (label.startsWith("* except led bar")) keep = prevKept;
      else keep = wanted.has(label) || wanted.has(full);
      if (keep) { SHOW.add(r); if (!label.startsWith("* except")) { if (firstW) firstW = false; else CLEAR_B.add(r); } }
      else DELETE.add(r);
      if (!label.startsWith("* except")) prevKept = keep;
    }
    for (const w of wanted) if (w !== "generic" && !known.has(w)) warnings.push(`Warranty line not found in template: ${w}`);

    // ---- product rows
    const plan = [], cats = [], blocks = [];
    for (const p of spec.products) if (!cats.includes(p.category)) cats.push(p.category);
    cats.forEach((cat, ci) => {
      if (ci) plan.push(["blank"]);
      plan.push(["header", cat]); let bstart = plan.length - 1;
      spec.products.filter(p => p.category === cat).forEach((p, pi) => {
        if (pi) { plan.push(["blank"]); plan.push(["sep"]); bstart = plan.length - 1; }
        plan.push(["product", p]);
        (p.features || []).forEach(f => plan.push(["line", f]));
        if (p.capacity) plan.push(["line", "Capacity:", p.capacity]);
        if (p.dims && p.dims.length) { plan.push(["line", "Dimensions & Weight:"]); p.dims.forEach(([a, b]) => plan.push(["line", a, b])); }
        for (const f of p.notes || []) {
          if (f.length > 105) {
            let cut = f.lastIndexOf(". ", 104); cut = cut >= 0 ? cut + 1 : f.lastIndexOf(" ", 104);
            plan.push(["line", f.slice(0, cut).trimEnd()]); plan.push(["line", "    " + f.slice(cut).trim()]);
          } else plan.push(["line", f]);
        }
        blocks.push([bstart, plan.length - 1]);
      });
    });
    const N = plan.length;
    const delSorted = [...DELETE].sort((a, b) => a - b);
    const newrow = r => {
      if (r <= HEADER_ROW) return r;
      if (DELETE.has(r)) throw new Error("internal: deleted row referenced " + r);
      let k = 0; for (const d of delSorted) { if (d < r) k++; else break; }
      return r + N - k;
    };
    const cell = (ref, style, o = {}) => {
      if (o.text != null) return `<c r="${ref}" s="${style}" t="s"><v>${sidx(o.text)}</v></c>`;
      if (o.formula != null) return `<c r="${ref}" s="${style}"><f>${o.formula}</f><v>${o.num}</v></c>`;
      if (o.num != null) return `<c r="${ref}" s="${style}"><v>${o.num}</v></c>`;
      return `<c r="${ref}" s="${style}"/>`;
    };
    const newRows = [], merges = [], productRows = [];
    plan.forEach((item, k) => {
      const r = HEADER_ROW + 1 + k, kind = item[0];
      if (kind === "header" || kind === "sep") {
        const cells = (kind === "header" ? cell(`B${r}`, 87, { text: item[1] }) : cell(`B${r}`, 87)) + [..."CDEFGHIJ"].map(x => cell(`${x}${r}`, 87)).join("") + cell(`L${r}`, 91);
        newRows.push(`<row r="${r}" s="85" customFormat="1" ht="18" spans="2:12">${cells}</row>`); merges.push(`B${r}:J${r}`);
      } else if (kind === "product") {
        const p = item[1], qty = p.qty || 1, price = p.price;
        const cells = cell(`B${r}`, 13, { text: p.model }) + cell(`C${r}`, 4, { text: p.name }) + cell(`E${r}`, 23) + cell(`F${r}`, 23) + cell(`G${r}`, 26)
          + cell(`H${r}`, 26, { num: qty }) + cell(`I${r}`, 58, { num: price }) + cell(`J${r}`, 58, { num: qty * price, formula: `H${r}*I${r}` }) + cell(`L${r}`, 60);
        newRows.push(`<row r="${r}" s="4" customFormat="1" spans="2:12">${cells}</row>`);
        const nBody = (p.features || []).length + (p.capacity ? 1 : 0) + (p.dims && p.dims.length ? 1 + p.dims.length : 0);
        productRows.push([r, nBody, p]);
      } else {
        const ctext = item[1], dtext = item[2];
        const cells = cell(`B${r}`, 13) + cell(`C${r}`, 7, ctext != null ? { text: ctext } : {}) + cell(`D${r}`, 7, dtext != null ? { text: dtext } : {})
          + cell(`E${r}`, 7) + cell(`F${r}`, 7) + cell(`G${r}`, 26) + cell(`H${r}`, 90) + cell(`I${r}`, 90) + cell(`J${r}`, 90) + cell(`L${r}`, 60);
        newRows.push(`<row r="${r}" s="4" customFormat="1" spans="2:12">${cells}</row>`);
      }
    });

    // ---- rebuild sheetData
    const i1 = s.indexOf("<sheetData>"), i2 = s.indexOf("</sheetData>");
    let head = s.slice(0, i1), data = s.slice(i1 + 11, i2), tail = s.slice(i2 + 12);
    const rows = data.match(/<row r="\d+"[^>]*\/>|<row r="\d+"[\s\S]*?<\/row>/g) || [];
    if (rows.join("") !== data) throw new Error("internal: unexpected sheet layout");
    const shift = (m, col, row) => col + newrow(+row);
    const out = [];
    for (let rx of rows) {
      const r = +rx.match(/<row r="(\d+)"/)[1];
      if (DELETE.has(r)) continue;
      if (SHOW.has(r)) rx = rx.replace(' hidden="1"', "");
      if (r === TRANSPORT_ROW) rx = rx.replace(/(<c r="[C-G]221" s=")45(")/g, (m, a, b) => a + "7" + b);   // drop the yellow reminder highlight
      if (CLEAR_B.has(r)) rx = rx.replace(new RegExp(`<c r="B${r}" s="(\\d+)" t="s"><v>\\d+</v></c>`), `<c r="B${r}" s="$1"/>`);
      for (const [key, txt] of TEXT) {
        const [tr, col] = key.split("|");
        if (+tr === r) {
          const re = new RegExp(`<c r="${col}${r}" s="(\\d+)" t="s"><v>\\d+</v></c>`);
          if (!re.test(rx)) throw new Error(`internal: cell ${col}${r} not found`);
          rx = rx.replace(re, (m, st) => `<c r="${col}${r}" s="${st}" t="s"><v>${sidx(txt)}</v></c>`);
        }
      }
      if (r > HEADER_ROW) {
        rx = rx.replace(/<row r="\d+"/, `<row r="${newrow(r)}"`);
        rx = rx.replace(/(?<=<c r=")([A-Z]+)(\d+)/g, shift);
      }
      out.push(rx);
      if (r === HEADER_ROW) out.push(...newRows);
    }
    data = out.join("");

    // ---- customer block, quote number, date
    const cust = spec.customer || {};
    let addr = cust.address || []; if (typeof addr === "string") addr = addr.split(/\n|,\s*(?=[A-Z])/).map(x => x.trim()).filter(Boolean);
    const fields = { C9: cust.to, C10: addr[0], C11: addr.slice(1).join(", ") || null, C12: cust.attention, C13: cust.tel, C14: cust.mobile, C15: cust.email, J10: spec.quote_no };
    for (const [ref, v] of Object.entries(fields)) {
      if (!v) continue;
      const re = new RegExp(`<c r="${ref}" s="(\\d+)"/>`);
      if (!re.test(data)) throw new Error(`internal: cell ${ref} not found`);
      data = data.replace(re, (m, st) => `<c r="${ref}" s="${st}" t="s"><v>${sidx(v)}</v></c>`);
    }
    if (cust.mobile) data = data.replace('<row r="14" hidden="1"', '<row r="14"');
    if (spec.date) {
      const serial = Math.round((isoToDate(spec.date) - Date.UTC(1899, 11, 30)) / 86400000);
      data = data.replace(/<c r="J9" s="(\d+)"\/>/, `<c r="J9" s="$1"><v>${serial}</v></c>`);
    }

    // ---- merges
    const oldM = [...tail.matchAll(/<mergeCell ref="([^"]+)"\/>/g)].map(m => m[1]);
    const keepM = oldM.filter(m => !(m.match(/\d+/g) || []).some(x => DELETE.has(+x))).map(m => m.replace(/([A-Z]+)(\d+)/g, shift));
    const allM = keepM.concat(merges);
    tail = tail.replace(/<mergeCells count="\d+">[\s\S]*?<\/mergeCells>/, `<mergeCells count="${allM.length}">` + allM.map(m => `<mergeCell ref="${m}"/>`).join("") + "</mergeCells>");
    tail = tail.replace(/<rowBreaks[\s\S]*?<\/rowBreaks>/, "");

    // ---- page breaks: never split a product block (see build_quote.py)
    const heights = new Map();
    for (const m of data.matchAll(/<row r="(\d+)"([^>]*)>/g)) {
      const attrs = m[2];
      if (attrs.includes('hidden="1"')) heights.set(+m[1], 0);
      else { const h = attrs.match(/\sht="([\d.]+)"/); heights.set(+m[1], h ? +h[1] : 14.4); }
    }
    const H = r => (heights.has(r) ? heights.get(r) : 14.4);
    const sumH = (a, b) => { let x = 0; for (let r = a; r <= b; r++) x += H(r); return x; };
    const lastRow = newrow(276);
    const PAGE_PT = 841.89 - (0.275590551181102 + 0.393700787401575) * 72;
    const CAP = PAGE_PT / 0.767 * 0.96, TITLE = sumH(1, 6);
    const keep = blocks.map(([a, b]) => [HEADER_ROW + 1 + a, HEADER_ROW + 1 + b]);
    const terms = newrow(208);
    keep.push(sumH(terms, lastRow) <= CAP - TITLE ? [terms, lastRow] : [newrow(266), lastRow]);
    const starts = new Map(keep);
    const breaks = []; let used = 0;
    for (let r = 1; r <= lastRow; r++) {
      if (starts.has(r)) {
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
    head = head.replace(/<dimension ref="A1:N(\d+)"\/>/, (m, n) => `<dimension ref="A1:N${newrow(+n)}"/>`);
    put("xl/worksheets/sheet1.xml", head + "<sheetData>" + data + "</sheetData>" + tail);

    // ---- drawings: keep logo/signature (shift signature), add product photos
    let d = get("xl/drawings/drawing1.xml"), rels = get("xl/drawings/_rels/drawing1.xml.rels");
    const anchors = d.match(/<xdr:(?:twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/xdr:(?:twoCellAnchor|oneCellAnchor)>/g);
    const body = [];
    for (let a of anchors) {
      const r0 = +a.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/)[1];
      if (r0 >= HEADER_ROW && r0 < 205) continue;
      if (r0 >= HEADER_ROW) a = a.replace(/<xdr:row>(\d+)<\/xdr:row>/g, (m, n) => `<xdr:row>${newrow(+n + 1) - 1}</xdr:row>`);
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
    const pre = d.slice(0, d.indexOf(anchors[0])), post = d.slice(d.lastIndexOf(anchors[anchors.length - 1]) + anchors[anchors.length - 1].length);
    put("xl/drawings/drawing1.xml", pre + body.join("") + post);
    put("xl/drawings/_rels/drawing1.xml.rels", rels);
    const ct = get("[Content_Types].xml");
    if (!/Extension="jpeg"/i.test(ct)) put("[Content_Types].xml", ct.replace("<Default ", '<Default Extension="jpeg" ContentType="image/jpeg"/><Default '));

    // ---- shared strings, workbook
    sst = sst.replace("</sst>", newSi.join("") + "</sst>").replace(/uniqueCount="\d+"/, `uniqueCount="${nSi}"`);
    sst = sst.replace(/ count="(\d+)"/, (m, n) => ` count="${+n + newSi.length}"`);
    put("xl/sharedStrings.xml", sst);
    let w = get("xl/workbook.xml");
    w = w.replace(/('With Disc Column'!\$A\$1:\$K\$)(\d+)/, (m, a, n) => a + newrow(+n));
    w = w.replace(/<calcPr([^>]*?)\s*\/>/, (m, a) => "<calcPr" + a + (a.includes("fullCalcOnLoad") ? "" : ' fullCalcOnLoad="1"') + "/>");
    put("xl/workbook.xml", w);

    const total = spec.products.reduce((x, p) => x + (p.qty || 1) * p.price, 0);
    return { bytes: await zip(files), warnings, total };
  }

  // ------------------------------------------------------------------ business rules
  // Warranty line for a product (template labels). Order of checks matters.
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
    return null;
  }

  // Customer initials: person -> first + last name initials; company -> first letter of each main word.
  function initials(name) {
    const titles = /^(mr|mrs|ms|miss|dr|me|m|mme|mlle)\.?$/i, skip = /^(ltd|ltee|limited|co|company|plc|inc|and|&|the|of)\.?$/i;
    const words = String(name || "").replace(/[()]/g, " ").split(/\s+/).filter(w => w && !titles.test(w));
    const isCompany = words.some(w => skip.test(w)) || words.length > 3;
    if (isCompany) return words.filter(w => !skip.test(w)).map(w => w[0]).join("").toUpperCase().slice(0, 4) || "XX";
    if (words.length >= 2) return (words[0][0] + words[words.length - 1][0]).toUpperCase();
    return (words[0] || "XX").slice(0, 2).toUpperCase();
  }

  const API = { buildQuote, warrantyFor, initials, unzip, zip };
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
  #dq-panel details{margin-top:12px}#dq-panel summary{cursor:pointer;color:#c8102e}`;
  const st = document.createElement("style"); st.textContent = css; document.head.appendChild(st);

  const today = new Date();
  const until = new Date(today); until.setDate(until.getDate() + 14);
  const panel = document.createElement("div"); panel.id = "dq-panel";
  panel.innerHTML = `
  <header><b>Quote Builder \u00b7 Dragon Electronics</b><button id="dq-close" title="Close">\u00d7</button></header>
  <div id="dq-body">
    <fieldset><legend>Customer</legend>
      <label>To (name or company) *</label><input type="text" id="dq-to" placeholder="Mr. Jean DOE">
      <label>Address</label><input type="text" id="dq-addr1" placeholder="Street"><input type="text" id="dq-addr2" placeholder="Town" style="margin-top:4px">
      <label>Attention (name, title)</label><input type="text" id="dq-att">
      <div class="row"><div><label>Tel</label><input type="text" id="dq-tel"></div><div><label>Mobile</label><input type="text" id="dq-mob"></div></div>
      <label>Email</label><input type="text" id="dq-email">
    </fieldset>
    <fieldset><legend>Products</legend>
      <label>One model per line, quantity after it (default 1)</label>
      <textarea id="dq-models" placeholder="OLED65B66LA&#10;GC-X31FFCRB x2&#10;MS4295CIS 3"></textarea>
      <label>Price</label>
      <select id="dq-price"><option value="retail">Normal retail price (incl. VAT)</option><option value="promo">Promo price where available</option></select>
    </fieldset>
    <fieldset><legend>Terms</legend>
      <label>Payment</label>
      <div class="row"><select id="dq-pay">
        <option value="days">Within \u2026 days of invoice</option>
        <option value="down_delivery">\u2026% down payment, balance before delivery</option>
        <option value="down_week">\u2026% down payment, within 1 week (Kitchen Maker)</option>
        <option value="custom">Custom text</option></select>
        <input type="number" id="dq-payn" value="30" style="max-width:80px"></div>
      <input type="text" id="dq-paytxt" placeholder="Custom payment text" style="display:none;margin-top:4px">
      <label>Delivery</label>
      <select id="dq-deliv"><option value="stock">Ex-stock, subject unsold</option><option value="3-4 months">3~4 months</option><option value="6 months">6 months + free storage (Kitchen Maker)</option></select>
      <div class="chk"><input type="checkbox" id="dq-transport"> Add the transport fee line</div>
      <label>Validity</label>
      <div class="row"><select id="dq-valid"><option value="until">Valid until a date</option><option value="days">\u2026 days from quote date</option></select>
        <input type="date" id="dq-until" value="${isoLocal(until)}"><input type="number" id="dq-vdays" value="14" style="display:none"></div>
      <div class="hint">A Saturday or Sunday is moved back to the Friday before. Check public holidays yourself.</div>
      <label>Exchange / return policy</label>
      <select id="dq-policy"><option value="standard">Standard</option><option value="project">Project / bulk quantities</option></select>
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
    $("dq-payn").style.display = $("dq-pay").value === "custom" ? "none" : "block";
    $("dq-until").style.display = $("dq-valid").value === "until" ? "block" : "none";
    $("dq-vdays").style.display = $("dq-valid").value === "days" ? "block" : "none";
  };
  ["dq-pay", "dq-valid"].forEach(id => $(id).addEventListener("change", () => {
    if (id === "dq-pay") $("dq-payn").value = $("dq-pay").value === "days" ? 30 : 50;
    upd();
  }));
  let noEdited = false;
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
    $("dq-tplstat").innerHTML = saved ? `\u2713 Saved on this computer: ${h(saved.name)} (${h(saved.date)}). Choose a file only to replace it.` : "<span style='color:#c8102e'>One-time step: choose your quotation template file (template.xlsx).</span>";
  };
  tplStatus();
  $("dq-tpl").addEventListener("change", async e => {
    const file = e.target.files[0]; if (!file) return;
    try {
      const buf = await file.arrayBuffer(), files = await DQCore.unzip(buf);
      const wb = new TextDecoder().decode(files.get("xl/workbook.xml") || new Uint8Array());
      if (!files.has("xl/worksheets/sheet1.xml") || !wb.includes("With Disc Column")) throw new Error("this is not the DP_ISL_100 quotation template");
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
    const items = $("dq-models").value.split(/\n|;/).map(l => l.trim()).filter(Boolean).map(l => {
      const m = l.match(/^([A-Za-z0-9.\-\/]+)\s*(?:[x\u00d7*]\s*|qty\s*|\s)?(\d+)?$/i);
      return m ? { model: m[1].toUpperCase(), qty: +(m[2] || 1) } : { bad: l };
    });
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
          const usePromo = $("dq-price").value === "promo" && d.promo;
          const price = usePromo ? d.promo : d.retail;
          products.push({ category: d.category, model: d.model, name: d.name, qty: it.qty, price, features: d.features, capacity: d.capacity, dims: d.dims, notes: d.notes, _d: d, _promo: usePromo });
          try { if (d.image) images[d.model] = await photo(d.image); } catch (e) { problems.push(`${d.model}: photo could not be read`); }
          line.className = "ok"; line.innerHTML = `\u2713 ${h(d.model)} \u2013 ${h(d.name)} \u2013 ${fmt(price)}${d.noStock ? " <span class='warn'>(No Stock on website)</span>" : ""}`;
          if (!d.dims.length) problems.push(`${d.model}: no dimensions on the website`);
        } catch (e) {
          line.className = "err"; line.innerHTML = `\u2717 ${h(it.model)}: ${h(e.message)}`; problems.push(`${it.model}: ${e.message}`);
        }
      }
      if (!products.length) throw new Error("No product could be read, nothing to build.");
      const t = {};
      const pv = $("dq-pay").value, pn = +$("dq-payn").value || 0;
      t.payment = pv === "custom" ? { type: "custom", text: $("dq-paytxt").value || "As agreed" } : pv === "days" ? { type: "days", days: pn } : { type: pv, percent: pn };
      t.delivery = $("dq-deliv").value; t.transport_fee = $("dq-transport").checked;
      if ($("dq-valid").value === "until") {
        const d = new Date($("dq-until").value + "T12:00:00");
        if (d.getDay() === 6) d.setDate(d.getDate() - 1); if (d.getDay() === 0) d.setDate(d.getDate() - 2);
        if (isoLocal(d) !== $("dq-until").value) log(`Validity moved to Friday ${isoLocal(d)} (the date fell on a weekend).`, "warn");
        t.validity = { until: isoLocal(d) };
      } else t.validity = { days: +$("dq-vdays").value || 14 };
      t.policy = $("dq-policy").value; t.installation = $("dq-install").checked;
      const spec = {
        customer: { to, address: [$("dq-addr1").value.trim(), $("dq-addr2").value.trim()].filter(Boolean), attention: $("dq-att").value.trim(), tel: $("dq-tel").value.trim(), mobile: $("dq-mob").value.trim(), email: $("dq-email").value.trim() },
        quote_no: $("dq-no").value.trim(), date: $("dq-date").value, terms: t,
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
      const promoRows = products.filter(p => !p._promo && p._d.promo);
      let html = `<table><tr><th>Model</th><th class="n">Qty</th><th class="n">Price</th><th class="n">Amount</th></tr>${products.map(p => `<tr><td>${h(p.model)}</td><td class="n">${p.qty}</td><td class="n">${fmt(p.price)}</td><td class="n">${fmt(p.qty * p.price)}</td></tr>`).join("")}<tr><th colspan="3">Total incl. VAT</th><th class="n">${fmt(res.total)}</th></tr></table>`;
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
