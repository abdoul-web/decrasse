/* Décrasse — lecture des fichiers : CSV / TSV / TXT, Excel .xlsx (lecteur maison), .xls (SheetJS chargé au besoin). */
(function (D) {
  'use strict';

  const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  const unxml = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) =>
    e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : +e.slice(1)) : (ENT[e] ?? m));
  const attr = (s, name) => { const m = s.match(new RegExp('(?:^|\\s)' + name + '="([^"]*)"')); return m ? unxml(m[1]) : null; };
  const stripNs = (xml) => xml.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, '<$1');

  /* ---------- CSV ---------- */
  function sniffDelimiter(text) {
    const lines = text.split(/\r?\n/).filter((l) => l.trim()).slice(0, 20);
    let best = ',', bestScore = -1;
    for (const d of [';', ',', '\t', '|']) {
      const counts = lines.map((l) => splitLine(l, d).length);
      if (!counts.length) continue;
      const max = Math.max(...counts);
      if (max < 2) continue;
      const consistent = counts.filter((c) => c === max).length;
      const score = consistent * 10 + max;
      if (score > bestScore) { bestScore = score; best = d; }
    }
    return best;
  }
  function splitLine(line, d) { // découpe approximative pour le flair (ignore les guillemets imbriqués)
    const out = []; let f = '', q = false;
    for (const ch of line) { if (ch === '"') q = !q; else if (ch === d && !q) { out.push(f); f = ''; } else f += ch; }
    out.push(f); return out;
  }
  function parseCSV(text, delim) {
    text = text.replace(/^﻿/, '');
    const d = delim || sniffDelimiter(text);
    const rows = []; let row = [], f = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; }
        else f += ch;
      } else if (ch === '"' && f === '') q = true;
      else if (ch === d) { row.push(f); f = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(f); rows.push(row); row = []; f = '';
      } else f += ch;
    }
    if (f !== '' || row.length) { row.push(f); rows.push(row); }
    return rows;
  }
  function decodeText(buf) {
    const u8 = new Uint8Array(buf);
    if (u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8);
    if (u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8);
    try { return new TextDecoder('utf-8', { fatal: true }).decode(u8); }
    catch { return new TextDecoder('windows-1252').decode(u8); }
  }

  /* ---------- XLSX ---------- */
  const BUILTIN_DATE_FMTS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57]);
  function isDateFormat(code) {
    const s = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '').replace(/_./g, '');
    return /[dmyhs]/i.test(s) && !/^general$/i.test(s.trim());
  }
  const DAY = 86400000, EPOCH_1900 = Date.UTC(1899, 11, 30), EPOCH_1904 = Date.UTC(1904, 0, 1);
  function serialToDate(n, date1904) {
    const ms = Math.round(n * DAY);
    return new Date((date1904 ? EPOCH_1904 : EPOCH_1900) + ms);
  }
  function colIndex(ref) {
    let n = 0;
    for (const ch of ref) { const c = ch.charCodeAt(0); if (c < 65 || c > 90) break; n = n * 26 + (c - 64); }
    return n - 1;
  }

  async function readXlsx(buf) {
    const z = D.zip.readZip(buf);
    const wbXml = await z.text('xl/workbook.xml');
    if (!wbXml) throw new Error("Classeur Excel incomplet : la liste des feuilles est introuvable.");
    const wb = stripNs(wbXml);
    const date1904 = /<workbookPr\b[^>]*date1904="(1|true)"/.test(wb);
    const relsXml = stripNs((await z.text('xl/_rels/workbook.xml.rels')) || '');
    const rels = {};
    for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) rels[attr(m[1], 'Id')] = attr(m[1], 'Target');

    const shared = [];
    const ssXml = await z.text('xl/sharedStrings.xml');
    if (ssXml) {
      for (const m of stripNs(ssXml).matchAll(/<si>([\s\S]*?)<\/si>/g)) {
        const body = m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
        let t = '';
        for (const tm of body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) t += tm[1];
        shared.push(unxml(t));
      }
    }

    const dateStyles = new Set();
    const stXml = await z.text('xl/styles.xml');
    if (stXml) {
      const st = stripNs(stXml);
      const custom = {};
      for (const m of st.matchAll(/<numFmt\b([^>]*)\/?>/g)) custom[+attr(m[1], 'numFmtId')] = attr(m[1], 'formatCode') || '';
      const xfs = (st.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/) || [])[1] || '';
      let i = 0;
      for (const m of xfs.matchAll(/<xf\b([^>]*?)(?:\/>|>)/g)) {
        const id = +(attr(m[1], 'numFmtId') || 0);
        if (BUILTIN_DATE_FMTS.has(id) || (custom[id] != null && isDateFormat(custom[id]))) dateStyles.add(i);
        i++;
      }
    }

    const sheets = [];
    for (const m of wb.matchAll(/<sheet\b([^>]*)\/?>/g)) {
      const name = attr(m[1], 'name') || 'Feuille';
      const state = attr(m[1], 'state');
      let target = rels[attr(m[1], 'r:id') || attr(m[1], 'id')];
      if (!target) continue;
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const xml = await z.text(target);
      if (!xml) continue;
      sheets.push({ name, hidden: state === 'hidden' || state === 'veryHidden', rows: parseSheet(stripNs(xml), shared, dateStyles, date1904) });
    }
    return sheets;
  }

  function parseSheet(xml, shared, dateStyles, date1904) {
    const rows = [];
    let r = -1, c = -1;
    const data = (xml.match(/<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/) || [])[1] || '';
    const re = /<row\b([^>]*?)(\/?)>|<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = re.exec(data))) {
      if (m[1] !== undefined) {
        const rr = attr(m[1], 'r');
        r = rr ? +rr - 1 : r + 1; c = -1;
        continue;
      }
      const a = m[3], body = m[4] || '';
      const ref = attr(a, 'r');
      if (ref) { c = colIndex(ref); const rn = ref.match(/\d+/); if (rn) r = +rn[0] - 1; } else c++;
      if (r < 0) r = 0;
      const t = attr(a, 't'), s = +(attr(a, 's') || 0);
      const vm = body.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/);
      let v = '';
      if (t === 'inlineStr') {
        let txt = '';
        for (const tm of body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) txt += tm[1];
        v = unxml(txt);
      } else if (vm) {
        const raw = unxml(vm[1]);
        if (t === 's') v = shared[+raw] ?? '';
        else if (t === 'str' || t === 'e') v = raw;
        else if (t === 'b') v = raw === '1';
        else if (t === 'd') v = new Date(raw);
        else {
          const n = Number(raw);
          v = isFinite(n) ? (dateStyles.has(s) ? serialToDate(n, date1904) : n) : raw;
        }
      }
      if (v === '' || c < 0 || c > 16383) continue;
      while (rows.length <= r) rows.push([]);
      const row = rows[r];
      while (row.length < c) row.push('');
      row[c] = v;
    }
    return rows;
  }

  /* ---------- .xls via SheetJS (secours) ---------- */
  let sheetjs = null;
  function loadScript(src) {
    return new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = src; s.async = true; s.crossOrigin = 'anonymous';
      s.onload = res; s.onerror = () => { s.remove(); rej(new Error(src)); };
      document.head.appendChild(s);
    });
  }
  async function getSheetJS() {
    if (globalThis.XLSX) return globalThis.XLSX;
    if (!sheetjs) {
      sheetjs = (async () => {
        for (const src of D.config.sheetjsSources) {
          try { await loadScript(src); if (globalThis.XLSX) return globalThis.XLSX; } catch { /* source suivante */ }
        }
        sheetjs = null;
        throw new Error("Impossible de charger le lecteur de fichiers .xls. Enregistrez le fichier en .xlsx ou en .csv depuis Excel, puis réessayez.");
      })();
    }
    return sheetjs;
  }
  async function readLegacy(buf) {
    const X = await getSheetJS();
    const wb = X.read(new Uint8Array(buf), { type: 'array', cellDates: true });
    return wb.SheetNames.map((name) => {
      const rows = X.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '', blankrows: true });
      // SheetJS crée des dates à minuit local : on les ramène en UTC pour rester cohérent avec le lecteur maison.
      for (const row of rows) for (let i = 0; i < row.length; i++) {
        const v = row[i];
        if (v instanceof Date && !isNaN(v)) row[i] = new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds()));
      }
      return { name, hidden: false, rows };
    });
  }

  /** Lit un File du navigateur. Retourne { name, sheets: [{ name, rows, hidden }] } */
  async function readFile(file) {
    const name = file.name;
    const ext = (name.match(/\.([^.]+)$/) || [])[1]?.toLowerCase() || '';
    const buf = await file.arrayBuffer();
    const head = new Uint8Array(buf.slice(0, 8));
    const isZip = head[0] === 0x50 && head[1] === 0x4b;
    const isOle = head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0;
    if (ext === 'xlsx' || ext === 'xlsm' || (isZip && ext !== 'csv')) {
      try { return { name, sheets: await readXlsx(buf) }; }
      catch (e) {
        try { return { name, sheets: await readLegacy(buf) }; } catch { throw e; }
      }
    }
    if (ext === 'xls' || ext === 'ods' || isOle) return { name, sheets: await readLegacy(buf) };
    if (ext && !['csv', 'tsv', 'txt', 'tab'].includes(ext)) {
      throw new Error(`« ${name} » n'est pas un tableur. Formats acceptés : .xlsx, .xls, .csv, .tsv, .txt.`);
    }
    const text = decodeText(buf);
    return { name, sheets: [{ name, rows: parseCSV(text, ext === 'tsv' || ext === 'tab' ? '\t' : null) }] };
  }

  D.io = { parseCSV, decodeText, readXlsx, readFile, serialToDate, isDateFormat };
})(globalThis.Decrasse || (globalThis.Decrasse = {}));
