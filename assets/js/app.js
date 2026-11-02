/* Décrasse — interface. Relie les fichiers, le moteur de nettoyage, l'aperçu et les exports. */
(function (D) {
  'use strict';
  const E = D.engine;
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf = new Intl.NumberFormat('fr-CA');
  const pct = (a, b) => (b ? Math.round((100 * a) / b) : 0);
  const plural = (n, one, many) => `${nf.format(n)} ${n > 1 ? many : one}`;
  const nextFrame = () => new Promise((r) => requestAnimationFrame(() => setTimeout(r, 0)));

  const state = {
    files: [],
    options: { dedupe: true, unifySpelling: true, dropTotals: true, dropEmptyColumns: true, sortByDate: false, addSource: true, skipTitleRows: true },
    overrides: {},
    result: null,
    ui: { tab: 'preview', search: '', filter: 'all', changeFilter: 'all', dateFormat: 'iso', csv: 'std', ch: {} },
  };
  let seq = 0;

  /* ---------- Messages ---------- */
  let toastTimer;
  function toast(msg, isError) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.toggle('err', !!isError);
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, isError ? 7000 : 3500);
  }
  async function busy(on, text) {
    $('#busyText').textContent = text || 'Nettoyage en cours…';
    $('#busy').hidden = !on;
    if (on) await nextFrame();
  }

  /* ---------- Fichiers ---------- */
  const hasData = (sheet) => sheet.rows.some((r) => r.some((c) => c !== '' && c != null));
  function inputs() {
    const out = [];
    for (const f of state.files) {
      const usable = f.sheets.filter((s) => !s.hidden && hasData(s));
      const list = f.sheet === 'all' ? usable : [f.sheets[f.sheet]].filter(Boolean);
      for (const s of list) out.push({ name: f.sheets.length > 1 ? `${f.name} › ${s.name}` : f.name, rows: s.rows });
    }
    return out;
  }

  async function addFiles(list) {
    const files = [...list];
    if (!files.length) return;
    const tooBig = files.filter((f) => f.size > D.config.maxFileMB * 1024 * 1024);
    if (tooBig.length) toast(`« ${tooBig[0].name} » dépasse ${D.config.maxFileMB} Mo. Le traitement peut être lent.`, false);
    await busy(true, files.length > 1 ? `Lecture de ${files.length} fichiers…` : 'Lecture du fichier…');
    const errors = [];
    for (const file of files) {
      try {
        const read = await D.io.readFile(file);
        const first = read.sheets.findIndex((s) => !s.hidden && hasData(s));
        if (first < 0) { errors.push(`« ${file.name} » est vide.`); continue; }
        state.files.push({ id: ++seq, name: read.name, sheets: read.sheets, sheet: first });
      } catch (e) {
        errors.push(e && e.message ? e.message : `Impossible de lire « ${file.name} ».`);
      }
    }
    if (errors.length) toast(errors.join(' '), true);
    await run();
  }

  function loadSample() {
    state.files = D.sample.map((s) => ({ id: ++seq, name: s.name, sheets: [{ name: s.name, rows: D.io.parseCSV(s.text) }], sheet: 0 }));
    state.overrides = {};
    run();
  }

  function reset() {
    state.files = []; state.overrides = {}; state.result = null; state.ui.ch = {};
    $('#search').value = ''; state.ui.search = '';
    show(false);
    window.scrollTo({ top: 0 });
    if (location.hash) history.replaceState(null, '', location.pathname + location.search);
  }

  /* ---------- Nettoyage ---------- */
  async function run() {
    if (!state.files.length) { show(false); return; }
    const ins = inputs();
    const total = ins.reduce((a, i) => a + i.rows.length, 0);
    if (total > 15000) await busy(true, `Nettoyage de ${nf.format(total)} lignes…`);
    try {
      state.result = E.clean(ins, state.options, state.overrides);
    } catch (e) {
      console.error(e);
      toast("Le nettoyage a échoué sur ce fichier. Vérifiez qu'il s'agit bien d'un tableau, puis réessayez.", true);
      await busy(false);
      return;
    }
    await busy(false);
    show(true);
    render();
  }

  function show(ws) {
    const was = !$('#workspace').hidden;
    $('#landing').hidden = ws;
    $('#workspace').hidden = !ws;
    if (ws && !was) window.scrollTo({ top: 0 });
  }

  /* ---------- Rendu ---------- */
  function render() {
    renderTop(); renderKpis(); renderNotes(); renderJournal(); renderTabs();
    $('#nChanges').textContent = nf.format(state.result.changes.length);
    $('#nColumns').textContent = nf.format(state.result.columns.length);
  }

  function renderTop() {
    const r = state.result;
    $('#wsTitle').textContent = state.files.length > 1 ? `${state.files.length} fichiers fusionnés` : state.files[0].name;
    $('#wsSub').textContent = `${plural(r.meta.outputRows, 'ligne propre', 'lignes propres')} sur ${nf.format(r.meta.inputRows)} lues · ${plural(r.columns.length, 'colonne', 'colonnes')}`;
    $('#chips').innerHTML = state.files.map((f) => {
      const usable = f.sheets.map((s, i) => ({ s, i })).filter(({ s }) => !s.hidden && hasData(s));
      const info = r.meta.inputs.filter((x) => x.name === f.name || x.name.startsWith(f.name + ' › '));
      const rows = info.reduce((a, x) => a + x.rows, 0);
      const sel = usable.length > 1
        ? `<select data-sheet="${f.id}" aria-label="Feuille à utiliser pour ${esc(f.name)}">${usable.map(({ s, i }) => `<option value="${i}"${f.sheet === i ? ' selected' : ''}>${esc(s.name)}</option>`).join('')}<option value="all"${f.sheet === 'all' ? ' selected' : ''}>Toutes les feuilles</option></select>`
        : '';
      return `<span class="chip"><span class="nm" title="${esc(f.name)}">${esc(f.name)}</span>${sel}<span class="ct">${plural(rows, 'ligne', 'lignes')}</span><button type="button" data-remove="${f.id}" aria-label="Retirer ${esc(f.name)}">×</button></span>`;
    }).join('');
  }

  function renderKpis() {
    const r = state.result, L = Object.fromEntries(r.log.map((l) => [l.id, l.count]));
    const attention = L.warn + L.removed;
    $('#kpis').innerHTML = [
      ['', nf.format(r.meta.outputRows), 'Lignes propres', `sur ${nf.format(r.meta.inputRows)} lues`],
      ['k-fix', `<span>${nf.format(r.meta.cellsFixed)}</span>`, 'Cellules corrigées', `${pct(r.meta.cellsFixed, r.meta.cellsTotal)} % du tableau`],
      ['', nf.format(L.dups + L.empty + L.totals + L.repeatHeader), 'Lignes retirées', `${plural(L.dups, 'doublon', 'doublons')}, ${plural(L.totals, 'total', 'totaux')}, ${plural(L.empty, 'vide', 'vides')}`],
      ['k-warn' + (attention ? '' : ' zero'), nf.format(attention), 'À vérifier', attention ? 'Valeurs signalées ou illisibles' : 'Aucune valeur douteuse'],
    ].map(([cls, v, label, sub]) => `<div class="kpi ${cls}"><span class="eyebrow">${label}</span><span class="v">${v}</span><span class="s">${esc(sub)}</span></div>`).join('');
  }

  function renderNotes() {
    const n = state.result.notes;
    $('#notes').hidden = !n.length;
    $('#notes').innerHTML = n.map((x) => `<p>${esc(x)}</p>`).join('');
  }

  function renderJournal() {
    const items = state.result.log;
    const zero = items.filter((l) => !l.count).length;
    $('#journal').innerHTML = items.map((l) => `<li class="k-${l.kind}${l.count ? '' : ' zero'}"><span>${esc(l.label)}</span><span class="c">${nf.format(l.count)}</span></li>`).join('')
      + (zero === items.length ? '<li><span>Aucune correction nécessaire : ce fichier était déjà propre.</span></li>' : '');
  }

  /* ----- Onglets ----- */
  const TABS = ['preview', 'changes', 'columns', 'charts'];
  function setTab(t, focus) {
    state.ui.tab = t;
    for (const x of TABS) {
      const b = $('#tab-' + x);
      b.setAttribute('aria-selected', String(x === t));
      b.tabIndex = x === t ? 0 : -1;
      $('#panel-' + x).hidden = x !== t;
    }
    if (focus) $('#tab-' + t).focus();
    renderTabs();
  }
  function renderTabs() {
    if (!state.result) return;
    ({ preview: renderPreview, changes: renderChanges, columns: renderColumns, charts: renderCharts })[state.ui.tab]();
  }

  const cellClass = (col) => (['number', 'money', 'percent'].includes(col.type) ? 'num' : ['date', 'phone', 'postal'].includes(col.type) ? 'mono' : '');
  const show_ = (col, v) => E.display(col, v, { dateFormat: state.ui.dateFormat });

  function renderPreview() {
    const r = state.result, cols = r.columns;
    const q = E.norm(state.ui.search || '');
    const f = state.ui.filter;
    const matches = [];
    for (let i = 0; i < r.rows.length; i++) {
      const row = r.rows[i];
      if (f === 'fixed' && !row.cells.some((c) => c.fix)) continue;
      if (f === 'attention' && !row.cells.some((c) => c.status === 'warn' || c.status === 'removed')) continue;
      if (q && !row.cells.some((c, ci) => c.v != null && E.norm(show_(cols[ci], c.v)).includes(q))) continue;
      matches.push(i);
    }
    const lim = D.config.previewRows;
    const shown = matches.slice(0, lim);
    $('#previewCount').textContent = matches.length > lim
      ? `${nf.format(lim)} premières lignes sur ${nf.format(matches.length)} · le fichier téléchargé contient tout`
      : plural(matches.length, 'ligne', 'lignes');
    if (!shown.length) { $('#previewTable').innerHTML = `<tbody><tr><td class="empty-state">${q || f !== 'all' ? 'Aucune ligne ne correspond.' : 'Aucune donnée.'}</td></tr></tbody>`; return; }
    let h = `<thead><tr><th class="rn" scope="col">Ligne</th>${cols.map((c) => `<th scope="col" class="${cellClass(c) === 'num' ? 'num' : ''}">${esc(c.label)}${E.TYPE_LABEL[c.type] !== c.label ? `<span class="t">${E.TYPE_LABEL[c.type]}</span>` : ''}</th>`).join('')}</tr></thead><tbody>`;
    for (const i of shown) {
      const row = r.rows[i];
      h += `<tr><td class="rn">${i + 2}</td>`;
      row.cells.forEach((c, ci) => {
        const col = cols[ci];
        const st = c.status === 'removed' ? 'bad' : c.status === 'warn' ? 'warn' : c.fix ? 'fx' : c.v == null ? 'nul' : '';
        const tip = c.fix ? ` title="${esc(c.fix + ' · avant : « ' + E.rawDisplay(c.o) + ' »')}"` : '';
        const txt = c.v == null ? (c.status === 'removed' ? E.rawDisplay(c.o) : '') : show_(col, c.v);
        h += `<td class="${cellClass(col)} ${st}"${tip}>${esc(txt)}</td>`;
      });
      h += '</tr>';
    }
    $('#previewTable').innerHTML = h + '</tbody>';
  }

  function renderChanges() {
    const r = state.result;
    const sel = $('#changeFilter');
    const reasons = new Map();
    for (const ch of r.changes) for (const p of ch.reason.split(' · ')) reasons.set(p, (reasons.get(p) || 0) + 1);
    const warnN = r.changes.filter((c) => c.status === 'warn').length, badN = r.changes.filter((c) => c.status === 'removed').length;
    const opts = [['all', `Toutes les corrections (${nf.format(r.changes.length)})`]];
    if (warnN) opts.push(['@warn', `À vérifier (${nf.format(warnN)})`]);
    if (badN) opts.push(['@removed', `Illisibles, retirées (${nf.format(badN)})`]);
    for (const [k, n] of [...reasons].sort((a, b) => b[1] - a[1])) opts.push([k, `${k} (${nf.format(n)})`]);
    if (!opts.some(([k]) => k === state.ui.changeFilter)) state.ui.changeFilter = 'all';
    sel.innerHTML = opts.map(([k, l]) => `<option value="${esc(k)}"${k === state.ui.changeFilter ? ' selected' : ''}>${esc(l)}</option>`).join('');
    const f = state.ui.changeFilter;
    const list = r.changes.filter((c) => f === 'all' || (f === '@warn' ? c.status === 'warn' : f === '@removed' ? c.status === 'removed' : c.reason.split(' · ').includes(f)));
    const lim = 500;
    $('#changesCount').textContent = list.length > lim ? `${lim} premières sur ${nf.format(list.length)} · toutes sont dans le fichier Excel` : plural(list.length, 'correction', 'corrections');
    if (!list.length) { $('#changesTable').innerHTML = '<tbody><tr><td class="empty-state">Aucune correction.</td></tr></tbody>'; return; }
    $('#changesTable').innerHTML = '<thead><tr><th class="rn" scope="col">Ligne</th><th scope="col">Colonne</th><th scope="col">Avant</th><th scope="col">Après</th><th scope="col">Correction</th></tr></thead><tbody>'
      + list.slice(0, lim).map((ch) => {
        const col = r.columns[ch.col];
        const cls = ch.status === 'removed' ? 'bad' : ch.status === 'warn' ? 'warn' : 'fx';
        const before = E.rawDisplay(ch.before);
        return `<tr><td class="rn">${ch.row}</td><td>${esc(col.label)}</td><td class="mono">${esc(before === '' ? '(vide)' : before.replace(/^\s|\s$/g, '␣'))}</td><td class="${cls} ${cellClass(col)}">${esc(ch.after == null ? '(vide)' : show_(col, ch.after))}</td><td class="wrap">${esc(ch.reason)}</td></tr>`;
      }).join('') + '</tbody>';
  }

  function renderColumns() {
    const r = state.result;
    const stats = new Map(r.columns.map((c) => [c.key, c]));
    $('#colGrid').innerHTML = '<div class="colrow head"><span></span><span>Nom</span><span>Type</span><span>Contenu</span></div>' + r.allColumns.map((c) => {
      const s = stats.get(c.key);
      const on = c.source ? state.options.addSource : !c.exclude;
      const types = E.TYPES.map((t) => `<option value="${t}"${t === c.type ? ' selected' : ''}>${E.TYPE_LABEL[t]}${t === c.detected ? ' (détecté)' : ''}</option>`).join('');
      const content = !on ? 'Exclue du fichier final' : s
        ? `${plural(s.unique, 'valeur unique', 'valeurs uniques')} · ${plural(s.missing, 'vide', 'vides')}<span class="ex">${esc(s.samples.map((v) => show_(s, v)).join(' · '))}</span>`
        : 'Vide, retirée';
      return `<div class="colrow${on ? '' : ' off'}">
        <input type="checkbox" data-include="${esc(c.key)}"${on ? ' checked' : ''} aria-label="Inclure ${esc(c.label)}">
        <input type="text" id="col-${esc(c.key)}" data-label="${esc(c.key)}" value="${esc(c.label)}" aria-label="Nom de la colonne"${c.source ? ' disabled' : ''}>
        <select data-type="${esc(c.key)}" aria-label="Type de ${esc(c.label)}"${c.source ? ' disabled' : ''}>${types}</select>
        <span class="stats">${content}</span>
      </div>`;
    }).join('');
  }

  /* ----- Graphiques ----- */
  const PREF_DIM = ['ville', 'produit', 'categorie', 'region', 'statut', 'representant', 'client', 'province', '__src'];
  function renderCharts() {
    const r = state.result, cols = r.columns, ch = state.ui.ch;
    const measures = cols.map((c, i) => ({ c, i })).filter(({ c }) => ['money', 'number', 'percent'].includes(c.type));
    const dims = cols.map((c, i) => ({ c, i })).filter(({ c }) => ['text', 'bool', 'postal'].includes(c.type) && c.unique >= 2 && c.unique <= 200);
    const dates = cols.map((c, i) => ({ c, i })).filter(({ c }) => c.type === 'date');
    const ok = (list, k) => list.some(({ c }) => c.key === k);
    if (!(ch.m === '__count' || ok(measures, ch.m))) ch.m = (measures.find(({ c }) => c.type === 'money') || measures.find(({ c }) => c.type === 'number') || { c: { key: '__count' } }).c.key;
    if (!ok(dims, ch.d)) ch.d = ([...dims].sort((a, b) => {
      const pa = PREF_DIM.indexOf(a.c.key), pb = PREF_DIM.indexOf(b.c.key);
      return (pa < 0 ? 99 : pa) - (pb < 0 ? 99 : pb) || a.c.unique - b.c.unique;
    })[0] || { c: { key: '' } }).c.key;
    if (!(ch.t === '' || ok(dates, ch.t))) ch.t = dates[0] ? dates[0].c.key : '';
    $('#chMeasure').innerHTML = measures.map(({ c }) => `<option value="${esc(c.key)}"${c.key === ch.m ? ' selected' : ''}>${esc(c.label)} (somme)</option>`).join('') + `<option value="__count"${ch.m === '__count' ? ' selected' : ''}>Nombre de lignes</option>`;
    $('#chDim').innerHTML = dims.length ? dims.map(({ c }) => `<option value="${esc(c.key)}"${c.key === ch.d ? ' selected' : ''}>${esc(c.label)}</option>`).join('') : '<option value="">Aucune colonne adaptée</option>';
    $('#chDate').innerHTML = dates.map(({ c }) => `<option value="${esc(c.key)}"${c.key === ch.t ? ' selected' : ''}>${esc(c.label)}</option>`).join('') + `<option value=""${ch.t === '' ? ' selected' : ''}>Aucune</option>`;

    const mi = cols.findIndex((c) => c.key === ch.m), mcol = cols[mi];
    const val = (row) => (mi < 0 ? 1 : row.cells[mi].v || 0);
    const fmt = (v) => (mi < 0 ? nf.format(Math.round(v)) : E.display(mcol, Math.round(v * 100) / 100));
    const mlabel = mi < 0 ? 'Nombre de lignes' : mcol.label;
    const out = [];
    const totalV = r.rows.reduce((a, row) => a + val(row), 0);

    const di = cols.findIndex((c) => c.key === ch.d);
    if (di >= 0) {
      const agg = new Map();
      for (const row of r.rows) { const k = row.cells[di].v ?? '(vide)'; agg.set(k, (agg.get(k) || 0) + val(row)); }
      let arr = [...agg].sort((a, b) => b[1] - a[1]);
      if (arr.length > 12) { const rest = arr.slice(11).reduce((a, x) => a + x[1], 0); arr = [...arr.slice(0, 11), [`Autres (${arr.length - 11})`, rest]]; }
      const max = Math.max(...arr.map((x) => Math.abs(x[1])), 1e-9);
      out.push(`<section class="chart-card"><div><h3>${esc(mlabel)} par ${esc(cols[di].label.toLowerCase())}</h3><p class="sub">Total : ${esc(fmt(totalV))}</p></div><div class="hbars">${arr.map(([k, v]) => `<div class="hb"><span class="lab" title="${esc(k)}">${esc(k)}</span><span class="val">${esc(fmt(v))}</span><div class="track"><div class="fill" style="width:${Math.max(0.5, (100 * Math.abs(v)) / max)}%"></div></div></div>`).join('')}</div></section>`);
    }
    const ti = cols.findIndex((c) => c.key === ch.t);
    if (ti >= 0) out.push(timeChart(r.rows, ti, val, fmt, mlabel));
    $('#charts').innerHTML = out.join('') || '<p class="empty-state">Ajoutez une colonne de texte ou de date pour obtenir des graphiques.</p>';
  }

  function timeChart(rows, ti, val, fmt, mlabel) {
    const ds = rows.map((r) => r.cells[ti].v).filter(Boolean).sort();
    if (!ds.length) return '';
    const toD = (s) => new Date(s + 'T00:00:00Z');
    const span = (toD(ds[ds.length - 1]) - toD(ds[0])) / 86400000;
    const unit = span <= 45 ? 'day' : span <= 200 ? 'week' : span <= 365 * 6 ? 'month' : 'year';
    const keyOf = (s) => {
      const d = toD(s);
      if (unit === 'day') return s;
      if (unit === 'week') { d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7)); return d.toISOString().slice(0, 10); }
      if (unit === 'month') return s.slice(0, 7) + '-01';
      return s.slice(0, 4) + '-01-01';
    };
    const step = (s) => {
      const d = toD(s);
      if (unit === 'day') d.setUTCDate(d.getUTCDate() + 1);
      else if (unit === 'week') d.setUTCDate(d.getUTCDate() + 7);
      else if (unit === 'month') d.setUTCMonth(d.getUTCMonth() + 1);
      else d.setUTCFullYear(d.getUTCFullYear() + 1);
      return d.toISOString().slice(0, 10);
    };
    const agg = new Map();
    for (const r of rows) { const v = r.cells[ti].v; if (v) { const k = keyOf(v); agg.set(k, (agg.get(k) || 0) + val(r)); } }
    const buckets = [];
    for (let k = keyOf(ds[0]), end = keyOf(ds[ds.length - 1]); k <= end && buckets.length < 400; k = step(k)) buckets.push([k, agg.get(k) || 0]);
    const max = Math.max(...buckets.map((b) => b[1]), 0);
    const nice = (m) => { if (m <= 0) return 1; const p = 10 ** Math.floor(Math.log10(m)); const f = m / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p; };
    const top = nice(max);
    const W = 800, H = 220, L = 64, B = 26, T = 8, plotW = W - L - 8, plotH = H - B - T;
    const bw = plotW / buckets.length;
    const label = (k) => {
      const d = toD(k);
      const o = unit === 'month' ? { month: 'short', year: '2-digit' } : unit === 'year' ? { year: 'numeric' } : { day: 'numeric', month: 'short' };
      return d.toLocaleDateString('fr-CA', { ...o, timeZone: 'UTC' });
    };
    const every = Math.max(1, Math.ceil(buckets.length / 10));
    let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(mlabel)} dans le temps">`;
    for (let i = 0; i <= 4; i++) {
      const y = T + plotH - (plotH * i) / 4, v = (top * i) / 4;
      svg += `<line class="grid" x1="${L}" x2="${W - 8}" y1="${y}" y2="${y}"/><text class="axis" x="${L - 8}" y="${y + 3.5}" text-anchor="end">${esc(compact(v))}</text>`;
    }
    buckets.forEach(([k, v], i) => {
      const h = top ? (plotH * Math.max(0, v)) / top : 0;
      const x = L + i * bw + Math.min(2, bw * 0.15);
      svg += `<rect class="bar" x="${x.toFixed(1)}" y="${(T + plotH - h).toFixed(1)}" width="${Math.max(1, bw - 2 * Math.min(2, bw * 0.15)).toFixed(1)}" height="${h.toFixed(1)}" rx="${Math.min(3, bw / 4).toFixed(1)}"><title>${esc(label(k))} : ${esc(fmt(v))}</title></rect>`;
      if (i % every === 0) svg += `<text class="axis" x="${(L + i * bw + bw / 2).toFixed(1)}" y="${H - 8}" text-anchor="middle">${esc(label(k))}</text>`;
    });
    svg += '</svg>';
    const unitTxt = { day: 'par jour', week: 'par semaine', month: 'par mois', year: 'par année' }[unit];
    return `<section class="chart-card wide"><div><h3>${esc(mlabel)} ${unitTxt}</h3><p class="sub">Du ${esc(ds[0])} au ${esc(ds[ds.length - 1])}</p></div><div class="ts">${svg}</div></section>`;
  }
  const compact = (v) => new Intl.NumberFormat('fr-CA', { notation: 'compact', maximumFractionDigits: 1 }).format(v);

  /* ---------- Exports ---------- */
  function baseName() {
    if (state.files.length === 1) return state.files[0].name.replace(/\.[^.]+$/, '') + '-propre';
    return 'decrasse-fusion-' + new Date().toISOString().slice(0, 10);
  }
  // Dans un aperçu intégré (artifact claude.ai), le téléchargement passe par la capacité « downloads » de l'hôte.
  const hostDownloads = globalThis.claude && typeof globalThis.claude.use === 'function'
    ? globalThis.claude.use('downloads').catch(() => null) : Promise.resolve(null);
  async function save(blob, name) {
    const host = await hostDownloads;
    if (host) {
      try { await host.save({ filename: name, data: blob }); toast(`Téléchargé : ${name}`); }
      catch (e) { if (e && e.code !== 'declined') toast("Le téléchargement n'est pas disponible dans cet aperçu.", true); }
      return;
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name; a.rel = 'noopener';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    toast(`Téléchargé : ${name}`);
  }
  async function downloadXlsx() {
    if (!state.result) return;
    const big = state.result.rows.length > 20000;
    if (big) await busy(true, 'Préparation du fichier Excel…');
    try {
      const bytes = await D.xlsx.writeXlsx(E.workbookSheets(state.result));
      save(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + '.xlsx');
    } catch (e) {
      console.error(e);
      toast("Le fichier Excel n'a pas pu être créé. Essayez l'export CSV.", true);
    } finally { if (big) busy(false); }
  }
  function downloadCsv() {
    if (!state.result) return;
    const fr = $('#csvFormat').value === 'fr';
    const text = E.toCSV(state.result, { delimiter: fr ? ';' : ',', decimal: fr ? ',' : '.', dateFormat: state.ui.dateFormat });
    save(new Blob([text], { type: 'text/csv;charset=utf-8' }), baseName() + '.csv');
  }

  /* ---------- Événements ---------- */
  function bind() {
    $('#fileInput').addEventListener('change', (e) => { addFiles(e.target.files); e.target.value = ''; });
    $('#demoBtn').addEventListener('click', loadSample);
    $('#resetBtn').addEventListener('click', reset);
    $('#dlXlsx').addEventListener('click', downloadXlsx);
    $('#dlCsv').addEventListener('click', downloadCsv);

    for (const k of Object.keys(state.options)) {
      const el = document.getElementById('opt-' + k);
      if (!el) continue;
      el.checked = state.options[k];
      el.addEventListener('change', () => { state.options[k] = el.checked; run(); });
    }
    $('#dateFormat').addEventListener('change', (e) => { state.ui.dateFormat = e.target.value; renderTabs(); });

    $('#chips').addEventListener('click', (e) => {
      const id = e.target.closest('[data-remove]')?.dataset.remove;
      if (!id) return;
      state.files = state.files.filter((f) => String(f.id) !== id);
      state.files.length ? run() : reset();
    });
    $('#chips').addEventListener('change', (e) => {
      const id = e.target.dataset.sheet;
      if (!id) return;
      const f = state.files.find((x) => String(x.id) === id);
      f.sheet = e.target.value === 'all' ? 'all' : +e.target.value;
      run();
    });

    const tabs = $('.tabs');
    tabs.addEventListener('click', (e) => { const b = e.target.closest('[role="tab"]'); if (b) setTab(b.id.slice(4)); });
    tabs.addEventListener('keydown', (e) => {
      const i = TABS.indexOf(state.ui.tab);
      if (e.key === 'ArrowRight') { e.preventDefault(); setTab(TABS[(i + 1) % TABS.length], true); }
      if (e.key === 'ArrowLeft') { e.preventDefault(); setTab(TABS[(i + TABS.length - 1) % TABS.length], true); }
    });

    let st;
    $('#search').addEventListener('input', (e) => { clearTimeout(st); st = setTimeout(() => { state.ui.search = e.target.value; renderPreview(); }, 150); });
    $('#rowFilter').addEventListener('change', (e) => { state.ui.filter = e.target.value; renderPreview(); });
    $('#changeFilter').addEventListener('change', (e) => { state.ui.changeFilter = e.target.value; renderChanges(); });

    const grid = $('#colGrid');
    grid.addEventListener('change', (e) => {
      const t = e.target, ov = (k) => (state.overrides[k] = state.overrides[k] || {});
      if (t.dataset.include) {
        if (t.dataset.include === '__src') state.options.addSource = t.checked, ($('#opt-addSource').checked = t.checked);
        else ov(t.dataset.include).exclude = !t.checked;
      } else if (t.dataset.type) ov(t.dataset.type).type = t.value;
      else if (t.dataset.label) { const v = t.value.trim(); if (!v) { renderColumns(); return; } ov(t.dataset.label).label = v; }
      else return;
      run();
    });
    grid.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.dataset.label) e.target.blur(); });

    ['chMeasure', 'chDim', 'chDate'].forEach((id, i) => $('#' + id).addEventListener('change', (e) => { state.ui.ch[['m', 'd', 't'][i]] = e.target.value; renderCharts(); }));

    // Glisser-déposer n'importe où sur la page
    let depth = 0;
    const isFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
    window.addEventListener('dragenter', (e) => { if (!isFiles(e)) return; e.preventDefault(); depth++; $('#dragVeil').hidden = false; });
    window.addEventListener('dragover', (e) => { if (isFiles(e)) e.preventDefault(); });
    window.addEventListener('dragleave', (e) => { if (!isFiles(e)) return; depth = Math.max(0, depth - 1); if (!depth) $('#dragVeil').hidden = true; });
    window.addEventListener('drop', (e) => {
      if (!isFiles(e)) return;
      e.preventDefault(); depth = 0; $('#dragVeil').hidden = true;
      addFiles(e.dataTransfer.files);
    });
    // Coller un tableau copié depuis Excel ou Google Sheets
    window.addEventListener('paste', (e) => {
      if (e.target.closest && e.target.closest('input, textarea, select')) return;
      const text = e.clipboardData?.getData('text/plain');
      if (!text || !text.includes('\t') || text.split(/\r?\n/).length < 2) return;
      e.preventDefault();
      state.files.push({ id: ++seq, name: 'Tableau collé', sheets: [{ name: 'Tableau collé', rows: D.io.parseCSV(text, '\t') }], sheet: 0 });
      run();
    });

    const y = document.getElementById('year');
    if (y) y.textContent = new Date().getFullYear();
    if (location.hash === '#exemple') loadSample();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
  else bind();
})(globalThis.Decrasse || (globalThis.Decrasse = {}));
