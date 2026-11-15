/* Décrasse — calculs du rapport. Aucune donnée n'est transmise au réseau. */
(function (D) {
  'use strict';
  function createReportEngine() {
    const numericTypes = ['number', 'money', 'percent'];
    const finite = (v) => typeof v === 'number' && Number.isFinite(v);
    const missing = (v) => v == null || v === '';
    const text = (v) => missing(v) ? 'Non renseigné' : String(v);
    function day(v) {
      if (typeof v !== 'string') return null;
      const m = v.match(/^(\d{4}-\d{2}-\d{2})(?:[T ](?:[01]\d|2[0-3]):[0-5]\d(?::[0-5]\d(?:\.\d+)?)?(?:Z|[+-](?:[01]\d|2[0-3]):?[0-5]\d)?)?$/);
      if (!m) return null;
      const time = Date.parse(m[1] + 'T00:00:00Z');
      return Number.isFinite(time) && new Date(time).toISOString().slice(0, 10) === m[1] ? m[1] : null;
    }
    function sum(values) {
      // Somme compensée : les cents ne disparaissent pas dans les arrondis successifs.
      let n = 0, correction = 0;
      for (const v of values) { const y = v - correction, t = n + y; correction = (t - n) - y; n = t; }
      return Number.isFinite(n) ? n : null;
    }
    function stats(values) {
      const sorted = values.filter(finite).sort((a, b) => a - b), n = sorted.length;
      if (!n) return { n: 0, sum: null, mean: null, median: null, min: null, max: null, negative: 0 };
      const total = sum(sorted), mid = Math.floor(n / 2);
      return { n, sum: total, mean: total == null ? null : total / n,
        median: n % 2 ? sorted[mid] : sorted[mid - 1] / 2 + sorted[mid] / 2,
        min: sorted[0], max: sorted[n - 1], negative: sorted.filter((v) => v < 0).length };
    }
    function histogram(values, summary) {
      if (!summary.n) return [];
      const min = summary.min, max = summary.max, count = min === max ? 1 : Math.min(8, Math.max(2, Math.ceil(Math.sqrt(summary.n))));
      const scale = Math.max(Math.abs(min), Math.abs(max), 1), lo = min / scale, hi = max / scale;
      const boundary = (i) => min * (1 - i / count) + max * (i / count);
      const bins = Array.from({ length: count }, (_, i) => ({ lo: boundary(i), hi: boundary(i + 1), n: 0 }));
      for (const v of values) { const i = count === 1 ? 0 : Math.min(count - 1, Math.floor(((v / scale - lo) / (hi - lo)) * count)); bins[i].n++; }
      return bins;
    }
    function fromResult(result, title) {
      return { version: 1, title: title || 'Mon rapport',
        columns: result.columns.map((c) => ({ key: c.key, label: c.label, type: c.type, source: !!c.source })),
        rows: result.rows.map((r) => r.cells.map((c) => c.v)),
        warnings: result.rows.map((r) => r.cells.filter((c) => c.status === 'warn' || c.status === 'removed').length),
        preparation: { read: result.meta.inputRows, removed: result.meta.removedRows, fixed: result.meta.cellsFixed } };
    }
    function choices(data) {
      const metrics = [], dimensions = [], dates = [];
      data.columns.forEach((c, i) => {
        const entry = { ...c, id: String(i) };
        if (numericTypes.includes(c.type)) metrics.push(entry);
        if (c.type === 'date') dates.push(entry);
        if (['text', 'bool', 'postal'].includes(c.type)) dimensions.push(entry);
      });
      dimensions.sort((a, b) => {
        const rank = (c) => /produit|cat[ée]gorie|statut|service|ville|r[ée]gion|[ée]quipe/i.test(c.label) ? 0 : c.source ? 3 : 1;
        return rank(a) - rank(b);
      });
      return { metrics, dimensions, dates };
    }
    function defaults(data) {
      const c = choices(data), m = c.metrics.find((x) => x.type === 'money') || c.metrics[0];
      return { title: 'La vue d’ensemble', measure: m?.id ?? '', dimension: c.dimensions[0]?.id ?? '', date: c.dates[0]?.id ?? '',
        aggregate: m ? (m.type === 'percent' ? 'mean' : 'sum') : 'count', unit: m?.type === 'percent' ? 'percent' : 'number',
        grain: 'month', start: '', end: '', category: null, query: '' };
    }
    function normalize(data, config) {
      const d = defaults(data), c = choices(data), out = { ...d, ...config };
      if (!c.metrics.some((x) => x.id === out.measure)) out.measure = d.measure;
      if (!c.dimensions.some((x) => x.id === out.dimension)) out.dimension = d.dimension;
      if (!c.dates.some((x) => x.id === out.date)) out.date = d.date;
      if (!['sum', 'mean', 'count'].includes(out.aggregate)) out.aggregate = d.aggregate;
      if (!out.measure && out.aggregate !== 'count') out.aggregate = 'count';
      if (!['number', 'CAD', 'USD', 'EUR', 'percent'].includes(out.unit)) out.unit = 'number';
      if (out.aggregate === 'count') out.unit = 'number';
      if (!['day', 'month'].includes(out.grain)) out.grain = 'month';
      out.start = day(out.start) || ''; out.end = day(out.end) || '';
      if (!out.date) out.start = out.end = '';
      out.title = String(out.title || d.title).slice(0, 120);
      out.query = String(out.query || '').slice(0, 200);
      if (out.category !== null) out.category = String(out.category);
      return out;
    }
    function extent(data, date) {
      let min = null, max = null, n = 0;
      if (date === '') return { min, max, n };
      for (const row of data.rows) { const d = day(row[+date]); if (!d) continue; n++; if (!min || d < min) min = d; if (!max || d > max) max = d; }
      return { min, max, n };
    }
    function preset(data, config, which) {
      const range = extent(data, config.date);
      if (!range.max || which === 'all') return { start: '', end: '' };
      if (which === 'month') return { start: range.max.slice(0, 7) + '-01', end: range.max };
      const start = new Date(range.max + 'T00:00:00Z'); start.setUTCDate(start.getUTCDate() - 29);
      return { start: start.toISOString().slice(0, 10), end: range.max };
    }
    function analyze(data, input) {
      const config = normalize(data, input), m = +config.measure, d = +config.dimension;
      const dated = extent(data, config.date), hasDates = config.date !== '', hasDimension = config.dimension !== '';
      const isCount = config.aggregate === 'count', query = config.query.trim().toLocaleLowerCase('fr');
      const invalidRange = !!(config.start && config.end && config.start > config.end);
      const selected = [], indices = [];
      data.rows.forEach((row, i) => {
        if (invalidRange) return;
        if (hasDates && (config.start || config.end)) {
          const v = day(row[+config.date]);
          if (!v || (config.start && v < config.start) || (config.end && v > config.end)) return;
        }
        if (hasDimension && config.category !== null && text(row[d]) !== config.category) return;
        if (query && !row.some((v) => !missing(v) && String(v).toLocaleLowerCase('fr').includes(query))) return;
        selected.push(row); indices.push(i);
      });
      const values = isCount ? [] : selected.map((r) => r[m]).filter(finite), summary = stats(values);
      const valueOf = (s, n) => isCount ? n : config.aggregate === 'mean' ? s.mean : s.sum;
      function group(keyFn) {
        const groups = new Map();
        for (const row of selected) {
          const key = keyFn(row); if (key == null) continue;
          if (!groups.has(key)) groups.set(key, { key, rows: 0, values: [] });
          const g = groups.get(key); g.rows++; if (!isCount && finite(row[m])) g.values.push(row[m]);
        }
        return [...groups.values()].map((g) => { const s = stats(g.values); return { key: g.key, rows: g.rows, valid: isCount ? g.rows : s.n, value: valueOf(s, g.rows) }; });
      }
      const categories = hasDimension ? group((r) => text(r[d])).sort((a, b) => (b.value ?? -Infinity) - (a.value ?? -Infinity) || a.key.localeCompare(b.key, 'fr')) : [];
      const series = hasDates ? group((r) => { const v = day(r[+config.date]); return v && v.slice(0, config.grain === 'month' ? 7 : 10); }).sort((a, b) => a.key.localeCompare(b.key)) : [];
      const datedSelected = selected.filter((r) => hasDates && day(r[+config.date])).length;
      const missingCells = selected.reduce((a, r) => a + r.filter(missing).length, 0);
      const warnedCells = indices.reduce((a, i) => a + (data.warnings?.[i] || 0), 0);
      const valid = isCount ? selected.length : summary.n, total = valueOf(summary, selected.length);
      const shareAllowed = config.aggregate !== 'mean' && (isCount || summary.negative === 0) && total > 0;
      const insights = [];
      if (categories.length && categories[0].value != null) {
        const top = categories[0];
        insights.push({ kind: 'leader', label: 'En tête', category: top.key, value: top.value,
          share: shareAllowed ? top.value / total : null, rows: top.rows });
      }
      if (series.length > 1) {
        const available = series.filter((p) => p.value != null);
        if (available.length) { const peak = available.reduce((a, b) => b.value > a.value ? b : a); insights.push({ kind: 'peak', label: 'Le point haut observé', period: peak.key, value: peak.value }); }
      }
      if (valid < selected.length) insights.push({ kind: 'missing', label: 'À garder en tête', missing: selected.length - valid });
      if (!insights.length) insights.push({ kind: 'scope', label: 'Votre périmètre', rows: selected.length });
      return { config, indices, count: selected.length, total, summary, valid, categories, series, dated, datedSelected, distribution: histogram(values, summary),
        missingCells, warnedCells, shareAllowed, insights, invalidRange, hasDimension, hasDates,
        measureLabel: isCount ? 'Nombre de lignes' : data.columns[m]?.label || 'Mesure',
        dimensionLabel: hasDimension ? data.columns[d].label : '', dateLabel: hasDates ? data.columns[+config.date].label : '' };
    }
    return { fromResult, choices, defaults, normalize, extent, preset, analyze, stats, day, text, finite, factory: createReportEngine };
  }
  D.reportEngine = createReportEngine();
})(globalThis.Decrasse || (globalThis.Decrasse = {}));
