import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import '../assets/js/zip.js';
import '../assets/js/xlsx-write.js';
import '../assets/js/io.js';
import '../assets/js/engine.js';
import '../assets/js/sample.js';
import '../assets/js/report-engine.js';
import '../assets/js/report-ui.js';

const D = globalThis.Decrasse, R = D.reportEngine;
const fixture = (rows, types = ['date', 'text', 'money']) => ({ title: 'Test',
  columns: types.map((type, i) => ({ key: String(i), type, label: ['Date', 'Groupe', 'Montant'][i] || 'Mesure' })), rows });
const base = { measure: '2', dimension: '1', date: '0', aggregate: 'sum' };

test('rapport : l’exemple réconcilie le total, les groupes et les périodes sur toutes les lignes', () => {
  const clean = D.engine.clean(D.sample.map((s) => ({ name: s.name, rows: D.io.parseCSV(s.text) })));
  const data = R.fromResult(clean, 'Exemple'), a = R.analyze(data);
  assert.equal(a.count, 24); assert.equal(a.valid, 22);
  assert.equal(Math.round(a.total * 100), 1220403);
  assert.equal(Math.round(a.categories.reduce((n, g) => n + g.value, 0) * 100), 1220403);
  assert.equal(Math.round(a.series.reduce((n, g) => n + g.value, 0) * 100), 1220403);
  assert.deepEqual(data.preparation, { read: 29, removed: 5, fixed: clean.meta.cellsFixed });
  assert.ok(!JSON.stringify(data).includes('"o":'), 'l’export ne contient pas les anciennes valeurs brutes');
});

test('rapport : zéro est une valeur ; absences, texte et non-finis sont exclus sans devenir zéro', () => {
  const data = fixture([['2026-01-01', 'A', 0], ['2026-01-02', 'A', null], ['2026-01-03', 'B', '12'], ['2026-01-04', 'B', -10], ['2026-01-05', 'C', Infinity]]);
  const a = R.analyze(data, { ...base, aggregate: 'mean' });
  assert.equal(a.count, 5); assert.equal(a.valid, 2); assert.equal(a.total, -5); assert.equal(a.summary.median, -5);
  assert.equal(a.summary.min, -10); assert.equal(a.summary.max, 0); assert.equal(a.summary.negative, 1);
  assert.equal(a.categories.find((c) => c.key === 'C').value, null);
  assert.equal(a.shareAllowed, false);
});

test('rapport : la moyenne globale n’est pas une moyenne des moyennes de groupes', () => {
  const a = R.analyze(fixture([['2026-01-01', 'A', 100], ['2026-01-01', 'B', 0], ['2026-01-01', 'B', 0], ['2026-01-01', 'B', 0]]), { ...base, aggregate: 'mean' });
  assert.equal(a.total, 25); assert.equal(a.summary.median, 0); assert.equal(a.shareAllowed, false);
});

test('rapport : périodes inclusives, jours du fichier et heures/fuseaux conservés', () => {
  const data = fixture([['2026-01-01T23:59:00-08:00', 'A', 1], ['2026-01-31T23:59:59.123Z', 'A', 2], ['2026-02-01', 'A', 3], ['2026-02-30', 'A', 4], ['2026-02-01T29:00:00', 'A', 5], [null, 'A', 6]]);
  const a = R.analyze(data, { ...base, start: '2026-01-01', end: '2026-01-31', grain: 'day' });
  assert.equal(a.count, 2); assert.equal(a.total, 3); assert.equal(a.dated.n, 3);
  assert.deepEqual(a.series.map((p) => p.key), ['2026-01-01', '2026-01-31']);
  assert.equal(R.day('2024-02-29'), '2024-02-29'); assert.equal(R.day('2026-02-29'), null);
  assert.equal(R.analyze(data, { ...base, start: '2026-03-01', end: '2026-01-01' }).invalidRange, true);
});

test('rapport : présélections basées sur la dernière date du fichier, jamais sur aujourd’hui', () => {
  const data = fixture([['2020-01-01', 'A', 1], ['2020-03-05', 'A', 2]]);
  assert.deepEqual(R.preset(data, base, 'month'), { start: '2020-03-01', end: '2020-03-05' });
  assert.deepEqual(R.preset(data, base, '30'), { start: '2020-02-05', end: '2020-03-05' });
  assert.deepEqual(R.preset(data, base, 'all'), { start: '', end: '' });
});

test('rapport : filtre de groupe et recherche changent tous les chiffres du même périmètre', () => {
  const data = fixture([['2026-01-01', 'A', 10], ['2026-01-02', 'B', 20], ['2026-02-01', 'A', 30], [null, null, 5]]);
  const a = R.analyze(data, { ...base, category: 'A', query: '2026-02' });
  assert.equal(a.total, 30); assert.equal(a.count, 1); assert.equal(a.categories.length, 1); assert.equal(a.series.length, 1);
  const all = R.analyze(data, base);
  assert.equal(all.categories.find((c) => c.key === 'Non renseigné').value, 5);
  assert.equal(all.datedSelected, 3); assert.equal(all.count, 4);
});

test('rapport : les pourcentages proposent une moyenne simple et l’unité pourcentage', () => {
  const data = fixture([[null, 'A', .1], [null, 'B', .2]], ['date', 'text', 'percent']);
  const a = R.analyze(data);
  assert.equal(a.config.aggregate, 'mean'); assert.equal(a.config.unit, 'percent');
  assert.ok(Math.abs(a.total - .15) < 1e-12);
  assert.match(D.reportUI.format(a.total, a.config), /15/);
});

test('rapport : sans nombres, sans dates ou avec zéro ligne, aucune métrique n’est inventée', () => {
  const data = fixture([['A'], ['B'], [null]], ['text']);
  const a = R.analyze(data);
  assert.equal(a.config.aggregate, 'count'); assert.equal(a.total, 3); assert.equal(a.hasDates, false);
  assert.equal(a.categories.length, 3); assert.equal(a.series.length, 0);
  const empty = R.analyze(fixture([])); assert.equal(empty.count, 0); assert.equal(empty.total, null);
  assert.equal(R.analyze(fixture([[null, 'A', null]]), base).total, null);
});

test('rapport : sommes compensées, valeurs négatives et dépassement de capacité', () => {
  const a = R.analyze(fixture([[null, 'A', .1], [null, 'A', .2], [null, 'A', .3]]), base);
  assert.equal(a.total, .6);
  const negative = R.analyze(fixture([[null, 'A', -10], [null, 'B', 20]]), base);
  assert.equal(negative.total, 10); assert.equal(negative.shareAllowed, false);
  const huge = R.analyze(fixture([[null, 'A', 1e308], [null, 'A', 1e308]]), base);
  assert.equal(huge.total, null); assert.equal(huge.summary.median, 1e308);
});

test('rapport : la distribution conserve chaque valeur, y compris négatifs et maximum', () => {
  const a = R.analyze(fixture([-100, -5, 0, 0, 5, 100].map((v) => [null, 'A', v])), base);
  assert.equal(a.distribution.reduce((n, b) => n + b.n, 0), 6);
  assert.equal(a.distribution[0].lo, -100);
  assert.equal(a.distribution.at(-1).hi, 100);
  assert.equal(a.distribution[0].n, 1);
  assert.equal(a.distribution.at(-1).n, 1);
});

test('rapport : l’HTML autonome conserve toutes les données et les filtres, sans script injecté', () => {
  const hostile = '</script><img src=x onerror="globalThis.pwned=true">&\u2028';
  const data = fixture([['2026-01-01', hostile, 10], ['2026-01-02', 'B', 20]]);
  const config = R.normalize(data, { ...base, title: hostile, category: hostile });
  const html = D.reportUI.serialize(data, config);
  assert.ok(html.startsWith('<!doctype html>'));
  assert.equal((html.match(/<\/script>/gi) || []).length, 1);
  assert.ok(!html.includes('<img src=x'));
  assert.ok(!/<script\s+src=|<link[^>]+href=/i.test(html));
  // Exécute le script exporté dans un autre contexte, sans les modules de l’application.
  const root = { classList: { add() {} } };
  const document = { getElementById: () => root };
  let captured;
  // Remplace uniquement l’appel final au montage pour inspecter le payload réel.
  const originalScript = html.match(/<script>([\s\S]*)<\/script>/)[1];
  const call = "ui.mount(document.getElementById('report'),", offset = originalScript.lastIndexOf(call);
  const script = originalScript.slice(0, offset) + 'capture(' + originalScript.slice(offset + call.length);
  const context = vm.createContext({ document, Intl, console, capture: (d, c) => { captured = { data: d, config: c }; } });
  vm.runInContext(script, context);
  assert.equal(captured.data.rows.length, 2);
  assert.equal(captured.data.rows[0][1], hostile);
  assert.equal(captured.config.category, hostile);
  assert.equal(context.pwned, undefined);
  assert.equal(R.analyze(captured.data, captured.config).total, 10);
});
