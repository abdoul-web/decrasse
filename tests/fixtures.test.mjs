import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import '../assets/js/zip.js';
import '../assets/js/xlsx-write.js';
import '../assets/js/io.js';
import '../assets/js/engine.js';

const D = globalThis.Decrasse;
const here = new URL('./fixtures/', import.meta.url);

for (const file of ['openpyxl.xlsx', 'lo/openpyxl.xlsx']) {
  test(`lecture d'un vrai classeur : ${file}`, { skip: !existsSync(new URL(file, here)) }, async () => {
    const sheets = await D.io.readXlsx(readFileSync(new URL(file, here)));
    assert.deepEqual(sheets.map((s) => s.name), ['Ventes T1', 'Notes']);
    assert.equal(sheets[1].rows[1][0], 'Bonjour & <bienvenue>');
    const res = D.engine.clean([{ name: file, rows: sheets[0].rows }]);
    const by = (l) => res.columns.findIndex((c) => c.label === l);
    assert.equal(res.log.find((l) => l.id === 'title').count, 1);
    assert.deepEqual(res.columns.map((c) => c.type), ['date', 'text', 'money', 'percent', 'bool']);
    assert.deepEqual(res.rows.map((r) => r.cells[by('Date')].v), ['2026-01-03', '2026-01-15', '2026-02-01']);
    assert.deepEqual(res.rows.map((r) => r.cells[by('Montant')].v), [1250.5, 1300, 99.99]);
    assert.deepEqual(res.rows.map((r) => r.cells[by('Taux de commande') >= 0 ? 0 : by('Taux')].v), [0.12, 0.07, 0.05]);
    assert.deepEqual(res.rows.map((r) => r.cells[by('Client')].v), ['Alpha inc.', 'Alpha inc.', 'Bêta ltée']);
  });
}
