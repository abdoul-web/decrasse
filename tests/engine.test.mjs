import { test } from 'node:test';
import assert from 'node:assert/strict';
import '../assets/js/zip.js';
import '../assets/js/xlsx-write.js';
import '../assets/js/io.js';
import '../assets/js/engine.js';
import '../assets/js/sample.js';

const D = globalThis.Decrasse;
const E = D.engine;
const run = (text, opts) => E.clean([{ name: 't.csv', rows: D.io.parseCSV(text) }], opts);
const col = (res, label) => res.columns.findIndex((c) => c.label === label);
const values = (res, label) => res.rows.map((r) => r.cells[col(res, label)].v);

test('nombres : formats québécois, américains, négatifs, pourcentages', () => {
  const p = (s, ctx) => E.parseNumber(s, ctx)?.n ?? null;
  assert.equal(p('1 250,00 $'), 1250);
  assert.equal(p('$1,240.00'), 1240);
  assert.equal(p('1.234,56 €'), 1234.56);
  assert.equal(p('(430,50 $)'), -430.5);
  assert.equal(p('-12'), -12);
  assert.equal(p('12,5'), 12.5);
  assert.equal(p('1,500', { thousandsComma: true }), 1500);
  assert.equal(p('1,500'), 1.5);
  assert.equal(p('abc'), null);
  assert.equal(p('2026-01-03'), null);
  assert.deepEqual(E.parseNumber('12,5 %'), { n: 12.5, pct: true });
});

test('dates : ISO, JJ/MM, MM/JJ par contexte, texte français et anglais, séries Excel', () => {
  assert.equal(E.parseDate('2026-01-03'), '2026-01-03');
  assert.equal(E.parseDate('05/01/2026'), '2026-01-05');
  assert.equal(E.parseDate('05/01/2026', { order: 'mdy' }), '2026-05-01');
  assert.equal(E.parseDate('01/25/2026'), '2026-01-25');
  assert.equal(E.parseDate('12 janv. 2026'), '2026-01-12');
  assert.equal(E.parseDate('1er février 2026'), '2026-02-01');
  assert.equal(E.parseDate('Jan 12, 2026'), '2026-01-12');
  assert.equal(E.parseDate('lundi 2 mars 2026'), '2026-03-02');
  assert.equal(E.parseDate('31/02/2026'), null);
  assert.equal(E.parseDate(46025, { serial: true }), '2026-01-03');
  assert.equal(E.parseDate('20260103', { compact: true }), '2026-01-03');
  assert.equal(E.parseDate('Boulangerie'), null);
  assert.equal(E.dateContext(['03/25/2026', '01/02/2026']).order, 'mdy');
});

test('en-têtes : synonymes FR/EN fusionnés', () => {
  assert.equal(E.canonHeader('Qté').key, 'quantite');
  assert.equal(E.canonHeader('Montant ($)').key, 'montant');
  assert.equal(E.canonHeader('E-mail').key, 'courriel');
  assert.equal(E.canonHeader('N° de facture').key, 'no_facture');
  assert.equal(E.canonHeader('Couleur préférée :').label, 'Couleur préférée');
});

test("l'exemple complet est nettoyé comme prévu", () => {
  const res = E.clean(D.sample.map((s) => ({ name: s.name, rows: D.io.parseCSV(s.text) })));
  const log = Object.fromEntries(res.log.map((l) => [l.id, l.count]));
  assert.deepEqual(res.columns.map((c) => c.label), ['Date', 'Client', 'Courriel', 'Téléphone', 'Ville', 'Produit', 'Quantité', 'Montant', 'Source']);
  assert.deepEqual(res.columns.map((c) => c.type), ['date', 'text', 'email', 'phone', 'text', 'text', 'number', 'money', 'text']);
  assert.equal(res.rows.length, 24);
  assert.equal(log.title, 2);
  assert.equal(log.totals, 1);
  assert.equal(log.dups, 2);
  assert.equal(log.empty, 2);
  assert.deepEqual([...new Set(values(res, 'Ville'))].sort(), ['Montréal', 'Québec', 'Shawinigan', 'Trois-Rivières']);
  assert.deepEqual([...new Set(values(res, 'Produit'))].sort(), ['Alignement', 'Batterie', "Changement d'huile", 'Freins avant', 'Pneus hiver']);
  assert.ok(values(res, 'Client').includes('Boulangerie Lavoie'));
  assert.ok(!values(res, 'Client').includes('Boulangerie  Lavoie'));
  const tel = values(res, 'Téléphone');
  assert.ok(tel.includes('819 555-0142'));
  assert.ok(tel.includes('418 555-0190 poste 12'));
  assert.ok(tel.includes('555-0190'), 'numéro incomplet conservé');
  const mails = values(res, 'Courriel');
  assert.ok(mails.includes('info@boulangerielavoie.ca'));
  assert.ok(!mails.some((m) => m && /[A-Z]/.test(m)));
  const warn = res.changes.filter((c) => c.status === 'warn').map((c) => String(c.after));
  assert.ok(warn.includes('dispatch@transportgagnon'));
  assert.ok(warn.includes('555-0190'));
  const total = values(res, 'Montant').reduce((a, v) => a + (v || 0), 0);
  assert.equal(Math.round(total * 100) / 100, 12204.03);
});

test('codes avec zéros initiaux et identifiants restent du texte', () => {
  const res = run('Code client;Montant\n00123;10\n00456;20\n00789;30');
  assert.equal(res.columns[0].type, 'text');
  assert.equal(res.columns[0].label, 'N° de client');
  assert.deepEqual(values(res, 'N° de client'), ['00123', '00456', '00789']);
});

test('codes postaux, oui/non, pourcentages', () => {
  const res = run('Code postal;Actif;Taux\ng9a5h7;oui;12 %\nG8T-2K1;NON;7,5 %\nh2x 1y4;Oui;10 %');
  assert.deepEqual(values(res, 'Code postal'), ['G9A 5H7', 'G8T 2K1', 'H2X 1Y4']);
  assert.deepEqual(values(res, 'Actif'), ['Oui', 'Non', 'Oui']);
  assert.deepEqual(values(res, 'Taux'), [0.12, 0.075, 0.1]);
});

test('en-têtes répétés et options désactivées', () => {
  const text = 'Nom;Ville\nA;québec\nNom;Ville\nA;Québec\nB;Québec';
  const on = run(text);
  assert.equal(on.log.find((l) => l.id === 'repeatHeader').count, 1);
  assert.equal(on.rows.length, 2);
  const off = run(text, { dedupe: false, unifySpelling: false });
  assert.equal(off.rows.length, 3);
  assert.deepEqual(values(off, 'Ville'), ['québec', 'Québec', 'Québec']);
});

test('remplacements utilisateur : renommer, forcer un type, exclure', () => {
  const rows = D.io.parseCSV('Réf;Prix\n0012;5\n0034;6');
  const res = E.clean([{ name: 'x', rows }], {}, { 'x:ref': { label: 'Référence' }, prix: { exclude: true } });
  assert.deepEqual(res.columns.map((c) => c.label), ['Référence']);
  const forced = E.clean([{ name: 'x', rows }], {}, { 'x:ref': { type: 'number' } });
  assert.deepEqual(values(forced, 'Réf'), [12, 34]);
});

test('CSV exporté : BOM, guillemets, format Excel français', () => {
  const res = run('Client;Montant\n"Rose, & Co";1 250,50 $');
  const std = E.toCSV(res);
  assert.ok(std.startsWith('﻿Client,Montant'));
  assert.ok(std.includes('"Rose, & Co",1250.5'));
  const fr = E.toCSV(res, { delimiter: ';', decimal: ',' });
  assert.ok(fr.includes('Rose, & Co;1250,5'));
});

test('xlsx : écriture puis relecture par le lecteur maison', async () => {
  const res = E.clean(D.sample.map((s) => ({ name: s.name, rows: D.io.parseCSV(s.text) })));
  const bytes = await D.xlsx.writeXlsx(E.workbookSheets(res));
  const sheets = await D.io.readXlsx(bytes);
  assert.deepEqual(sheets.map((s) => s.name), ['Données propres', 'Corrections', 'Rapport']);
  const back = E.clean([{ name: 'back.xlsx', rows: sheets[0].rows }]);
  assert.equal(back.rows.length, res.rows.length);
  assert.deepEqual(values(back, 'Date'), values(res, 'Date'));
  assert.deepEqual(values(back, 'Montant'), values(res, 'Montant'));
  assert.ok(sheets[0].rows[1][0] instanceof Date);
});

test('sens jour/mois décidé fichier par fichier lors d\'une fusion', () => {
  const qc = D.io.parseCSV('Date;Montant\n05/01/2026;10\n25/01/2026;20');
  const us = D.io.parseCSV('Date;Montant\n02/04/2026;30\n03/20/2026;40');
  const res = E.clean([{ name: 'qc.csv', rows: qc }, { name: 'us.csv', rows: us }]);
  assert.deepEqual(values(res, 'Date'), ['2026-01-05', '2026-01-25', '2026-02-04', '2026-03-20']);
});

test('zéros initiaux rétablis et escompte en pourcentage', () => {
  const res = E.clean([{ name: 'x', rows: [['No client', 'Escompte'], ['00412', '5 %'], ['00418', 0.1], [412, '10%'], ['00433', '0 %']] }]);
  assert.deepEqual(values(res, 'N° de client'), ['00412', '00418', '00412', '00433']);
  assert.deepEqual(values(res, 'Escompte'), [0.05, 0.1, 0.1, 0]);
});

test('forme canonique : accents, casse mixte, codes courts en majuscules', () => {
  const res = run('Statut;Province;Catégorie\nlivre;qc;PIÈCES\nLivré;Qc;pièces\nLIVRÉ;QC;Pièces\nLivre;qc;pièces');
  assert.deepEqual([...new Set(values(res, 'Statut'))], ['Livré']);
  assert.deepEqual([...new Set(values(res, 'Province'))], ['QC']);
  assert.deepEqual([...new Set(values(res, 'Catégorie'))], ['Pièces']);
});

test('quantités avec unités, notes de bas de page, en-tête sur deux lignes', () => {
  const rows = [['Commande', '', 'Client', '', ''], ['N° commande', 'Qté', 'Nom', 'Ville', 'Montant'],
    ['A1', '2 unités', 'Alpha', 'Laval', 10], ['A2', '1 unité', 'Bêta', 'Laval', 20], ['A3', 3, 'Gamma', 'Lévis', 30],
    [], ['* Montants avant taxes'], ['Préparé par : Comptabilité']];
  const res = E.clean([{ name: 'x', rows }]);
  assert.deepEqual(res.columns.map((c) => c.label), ['N° de commande', 'Quantité', 'Client', 'Ville', 'Montant']);
  assert.deepEqual(values(res, 'Quantité'), [2, 1, 3]);
  assert.equal(res.rows.length, 3);
  assert.equal(res.log.find((l) => l.id === 'footer').count, 2);
});

test('deux colonnes renommées pareil sont fusionnées', () => {
  const res = E.clean([{ name: 'x', rows: [['Courriel', 'Email perso'], ['a@x.ca', ''], ['', 'b@y.ca']] }], {}, { 'x:email perso': { label: 'Courriel' } });
  assert.deepEqual(res.columns.map((c) => c.label), ['Courriel']);
  assert.deepEqual(values(res, 'Courriel'), ['a@x.ca', 'b@y.ca']);
});
