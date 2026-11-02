/* Décrasse — moteur de nettoyage. Fonctions pures, sans accès au DOM (testées avec node --test). */
(function (D) {
  'use strict';

  /* ---------- Outils texte ---------- */
  const deacc = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
  const norm = (s) => deacc(s).toLowerCase().replace(/[\s  ]+/g, ' ').trim();
  const groupKey = (s) => norm(s).replace(/[-_.,'’]+/g, ' ').replace(/\s+/g, ' ').trim();
  const isBlank = (v) => v === null || v === undefined || (typeof v === 'string' && v.replace(/[\s  ​﻿]/g, '') === '');
  const MISSING = new Set(['n/a', 'na', 'n.a.', 'n/d', 'nd', 'n.d.', '-', '--', '---', '—', 'null', 'nil', 'none', 'nan', 'aucun', 'aucune', '?', '??', '#n/a', '#na', '#valeur!', '#value!', '#ref!', '#div/0!', 'vide', 'empty', 'inconnu', 'unknown', 's/o']);
  const isMissingToken = (v) => typeof v === 'string' && MISSING.has(norm(v));

  /* ---------- En-têtes ---------- */
  const SYN = {
    date: ['Date', 'date', 'date vente', 'date de vente', 'date de la vente', 'date transaction', 'date de transaction', 'date facture', 'date de facture', 'date commande', 'date de commande', 'jour', 'order date', 'invoice date', 'transaction date', 'sale date'],
    client: ['Client', 'client', 'nom client', 'nom du client', 'customer', 'customer name', 'client name', 'compte', 'nom du compte', 'account', 'account name', 'entreprise', 'company', 'societe'],
    prenom: ['Prénom', 'prenom', 'first name', 'firstname', 'given name'],
    nom: ['Nom', 'nom', 'nom de famille', 'last name', 'lastname', 'surname', 'family name'],
    courriel: ['Courriel', 'courriel', 'email', 'e mail', 'mail', 'adresse courriel', 'adresse email', 'adresse e mail', 'adresse electronique', 'email address'],
    telephone: ['Téléphone', 'telephone', 'tel', 'no tel', 'no de tel', 'no de telephone', 'numero de telephone', 'phone', 'phone number', 'cellulaire', 'cell', 'mobile', 'tel cellulaire'],
    adresse: ['Adresse', 'adresse', 'address', 'rue', 'street', 'adresse civique'],
    ville: ['Ville', 'ville', 'city', 'municipalite', 'localite', 'town'],
    province: ['Province', 'province', 'prov', 'state', 'etat province'],
    code_postal: ['Code postal', 'code postal', 'cp', 'postal code', 'zip', 'zip code', 'postcode'],
    pays: ['Pays', 'pays', 'country'],
    produit: ['Produit', 'produit', 'article', 'item', 'product', 'service', 'description', 'designation'],
    categorie: ['Catégorie', 'categorie', 'category', 'famille', 'type de produit', 'product category'],
    quantite: ['Quantité', 'quantite', 'qte', 'qty', 'quantity', 'nb', 'nombre', 'unites', 'units', 'nb unites'],
    prix: ['Prix unitaire', 'prix', 'prix unitaire', 'prix u', 'unit price', 'price'],
    montant: ['Montant', 'montant', 'total', 'prix total', 'amount', 'montant total', 'ventes', 'sales', 'revenu', 'revenus', 'revenue', 'chiffre d affaires', 'ca', 'total amount'],
    statut: ['Statut', 'statut', 'status', 'etat'],
    representant: ['Représentant', 'representant', 'vendeur', 'rep', 'sales rep', 'conseiller', 'agent', 'salesperson'],
    region: ['Région', 'region', 'territoire', 'territory'],
    no_facture: ['N° de facture', 'no facture', 'no de facture', 'numero facture', 'numero de facture', 'facture', 'invoice', 'invoice number', 'invoice no'],
    id: ['Identifiant', 'id', 'identifiant'],
    no_client: ['N° de client', 'no client', 'no de client', 'numero client', 'numero de client', 'code client', 'customer id', 'client id', 'customer number'],
    commentaire: ['Commentaire', 'commentaire', 'commentaires', 'note', 'notes', 'remarque', 'remarques', 'comment', 'comments'],
  };
  const SYN_INDEX = new Map();
  for (const [k, [label, ...words]] of Object.entries(SYN)) for (const w of words) SYN_INDEX.set(w, { key: k, label });
  const headerNorm = (h) => norm(h).replace(/\(.*?\)/g, ' ').replace(/\bn[°º]\s*/g, 'no ').replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();

  function canonHeader(raw, index) {
    const k = headerNorm(raw);
    const hit = SYN_INDEX.get(k);
    if (hit) return { key: hit.key, label: hit.label };
    const t = String(raw).replace(/[\s  ]+/g, ' ').replace(/\s*:\s*$/, '').trim();
    if (!t) { const L = D.xlsx ? D.xlsx.colName(index) : String(index + 1); return { key: 'x:#' + L, label: 'Colonne ' + L }; }
    return { key: 'x:' + k, label: t[0].toUpperCase() + t.slice(1) };
  }

  /* ---------- Nombres ---------- */
  function numberContext(values) {
    let dotDec = 0, commaDec = 0;
    for (const v of values) {
      if (typeof v !== 'string') continue;
      const s = v.replace(/[\s  $€£%]/g, '');
      const c = s.lastIndexOf(','), d = s.lastIndexOf('.');
      if (c >= 0 && d >= 0) { if (c > d) commaDec++; else dotDec++; }
      else if (/,\d{1,2}$|,\d{4,}$/.test(s)) commaDec++;
      else if (/\.\d{1,2}$|\.\d{4,}$/.test(s)) dotDec++;
    }
    return { thousandsComma: dotDec > commaDec, thousandsDot: commaDec > dotDec };
  }

  function parseNumber(v, ctx = {}) {
    if (typeof v === 'number') return isFinite(v) ? { n: v, pct: false } : null;
    if (typeof v !== 'string') return null;
    let s = v.trim();
    if (!s) return null;
    let neg = false, pct = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/[\s  ']/g, '').replace(/^(CA\$|US\$|CAD|USD|EUR)|(CAD|USD|EUR)$/i, '').replace(/[$€£¥]/g, '');
    if (s.endsWith('%')) { pct = true; s = s.slice(0, -1); }
    if (/^[-+]/.test(s)) { if (s[0] === '-') neg = !neg; s = s.slice(1); }
    else if (/\d-$/.test(s)) { neg = !neg; s = s.slice(0, -1); }
    if (!/^(\d[\d.,]*|[.,]\d+)$/.test(s)) return null;
    const c = s.lastIndexOf(','), d = s.lastIndexOf('.');
    if (c >= 0 && d >= 0) s = c > d ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    else if (c >= 0) {
      const p = s.split(',');
      s = p.length > 2 || (p[1].length === 3 && ctx.thousandsComma && p[0] !== '0') ? p.join('') : p.join('.');
    } else if (d >= 0) {
      const p = s.split('.');
      if (p.length > 2 || (p[1].length === 3 && ctx.thousandsDot && p[0] !== '0')) s = p.join('');
    }
    const n = Number(s);
    if (!isFinite(n)) return null;
    return { n: neg ? -n : n, pct };
  }

  /* ---------- Dates ---------- */
  const MONTHS = { janv: 1, janvier: 1, jan: 1, january: 1, fev: 2, fevr: 2, fevrier: 2, feb: 2, february: 2, mars: 3, mar: 3, march: 3, avr: 4, avril: 4, apr: 4, april: 4, mai: 5, may: 5, juin: 6, jun: 6, june: 6, juil: 7, juillet: 7, jul: 7, july: 7, aout: 8, aou: 8, aug: 8, august: 8, sep: 9, sept: 9, septembre: 9, september: 9, oct: 10, octobre: 10, october: 10, nov: 11, novembre: 11, november: 11, dec: 12, decembre: 12, december: 12 };
  const WEEKDAYS = new Set(['lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi', 'dimanche', 'lun', 'mer', 'jeu', 'ven', 'sam', 'dim', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun', 'le']);
  const DAY = 86400000, EPOCH = Date.UTC(1899, 11, 30);
  function iso(y, m, d) {
    if (y < 1900 || y > 2100) return null;
    const t = new Date(Date.UTC(y, m - 1, d));
    if (t.getUTCFullYear() !== y || t.getUTCMonth() !== m - 1 || t.getUTCDate() !== d) return null;
    return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }
  const DMY_RE = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4}|\d{2})(?:[\sT].*)?$/;
  function dateContext(values) {
    let aBig = 0, bBig = 0, slashes = 0;
    for (const v of values) {
      if (typeof v !== 'string') continue;
      const m = v.trim().match(DMY_RE);
      if (!m) continue;
      slashes++;
      if (+m[1] > 12) aBig++;
      if (+m[2] > 12) bBig++;
    }
    return { order: bBig > aBig ? 'mdy' : 'dmy', ambiguous: slashes > 0 && aBig === 0 && bBig === 0, conflict: aBig > 0 && bBig > 0 };
  }
  function parseDate(v, ctx = {}) {
    if (v instanceof Date) return isNaN(v) ? null : iso(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
    if (typeof v === 'number') {
      if (ctx.serial && v > 1 && v < 73051) { const t = new Date(EPOCH + Math.floor(v) * DAY); return iso(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); }
      if (ctx.compact && Number.isInteger(v) && v >= 19000101 && v <= 21001231) return parseDate(String(v), ctx);
      return null;
    }
    if (typeof v !== 'string') return null;
    const s = v.trim();
    let m;
    if ((m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:[\sT].*)?$/))) return iso(+m[1], +m[2], +m[3]);
    if (ctx.compact && (m = s.match(/^(\d{4})(\d{2})(\d{2})$/))) return iso(+m[1], +m[2], +m[3]);
    if ((m = s.match(DMY_RE))) {
      let a = +m[1], b = +m[2], y = +m[3];
      if (m[3].length === 2) y += y < 50 ? 2000 : 1900;
      if (a > 12) return iso(y, b, a);
      if (b > 12) return iso(y, a, b);
      return ctx.order === 'mdy' ? iso(y, a, b) : iso(y, b, a);
    }
    if (/[a-z]/i.test(s)) {
      const toks = norm(s).replace(/(\d)(er|e|st|nd|rd|th)\b/g, '$1').split(/[\s,.\-\/]+/).filter((t) => t && !t.includes(':') && !WEEKDAYS.has(t) && t !== 'h');
      let y = null, mo = null, d = null, junk = 0;
      for (const t of toks) {
        if (/^\d{4}$/.test(t) && y === null) y = +t;
        else if (MONTHS[t] && mo === null) mo = MONTHS[t];
        else if (/^\d{1,2}$/.test(t) && d === null) d = +t;
        else if (/^\d{1,2}h\d{0,2}$/.test(t)) continue;
        else junk++;
      }
      if (y && mo && d && junk === 0) return iso(y, mo, d);
    }
    return null;
  }

  /* ---------- Contacts ---------- */
  const EMAIL_RE = /^[a-z0-9._%+'-]+@[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}$/;
  const cleanEmail = (v) => String(v).trim().replace(/^mailto:/i, '').replace(/[\s<>]/g, '').replace(/[.,;]+$/, '').toLowerCase();
  const POSTAL_RE = /^[ABCEGHJ-NPRSTVXY]\d[ABCEGHJ-NPRSTV-Z]\d[ABCEGHJ-NPRSTV-Z]\d$/;
  const cleanPostal = (v) => String(v).toUpperCase().replace(/[\s\- ]/g, '');
  function splitPhone(v) {
    let s = String(v).trim(), ext = '';
    const m = s.match(/\s*(?:poste|post\.?|p\.|ext\.?|extension|x|#)\s*(\d{1,6})\s*$/i);
    if (m) { ext = m[1]; s = s.slice(0, m.index); }
    let digits = s.replace(/\D/g, '');
    if (digits.length === 11 && digits[0] === '1') digits = digits.slice(1);
    return { digits, ext, plausible: /^[\d\s().+\-]+$/.test(s) };
  }
  const BOOL = { oui: true, non: false, yes: true, no: false, vrai: true, faux: false, true: true, false: false, o: true, n: false, y: true };

  /* ---------- Détection de type ---------- */
  const TYPES = ['text', 'number', 'money', 'percent', 'date', 'email', 'phone', 'postal', 'bool'];
  const TYPE_LABEL = { text: 'Texte', number: 'Nombre', money: 'Montant', percent: 'Pourcentage', date: 'Date', email: 'Courriel', phone: 'Téléphone', postal: 'Code postal', bool: 'Oui / Non', empty: 'Vide' };

  function detectType(col, present) {
    const n = present.length;
    if (!n) return 'empty';
    const k = col.key, L = headerNorm(col.label);
    const strs = present.filter((v) => typeof v === 'string').map((s) => s.trim());
    const share = (c) => c / n;

    const dateish = k === 'date' || /\b(date|jour|echeance|naissance|debut|fin|expiration|livraison|day|created|updated|modifie|cree)\b/.test(L);
    const dctx = { ...dateContext(present), serial: dateish, compact: dateish };
    const dHits = present.filter((v) => parseDate(v, dctx)).length;
    if ((dateish && share(dHits) >= 0.5) || (share(dHits) >= 0.8 && present.some((v) => typeof v !== 'number'))) return 'date';

    const eHits = strs.filter((s) => EMAIL_RE.test(cleanEmail(s))).length;
    if ((k === 'courriel' && share(strs.filter((s) => s.includes('@')).length) >= 0.5) || share(eHits) >= 0.8) return 'email';

    const pHits = present.filter((v) => POSTAL_RE.test(cleanPostal(v))).length;
    if ((k === 'code_postal' && share(pHits) >= 0.5) || share(pHits) >= 0.8) return 'postal';

    const phones = present.map(splitPhone);
    const tHits = phones.filter((p) => p.digits.length === 10 && p.plausible).length;
    const formatted = strs.filter((s) => /[\s().\-]/.test(s)).length;
    if ((k === 'telephone' && share(tHits) >= 0.5) || (share(tHits) >= 0.8 && share(formatted) >= 0.5 && !/montant|prix|total|quant/.test(L))) return 'phone';

    const bHits = present.filter((v) => typeof v === 'boolean' || (typeof v === 'string' && norm(v) in BOOL)).length;
    if (share(bHits) >= 0.9 && new Set(present.map((v) => norm(v))).size <= 4 && present.some((v) => typeof v === 'boolean' || norm(v).length > 1)) return 'bool';

    const idLike = ['id', 'no_client', 'no_facture', 'code_postal', 'telephone'].includes(k) || /\b(id|code|no|numero|sku|ref|reference|matricule|nas)\b/.test(L);
    const zeroLead = strs.filter((s) => /^0\d+$/.test(s.replace(/\s/g, ''))).length;
    const longDigits = strs.filter((s) => /^\d{13,}$/.test(s)).length;
    if (!idLike && !zeroLead && !longDigits) {
      const nctx = numberContext(present);
      const parsed = present.map((v) => parseNumber(v, nctx));
      const nHits = parsed.filter(Boolean).length;
      if (share(nHits) >= 0.8) {
        const pctShare = share(parsed.filter((p) => p && p.pct).length);
        const pctHeader = /\b(taux|pourcentage|pourcent|percent|rate|marge|ratio)\b/.test(L) || String(col.label).includes('%');
        if (pctShare >= 0.5 || (pctHeader && (pctShare > 0 || parsed.every((p) => !p || Math.abs(p.n) <= 1)))) return 'percent';
        const curr = share(strs.filter((s) => /[$€£]|\b(CAD|USD|EUR)\b/i.test(s)).length);
        if (['montant', 'prix'].includes(k) || curr >= 0.2 || /\b(montant|prix|total|cout|couts|cost|price|amount|revenu|revenus|salaire|solde|budget|depense|depenses|frais|tarif)\b/.test(L) || L.includes('$')) return 'money';
        return 'number';
      }
    }
    return 'text';
  }

  /* ---------- Nettoyage d'une cellule ---------- */
  const R = {
    spaces: 'Espaces superflus retirés', invisible: 'Caractères invisibles retirés', missing: 'Valeur manquante uniformisée',
    number: 'Nombre converti', money: 'Montant converti', percent: 'Pourcentage converti', date: 'Date convertie',
    email: 'Courriel normalisé', phone: 'Téléphone formaté', postal: 'Code postal formaté', bool: 'Oui / Non uniformisé',
    badNumber: 'Valeur non numérique retirée', badDate: 'Date illisible retirée',
    warnEmail: 'Courriel invalide, à vérifier', warnPhone: 'Numéro incomplet, à vérifier', warnPostal: 'Code postal invalide, à vérifier',
    spelling: 'Orthographe uniformisée',
  };
  const LOGCAT = {
    [R.spaces]: 'spaces', [R.invisible]: 'spaces', [R.missing]: 'missing', [R.number]: 'numbers', [R.money]: 'numbers', [R.percent]: 'numbers',
    [R.date]: 'dates', [R.email]: 'contacts', [R.phone]: 'contacts', [R.postal]: 'contacts', [R.bool]: 'bools',
    [R.badNumber]: 'removed', [R.badDate]: 'removed', [R.warnEmail]: 'warn', [R.warnPhone]: 'warn', [R.warnPostal]: 'warn', [R.spelling]: 'spelling',
  };

  const asText = (v) => (v instanceof Date ? (parseDate(v) || '') : typeof v === 'boolean' ? (v ? 'Oui' : 'Non') : String(v));
  function cleanText(o) {
    const s = asText(o);
    const noInv = s.replace(/[​-‍⁠﻿­]/g, '');
    const t = noInv.replace(/[\s  ]+/g, ' ').trim();
    return { v: t || null, fix: t === s ? null : noInv !== s && t === noInv.trim() ? R.invisible : R.spaces };
  }

  function cleanCell(o, col) {
    if (isBlank(o)) return { v: null };
    if (isMissingToken(o)) return { v: null, fix: R.missing };
    const raw = typeof o === 'string' ? o : null;
    switch (col.type) {
      case 'number': case 'money': case 'percent': {
        const p = parseNumber(o, col.nctx);
        if (!p) return { v: null, fix: R.badNumber, status: 'removed' };
        let n = p.n, fix = null;
        if (col.type === 'percent') {
          if (p.pct) n = n / 100;
          else if (Math.abs(n) > 1) { n = n / 100; fix = R.percent; }
        }
        n = Math.round(n * 1e10) / 1e10;
        if (raw !== null && raw.trim() !== String(n)) fix = col.type === 'money' ? R.money : col.type === 'percent' ? R.percent : R.number;
        return { v: n, fix };
      }
      case 'date': {
        const d = parseDate(o, col.dctx);
        if (!d) return { v: null, fix: R.badDate, status: 'removed' };
        return { v: d, fix: raw !== null && raw !== d ? R.date : typeof o === 'number' ? R.date : null };
      }
      case 'email': {
        const e = cleanEmail(o);
        if (!EMAIL_RE.test(e)) { const t = cleanText(o); return { v: t.v, fix: R.warnEmail, status: 'warn' }; }
        return { v: e, fix: e !== asText(o) ? R.email : null };
      }
      case 'phone': {
        const p = splitPhone(o);
        if (p.digits.length !== 10) { const t = cleanText(o); return { v: t.v, fix: R.warnPhone, status: 'warn' }; }
        const f = `${p.digits.slice(0, 3)} ${p.digits.slice(3, 6)}-${p.digits.slice(6)}${p.ext ? ' poste ' + p.ext : ''}`;
        return { v: f, fix: f !== asText(o) ? R.phone : null };
      }
      case 'postal': {
        const c = cleanPostal(o);
        if (!POSTAL_RE.test(c)) { const t = cleanText(o); return { v: t.v, fix: R.warnPostal, status: 'warn' }; }
        const f = c.slice(0, 3) + ' ' + c.slice(3);
        return { v: f, fix: f !== asText(o) ? R.postal : null };
      }
      case 'bool': {
        const b = typeof o === 'boolean' ? o : BOOL[norm(o)];
        if (b === undefined) return cleanText(o);
        const f = b ? 'Oui' : 'Non';
        return { v: f, fix: f !== asText(o) ? R.bool : null };
      }
      default: return cleanText(o);
    }
  }

  /* ---------- Structure du fichier ---------- */
  const nonBlank = (row) => row.filter((c) => !isBlank(c)).length;
  const looksLikeValue = (c) => typeof c === 'number' || c instanceof Date || parseNumber(String(c)) !== null || parseDate(String(c)) !== null;
  function findHeader(rows) {
    const scan = rows.slice(0, 40);
    const width = Math.max(0, ...scan.map(nonBlank));
    if (!width) return -1;
    const need = Math.max(Math.min(2, width), Math.ceil(width * 0.6));
    for (let i = 0; i < scan.length; i++) {
      const cells = scan[i].filter((c) => !isBlank(c));
      if (cells.length < need) continue;
      const textual = cells.filter((c) => typeof c === 'string' && !looksLikeValue(c)).length;
      if (textual / cells.length >= 0.6) return i;
    }
    return scan.findIndex((r) => nonBlank(r) === width);
  }
  const TOTAL_RE = /^(grand\s+total|sous\s?-?\s?total|total(\s+(general|global|des ventes|final))?)\s*:?$/;
  const isTotalRow = (row) => { const first = row.find((c) => !isBlank(c)); return typeof first === 'string' && TOTAL_RE.test(norm(first)); };

  /* ---------- Pipeline ---------- */
  const DEFAULTS = { skipTitleRows: true, dropTotals: true, unifySpelling: true, dedupe: true, dropEmptyColumns: true, sortByDate: false, addSource: true };
  const LOG_DEF = [
    ['title', 'Lignes de titre ignorées', 'info'], ['headers', 'En-têtes harmonisés', 'fix'], ['repeatHeader', 'En-têtes répétés retirés', 'fix'],
    ['empty', 'Lignes vides retirées', 'fix'], ['totals', 'Lignes de total retirées', 'fix'], ['spaces', 'Espaces superflus retirés', 'fix'],
    ['missing', 'Valeurs manquantes uniformisées', 'fix'], ['numbers', 'Nombres et montants convertis', 'fix'], ['dates', 'Dates converties', 'fix'],
    ['contacts', 'Courriels, téléphones et codes postaux formatés', 'fix'], ['bools', 'Oui / Non uniformisés', 'fix'],
    ['spelling', 'Orthographe uniformisée', 'fix'], ['emptyCols', 'Colonnes vides retirées', 'fix'], ['dups', 'Doublons retirés', 'fix'],
    ['removed', 'Valeurs illisibles retirées', 'bad'], ['warn', 'Valeurs à vérifier', 'warn'],
  ];

  /**
   * inputs : [{ name, rows: any[][] }]  (une entrée par feuille ou par fichier CSV)
   * options : voir DEFAULTS ; overrides : { [cléColonne]: { label?, type?, exclude? } }
   */
  function clean(inputs, options = {}, overrides = {}) {
    const opt = { ...DEFAULTS, ...options };
    const log = Object.fromEntries(LOG_DEF.map(([id]) => [id, 0]));
    const notes = [];
    const columns = [], colIdx = new Map(), records = [], meta = { inputs: [] };

    for (const input of inputs) {
      const rows = (input.rows || []).map((r) => (Array.isArray(r) ? r : []));
      const h = findHeader(rows);
      const info = { name: input.name, rows: 0, header: h };
      meta.inputs.push(info);
      if (h < 0) { notes.push(`« ${input.name} » ne contient aucune donnée.`); continue; }
      if (opt.skipTitleRows) log.title += rows.slice(0, h).filter((r) => nonBlank(r)).length;
      const headerRow = rows[h];
      const width = Math.max(headerRow.length, ...rows.slice(h + 1, h + 200).map((r) => r.length));
      const seen = new Map();
      const hdr = [];
      for (let i = 0; i < width; i++) {
        const raw = headerRow[i] ?? '';
        if (isBlank(raw) && !rows.slice(h + 1).some((r) => !isBlank(r[i]))) { hdr.push(null); continue; }
        let c = canonHeader(isBlank(raw) ? '' : raw, i);
        const dup = seen.get(c.key) || 0;
        seen.set(c.key, dup + 1);
        if (dup) c = { key: c.key + '#' + (dup + 1), label: c.label + ' (' + (dup + 1) + ')' };
        if (String(raw) !== c.label) log.headers++;
        hdr.push(c);
        if (!colIdx.has(c.key)) { colIdx.set(c.key, columns.length); columns.push({ key: c.key, label: c.label }); }
      }
      const headerSig = headerRow.map((c) => norm(c ?? '')).join('|');
      let lastData = rows.length - 1;
      while (lastData > h && !nonBlank(rows[lastData])) lastData--;
      for (let r = h + 1; r <= lastData; r++) {
        const row = rows[r];
        if (!nonBlank(row)) { log.empty++; continue; }
        if (row.map((c) => norm(c ?? '')).join('|') === headerSig) { log.repeatHeader++; continue; }
        if (opt.dropTotals && isTotalRow(row)) { log.totals++; continue; }
        const rec = { src: input.name, srcRow: r + 1, raw: {} };
        hdr.forEach((c, i) => { if (c) rec.raw[c.key] = row[i]; });
        records.push(rec);
        info.rows++;
      }
    }
    if (opt.addSource && meta.inputs.length > 1) columns.push({ key: '__src', label: 'Source', source: true });

    // Types et contextes
    for (const col of columns) {
      const ov = overrides[col.key] || {};
      if (ov.label) col.label = ov.label;
      col.exclude = !!ov.exclude;
      if (col.source) { col.type = 'text'; continue; }
      const present = records.map((r) => r.raw[col.key]).filter((v) => !isBlank(v) && !isMissingToken(v));
      col.detected = detectType(col, present);
      col.type = ov.type && TYPES.includes(ov.type) ? ov.type : col.detected === 'empty' ? 'text' : col.detected;
      col.nctx = numberContext(present);
      const dc = dateContext(present);
      col.dctx = { ...dc, serial: true, compact: col.type === 'date' };
      if (col.type === 'date' && dc.ambiguous) notes.push(`${col.label} : les dates comme 05/01/2026 ont été lues jour/mois/année (format québécois).`);
      if (col.type === 'date' && dc.conflict) notes.push(`${col.label} : le fichier mélange les formats jour/mois et mois/jour. Vérifiez les dates en surbrillance.`);
    }
    const active = columns.filter((c) => !c.exclude);

    // Nettoyage cellule par cellule
    for (const rec of records) {
      rec.cells = active.map((col) => {
        if (col.source) return { v: rec.src };
        const o = rec.raw[col.key];
        const res = cleanCell(o, col);
        return { o, v: res.v, fix: res.fix || null, status: res.status || (res.fix ? 'fixed' : null) };
      });
    }

    // Orthographe : regroupe les variantes (casse, accents, espaces, traits d'union) d'une même valeur
    if (opt.unifySpelling) {
      active.forEach((col, ci) => {
        if (col.type !== 'text' || col.source) return;
        const vals = records.map((r) => r.cells[ci].v).filter((v) => v != null);
        if (!vals.length || vals.reduce((a, v) => a + v.length, 0) / vals.length > 40) return;
        const groups = new Map();
        for (const v of vals) { const k = groupKey(v); if (!groups.has(k)) groups.set(k, new Map()); const g = groups.get(k); g.set(v, (g.get(v) || 0) + 1); }
        const quality = (s) => (/[^\x00-\x7f]/.test(s) ? 2 : 0) + (/^\p{Lu}/u.test(s) && s !== s.toUpperCase() ? 1 : 0) + (s !== s.toUpperCase() && s !== s.toLowerCase() ? 0.5 : 0);
        const canon = new Map();
        groups.forEach((g, k) => { if (g.size > 1) canon.set(k, [...g.entries()].sort((a, b) => b[1] - a[1] || quality(b[0]) - quality(a[0]) || a[0].localeCompare(b[0]))[0][0]); });
        if (!canon.size) return;
        for (const r of records) {
          const cell = r.cells[ci];
          if (cell.v == null) continue;
          const best = canon.get(groupKey(cell.v));
          if (best && best !== cell.v) { cell.v = best; cell.fix = cell.fix ? cell.fix + ' · ' + R.spelling : R.spelling; cell.status = cell.status || 'fixed'; }
        }
      });
    }

    // Colonnes vides
    let keep = active.map((_, i) => i);
    if (opt.dropEmptyColumns) {
      keep = keep.filter((i) => active[i].source || records.some((r) => r.cells[i].v != null || r.cells[i].status === 'removed'));
      log.emptyCols = active.length - keep.length;
    }
    const outCols = keep.map((i) => active[i]);
    for (const r of records) r.cells = keep.map((i) => r.cells[i]);

    // Doublons (la colonne Source n'entre pas dans la comparaison)
    let out = records;
    if (opt.dedupe) {
      const seen = new Set();
      out = [];
      for (const r of records) {
        const sig = JSON.stringify(r.cells.map((c, i) => (outCols[i].source ? '' : c.v)));
        if (seen.has(sig)) { log.dups++; continue; }
        seen.add(sig); out.push(r);
      }
    }

    if (opt.sortByDate) {
      const di = outCols.findIndex((c) => c.type === 'date');
      if (di >= 0) out = out.map((r, i) => [r, i]).sort((a, b) => {
        const x = a[0].cells[di].v, y = b[0].cells[di].v;
        if (x === y) return a[1] - b[1];
        if (x == null) return 1; if (y == null) return -1;
        return x < y ? -1 : 1;
      }).map((p) => p[0]);
    }

    // Journal et liste des corrections
    const changes = [];
    let cellsFixed = 0;
    out.forEach((r, ri) => r.cells.forEach((c, ci) => {
      if (!c.fix) return;
      cellsFixed++;
      for (const part of c.fix.split(' · ')) { const cat = LOGCAT[part]; if (cat) log[cat]++; }
      if (changes.length < 50000) changes.push({ row: ri + 2, col: ci, reason: c.fix, before: c.o, after: c.v, status: c.status });
    }));

    for (const [i, col] of outCols.entries()) {
      const vals = out.map((r) => r.cells[i].v);
      col.missing = vals.filter((v) => v == null).length;
      col.unique = new Set(vals.filter((v) => v != null)).size;
      col.samples = [...new Set(vals.filter((v) => v != null))].slice(0, 3);
      delete col.nctx; delete col.dctx;
    }

    return {
      columns: outCols,
      allColumns: columns.map(({ key, label, type, detected, exclude, source }) => ({ key, label, type, detected, exclude, source })),
      rows: out,
      changes,
      log: LOG_DEF.map(([id, label, kind]) => ({ id, label, kind, count: log[id] })),
      notes,
      meta: { ...meta, inputRows: records.length, outputRows: out.length, cellsFixed, cellsTotal: Math.max(1, records.length * Math.max(1, outCols.length)) },
    };
  }

  /* ---------- Affichage et export ---------- */
  const nf = new Intl.NumberFormat('fr-CA', { maximumFractionDigits: 4 });
  const cf = new Intl.NumberFormat('fr-CA', { style: 'currency', currency: 'CAD' });
  const pf = new Intl.NumberFormat('fr-CA', { style: 'percent', maximumFractionDigits: 2 });
  const dmy = (isoStr) => { const [y, m, d] = isoStr.split('-'); return `${d}/${m}/${y}`; };
  function display(col, v, opts = {}) {
    if (v == null) return '';
    if (typeof v === 'number') return col.type === 'money' ? cf.format(v) : col.type === 'percent' ? pf.format(v) : nf.format(v);
    if (col.type === 'date' && opts.dateFormat === 'dmy') return dmy(v);
    return String(v);
  }
  function rawDisplay(o) {
    if (o === undefined || o === null) return '';
    if (o instanceof Date) return parseDate(o) || '';
    if (typeof o === 'boolean') return o ? 'VRAI' : 'FAUX';
    return String(o);
  }

  function toCSV(result, { delimiter = ',', decimal = '.', dateFormat = 'iso' } = {}) {
    const q = (s) => { s = String(s); return /["\r\n]/.test(s) || s.includes(delimiter) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    const val = (col, v) => {
      if (v == null) return '';
      if (typeof v === 'number') return decimal === ',' ? String(v).replace('.', ',') : String(v);
      if (col.type === 'date' && dateFormat === 'dmy') return dmy(v);
      return v;
    };
    const lines = [result.columns.map((c) => q(c.label)).join(delimiter)];
    for (const r of result.rows) lines.push(r.cells.map((c, i) => q(val(result.columns[i], c.v))).join(delimiter));
    return '﻿' + lines.join('\r\n') + '\r\n';
  }

  function workbookSheets(result) {
    const kindOf = (c) => (['money', 'percent', 'date'].includes(c.type) ? c.type : c.type === 'number' ? 'number' : 'text');
    const data = {
      name: 'Données propres', autofilter: true,
      columns: result.columns.map((c) => ({ label: c.label, kind: kindOf(c) })),
      rows: result.rows.map((r) => r.cells.map((c) => c.v)),
    };
    const corrections = {
      name: 'Corrections', autofilter: true,
      columns: [{ label: 'Ligne', kind: 'number' }, { label: 'Colonne', kind: 'text' }, { label: 'Avant', kind: 'text' }, { label: 'Après', kind: 'text' }, { label: 'Correction', kind: 'text', maxWidth: 70 }],
      rows: result.changes.map((ch) => { const col = result.columns[ch.col]; return [ch.row, col.label, rawDisplay(ch.before), display(col, ch.after), ch.reason]; }),
    };
    const report = {
      name: 'Rapport',
      columns: [{ label: 'Élément', kind: 'text', maxWidth: 52 }, { label: 'Valeur', kind: 'text', maxWidth: 70 }],
      rows: [
        ['Généré par', 'Décrasse'],
        ['Date', new Date().toISOString().slice(0, 10)],
        ['Fichiers', result.meta.inputs.map((i) => `${i.name} (${i.rows} lignes)`).join(', ')],
        ['Lignes lues', String(result.meta.inputRows)],
        ['Lignes propres', String(result.meta.outputRows)],
        ['Cellules corrigées', String(result.meta.cellsFixed)],
        ['', ''],
        ['Étape', 'Nombre'],
        ...result.log.map((l) => [l.label, String(l.count)]),
        ['', ''],
        ['Colonne', 'Type · valeurs vides'],
        ...result.columns.map((c) => [c.label, `${TYPE_LABEL[c.type]} · ${c.missing} vide${c.missing > 1 ? 's' : ''}`]),
        ...(result.notes.length ? [['', ''], ['Remarques', ''], ...result.notes.map((n) => ['', n])] : []),
      ],
    };
    return [data, corrections, report];
  }

  D.engine = { clean, parseNumber, parseDate, dateContext, numberContext, detectType, canonHeader, findHeader, display, rawDisplay, toCSV, workbookSheets, TYPES, TYPE_LABEL, norm };
})(globalThis.Decrasse || (globalThis.Decrasse = {}));
