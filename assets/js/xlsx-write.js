/* Décrasse — écriture de classeurs .xlsx (OOXML), sans dépendance.
   Feuilles typées : texte, nombre, montant, pourcentage, date ; en-tête figé et filtre automatique. */
(function (D) {
  'use strict';

  const esc = (s) => String(s)
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colName = (i) => { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const EPOCH = Date.UTC(1899, 11, 30);

  // Index des styles (cellXfs) déclarés dans styles.xml ci-dessous
  const S = { text: 0, header: 1, date: 2, money: 3, percent: 4, number: 5, wrap: 6 };

  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/><numFmt numFmtId="165" formatCode="#,##0.00\\ &quot;$&quot;;\\-#,##0.00\\ &quot;$&quot;"/></numFmts>
<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2A55C9"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  function cellXml(ref, v, kind) {
    if (v === null || v === undefined || v === '') return '';
    if (kind === 'date') {
      const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m) return `<c r="${ref}" s="${S.date}"><v>${(Date.UTC(+m[1], +m[2] - 1, +m[3]) - EPOCH) / 86400000}</v></c>`;
    }
    if (typeof v === 'number' && isFinite(v)) {
      const s = kind === 'money' ? S.money : kind === 'percent' ? S.percent : S.number;
      return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
    }
    if (typeof v === 'boolean') return `<c r="${ref}" t="inlineStr"><is><t>${v ? 'Oui' : 'Non'}</t></is></c>`;
    const s = kind === 'wrap' ? ` s="${S.wrap}"` : '';
    return `<c r="${ref}" t="inlineStr"${s}><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  }

  function sheetXml(sheet) {
    const cols = sheet.columns;
    const nRows = sheet.rows.length + 1;
    const last = colName(Math.max(0, cols.length - 1));
    const widths = cols.map((c, i) => {
      let w = String(c.label).length + 4;
      for (let r = 0; r < Math.min(sheet.rows.length, 300); r++) {
        const v = sheet.rows[r][i];
        const len = c.kind === 'date' ? 11 : c.kind === 'money' ? String(Math.round(Math.abs(v || 0))).length + 7 : String(v ?? '').length + 2;
        if (len > w) w = len;
      }
      return Math.min(c.maxWidth || 60, Math.max(8, w));
    });
    let x = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:${last}${nRows}"/>
<sheetViews><sheetView workbookViewId="0"${sheet.selected ? ' tabSelected="1"' : ''}><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>
<sheetData>`;
    x += `<row r="1">${cols.map((c, i) => `<c r="${colName(i)}1" t="inlineStr" s="${S.header}"><is><t xml:space="preserve">${esc(c.label)}</t></is></c>`).join('')}</row>`;
    const parts = [x];
    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r], rn = r + 2;
      let cells = '';
      for (let i = 0; i < cols.length; i++) cells += cellXml(colName(i) + rn, row[i], cols[i].kind);
      parts.push(`<row r="${rn}">${cells}</row>`);
    }
    parts.push('</sheetData>');
    if (sheet.autofilter && cols.length) parts.push(`<autoFilter ref="A1:${last}${nRows}"/>`);
    parts.push('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>');
    return parts.join('');
  }

  function definedNames(sheets, names) {
    const defs = sheets.map((s, i) => s.autofilter && s.columns.length
      ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${esc(names[i].replace(/'/g, "''"))}'!$A$1:$${colName(s.columns.length - 1)}$${s.rows.length + 1}</definedName>`
      : '').join('');
    return defs ? `<definedNames>${defs}</definedNames>` : '';
  }

  /** sheets : [{ name, columns: [{ label, kind: 'text'|'number'|'money'|'percent'|'date'|'wrap' }], rows: [[...]], autofilter }] */
  async function writeXlsx(sheets) {
    const used = new Set();
    const names = sheets.map((s) => {
      let n = String(s.name).replace(/[\\/?*[\]:]/g, ' ').slice(0, 31).trim() || 'Feuille';
      let k = n, i = 2;
      while (used.has(k.toLowerCase())) k = n.slice(0, 28) + ' ' + i++;
      used.add(k.toLowerCase());
      return k;
    });
    sheets[0].selected = true;
    const files = [
      { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>` },
      { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>` },
      { name: 'docProps/app.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Décrasse</Application></Properties>` },
      { name: 'docProps/core.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Décrasse</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}</dcterms:created></cp:coreProperties>` },
      { name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${names.map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${definedNames(sheets, names)}</workbook>` },
      { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
      { name: 'xl/styles.xml', data: STYLES },
      ...sheets.map((s, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s) })),
    ];
    return D.zip.writeZip(files);
  }

  D.xlsx = { writeXlsx, colName };
})(globalThis.Decrasse || (globalThis.Decrasse = {}));
