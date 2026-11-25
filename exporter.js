/* Waqid: fill the input columns of the user's own workbook with the app's data.
   Formulas, styles, lists and sheets stay as they are. Works in browser and node. */
(function (root) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const serial = iso => { // 'YYYY-MM-DD' -> Excel serial
    const [y, m, d] = iso.split('-').map(Number);
    return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  };
  // sheet name -> {first data row, capacity, fields: [column letter, key, kind]}
  const MAP = {
    'العملاء': { cap: 100, f: [['B','name','s'],['C','phone','s'],['D','entity','s'],['E','address','s'],['F','agreeDate','d'],['G','endDate','d'],['H','notes','s']], data: 'clients' },
    'العقود': { cap: 100, f: [['B','name','s'],['C','client','s'],['D','service','s'],['E','date','d'],['F','status','s'],['G','value','n'],['H','retention','p'],['U','notes','s']], data: 'contracts' },
    'المستخلصات': { cap: 300, f: [['B','project','s'],['C','item','s'],['D','no','s'],['E','date','d'],['F','days','n'],['G','value','n'],['H','notes','s']], data: 'bills' },
    'الحركات': { cap: 1000, f: [['B','date','d'],['C','type','s'],['D','project','s'],['E','item','s'],['F','amount','n'],['G','method','s'],['H','docNo','s'],['I','notes','s']], data: 'journal' }
  };
  function cellXml(ref, style, kind, v) {
    if (v === undefined || v === null || v === '') return null;
    const s = style ? ` s="${style}"` : '';
    if (kind === 'n') return `<c r="${ref}"${s}><v>${Number(v)}</v></c>`;
    if (kind === 'p') return `<c r="${ref}"${s}><v>${Number(v) / 100}</v></c>`;
    if (kind === 'd') { if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return `<c r="${ref}"${s} t="inlineStr"><is><t>${esc(v)}</t></is></c>`; return `<c r="${ref}"${s}><v>${serial(v)}</v></c>`; }
    return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  }
  function patchSheet(xml, spec, rows) {
    rows.forEach((rec, i) => {
      const r = 7 + i;
      spec.f.forEach(([col, key, kind]) => {
        const ref = col + r;
        const out = cellXml(ref, null, kind, rec[key]);
        if (!out) return;
        const re = new RegExp(`<c r="${ref}"((?:\\s+[a-z]+="[^"]*")*)\\s*(?:/>|>(?:<v>[^<]*</v>)?</c>)`);
        const m = xml.match(re);
        if (!m) throw new Error('cell not found ' + ref);
        const sm = m[1].match(/\ss="(\d+)"/);
        xml = xml.replace(re, cellXml(ref, sm ? sm[1] : null, kind, rec[key]));
      });
    });
    return xml;
  }
  async function exportWorkbook(JSZip, templateBytes, data, settings) {
    const zip = await JSZip.loadAsync(templateBytes);
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const sheets = {};
    for (const m of wbXml.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
      const t = rels.match(new RegExp(`<Relationship [^>]*Id="${m[2]}"[^>]*Target="([^"]+)"`)) || rels.match(new RegExp(`<Relationship [^>]*Target="([^"]+)"[^>]*Id="${m[2]}"`));
      sheets[m[1].replace(/&quot;/g, '"')] = 'xl/' + t[1].replace(/^\/?(xl\/)?/, '');
    }
    const report = {};
    for (const [name, spec] of Object.entries(MAP)) {
      if (!sheets[name]) throw new Error('sheet missing: ' + name);
      const rows = data[spec.data] || [];
      if (rows.length > spec.cap) throw new Error(`${name}: ${rows.length} rows exceed the workbook table size ${spec.cap}`);
      let xml = await zip.file(sheets[name]).async('string');
      xml = patchSheet(xml, spec, rows);
      zip.file(sheets[name], xml);
      report[name] = rows.length;
    }
    // Services list: the workbook's lists sheet feeds its dropdown and its profit-by-service columns (12 slots, B5:B16).
    if (Array.isArray(data.services) && sheets['القوائم']) {
      if (data.services.length > 12) throw new Error('الحد الأقصى 12 خدمة في ملف الإكسل');
      let xml = await zip.file(sheets['القوائم']).async('string');
      for (let i = 0; i < 12; i++) {
        const ref = 'B' + (5 + i); const re = new RegExp(`<c r="${ref}"((?:\\s+[a-z]+="[^"]*")*)\\s*(?:/>|>(?:<v>[^<]*</v>)?</c>)`);
        const m = xml.match(re); if (!m) throw new Error('cell not found ' + ref);
        const sm = m[1].match(/\ss="(\d+)"/); const st = sm ? sm[1] : null;
        const v = data.services[i];
        xml = xml.replace(re, v ? cellXml(ref, st, 's', v) : `<c r="${ref}"${st ? ` s="${st}"` : ''}/>`);
      }
      zip.file(sheets['القوائم'], xml); report['القوائم (الخدمات)'] = data.services.length;
    }
    // Drop stale cached results of formulas in every sheet, so any viewer recalculates from the new data.
    for (const path of Object.values(sheets)) {
      let xml = await zip.file(path).async('string');
      xml = xml.replace(/(<f(?:\s[^>]*)?>[^<]*<\/f>|<f(?:\s[^>]*)?\/>)<v>[^<]*<\/v>/g, '$1');
      zip.file(path, xml);
    }
    // optional settings sheet cells (القوائم!K): not touched. Ask Excel to recalculate everything on open.
    zip.file('xl/workbook.xml', wbXml.replace(/<calcPr[^>]*\/>/, '<calcPr fullCalcOnLoad="1"/>'));
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { bytes, report };
  }

  /* ---------- paste-ready workbook (no template needed) ----------
     Each sheet reproduces the cell addresses of the user's financial workbook: headers on row 6, records from row 7,
     the same column letters. The user selects the data block and pastes it (values only) at the stated cell. */
  const PASTE = {
    'العملاء': { arr: 'clients', cap: 100, cols: [['B','اسم العميل','name','s'],['C','الهاتف','phone','s'],['D','الجهة أو المهنة','entity','s'],['E','العنوان','address','s'],['F','تاريخ الاتفاق','agreeDate','d'],['G','تاريخ النهاية','endDate','d'],['H','ملاحظات','notes','s']], paste: 'B7', last: 'H' },
    'العقود': { arr: 'contracts', cap: 100, cols: [['B','اسم المشروع','name','s'],['C','العميل','client','s'],['D','نوع الخدمة','service','s'],['E','تاريخ العقد','date','d'],['F','حالة العقد','status','s'],['G','قيمة العقد','value','n'],['H','نسبة المحتجز','retention','p'],['U','ملاحظات','notes','s']], paste: 'B7', last: 'H', extra: 'U7' },
    'المستخلصات': { arr: 'bills', cap: 300, cols: [['B','المشروع','project','s'],['C','البند','item','s'],['D','رقم المستخلص','no','s'],['E','تاريخ الإصدار','date','d'],['F','مدة السداد بالأيام','days','n'],['G','قيمة المستخلص','value','n'],['H','ملاحظات','notes','s']], paste: 'B7', last: 'H' },
    'الحركات': { arr: 'journal', cap: 1000, cols: [['B','التاريخ','date','d'],['C','نوع الحركة','type','s'],['D','المشروع','project','s'],['E','البند','item','s'],['F','المبلغ','amount','n'],['G','طريقة الدفع','method','s'],['H','رقم المستند','docNo','s'],['I','ملاحظات','notes','s']], paste: 'B7', last: 'I' }
  };
  // cellXfs: 0 default,1 header,2 text,3 date,4 money,5 pct,6 note,7 warn, 8..11 new-row text/date/money/pct, 12 title
  const XF = { s: 2, d: 3, n: 4, p: 5 };
  const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="#,##0.00"/></numFmts>' +
    '<fonts count="4"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FF1F2A44"/><name val="Arial"/></font><font><b/><sz val="14"/><color rgb="FF1F2A44"/><name val="Arial"/></font></fonts>' +
    '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F2A44"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="13">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="9" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="2" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="9" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '</cellXfs></styleSheet>';
  const NEW_XF = { s: 8, d: 9, n: 10, p: 11 };
  const colNum = c => c.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const tc = (ref, st, text) => `<c r="${ref}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${esc(text)}</t></is></c>`;
  function pasteSheetXml(title, spec, rows, mark, extraNotes) {
    const rowsXml = [];
    const n = rows.length, nNew = Math.max(0, n - Math.min(mark || 0, n));
    const notes = [
      [1, 12, title],
      [2, 6, `الصق هنا: تبويب «${title}» في ملفك، من الخلية ${spec.paste} ${spec.extra ? '(الأعمدة B إلى H)' : ''}. حدد الخلايا من B7 إلى ${spec.last}${6 + Math.max(n, 1)} ثم انسخ، وفي ملفك اضغط Ctrl+Shift+V (لصق القيم فقط) حتى تبقى معادلاتك وقوائمك كما هي.`],
      [3, 6, `عدد السجلات في هذا الملف: ${n}. الجديد منذ آخر تصدير: ${nNew} (مظلل بالأخضر). هذا الملف يحتوي كل السجلات من الأول، فالصقه دائماً من الصف 7 وسيتجدد الموجود ويضاف الجديد في الصفوف التالية بدون حذف شيء.`]
    ];
    if (spec.extra) notes.push([4, 6, `ملاحظات العقود في العمود U: حدد U7 إلى U${6 + Math.max(n, 1)} وانسخ والصق في الخلية ${spec.extra} من تبويب «العقود» (قيم فقط).`]);
    if (extraNotes) notes.push([5, 7, extraNotes]);
    const merges = [];
    notes.forEach(([r, st, text]) => { const ht = st === 12 ? 24 : 34; merges.push(`<mergeCell ref="B${r}:${spec.last}${r}"/>`); rowsXml.push(`<row r="${r}" ht="${ht}" customHeight="1">${tc('B' + r, st, text)}</row>`); });
    rowsXml.push(`<row r="6" ht="32" customHeight="1">${spec.cols.map(([c, h]) => tc(c + 6, 1, h)).join('')}</row>`);
    rows.forEach((rec, i) => {
      const r = 7 + i, isNew = i >= n - nNew; const cells = [];
      spec.cols.forEach(([c, , key, kind]) => {
        const v = rec[key]; const st = (isNew ? NEW_XF : XF)[kind];
        if (v === undefined || v === null || v === '') { cells.push(`<c r="${c}${r}" s="${st}"/>`); return; }
        if (kind === 'n') cells.push(`<c r="${c}${r}" s="${st}"><v>${Number(v)}</v></c>`);
        else if (kind === 'p') cells.push(`<c r="${c}${r}" s="${st}"><v>${Number(v) / 100}</v></c>`);
        else if (kind === 'd' && /^\d{4}-\d{2}-\d{2}$/.test(v)) cells.push(`<c r="${c}${r}" s="${st}"><v>${serial(v)}</v></c>`);
        else cells.push(tc(c + r, st, v));
      });
      rowsXml.push(`<row r="${r}">${cells.join('')}</row>`);
    });
    const maxCol = Math.max(...spec.cols.map(c => colNum(c[0])));
    const widths = spec.cols.map(([c]) => `<col min="${colNum(c)}" max="${colNum(c)}" width="${c === 'B' ? 34 : 20}" customWidth="1"/>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView rightToLeft="1" workbookViewId="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="6" customWidth="1"/>${widths}</cols><sheetData>${rowsXml.join('')}</sheetData><mergeCells count="${merges.length}">${merges.join('')}</mergeCells><pageMargins left="0.3" right="0.3" top="0.3" bottom="0.3" header="0.2" footer="0.2"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`;
  }
  function servicesSheetXml(services) {
    const rowsXml = [`<row r="1">${tc('B1', 12, 'القوائم (اختياري)')}</row>`,
      `<row r="2">${tc('B2', 6, 'الصق في تبويب «القوائم» من الخلية B5 (قيم فقط) فقط إذا أضفت خدمة جديدة في التطبيق. هذه قائمة الخدمات التي تظهر في قائمة العقود. لا تلصقها إن كانت قائمتك فيها خدمات أخرى لا تريد استبدالها.')}</row>`,
      `<row r="4">${tc('B4', 1, 'الخدمات')}</row>`];
    (services || []).slice(0, 12).forEach((s, i) => rowsXml.push(`<row r="${5 + i}">${tc('B' + (5 + i), 2, s)}</row>`));
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews><cols><col min="1" max="1" width="6" customWidth="1"/><col min="2" max="2" width="34" customWidth="1"/></cols><sheetData>${rowsXml.join('')}</sheetData></worksheet>`;
  }
  function trashSheetXml(trash) {
    const lab = { clients: 'عميل', contracts: 'عقد', bills: 'مستخلص', journal: 'حركة' };
    const rowsXml = [`<row r="1">${tc('B1', 12, 'سجلات محذوفة من التطبيق (للمراجعة فقط، لا تُلصق)')}</row>`, `<row r="2">${tc('B2', 6, 'لا يُحذف أي سجل نهائياً. أي سجل يُحذف من التطبيق يبقى هنا مع تاريخ الحذف.')}</row>`,
      `<row r="4">${tc('B4', 1, 'النوع')}${tc('C4', 1, 'تاريخ الحذف')}${tc('D4', 1, 'البيانات')}</row>`];
    (trash || []).forEach((x, i) => rowsXml.push(`<row r="${5 + i}">${tc('B' + (5 + i), 2, lab[x.kind] || x.kind)}${tc('C' + (5 + i), 2, String(x.t || '').slice(0, 10))}${tc('D' + (5 + i), 2, JSON.stringify(x.rec))}</row>`));
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews><cols><col min="2" max="3" width="16" customWidth="1"/><col min="4" max="4" width="90" customWidth="1"/></cols><sheetData>${rowsXml.join('')}</sheetData></worksheet>`;
  }
  async function exportPaste(JSZip, data, marks) {
    const names = Object.keys(PASTE), all = names.concat(['القوائم', 'المحذوفات']);
    const zip = new JSZip(); const report = {};
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + all.map((n, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '</Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + all.map((n, i) => `<sheet name="${n}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + all.map((n, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') + `<Relationship Id="rId${all.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    zip.file('xl/styles.xml', STYLES);
    names.forEach((n, i) => {
      const spec = PASTE[n], rows = data[spec.arr] || [];
      if (rows.length > spec.cap) throw new Error(`${n}: ${rows.length} سجل أكثر من سعة الجدول ${spec.cap}`);
      zip.file(`xl/worksheets/sheet${i + 1}.xml`, pasteSheetXml(n, spec, rows, (marks || {})[spec.arr] || 0,
        n === 'المستخلصات' || n === 'الحركات' || n === 'العقود' ? 'أسماء العميل والمشروع والخدمة والبند ونوع الحركة يجب أن تطابق قوائم ملفك حرفياً حتى تعمل القوائم المنسدلة والمعادلات.' : ''));
      report[n] = rows.length;
    });
    zip.file(`xl/worksheets/sheet${names.length + 1}.xml`, servicesSheetXml(data.services));
    zip.file(`xl/worksheets/sheet${names.length + 2}.xml`, trashSheetXml(data.trash));
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { bytes, report };
  }
  const api = { exportWorkbook, exportPaste, PASTE, MAP, serial };
  if (typeof module !== 'undefined') module.exports = api; else root.WaqidExport = api;
})(typeof self !== 'undefined' ? self : this);
