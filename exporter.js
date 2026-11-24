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
  const api = { exportWorkbook, MAP, serial };
  if (typeof module !== 'undefined') module.exports = api; else root.WaqidExport = api;
})(typeof self !== 'undefined' ? self : this);
