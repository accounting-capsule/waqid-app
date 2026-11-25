(function () {
'use strict';
/* ---------- lists (same as the القوائم sheet of the workbook) ---------- */
const LISTS = {
  types: ['إيراد', 'مصروف مشروع', 'مصروف عمومي'],
  services: ['ترخيص', 'إنشاءات', 'تشطيبات', 'تصالح', 'عقود أراضي أملاك دولة'],
  revItems: ['دفعة مقدمة', 'مستخلص محصل', 'إفراج عن محتجز'],
  projItems: ['استشاري فرعي', 'رخص وجهات حكومية', 'رفع مساحي', 'رسوم نقابة', 'طباعة ومستندات', 'انتقالات وسفر', 'مقاول باطن', 'مصروف مشروع آخر'],
  genItems: ['إيجار المكتب', 'كهرباء ومياه وإنترنت', 'مرتبات وأجور', 'أدوات مكتبية', 'تسويق ودعاية', 'ضرائب وتأمينات', 'صيانة', 'مصروف عمومي آخر'],
  pay: ['نقدي', 'تحويل بنكي', 'شيك', 'إنستاباي / محفظة'],
  status: ['قيد المعالجة', 'مكتمل', 'مرفوض', 'متوقف'],
  billItems: ['دفعة مقدمة', 'مستخلص']
};
const DEF_DAYS = 30;
const KEY = 'waqid_proto_v1';
const $ = s => document.querySelector(s);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const AR = '٠١٢٣٤٥٦٧٨٩';
const norm = s => String(s == null ? '' : s).replace(/[٠-٩]/g, d => AR.indexOf(d)).replace(/[٬,\s]/g, '').replace('٫', '.');
const money = n => (Math.round((+n || 0) * 100) / 100).toLocaleString('en-US', { maximumFractionDigits: 2 });
const today = () => new Date().toISOString().slice(0, 10);
const fmtDate = iso => { if (!iso) return '—'; const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso); return m ? `${m[3]}/${m[2]}/${m[1]}` : iso; };
const dayNum = iso => Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / 86400000;
const addDays = (iso, n) => new Date((dayNum(iso) + n) * 86400000).toISOString().slice(0, 10);

/* ---------- storage ---------- */
let db;
function load() {
  try { db = JSON.parse(localStorage.getItem(KEY)); } catch (e) { db = null; }
  if (!db) { db = Object.assign(WaqidDemo(), { log: [], demo: true }); save(); }
  if (!Array.isArray(db.services)) { db.services = LISTS.services.slice(); save(); }
}
function save() { localStorage.setItem(KEY, JSON.stringify(db)); if (typeof SYNC !== 'undefined') SYNC.schedule(); }
const SYNC_TXT = { ok: 'تمت المزامنة ✓', pending: 'بانتظار المزامنة…', offline: 'لا يوجد اتصال، سيُرفع لاحقاً', conflict: 'توجد نسخة أحدث على الشيت', err: 'تعذرت المزامنة', warn: 'لم يتم الرفع' };
let syncState = '';
function syncBanner(s) {
  let b = document.getElementById('syncbanner');
  const bad = { conflict: 'تنبيه: تعديلاتك على هذا الجهاز لم تُرفع. يوجد تعديل أحدث من جهاز آخر. افتح «التصدير» واضغط «استرجاع من الشيت» (سيستبدل تعديلات هذا الجهاز) أو سجّل التعديلات من جديد.', err: 'تنبيه: تعذرت المزامنة، التعديلات الأخيرة غير محفوظة في الشيت.', warn: 'تنبيه: لم يتم الرفع.' };
  if (!bad[s]) { if (b) b.remove(); return; }
  if (!b) { b = document.createElement('div'); b.id = 'syncbanner'; b.style.cssText = 'position:sticky;top:0;z-index:50;background:#b3261e;color:#fff;padding:10px 14px;font-size:14px;line-height:1.5'; document.body.prepend(b); }
  b.textContent = bad[s];
}
const SYNC = WaqidSync({ storage: localStorage, fetch: (u, o) => fetch(u, o), getDb: () => db, setDb: d => { db = d; localStorage.setItem(KEY, JSON.stringify(db)); }, onStatus: (s, m) => { syncState = s; syncBanner(s); const e = document.getElementById('syncst'); if (e) e.textContent = SYNC_TXT[s] || s; } });
window.addEventListener('online', () => SYNC.flush());
function logChange(what, kind, name) { db.log.unshift({ t: new Date().toISOString(), what, kind, name }); db.log = db.log.slice(0, 300); }

/* ---------- calculations (mirror the workbook formulas) ---------- */
function calc() {
  const asOf = today();
  const rev = db.journal.filter(j => j.type === 'إيراد'), pe = db.journal.filter(j => j.type === 'مصروف مشروع'), ge = db.journal.filter(j => j.type === 'مصروف عمومي');
  const sum = a => a.reduce((s, x) => s + (+x.amount || 0), 0);
  const t = { rev: sum(rev), direct: sum(pe), general: sum(ge) };
  t.exp = t.direct + t.general; t.net = t.rev - t.exp;
  const contracts = {};
  db.contracts.forEach(c => {
    const mine = db.journal.filter(j => j.project === c.name);
    const r = mine.filter(j => j.type === 'إيراد');
    const x = { c, adv: sum(r.filter(j => j.item === 'دفعة مقدمة')), collected: sum(r), released: sum(r.filter(j => j.item === 'إفراج عن محتجز')), cost: sum(mine.filter(j => j.type === 'مصروف مشروع')), retHeld: 0, unpaid: 0 };
    x.remain = (+c.value || 0) - x.collected; x.pct = c.value ? x.collected / c.value : 0; x.net = x.collected - x.cost; x.exp = (+c.value || 0) - x.cost;
    contracts[c.name] = x;
  });
  // bills: retention, net, sequential allocation of collected money per project
  const alloc = {}; const bills = db.bills.map((b, i) => {
    const ct = db.contracts.find(c => c.name === b.project);
    const ret = b.item === 'دفعة مقدمة' ? 0 : (+b.value || 0) * ((ct && +ct.retention) || 0) / 100;
    const net = (+b.value || 0) - ret;
    const paidPool = db.journal.filter(j => j.project === b.project && j.type === 'إيراد' && j.item !== 'إفراج عن محتجز').reduce((s, j) => s + (+j.amount || 0), 0);
    const used = alloc[b.project] || 0;
    const collected = Math.min(net, Math.max(0, paidPool - used)); alloc[b.project] = used + net;
    const remaining = net - collected;
    const due = b.date ? addDays(b.date, b.days === '' || b.days == null ? DEF_DAYS : +b.days) : '';
    const late = (remaining <= 0.005 || !due) ? 0 : Math.max(0, dayNum(asOf) - dayNum(due));
    const cls = remaining <= 0.005 ? 'مسدد' : late === 0 ? 'لم يستحق بعد' : late <= 30 ? 'متأخر 1-30' : late <= 60 ? 'متأخر 31-60' : late <= 90 ? 'متأخر 61-90' : 'متأخر +90';
    return { b, i, ret, net, collected, remaining, due, late, cls, client: ct ? ct.client : '' };
  });
  bills.forEach(x => { const c = contracts[x.b.project]; if (c) { c.retHeld += x.ret; c.unpaid += x.remaining; } });
  Object.values(contracts).forEach(c => { c.retLeft = c.retHeld - c.released; });
  const clients = {};
  db.clients.forEach(cl => {
    const cs = db.contracts.filter(c => c.client === cl.name).map(c => contracts[c.name]);
    const bs = bills.filter(x => x.client === cl.name);
    clients[cl.name] = { cl, n: cs.length, total: cs.reduce((s, c) => s + (+c.c.value || 0), 0), collected: cs.reduce((s, c) => s + c.collected, 0), due: bs.reduce((s, x) => s + x.remaining, 0) };
    clients[cl.name].remain = clients[cl.name].total - clients[cl.name].collected;
  });
  const months = {};
  db.journal.forEach(j => { if (!j.date) return; const k = j.date.slice(0, 7); months[k] = months[k] || { r: 0, e: 0 }; if (j.type === 'إيراد') months[k].r += +j.amount || 0; else months[k].e += +j.amount || 0; });
  const overdue = bills.filter(x => x.late > 0);
  return { asOf, t, contracts, bills, clients, months, overdue, cash: t.rev - t.exp };
}

/* ---------- schemas ---------- */
const KINDS = {
  clients: { label: 'عميل', plural: 'العملاء', arr: 'clients', icon: '☺', title: x => x.name, sub: x => [x.entity, x.phone].filter(Boolean).join(' · '),
    fields: () => [
      { k: 'name', l: 'اسم العميل', t: 'text', req: 1, hint: 'يُكتب مرة واحدة وبنفس الشكل دائماً' },
      { k: 'phone', l: 'الهاتف', t: 'tel', hint: 'أرقام فقط (7 إلى 15 رقماً)' },
      { k: 'entity', l: 'الجهة أو المهنة', t: 'text' }, { k: 'address', l: 'العنوان', t: 'text' },
      { k: 'agreeDate', l: 'تاريخ الاتفاق', t: 'date', def: today() }, { k: 'endDate', l: 'تاريخ النهاية', t: 'date' }, { k: 'notes', l: 'ملاحظات', t: 'area' }] },
  contracts: { label: 'عقد / مشروع', plural: 'العقود', arr: 'contracts', icon: '▤', title: x => x.name, sub: x => `${x.client} · ${x.service}`,
    fields: () => [
      { k: 'client', l: 'العميل', t: 'select', req: 1, opts: db.clients.map(c => c.name), hint: db.clients.length ? '' : 'أضف عميلاً أولاً' },
      { k: 'service', l: 'نوع الخدمة', t: 'select', req: 1, opts: db.services },
      { k: 'name', l: 'اسم المشروع', t: 'text', req: 1, hint: 'مثال: ترخيص - اسم العميل. لا يتكرر.', auto: f => f.service && f.client ? `${f.service} - ${f.client}` : '' },
      { k: 'date', l: 'تاريخ العقد', t: 'date', def: today() },
      { k: 'status', l: 'حالة العقد', t: 'select', req: 1, opts: LISTS.status, def: 'قيد المعالجة' },
      { k: 'value', l: 'قيمة العقد (إجمالي الأتعاب بالجنيه)', t: 'num', req: 1 },
      { k: 'retention', l: 'نسبة المحتجز %', t: 'num', hint: 'إن وجدت. اتركها 0 إن لم يوجد محتجز', def: '0' },
      { k: 'notes', l: 'وصف المشروع / ملاحظات', t: 'area' }] },
  bills: { label: 'مستخلص / دفعة', plural: 'المستخلصات', arr: 'bills', icon: '✎', title: x => `${x.item} ${x.no ? '#' + x.no : ''} · ${money(x.value)}`, sub: x => x.project,
    fields: () => [
      { k: 'project', l: 'المشروع', t: 'select', req: 1, opts: db.contracts.map(c => c.name) },
      { k: 'item', l: 'البند', t: 'select', req: 1, opts: LISTS.billItems, hint: 'الدفعة الأولى عند التعاقد = دفعة مقدمة' },
      { k: 'no', l: 'رقم المستخلص', t: 'text' },
      { k: 'date', l: 'تاريخ الإصدار', t: 'date', def: today(), req: 1 },
      { k: 'days', l: 'مدة السداد (يوم)', t: 'num', def: String(DEF_DAYS) },
      { k: 'value', l: 'القيمة', t: 'num', req: 1 },
      { k: 'notes', l: 'ملاحظات', t: 'area' }] },
  journal: { label: 'حركة (إيراد / مصروف)', plural: 'الحركات', arr: 'journal', icon: '⇄', title: x => `${x.item}`, sub: x => `${x.type}${x.project ? ' · ' + x.project : ''}`,
    fields: f => { const t = f && f.type;
      return [
      { k: 'date', l: 'التاريخ', t: 'date', def: today(), req: 1 },
      { k: 'type', l: 'نوع الحركة', t: 'seg', req: 1, opts: LISTS.types, def: 'إيراد' },
      t !== 'مصروف عمومي' ? { k: 'project', l: 'المشروع', t: 'select', req: 1, opts: db.contracts.map(c => c.name), hint: 'كما في العقود' } : null,
      { k: 'item', l: 'البند', t: 'select', req: 1, opts: t === 'مصروف عمومي' ? LISTS.genItems : t === 'مصروف مشروع' ? LISTS.projItems : LISTS.revItems },
      { k: 'amount', l: 'المبلغ بالجنيه', t: 'num', req: 1, hint: 'أرقام فقط' },
      { k: 'method', l: 'طريقة الدفع', t: 'select', opts: LISTS.pay },
      { k: 'docNo', l: 'رقم المستند', t: 'text' }, { k: 'notes', l: 'ملاحظات', t: 'area' }].filter(Boolean); } }
};

/* ---------- ui helpers ---------- */
function toast(m) { const t = $('#toast'); t.textContent = m; t.classList.add('show'); clearTimeout(toast.h); toast.h = setTimeout(() => t.classList.remove('show'), 2400); }
function setTitle(t, back) { $('#title').textContent = t; const b = $('#backBtn'); b.hidden = !back; b.onclick = () => history.back(); }
const kpi = (l, v, cls = '') => `<div class="kpi ${cls}"><div class="l">${l}</div><div class="v">${v}</div></div>`;
function pill(cls) { const m = { 'مسدد': 'ok', 'لم يستحق بعد': '', 'متأخر 1-30': 'warn', 'متأخر 31-60': 'warn', 'متأخر 61-90': 'bad', 'متأخر +90': 'bad', 'مكتمل': 'ok', 'قيد المعالجة': 'warn', 'مرفوض': 'bad', 'متوقف': 'bad' }; return `<span class="pill ${m[cls] || ''}">${esc(cls)}</span>`; }

/* ---------- views ---------- */
const V = {};
V.home = () => {
  setTitle('الرئيسية');
  const c = calc();
  const mk = Object.keys(c.months).sort().slice(-6);
  const mx = Math.max(1, ...mk.map(k => Math.max(c.months[k].r, c.months[k].e)));
  const chart = mk.length ? `<div class="card"><div class="chart">${mk.map(k => `<div class="col"><div class="bars"><div class="b r" style="height:${c.months[k].r / mx * 100}%"></div><div class="b e" style="height:${c.months[k].e / mx * 100}%"></div></div>${k.slice(5)}/${k.slice(2, 4)}</div>`).join('')}</div><div class="legend"><span><i style="background:#C9A84C"></i>إيرادات</span><span><i style="background:#0D1B2A"></i>مصروفات</span></div></div>` : '';
  const due = c.bills.reduce((s, x) => s + x.remaining, 0), over = c.overdue.reduce((s, x) => s + x.remaining, 0);
  const projs = db.contracts.filter(x => x.status === 'قيد المعالجة').map(x => { const k = c.contracts[x.name]; return `<a class="card item" href="#/view/contracts/${db.contracts.indexOf(x)}"><div class="row"><div><div class="t">${esc(x.name)}</div><div class="s">${esc(x.client)}</div></div><div class="amt">${money(k.collected)} / ${money(x.value)}</div></div><div class="bar"><i style="width:${Math.min(100, k.pct * 100)}%"></i></div></a>`; }).join('');
  $('#app').innerHTML = `
  <div class="grid">
    <div class="kpi big"><div class="l">صافي ربح المكتب</div><div class="v">${money(c.t.net)} ج</div></div>
    ${kpi('الإيرادات المحصلة', money(c.t.rev), 'ok')}${kpi('إجمالي المصروفات', money(c.t.exp))}
    ${kpi('تكاليف المشاريع', money(c.t.direct))}${kpi('مصروفات عمومية', money(c.t.general))}
    ${kpi('مستحق للتحصيل', money(due))}${kpi('متأخر السداد', money(over), over > 0 ? 'bad' : '')}
    ${kpi('رصيد الخزنة', money(c.cash))}${kpi('مشاريع جارية', db.contracts.filter(x => x.status === 'قيد المعالجة').length)}
  </div>
  <h2>الأداء الشهري</h2>${chart || '<div class="card empty">لا توجد حركات بعد</div>'}
  ${c.overdue.length ? `<h2>مستخلصات متأخرة</h2>${c.overdue.sort((a, b) => b.late - a.late).slice(0, 4).map(x => `<a class="card item" href="#/view/bills/${x.i}"><div class="row"><div><div class="t">${esc(x.b.project)}</div><div class="s">${esc(x.client)} · استحقاق ${fmtDate(x.due)}</div></div><div style="text-align:left"><div class="amt">${money(x.remaining)}</div>${pill(x.cls)}</div></div></a>`).join('')}` : ''}
  <h2>المشاريع الجارية</h2>${projs || '<div class="card empty">لا توجد مشاريع جارية</div>'}`;
};

V.add = () => {
  setTitle('إضافة');
  $('#app').innerHTML = `<div class="menu">
    <a href="#/new/journal"><span class="big">⇄</span><div><b>حركة جديدة</b><small>إيراد أو مصروف مشروع أو مصروف عمومي</small></div></a>
    <a href="#/new/bills"><span class="big">✎</span><div><b>مستخلص / دفعة</b><small>مطلوبة من عميل</small></div></a>
    <a href="#/new/contracts"><span class="big">▤</span><div><b>عقد / مشروع</b><small>بطاقة المشروع وقيمة العقد</small></div></a>
    <a href="#/new/clients"><span class="big">☺</span><div><b>عميل جديد</b><small>الاسم والهاتف والجهة</small></div></a></div>`;
};

function listRows(kind, q) {
  const K = KINDS[kind]; let rows = db[K.arr].map((x, i) => ({ x, i }));
  if (kind === 'journal' || kind === 'bills') rows.sort((a, b) => (b.x.date || '').localeCompare(a.x.date || ''));
  if (q) rows = rows.filter(r => JSON.stringify(r.x).toLowerCase().includes(q.toLowerCase()));
  return rows;
}
V.records = (kind) => {
  kind = kind || 'journal'; setTitle('السجلات');
  const seg = Object.keys(KINDS).map(k => `<button class="${k === kind ? 'on' : ''}" data-k="${k}">${KINDS[k].plural} (${db[KINDS[k].arr].length})</button>`).join('');
  $('#app').innerHTML = `<div class="seg">${seg}</div><input class="search" id="q" placeholder="بحث..."><div id="list"></div>`;
  const draw = () => {
    const rows = listRows(kind, $('#q').value.trim()); const c = calc();
    $('#list').innerHTML = rows.length ? rows.map(({ x, i }) => {
      const K = KINDS[kind]; let right = '';
      if (kind === 'journal') right = `<div class="amt ${x.type === 'إيراد' ? 'in' : 'out'}">${x.type === 'إيراد' ? '+' : '−'}${money(x.amount)}</div><div class="s">${fmtDate(x.date)}</div>`;
      if (kind === 'bills') { const b = c.bills[i]; right = `<div class="amt">${money(x.value)}</div>${pill(b.cls)}`; }
      if (kind === 'contracts') right = `<div class="amt">${money(x.value)}</div>${pill(x.status)}`;
      if (kind === 'clients') { const k = c.clients[x.name]; right = k ? `<div class="s">مستحق</div><div class="amt">${money(k.due)}</div>` : ''; }
      return `<a class="card item" href="#/view/${kind}/${i}"><div class="row"><div><div class="t">${esc(K.title(x))}</div><div class="s">${esc(K.sub(x))}</div></div><div style="text-align:left">${right}</div></div></a>`;
    }).join('') : '<div class="empty">لا توجد سجلات</div>';
  };
  $('#q').oninput = draw; draw();
  document.querySelectorAll('.seg button').forEach(b => b.onclick = () => { location.hash = '#/records/' + b.dataset.k; });
  const on = document.querySelector('.seg button.on'); if (on && on.scrollIntoView) on.scrollIntoView({ inline: 'center', block: 'nearest' });
};

function formView(kind, idx) {
  const K = KINDS[kind]; const editing = idx != null && idx !== '' ; const rec = editing ? db[K.arr][+idx] : null;
  setTitle((editing ? 'تعديل ' : 'إضافة ') + K.label, true);
  let state = rec ? Object.assign({}, rec, kind === 'contracts' ? { retention: String(rec.retention) } : {}) : {};
  const initState = () => { K.fields(state).forEach(f => { if (state[f.k] == null) state[f.k] = f.def != null ? f.def : ''; }); };
  initState();
  const draw = () => {
    const fs = K.fields(state);
    $('#app').innerHTML = `<form id="f" novalidate>${fs.map(f => {
      const v = state[f.k] == null ? '' : state[f.k]; const lab = `<label for="f_${f.k}">${f.l}${f.req ? ' <span class="req">*</span>' : ''}</label>`;
      let inp;
      if (f.t === 'select') inp = `<select id="f_${f.k}" data-k="${f.k}"><option value="">اختر...</option>${f.opts.map(o => `<option ${o === v ? 'selected' : ''}>${esc(o)}</option>`).join('')}${v && !f.opts.includes(v) ? `<option selected>${esc(v)}</option>` : ''}</select>`;
      else if (f.t === 'seg') inp = `<div class="seg" id="f_${f.k}">${f.opts.map(o => `<button type="button" data-seg="${f.k}" data-v="${esc(o)}" class="${o === v ? 'on' : ''}">${esc(o)}</button>`).join('')}</div>`;
      else if (f.t === 'date') inp = `<input type="date" id="f_${f.k}" data-k="${f.k}" value="${esc(v)}">`;
      else if (f.t === 'area') inp = `<textarea id="f_${f.k}" data-k="${f.k}">${esc(v)}</textarea>`;
      else inp = `<input type="text" id="f_${f.k}" data-k="${f.k}" value="${esc(v)}" ${f.t === 'num' ? 'inputmode="decimal"' : f.t === 'tel' ? 'inputmode="tel"' : ''} autocomplete="off">`;
      return `<div>${lab}${inp}${f.hint ? `<div class="hint">${f.hint}</div>` : ''}<div class="err" id="e_${f.k}"></div></div>`;
    }).join('')}<div id="warns"></div>
    <button class="btn gold" type="submit">${editing ? 'حفظ التعديل' : 'حفظ'}</button>
    ${editing ? '<button class="btn danger" type="button" id="del">حذف</button>' : ''}</form>`;
    document.querySelectorAll('[data-k]').forEach(el => {
      const evt = el.tagName === 'SELECT' || el.type === 'date' ? 'change' : 'input';
      el.addEventListener(evt, () => {
        state[el.dataset.k] = el.value;
        if (kind === 'contracts' && (el.dataset.k === 'service' || el.dataset.k === 'client') && !editing) {
          const f = K.fields(state).find(x => x.k === 'name'); if (!state._nameTouched) { state.name = f.auto(state); const n = $('#f_name'); if (n) n.value = state.name; }
        }
        if (el.dataset.k === 'name') state._nameTouched = true;
        if (kind === 'bills' && el.dataset.k === 'item' && !editing && el.value === 'دفعة مقدمة') { state.days = '7'; const d = $('#f_days'); if (d) d.value = '7'; }
        softWarn();
      });
    });
    document.querySelectorAll('[data-seg]').forEach(b => b.onclick = () => {
      state[b.dataset.seg] = b.dataset.v; if (kind === 'journal') { state.item = ''; if (b.dataset.v === 'مصروف عمومي') state.project = ''; } draw(); });
    $('#f').onsubmit = e => { e.preventDefault(); submit(); };
    if ($('#del')) $('#del').onclick = del;
    softWarn();
  };
  const softWarn = () => {
    const w = []; const s = state; const c = calc();
    if (kind === 'journal' && s.project && s.type === 'إيراد') {
      const ct = db.contracts.find(x => x.name === s.project);
      if (s.item === 'دفعة مقدمة' && db.journal.some((j, i) => j.project === s.project && j.item === 'دفعة مقدمة' && String(i) !== String(idx))) w.push('هذا المشروع له دفعة مقدمة مسجلة بالفعل. الدفعات بعد الأولى تُسجل «مستخلص محصل».');
      if (ct && c.contracts[s.project] && !editing && +norm(s.amount) > c.contracts[s.project].remain + 0.01) w.push('المبلغ أكبر من المتبقي على العقد.');
    }
    if (kind === 'bills' && s.project && !editing) {
      const ct = db.contracts.find(x => x.name === s.project); const have = db.bills.filter(b => b.project === s.project).reduce((a, b) => a + (+b.value || 0), 0);
      if (ct && have + +norm(s.value) > +ct.value + 0.01) w.push('مجموع المستخلصات يتجاوز قيمة العقد.');
    }
    const el = $('#warns'); if (el) el.innerHTML = w.map(x => `<div class="warnbox">${esc(x)}</div>`).join('');
  };
  const validate = () => {
    const errs = {}; const fs = K.fields(state); const out = {};
    fs.forEach(f => {
      let v = state[f.k]; v = v == null ? '' : String(v).trim();
      if (f.t === 'num') v = norm(v);
      if (f.t === 'tel') v = norm(v);
      if (f.req && !v) { errs[f.k] = 'مطلوب'; return; }
      if (f.t === 'num' && v) { if (!/^\d+(\.\d+)?$/.test(v)) errs[f.k] = 'أرقام فقط'; else if (['value', 'amount'].includes(f.k) && +v <= 0) errs[f.k] = 'يجب أن يكون أكبر من صفر'; else if (f.k === 'retention' && +v > 100) errs[f.k] = 'من 0 إلى 100'; else v = +v; }
      if (f.t === 'tel' && v && !/^\d{7,15}$/.test(v)) errs[f.k] = 'من 7 إلى 15 رقماً';
      out[f.k] = v;
    });
    if (kind === 'clients' && out.name && db.clients.some((c, i) => c.name === out.name && String(i) !== String(idx))) errs.name = 'العميل مسجل بالفعل';
    if (kind === 'contracts' && out.name && db.contracts.some((c, i) => c.name === out.name && String(i) !== String(idx))) errs.name = 'اسم المشروع مكرر';
    if (kind === 'contracts' && out.retention === '') out.retention = 0;
    if (kind === 'bills' && out.days === '') out.days = DEF_DAYS;
    if (kind === 'clients' && out.endDate && out.agreeDate && out.endDate < out.agreeDate) errs.endDate = 'قبل تاريخ الاتفاق';
    return { errs, out };
  };
  const submit = () => {
    const { errs, out } = validate();
    document.querySelectorAll('.err').forEach(e => e.textContent = '');
    const keys = Object.keys(errs); if (keys.length) { keys.forEach(k => { const e = $('#e_' + k); if (e) e.textContent = errs[k]; }); const first = $('#e_' + keys[0]); first && first.scrollIntoView({ block: 'center' }); return; }
    if (editing) {
      const old = db[K.arr][+idx];
      if (kind === 'clients' && old.name !== out.name) db.contracts.forEach(c => { if (c.client === old.name) c.client = out.name; });
      if (kind === 'contracts' && old.name !== out.name) { db.bills.forEach(b => { if (b.project === old.name) b.project = out.name; }); db.journal.forEach(j => { if (j.project === old.name) j.project = out.name; }); }
      db[K.arr][+idx] = out; logChange('تعديل', K.label, K.title(out));
    } else { db[K.arr].push(out); logChange('إضافة', K.label, K.title(out)); }
    save(); toast('تم الحفظ ✓'); location.hash = '#/records/' + kind;
  };
  const del = () => {
    const x = db[K.arr][+idx];
    if (kind === 'clients' && db.contracts.some(c => c.client === x.name)) return toast('لا يمكن حذف عميل له عقود');
    if (kind === 'contracts' && (db.bills.some(b => b.project === x.name) || db.journal.some(j => j.project === x.name))) return toast('لا يمكن حذف مشروع له مستخلصات أو حركات');
    if (!confirm('حذف هذا السجل؟ سيُسجل الحذف في سجل التعديلات.')) return;
    db[K.arr].splice(+idx, 1); logChange('حذف', K.label, K.title(x)); save(); toast('تم الحذف'); location.hash = '#/records/' + kind;
  };
  draw();
}
V.new = (kind) => {
  if (kind === 'contracts' && !db.clients.length) {
    setTitle('إضافة عقد / مشروع', true);
    $('#app').innerHTML = `<div class="card empty"><div class="t">لا يوجد عملاء بعد</div><p class="s">العقد يتبع عميلاً مسجلاً. أضف أول عميل (الاسم والهاتف والجهة) ثم ارجع لتسجيل العقد.</p></div><a class="btn gold" href="#/new/clients" style="text-align:center;text-decoration:none;display:block">إضافة أول عميل</a>`;
    return;
  }
  if (kind === 'bills' && !db.contracts.length) {
    setTitle('إضافة مستخلص / دفعة', true);
    $('#app').innerHTML = `<div class="card empty"><div class="t">لا توجد عقود بعد</div><p class="s">المستخلص يتبع عقداً مسجلاً. سجّل العقد أولاً ثم أضف مستخلصاته.</p></div><a class="btn gold" href="#/new/contracts" style="text-align:center;text-decoration:none;display:block">إضافة أول عقد</a>`;
    return;
  }
  formView(kind, null);
};
V.view = (kind, idx) => {
  // detail card for contracts/clients first, then edit button
  const K = KINDS[kind]; const x = db[K.arr][+idx]; if (!x) { location.hash = '#/records'; return; }
  if (kind === 'journal' || kind === 'bills') return formView(kind, idx);
  setTitle(K.label, true); const c = calc();
  let body = '';
  if (kind === 'contracts') {
    const k = c.contracts[x.name]; body = `<div class="card"><div class="t">${esc(x.name)}</div><div class="s">${esc(x.client)} · ${esc(x.service)} · ${fmtDate(x.date)} ${pill(x.status)}</div><div class="bar"><i style="width:${Math.min(100, k.pct * 100)}%"></i></div></div>
    <div class="grid">${kpi('قيمة العقد', money(x.value))}${kpi('إجمالي المحصل', money(k.collected), 'ok')}${kpi('المتبقي لدى العميل', money(k.remain))}${kpi('نسبة التحصيل', Math.round(k.pct * 100) + '%')}${kpi('التكاليف المباشرة', money(k.cost))}${kpi('صافي الربح الحالي', money(k.net), k.net < 0 ? 'bad' : 'ok')}${kpi('المحتجز المتبقي', money(k.retLeft))}${kpi('مستخلصات غير محصلة', money(k.unpaid))}</div>`;
  } else {
    const k = c.clients[x.name] || { n: 0, total: 0, collected: 0, due: 0, remain: 0 };
    body = `<div class="card"><div class="t">${esc(x.name)}</div><div class="s">${esc([x.entity, x.phone, x.address].filter(Boolean).join(' · '))}</div></div><div class="grid">${kpi('عدد المشاريع', k.n)}${kpi('إجمالي العقود', money(k.total))}${kpi('إجمالي المحصل', money(k.collected), 'ok')}${kpi('مستحق للتحصيل', money(k.due))}</div>`;
  }
  $('#app').innerHTML = body + `<div class="btns"><button class="btn ghost" id="ed">تعديل</button>${kind === 'clients' ? '<button class="btn gold" id="st">كشف حساب</button>' : '<button class="btn gold" id="ad">مستخلص جديد</button>'}</div>`;
  $('#ed').onclick = () => { formView(kind, idx); };
  if ($('#st')) $('#st').onclick = () => { location.hash = '#/reports/' + encodeURIComponent(x.name); };
  if ($('#ad')) $('#ad').onclick = () => { location.hash = '#/new/bills'; };
};

V.reports = (clientName) => {
  setTitle('التقارير'); const c = calc();
  const sel = clientName ? decodeURIComponent(clientName) : (db.clients[0] && db.clients[0].name) || '';
  const buckets = ['لم يستحق بعد', 'متأخر 1-30', 'متأخر 31-60', 'متأخر 61-90', 'متأخر +90'];
  const age = buckets.map(b => ({ b, v: c.bills.filter(x => x.cls === b).reduce((s, x) => s + x.remaining, 0) }));
  const stmt = (() => {
    const cl = c.clients[sel]; if (!cl) return '<div class="empty">اختر عميلاً</div>';
    const cts = db.contracts.filter(x => x.client === sel);
    return `<div class="card"><div class="row"><div><div class="t">كشف حساب: ${esc(sel)}</div><div class="s">حتى ${fmtDate(c.asOf)}</div></div></div>
    <table><tr><th>المشروع</th><th>العقد</th><th>المحصل</th><th>المتبقي</th></tr>${cts.map(x => { const k = c.contracts[x.name]; return `<tr><td>${esc(x.name)}</td><td class="n">${money(x.value)}</td><td class="n">${money(k.collected)}</td><td class="n">${money(k.remain)}</td></tr>`; }).join('')}
    <tr><th>الإجمالي</th><th class="n">${money(cl.total)}</th><th class="n">${money(cl.collected)}</th><th class="n">${money(cl.remain)}</th></tr></table>
    <h2>مستخلصات مستحقة</h2><table><tr><th>المشروع</th><th>الاستحقاق</th><th>المتبقي</th><th></th></tr>${c.bills.filter(x => x.client === sel && x.remaining > 0.005).map(x => `<tr><td>${esc(x.b.project)}</td><td>${fmtDate(x.due)}</td><td class="n">${money(x.remaining)}</td><td>${pill(x.cls)}</td></tr>`).join('') || '<tr><td colspan="4" class="empty">لا يوجد مستحق</td></tr>'}</table></div>
    <button class="btn ghost" onclick="window.print()">طباعة / حفظ PDF</button>`;
  })();
  $('#app').innerHTML = `<h2>كشف حساب عميل</h2><select id="cs">${db.clients.map(x => `<option ${x.name === sel ? 'selected' : ''}>${esc(x.name)}</option>`).join('')}</select><div style="height:10px"></div>${stmt}
  <h2>أعمار الديون</h2><div class="card"><table>${age.map(a => `<tr><td>${a.b}</td><td class="n">${money(a.v)}</td></tr>`).join('')}</table></div>
  <h2>ربحية المشاريع</h2><div class="card"><table><tr><th>المشروع</th><th>المحصل</th><th>التكاليف</th><th>الصافي</th></tr>${db.contracts.map(x => { const k = c.contracts[x.name]; return `<tr><td>${esc(x.name)}</td><td class="n">${money(k.collected)}</td><td class="n">${money(k.cost)}</td><td class="n">${money(k.net)}</td></tr>`; }).join('')}</table></div>
  <h2>سجل التعديلات</h2><div class="card">${db.log.length ? db.log.slice(0, 15).map(l => `<div class="s">${new Date(l.t).toLocaleString('ar-EG', { dateStyle: 'short', timeStyle: 'short' })} · ${l.what} ${esc(l.kind)}: ${esc(l.name)}</div>`).join('') : '<div class="s">لا توجد تعديلات بعد</div>'}</div>`;
  $('#cs').onchange = e => { location.hash = '#/reports/' + encodeURIComponent(e.target.value); };
};

V.export = () => {
  setTitle('التصدير إلى Excel');
  $('#app').innerHTML = `<div class="card"><div class="t">تحديث ملف واقد المالي</div>
  <p class="s">اختر ملف الإكسل المالي الخاص بك (.xlsx). التطبيق يضع بياناتك في جداول العملاء والعقود والمستخلصات والحركات ويترك كل المعادلات كما هي، ثم يحمّل لك نسخة جديدة. الملف الأصلي لا يتغير. افتح النسخة في Excel وكل الحسابات تتحدث تلقائياً.</p>
  <label for="wb">ملف الإكسل المالي</label><input type="file" id="wb" accept=".xlsx"><button class="btn gold" id="go" disabled>إنشاء ملف Excel محدّث</button><div class="hint" id="msg"></div></div>
  <div class="card"><div class="t">قائمة الخدمات</div><p class="s">أضف أو احذف أنواع الخدمات. التغيير يظهر في العقود، ويُكتب في ملف الإكسل عند التصدير.</p><a class="btn ghost" href="#/lists" style="text-align:center;text-decoration:none">تعديل قائمة الخدمات</a></div>
  <div class="card"><div class="t">المزامنة مع Google Sheet</div><p class="s">كل تعديل يُحفظ تلقائياً في الشيت الخاص بك، فلو اتمسح المتصفح ترجع البيانات بزر الاسترجاع.</p>
  <label for="su">رابط الربط (Web app URL)</label><input id="su" dir="ltr" placeholder="https://script.google.com/macros/s/.../exec"><label for="sk">مفتاح الربط</label><input id="sk" dir="ltr" type="password">
  <div class="btns"><button class="btn gold" id="ssv">حفظ واختبار</button><button class="btn ghost" id="spl">استرجاع من الشيت</button><button class="btn ghost" id="sph">رفع الآن</button></div><div class="hint" id="syncst"></div></div>
  <div class="card"><div class="t">نسخة احتياطية</div><p class="s">احفظ بيانات التطبيق كملف، أو استرجعها منه.</p><div class="btns"><button class="btn ghost" id="bk">حفظ نسخة</button><button class="btn ghost" id="rs">استرجاع</button></div><input type="file" id="rsf" accept=".json" hidden></div>
  <div class="card"><div class="t">بيانات تجريبية</div><p class="s">النسخة الحالية تعمل ببيانات وهمية للتجربة فقط.</p><button class="btn danger" id="clr">مسح كل البيانات والبدء من الصفر</button></div>`;
  let file = null; $('#wb').onchange = e => { file = e.target.files[0]; $('#go').disabled = !file; };
  $('#go').onclick = async () => {
    try { $('#msg').textContent = 'جاري التجهيز...';
      const buf = await file.arrayBuffer(); const r = await WaqidExport.exportWorkbook(JSZip, buf, db);
      const blob = new Blob([r.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'واقد-نظام-مالي-محدث-' + today() + '.xlsx'; a.click();
      $('#msg').textContent = `تم: ${r.report['العملاء']} عميل، ${r.report['العقود']} عقد، ${r.report['المستخلصات']} مستخلص، ${r.report['الحركات']} حركة.`;
    } catch (e) { $('#msg').textContent = 'تعذر التصدير: ' + e.message; }
  };
  { const c = SYNC.get(); $('#su').value = c.url || ''; $('#sk').value = c.key || ''; $('#syncst').textContent = c.last ? 'آخر مزامنة: ' + new Date(c.last).toLocaleString('ar-EG') : ''; }
  $('#ssv').onclick = async () => { SYNC.configure($('#su').value.trim(), $('#sk').value.trim()); try { const p = await SYNC.ping(); $('#syncst').textContent = p.ok ? (p.rev ? 'الربط شغال. يوجد بيانات محفوظة، اضغط استرجاع أو رفع الآن.' : 'الربط شغال. اضغط رفع الآن.') : 'المفتاح غير صحيح'; } catch (e) { $('#syncst').textContent = 'تعذر الاتصال بالرابط'; } };
  $('#spl').onclick = async () => { if (!confirm('استبدال بيانات هذا الجهاز بالنسخة الموجودة على الشيت؟')) return; const r = await SYNC.pull(); if (r.ok) { toast('تم الاسترجاع'); location.hash = '#/home'; } else $('#syncst').textContent = 'لا توجد بيانات على الشيت أو تعذر الاتصال'; };
  $('#sph').onclick = async () => { const r = await SYNC.push(SYNC.get().rev === undefined && confirm('الشيت فيه بيانات قديمة. استبدالها ببيانات هذا الجهاز؟')); if (r.ok) toast('تم الرفع'); else if (r.error === 'conflict') $('#syncst').textContent = 'الشيت فيه نسخة أحدث. استرجعها أولاً.'; };
  $('#bk').onclick = () => { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([JSON.stringify(db)], { type: 'application/json' })); a.download = 'waqid-backup-' + today() + '.json'; a.click(); };
  $('#rs').onclick = () => $('#rsf').click();
  $('#rsf').onchange = async e => { try { const d = JSON.parse(await e.target.files[0].text()); if (!d.clients || !d.journal) throw 0; if (!Array.isArray(d.services)) d.services = LISTS.services.slice(); db = d; save(); toast('تم الاسترجاع'); location.hash = '#/home'; } catch (x) { toast('ملف غير صالح'); } };
  $('#clr').onclick = () => { if (confirm('مسح كل البيانات؟')) { db = { clients: [], contracts: [], bills: [], journal: [], log: [], demo: false, services: LISTS.services.slice() }; save(); toast('تم المسح'); location.hash = '#/home'; } };
};

V.lists = () => {
  setTitle('قائمة الخدمات', true);
  const used = s => db.contracts.some(c => c.service === s);
  const extras = ['إشراف', 'إلكتروميكانيك'].filter(x => !db.services.includes(x));
  $('#app').innerHTML = `<div class="card"><p class="s">القائمة الافتراضية من ملف الإكسل. حد أقصى 12 خدمة (حجم القائمة في الملف).</p>
  ${db.services.map((s, i) => `<div class="row" style="padding:8px 0;border-bottom:1px solid var(--line)"><div>${esc(s)}${used(s) ? ' <span class="pill">مستخدمة</span>' : ''}</div><button class="pill bad" data-del="${i}" style="border:0;font:inherit">حذف</button></div>`).join('')}</div>
  <label for="ns">خدمة جديدة</label><input id="ns" placeholder="اسم الخدمة"><div class="err" id="nse"></div><button class="btn gold" id="addS">إضافة</button>
  ${extras.length ? `<h2>مقترحة من نموذج الورق</h2><div class="seg">${extras.map(x => `<button data-x="${esc(x)}">+ ${esc(x)}</button>`).join('')}</div>` : ''}
  <button class="btn ghost" id="rst">استرجاع القائمة الافتراضية</button>`;
  const add = v => { v = (v || '').trim(); const e = $('#nse'); if (!v) { e.textContent = 'اكتب اسم الخدمة'; return; } if (db.services.includes(v)) { e.textContent = 'موجودة بالفعل'; return; } if (db.services.length >= 12) { e.textContent = 'وصلت للحد الأقصى (12)'; return; } db.services.push(v); logChange('إضافة', 'خدمة', v); save(); V.lists(); };
  $('#addS').onclick = () => add($('#ns').value);
  document.querySelectorAll('[data-x]').forEach(b => b.onclick = () => add(b.dataset.x));
  document.querySelectorAll('[data-del]').forEach(b => b.onclick = () => { const s = db.services[+b.dataset.del]; if (used(s)) return toast('الخدمة مستخدمة في عقود. لا يمكن حذفها'); if (!confirm('حذف الخدمة؟')) return; db.services.splice(+b.dataset.del, 1); logChange('حذف', 'خدمة', s); save(); V.lists(); });
  $('#rst').onclick = () => { if (db.contracts.some(c => !LISTS.services.includes(c.service))) return toast('بعض العقود تستخدم خدمات غير موجودة بالقائمة الافتراضية'); db.services = LISTS.services.slice(); save(); V.lists(); };
};

/* ---------- router ---------- */
function route() {
  const parts = (location.hash || '#/home').slice(2).split('/'); const name = parts[0] || 'home';
  document.querySelectorAll('#tabs a').forEach(a => a.classList.toggle('on', a.dataset.t === name || (['new', 'view'].includes(name) && a.dataset.t === 'add') || (name === 'lists' && a.dataset.t === 'export')));
  $('#demoBanner').hidden = !db.demo;
  (V[name] || V.home)(...parts.slice(1)); window.scrollTo(0, 0);
}
load(); window.addEventListener('hashchange', route); route();
if ('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('sw.js').catch(() => {});
})();
