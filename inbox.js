/* Waqid inbox import: merges the client's events (from the office sheet inbox tabs) into the app data.
   Adds new records, updates edited ones, and only MARKS client deletes ("محذوف من العميل"). Nothing is removed until the owner approves. */
(function (root) {
  const ORDER = ['clients', 'contracts', 'bills', 'journal'];
  const FIELDS = { clients: ['name', 'phone', 'entity', 'address', 'agreeDate', 'endDate', 'notes'], contracts: ['name', 'client', 'service', 'date', 'status', 'value', 'retention', 'notes'], bills: ['project', 'item', 'no', 'date', 'days', 'value', 'notes'], journal: ['date', 'type', 'project', 'item', 'amount', 'method', 'docNo', 'notes'] };
  const NUM = { clients: [], contracts: ['value', 'retention'], bills: ['days', 'value'], journal: ['amount'] };
  const LABEL = { clients: 'عميل', contracts: 'عقد / مشروع', bills: 'مستخلص / دفعة', journal: 'حركة' };
  const TITLE = { clients: x => x.name, contracts: x => x.name, bills: x => `${x.item || ''} ${x.no ? '#' + x.no : ''}`.trim() + ' · ' + (x.project || ''), journal: x => `${x.item || ''} · ${x.date || ''}` };
  const FLAG = 'محذوف من العميل';
  const ID = /^[A-Za-z0-9_-]{4,40}$/, EV = /^[A-Za-z0-9_-]{6,40}$/;
  function clean(kind, rec) {
    const o = {}; rec = rec && typeof rec === 'object' ? rec : {};
    FIELDS[kind].forEach(k => {
      let v = rec[k]; if (v === undefined || v === null) v = '';
      if (NUM[kind].includes(k)) { const n = parseFloat(v); v = isFinite(n) ? n : (k === 'retention' ? 0 : (k === 'days' ? 30 : '')); } else v = String(v).slice(0, 500);
      o[k] = v;
    });
    return o;
  }
  function apply(db, events, now) {
    now = now || new Date().toISOString();
    ORDER.forEach(k => { db[k] = db[k] || []; }); db.trash = db.trash || []; db.log = db.log || []; db.services = db.services || [];
    db.inboxApplied = db.inboxApplied || []; db.inboxReview = db.inboxReview || [];
    const done = new Set(db.inboxApplied);
    const res = { added: 0, updated: 0, flagged: 0, review: 0, already: 0, superseded: 0, pending: 0, ackable: [] };
    const log = (what, kind, rec) => { db.log.unshift({ t: now, what, kind: LABEL[kind], name: TITLE[kind](rec) }); db.log = db.log.slice(0, 300); };
    const good = (events || []).filter(e => e && ORDER.includes(e.kind) && (e.op === 'upsert' || e.op === 'delete') && EV.test(String(e.evid)) && ID.test(String(e.id)));
    good.sort((a, b) => ORDER.indexOf(a.kind) - ORDER.indexOf(b.kind)); // clients first, so contracts and bills find their parent
    const finish = e => { db.inboxApplied.push(e.evid); done.add(e.evid); res.ackable.push({ kind: e.kind, evid: e.evid }); };
    const toReview = (e, reason, rec) => { db.inboxReview.push({ t: now, kind: e.kind, id: e.id, evid: e.evid, reason, rec }); res.review++; finish(e); };
    const last = {}; good.forEach((e, n) => { last[e.kind + ':' + e.id] = n; }); // several events for one record: only the newest counts, older ones are history
    good.forEach((e, n) => {
      if (done.has(e.evid)) { res.already++; res.ackable.push({ kind: e.kind, evid: e.evid }); return; }
      if (last[e.kind + ':' + e.id] !== n) { res.superseded++; finish(e); return; }
      const arr = db[e.kind], i = arr.findIndex(r => r._id === e.id);
      if (e.op === 'delete') {
        if (i >= 0 && !arr[i].clientDeleted) { arr[i].clientDeleted = FLAG; res.flagged++; log('طلب حذف من العميل', e.kind, arr[i]); }
        finish(e); return;
      }
      const rec = clean(e.kind, e.rec);
      if (i < 0 && db.trash.some(t => t.kind === e.kind && t.rec && t.rec._id === e.id)) return toReview(e, 'سبق حذفه من المكتب', rec);
      if (e.kind === 'contracts' && !db.clients.some(c => c.name === rec.client)) { res.pending++; return; }
      if (e.kind === 'bills' && !db.contracts.some(c => c.name === rec.project)) { res.pending++; return; }
      if (e.kind === 'journal' && rec.type !== 'مصروف عمومي' && !db.contracts.some(c => c.name === rec.project)) { res.pending++; return; }
      if (i < 0 && e.kind === 'clients' && db.clients.some(c => c.name === rec.name)) return toReview(e, 'عميل بنفس الاسم موجود بالفعل', rec);
      if (i < 0 && e.kind === 'contracts' && db.contracts.some(c => c.name === rec.name)) return toReview(e, 'مشروع بنفس الاسم موجود بالفعل', rec);
      if (i < 0 && e.kind === 'clients' && !rec.name) return toReview(e, 'بدون اسم', rec);
      if (e.kind === 'contracts' && rec.service && !db.services.includes(rec.service) && db.services.length < 12) db.services.push(rec.service);
      if (i >= 0) {
        const old = arr[i];
        if (e.kind === 'clients' && old.name !== rec.name) db.contracts.forEach(c => { if (c.client === old.name) c.client = rec.name; });
        if (e.kind === 'contracts' && old.name !== rec.name) { db.bills.forEach(b => { if (b.project === old.name) b.project = rec.name; }); db.journal.forEach(j => { if (j.project === old.name) j.project = rec.name; }); }
        const same = FIELDS[e.kind].every(k => old[k] === rec[k]) && !old.clientDeleted;
        Object.assign(old, rec); delete old.clientDeleted;
        if (!same) { res.updated++; log('تعديل من العميل', e.kind, old); }
      } else { arr.push(Object.assign({ _id: e.id }, rec)); res.added++; log('إضافة من العميل', e.kind, rec); }
      finish(e);
    });
    db.inboxApplied = db.inboxApplied.slice(-3000);
    return res;
  }
  const flagged = db => ORDER.reduce((a, k) => a.concat((db[k] || []).map((x, i) => ({ kind: k, i, x })).filter(r => r.x.clientDeleted)), []);
  if (typeof module !== 'undefined') module.exports = { apply, flagged, FLAG, LABEL, TITLE }; else root.WaqidInbox = { apply, flagged, FLAG, LABEL, TITLE };
})(this);
