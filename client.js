/* Waqid client sender: sends each new/changed/deleted record to the office sheet as an append-only event.
   The client key can only add rows. It cannot read, edit or delete anything on the sheet. */
(function (root) {
  const CFG = 'waqid-client-cfg', KINDS = ['clients', 'contracts', 'bills', 'journal'];
  const FIELDS = { clients: ['name', 'phone', 'entity', 'address', 'agreeDate', 'endDate', 'notes'], contracts: ['name', 'client', 'service', 'date', 'status', 'value', 'retention', 'notes'], bills: ['project', 'item', 'no', 'date', 'days', 'value', 'notes'], journal: ['date', 'type', 'project', 'item', 'amount', 'method', 'docNo', 'notes'] };
  const rid = () => { try { const a = new Uint8Array(9); crypto.getRandomValues(a); return Array.from(a, b => b.toString(16).padStart(2, '0')).join(''); } catch (e) { return (Math.random().toString(16).slice(2) + Math.random().toString(16).slice(2)).slice(0, 18).padEnd(18, '0'); } };
  const h53 = s => { let a = 0xdeadbeef, b = 0x41c6ce57; for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); a = Math.imul(a ^ c, 2654435761); b = Math.imul(b ^ c, 1597334677); } a = Math.imul(a ^ (a >>> 16), 2246822507) ^ Math.imul(b ^ (b >>> 13), 3266489909); b = Math.imul(b ^ (b >>> 16), 2246822507) ^ Math.imul(a ^ (a >>> 13), 3266489909); return (4294967296 * (2097151 & b) + (a >>> 0)).toString(16); };
  const pick = (kind, r) => { const o = {}; FIELDS[kind].forEach(k => { if (r[k] !== undefined && r[k] !== null) o[k] = r[k]; }); return o; };
  function make(env) {
    const st = env.storage; let busy = false, timer = null;
    const get = () => { try { return JSON.parse(st.getItem(CFG)) || {}; } catch (e) { return {}; } };
    const put = c => st.setItem(CFG, JSON.stringify(c));
    const status = (s, m) => env.onStatus && env.onStatus(s, m);
    function ensureIds(db) { let n = 0; KINDS.forEach(k => (db[k] || []).forEach(r => { if (!r._id) { r._id = rid(); n++; } })); if (n && env.persist) env.persist(); }
    /* compare with what was already queued, add events for new/changed/removed records */
    function collect() {
      const c = get(); if (!c.url || !c.key) return; const db = env.getDb(); ensureIds(db);
      c.sent = c.sent || {}; c.queue = c.queue || []; const cur = {};
      KINDS.forEach(k => (db[k] || []).forEach(r => { cur[k + ':' + r._id] = h53(JSON.stringify(pick(k, r))); }));
      Object.keys(cur).forEach(key => { if (c.sent[key] !== cur[key]) { const [k, id] = [key.slice(0, key.indexOf(':')), key.slice(key.indexOf(':') + 1)]; const r = db[k].find(x => x._id === id); c.queue.push({ evid: rid(), op: 'upsert', kind: k, id, rec: pick(k, r) }); c.sent[key] = cur[key]; } });
      Object.keys(c.sent).forEach(key => { if (!(key in cur)) { const i = key.indexOf(':'); c.queue.push({ evid: rid(), op: 'delete', kind: key.slice(0, i), id: key.slice(i + 1) }); delete c.sent[key]; } });
      put(c);
    }
    async function send(events) {
      const c = get(); const r = await env.fetch(c.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ key: c.key, action: 'append', events }) });
      return r.json();
    }
    async function flush() {
      const c0 = get(); if (!c0.url || !c0.key) return { ok: false, error: 'not_configured' };
      if (busy) return { ok: false, error: 'busy' }; busy = true;
      try {
        collect();
        for (;;) {
          const c = get(); const batch = (c.queue || []).slice(0, 40); if (!batch.length) break;
          let j; try { j = await send(batch); } catch (e) { status('offline'); return { ok: false, error: 'offline' }; }
          if (!j.ok) { status(j.error === 'unauthorized' ? 'unauth' : 'err', j.error); return j; }
          const c2 = get(); c2.queue = (c2.queue || []).slice(batch.length); c2.last = new Date().toISOString(); put(c2);
        }
        status('ok'); return { ok: true };
      } finally { busy = false; }
    }
    return { get, collect, flush, pending: () => (get().queue || []).length,
      configure(url, key) { const c = get(); c.url = url; c.key = key; c.sent = c.sent || {}; c.queue = c.queue || []; put(c); },
      clear() { st.removeItem(CFG); },
      async test() { try { const j = await send([]); return j.ok ? { ok: true } : { ok: false, error: j.error }; } catch (e) { return { ok: false, error: 'offline' }; } },
      schedule() { if (!get().url) return; try { collect(); } catch (e) {} clearTimeout(timer); status('pending'); timer = setTimeout(flush, env.delay === undefined ? 2500 : env.delay); } };
  }
  if (typeof module !== 'undefined') module.exports = make; else root.WaqidClient = make;
})(this);
