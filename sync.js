/* Waqid sync client: whole-database snapshot to a Google Apps Script web app. */
(function (root) {
  const CFG = 'waqid-sync-cfg';
  function make(env) {
    const st = env.storage; let timer = null, busy = false;
    const get = () => { try { return JSON.parse(st.getItem(CFG)) || {}; } catch (e) { return {}; } };
    const put = c => st.setItem(CFG, JSON.stringify(c));
    const status = (s, m) => env.onStatus && env.onStatus(s, m);
    async function call(action, extra) {
      const c = get(); if (!c.url || !c.key) throw new Error('not_configured');
      const r = await env.fetch(c.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ key: c.key, action }, extra)) });
      const j = await r.json(); return j;
    }
    const hasData = d => d && (d.clients.length || d.contracts.length || d.bills.length || d.journal.length);
    async function push(force) {
      const c = get(); if (!c.url) return { ok: false, error: 'not_configured' };
      if (busy) return { ok: false, error: 'busy' }; busy = true;
      try {
        const d = env.getDb();
        if (!hasData(d) && !force) { status('warn', 'البيانات فارغة، لم يتم الرفع حمايةً للنسخة المحفوظة'); return { ok: false, error: 'empty' }; }
        if (c.rev === undefined) { const p = await call('ping'); if (!p.ok) { status('err', p.error); return p; } if (p.rev > 0 && !force) { status('conflict'); return { ok: false, error: 'conflict' }; } c.rev = p.rev; }
        const j = await call('push', { baseRev: c.rev, db: d });
        if (j.ok) { c.rev = j.rev; c.last = new Date().toISOString(); put(c); status('ok'); } else status(j.error === 'conflict' ? 'conflict' : 'err', j.error);
        return j;
      } catch (e) { status('offline'); return { ok: false, error: 'offline' }; } finally { busy = false; }
    }
    async function pull() {
      try { const j = await call('pull'); if (!j.ok) { status('err', j.error); return j; }
        if (!j.db) return { ok: false, error: 'nothing' };
        const c = get(); c.rev = j.rev; c.last = new Date().toISOString(); put(c); env.setDb(j.db); status('ok'); return j;
      } catch (e) { status('offline'); return { ok: false, error: 'offline' }; }
    }
    return { get, put, push, pull, ping: () => call('ping'),
      configure(url, key) { const c = get(); c.url = url; c.key = key; delete c.rev; put(c); },
      schedule() { if (!get().url) return; clearTimeout(timer); status('pending'); timer = setTimeout(() => push(false), env.delay === undefined ? 3000 : env.delay); },
      flush: () => { clearTimeout(timer); return push(false); } };
  }
  if (typeof module !== 'undefined') module.exports = make; else root.WaqidSync = make;
})(this);
