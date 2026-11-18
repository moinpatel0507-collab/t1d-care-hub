// Backend for the family app: Supabase login, family membership and a small
// document store that mimics the API the app was written against
// (db.doc / db.collection / onSnapshot), so app.js needs almost no changes.
const { createClient } = window.supabase || {};

const RECOVERY = /type=recovery/.test(location.hash);   // opened from a 'reset password' email
const URL = (window.T1D_CONFIG || {}).url;
const KEY = (window.T1D_CONFIG || {}).key;
export const supabase = URL && KEY && createClient ? createClient(URL, KEY, { auth: { persistSession: true, autoRefreshToken: true } }) : null;

const ROOT = 'family/t1d';           // path prefix the app uses
let familyId = null, me = null, members = [], family = null;
const cache = new Map();             // "collection/id" -> data
const listeners = new Set();         // functions to call after any change
let readyResolve; const ready = new Promise(r => (readyResolve = r));

/* ---------------- path mapping ----------------
   "family/t1d"                -> collection "plan", id "plan"
   "family/t1d/<col>"          -> collection <col>
   "family/t1d/<col>/<id>"     -> collection <col>, id <id>          */
function splitPath(path) {
  if (path === ROOT) return { col: 'plan', id: 'plan' };
  if (!path.startsWith(ROOT + '/')) throw new TypeError('Unsupported path ' + path);
  const rest = path.slice(ROOT.length + 1).split('/');
  if (rest.length === 1) return { col: rest[0] };
  if (rest.length === 2) return { col: rest[0], id: rest[1] };
  throw new TypeError('Unsupported path ' + path);
}
const key = (col, id) => col + '/' + id;

/* ---------------- offline: copy of the data on this phone + a queue of unsent changes ---------------- */
const idb = (() => {
  let dbp = null;
  const open = () => dbp ??= new Promise((res, rej) => { const r = indexedDB.open('t1d-offline', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  const tx = async (mode, fn) => { const d = await open(); return new Promise((res, rej) => { const t = d.transaction('kv', mode); const q = fn(t.objectStore('kv')); t.oncomplete = () => res(q?.result); t.onerror = () => rej(t.error); }); };
  return { get: k => tx('readonly', s => s.get(k)).catch(() => undefined), set: (k, v) => tx('readwrite', s => s.put(v, k)).catch(() => {}), clear: () => tx('readwrite', s => s.clear()).catch(() => {}) };
})();
let queue = [];                       // [{op:'set'|'del', col, id, data, at}]
let offline = false, flushing = false;
const isNetErr = (e) => !navigator.onLine || /fetch|network|load failed|timed? ?out|ECONN|503|502|504/i.test(String(e?.message || e?.details || e || ''));
const saveQueue = () => idb.set('queue:' + familyId, queue);
let snapTimer = null;
const saveSnapSoon = () => { clearTimeout(snapTimer); snapTimer = setTimeout(() => { if (familyId) idb.set('snap:' + familyId, [...cache.entries()]); }, 1500); };
function applyQueue(map) { for (const q of queue) { const k = key(q.col, q.id); if (q.op === 'del') map.delete(k); else map.set(k, q.data); } }
function enqueue(op, col, id, data) {
  queue = queue.filter(q => !(q.col === col && q.id === id));      // newest change wins
  queue.push({ op, col, id, data: data ?? null, at: Date.now() }); saveQueue(); syncPill();
}
async function flushQueue() {
  if (flushing || !queue.length || !familyId) return; flushing = true; syncPill();
  try {
    while (queue.length) {
      const q = queue[0];
      const { error } = q.op === 'del'
        ? await supabase.from('records').delete().match({ family_id: familyId, collection: q.col, id: q.id })
        : await supabase.from('records').upsert({ family_id: familyId, collection: q.col, id: q.id, data: q.data, updated_by: me, updated_at: new Date(q.at).toISOString() }, { onConflict: 'family_id,collection,id' });
      if (error) { if (isNetErr(error)) break; console.warn('dropped unsyncable change', q, error); }
      queue.shift(); saveQueue();
    }
  } catch (e) { /* still offline */ }
  flushing = false;
  if (!queue.length && offline) await goOnline();
  syncPill();
}
function syncPill() {
  const top = document.querySelector('.hero-top'); if (!top) return;
  let el = document.getElementById('syncPill');
  const n = queue.length;
  if (!offline && !n) { el?.remove(); return; }
  if (!el) { el = document.createElement('span'); el.id = 'syncPill'; el.className = 'sync-pill'; el.setAttribute('role', 'status'); top.insertBefore(el, top.querySelector('.acct') || null); }
  el.classList.toggle('off', offline);
  el.textContent = offline ? (n ? `Offline · ${n} to sync` : 'Offline · saving on this phone') : (flushing ? `Syncing ${n}…` : `${n} to sync`);
}
setInterval(() => { if (queue.length) flushQueue(); }, 30000);
addEventListener('online', () => { flushQueue(); if (offline) goOnline(); });
addEventListener('offline', () => { offline = true; syncPill(); });
const meta = { fromCache: false, hasPendingWrites: false };
const snapDoc = (col, id) => { const d = cache.get(key(col, id)); return { id, exists: !!d, data: () => d, metadata: meta }; };
const notify = () => { listeners.forEach(f => { try { f(); } catch (e) { console.error(e); } }); saveSnapSoon(); };
let notifyTimer = null;   // live changes can arrive in bursts (e.g. a restore): redraw once
const notifySoon = () => { clearTimeout(notifyTimer); notifyTimer = setTimeout(notify, 60); };
const err = (e) => Object.assign(new Error(e?.message || 'Save failed'), { code: e?.code === '42501' ? 'not_granted' : (e?.code || 'unavailable') });

async function writeRow(col, id, data) {
  if (offline || !navigator.onLine || queue.length) { enqueue('set', col, id, data); flushQueue(); return; }
  let error;
  try { ({ error } = await supabase.from('records').upsert({ family_id: familyId, collection: col, id, data, updated_by: me, updated_at: new Date().toISOString() }, { onConflict: 'family_id,collection,id' })); }
  catch (e) { error = e; }
  if (error && isNetErr(error)) { offline = !navigator.onLine || offline; enqueue('set', col, id, data); return; }
  if (error) throw err(error);
}
async function deleteRow(col, id) {
  if (offline || !navigator.onLine || queue.length) { enqueue('del', col, id); flushQueue(); return; }
  let error;
  try { ({ error } = await supabase.from('records').delete().match({ family_id: familyId, collection: col, id })); } catch (e) { error = e; }
  if (error && isNetErr(error)) { enqueue('del', col, id); return; }
  if (error) throw err(error);
}
const newId = () => (crypto.randomUUID ? crypto.randomUUID().replace(/-/g, '').slice(0, 20) : Date.now().toString(36) + Math.random().toString(36).slice(2, 10));

function docRef(path) {
  const { col, id } = splitPath(path);
  return {
    id, path,
    async get() {
      // photos are big, so they are not loaded with everything else: fetch one when it is asked for
      if (col === 'photos' && !cache.has(key(col, id)) && supabase && familyId && !offline) {
        try { const { data } = await supabase.from('records').select('data').eq('family_id', familyId).eq('collection', 'photos').eq('id', id).limit(1); if (data && data[0]) cache.set(key(col, id), data[0].data); } catch {}
      }
      return snapDoc(col, id);
    },
    async set(data) { const prev = cache.get(key(col, id)); cache.set(key(col, id), JSON.parse(JSON.stringify(data))); notifySoon(); try { await writeRow(col, id, data); } catch (e) { prev ? cache.set(key(col, id), prev) : cache.delete(key(col, id)); notifySoon(); throw e; } },
    async update(data) { const merged = { ...(cache.get(key(col, id)) || {}), ...data }; return this.set(merged); },
    async delete() { const prev = cache.get(key(col, id)); cache.delete(key(col, id)); notifySoon(); try { await deleteRow(col, id); } catch (e) { if (prev) cache.set(key(col, id), prev); notifySoon(); throw e; } },
    onSnapshot(next) { const f = () => next(snapDoc(col, id)); listeners.add(f); queueMicrotask(f); return () => listeners.delete(f); },
    collection(sub) { return colRef(path + '/' + sub); },
  };
}
function colRef(path, order = null, lim = null) {
  // a sub-collection under a doc, e.g. family/t1d/meta/backup is a doc; only one level is used by the app
  const { col } = splitPath(path);
  const run = () => {
    let docs = []; for (const [k, v] of cache) { const i = k.indexOf('/'); if (k.slice(0, i) === col) docs.push({ id: k.slice(i + 1), exists: true, data: () => v, metadata: meta }); }
    if (order) { const [f, dir] = order; docs.sort((a, b) => { const x = a.data()[f], y = b.data()[f]; return (x > y ? 1 : x < y ? -1 : 0) * (dir === 'desc' ? -1 : 1); }); }
    else docs.sort((a, b) => (a.id > b.id ? 1 : -1));
    if (lim) docs = docs.slice(0, lim);
    return { docs, size: docs.length, empty: !docs.length, metadata: meta, docChanges: () => [] };
  };
  return {
    path,
    doc: (id) => docRef(ROOT + '/' + col + '/' + (id || newId())),
    async add(data) { const r = docRef(ROOT + '/' + col + '/' + newId()); await r.set(data); return r; },
    orderBy: (f, dir = 'asc') => colRef(path, [f, dir], lim),
    limit: (n) => colRef(path, order, n),
    where() { return this; },
    async get() { return run(); },
    onSnapshot(next) { const f = () => next(run()); listeners.add(f); queueMicrotask(f); return () => listeners.delete(f); },
  };
}
const db = { doc: docRef, collection: colRef };

/* ---------------- loading & live sync ---------------- */
async function loadAll() {
  const fresh = new Map(); const page = 1000;
  for (let from = 0; ; from += page) {
    const { data, error } = await supabase.from('records').select('collection,id,data').eq('family_id', familyId).neq('collection', 'photos').range(from, from + page - 1);
    if (error) throw error;
    data.forEach(r => fresh.set(key(r.collection, r.id), r.data));
    if (data.length < page) break;
  }
  for (const [k, v] of cache) if (k.startsWith('photos/')) fresh.set(k, v);   // keep photos already opened
  applyQueue(fresh);
  cache.clear(); fresh.forEach((v, k) => cache.set(k, v)); notify();
}
let subscribed = false;
function subscribe() {
  if (subscribed) return; subscribed = true;
  supabase.channel('family-' + familyId)
    .on('postgres_changes', { event: '*', schema: 'public', table: 'records', filter: 'family_id=eq.' + familyId }, (p) => {
      if (p.eventType === 'DELETE') { const o = p.old || {}; if (o.collection) cache.delete(key(o.collection, o.id)); }
      else { const n = p.new; cache.set(key(n.collection, n.id), n.data); }
      notifySoon();
    })
    .subscribe();
  // phones pause background connections: refresh when the app comes back
  document.addEventListener('visibilitychange', () => { if (!document.hidden) flushQueue().then(() => loadAll()).catch(() => {}); });
  addEventListener('online', () => flushQueue().then(() => loadAll()).catch(() => {}));
}

/* ---------------- capabilities the app asks for ---------------- */
const userCap = {
  id: async () => me,
  profiles: async (ids) => Object.fromEntries(ids.map(id => [id, { name: members.find(m => m.user_id === id)?.display_name || '' }])),
};
const downloads = {
  async save({ filename, data }) {
    const blob = data instanceof Blob ? data : new Blob([data], { type: 'application/json' });
    const file = new File([blob], filename, { type: blob.type || 'application/json' });
    if (navigator.canShare?.({ files: [file] }) && /iPhone|iPad|Android/i.test(navigator.userAgent)) {
      try { await navigator.share({ files: [file], title: filename }); return { status: 'saved' }; }
      catch (e) { if (e.name === 'AbortError') throw Object.assign(new Error('Cancelled'), { code: 'declined' }); }
    }
    const a = Object.assign(document.createElement('a'), { href: URL_.createObjectURL(blob), download: filename });
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL_.revokeObjectURL(a.href), 4000);
    return { status: 'saved' };
  },
};
const URL_ = globalThis.URL;
window.claude = {
  use: async (name) => {
    if (name === 'db') { await ready; return familyId ? db : null; }
    if (name === 'user') { await ready; return familyId ? userCap : null; }
    if (name === 'downloads') return downloads;
    return null;
  },
};

/* ---------------- sign-in and family screens ---------------- */
const gate = document.createElement('div');
gate.id = 'authGate';
document.body.appendChild(gate);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
function screen(html) { gate.hidden = false; gate.innerHTML = `<div class="gate-card"><div class="gate-brand"><span class="ring"></span><span>T1D Care Hub</span></div>${html}</div>`; }
const msg = (t, bad) => { const m = gate.querySelector('.gate-msg'); if (m) { m.textContent = t; m.className = 'gate-msg' + (bad ? ' bad' : ''); } };

const eye = (id) => `<button type="button" class="pwtoggle" data-for="${id}" aria-label="Show password">Show</button>`;
function wirePwToggles(){ gate.querySelectorAll('.pwtoggle').forEach(b => b.onclick = () => { const i = gate.querySelector('#'+b.dataset.for); const show = i.type === 'password'; i.type = show ? 'text' : 'password'; b.textContent = show ? 'Hide' : 'Show'; }); }

function showSignIn(mode = 'in') {
  if (mode === 'reset') return showForgot();
  const up = mode === 'up';
  screen(`<div class="gate-tabs" role="tablist"><button type="button" role="tab" class="${up?'':'on'}" id="tabIn">Sign in</button><button type="button" role="tab" class="${up?'on':''}" id="tabUp">Register</button></div>
    <h1>${up ? 'Create your account' : 'Welcome back'}</h1>
    <p class="muted">${up ? 'One account each for you and his mum. After this you create the family, or join it with a family code.' : 'Sign in to see his care hub.'}</p>
    <form id="authForm" class="gate-form" novalidate>
      <label>Email<input id="a_email" type="email" autocomplete="email" inputmode="email" required></label>
      <label>Password ${up?'<small>at least 8 characters</small>':''}<span class="pwrow"><input id="a_pass" type="password" autocomplete="${up ? 'new-password' : 'current-password'}" minlength="8" required>${eye('a_pass')}</span></label>
      ${up ? `<label>Confirm password<span class="pwrow"><input id="a_pass2" type="password" autocomplete="new-password" minlength="8" required>${eye('a_pass2')}</span></label>` : ''}
      <button class="primary" type="submit">${up ? 'Create account' : 'Sign in'}</button>
      <p class="gate-msg" aria-live="polite"></p>
    </form>
    ${up ? '' : '<button type="button" class="linkbtn" id="forgot">Forgot your password?</button>'}`);
  wirePwToggles();
  gate.querySelector('#tabIn').onclick = () => showSignIn('in');
  gate.querySelector('#tabUp').onclick = () => showSignIn('up');
  if (!up) gate.querySelector('#forgot').onclick = () => showForgot(gate.querySelector('#a_email').value.trim());
  gate.querySelector('#authForm').onsubmit = async (e) => {
    e.preventDefault();
    const email = gate.querySelector('#a_email').value.trim(), password = gate.querySelector('#a_pass').value;
    if (!/^\S+@\S+\.\S+$/.test(email)) return msg('Enter a valid email address.', true);
    if (password.length < 8) return msg('The password needs at least 8 characters.', true);
    if (up && password !== gate.querySelector('#a_pass2').value) return msg("The two passwords don't match.", true);
    const btn = gate.querySelector('button[type=submit]'); btn.disabled = true;
    msg(up ? 'Creating your account…' : 'Signing in…');
    const { data, error } = up ? await supabase.auth.signUp({ email, password, options: { emailRedirectTo: location.origin } }) : await supabase.auth.signInWithPassword({ email, password });
    btn.disabled = false;
    if (error) return msg(/invalid login/i.test(error.message) ? 'Wrong email or password. Try again, or reset your password.' : /already registered/i.test(error.message) ? 'That email already has an account. Sign in instead.' : error.message, true);
    if (up && !data.session) {
      screen(`<h1>Check your email</h1><p class="muted">We sent a confirmation link to <strong>${esc(email)}</strong>. Tap it, then come back here and sign in.</p><p class="muted">Can't find it? Check your spam folder.</p><button class="primary" id="toIn">Go to sign in</button>`);
      gate.querySelector('#toIn').onclick = () => showSignIn('in');
    }
  };
}
function showForgot(prefill = '') {
  screen(`<h1>Reset your password</h1><p class="muted">Enter your email and we'll send you a link to choose a new password.</p>
    <form id="fpForm" class="gate-form" novalidate><label>Email<input id="fp_email" type="email" autocomplete="email" value="${esc(prefill)}" required></label>
    <button class="primary" type="submit">Send reset link</button><p class="gate-msg" aria-live="polite"></p></form>
    <button type="button" class="linkbtn" id="back">Back to sign in</button>`);
  gate.querySelector('#back').onclick = () => showSignIn('in');
  gate.querySelector('#fpForm').onsubmit = async (e) => {
    e.preventDefault(); const email = gate.querySelector('#fp_email').value.trim();
    if (!/^\S+@\S+\.\S+$/.test(email)) return msg('Enter a valid email address.', true);
    msg('Sending…'); const { error } = await supabase.auth.resetPasswordForEmail(email, { redirectTo: location.origin });
    if (error) return msg(error.message, true);
    msg('If that email has an account, a reset link is on its way. Open it on this device.');
  };
}
function showNewPassword() {
  screen(`<h1>Choose a new password</h1><p class="muted">At least 8 characters.</p>
    <form id="npForm" class="gate-form" novalidate>
      <label>New password<span class="pwrow"><input id="np1" type="password" autocomplete="new-password" minlength="8" required>${eye('np1')}</span></label>
      <label>Confirm new password<span class="pwrow"><input id="np2" type="password" autocomplete="new-password" minlength="8" required>${eye('np2')}</span></label>
      <button class="primary" type="submit">Save new password</button><p class="gate-msg" aria-live="polite"></p></form>`);
  wirePwToggles();
  gate.querySelector('#npForm').onsubmit = async (e) => {
    e.preventDefault(); const a = gate.querySelector('#np1').value, b = gate.querySelector('#np2').value;
    if (a.length < 8) return msg('The password needs at least 8 characters.', true);
    if (a !== b) return msg("The two passwords don't match.", true);
    msg('Saving…'); const { error } = await supabase.auth.updateUser({ password: a });
    if (error) return msg(error.message, true);
    history.replaceState(null, '', location.pathname); recovering = false; start();
  };
}
function showCode(code, onDone) {
  screen(`<h1>Your family is ready</h1><p class="muted">Send this code to his mum. She creates her own account, then taps <strong>Join family</strong> and types it in.</p>
    <div class="codebox" aria-label="Family code">${esc(code).split('').map(c => `<span>${c}</span>`).join('')}</div>
    <div class="gate-form"><div class="btns" style="justify-content:center"><button type="button" id="cpCode">Copy code</button>${navigator.share ? '<button type="button" id="shCode">Share</button>' : ''}</div>
    <button class="primary" type="button" id="goApp">Continue to the app</button><p class="gate-msg" aria-live="polite"></p></div>
    <p class="muted" style="font-size:.85rem">You can find it again any time in Care plan → Family &amp; account.</p>`);
  gate.querySelector('#cpCode').onclick = async () => { try { await navigator.clipboard.writeText(code); msg('Copied.'); } catch { msg('Copy it by hand: ' + code); } };
  const sh = gate.querySelector('#shCode'); if (sh) sh.onclick = () => navigator.share({ text: `Join our T1D Care Hub. Open ${location.origin}, create an account, tap Join family and enter this code: ${code}` }).catch(() => {});
  gate.querySelector('#goApp').onclick = onDone;
}
function showFamilySetup() {
  screen(`<h1>Set up your family</h1>
    <p class="muted">The first parent creates the family and gets a code. The second parent joins with that code.</p>
    <div class="gate-form"><label>Your name <small>shown on entries, e.g. Dad or Mum</small><input id="f_name" maxlength="40" autocomplete="given-name" required></label></div>
    <div class="gate-split">
      <div><h2>First parent</h2><p class="muted" style="font-size:.85rem">Creates the family and a code to share.</p><button class="primary" id="mk">Create family</button></div>
      <div><h2>Second parent</h2><label>Family code<input id="f_code" maxlength="8" autocapitalize="characters" autocomplete="off" spellcheck="false" placeholder="e.g. 7F3A9C21"></label><button id="jn">Join family</button></div>
    </div>
    <p class="gate-msg" aria-live="polite"></p>
    <button type="button" class="linkbtn" id="out">Sign out</button>`);
  const name = () => gate.querySelector('#f_name').value.trim();
  gate.querySelector('#out').onclick = () => supabase.auth.signOut();
  gate.querySelector('#mk').onclick = async () => {
    if (!name()) return msg('Enter your name first.', true); msg('Creating your family…');
    const { data, error } = await supabase.rpc('create_family', { p_name: name() });
    if (error) return msg(error.message, true);
    const code = (Array.isArray(data) ? data[0] : data)?.invite_code;
    code ? showCode(code, start) : start();
  };
  gate.querySelector('#jn').onclick = async () => {
    const code = gate.querySelector('#f_code').value.replace(/\s/g, '').toUpperCase();
    if (!name()) return msg('Enter your name first.', true);
    if (!/^[0-9A-F]{8}$/.test(code)) return msg('Family codes are 8 letters and numbers (0–9, A–F). Check it and try again.', true);
    msg('Joining…'); const { error } = await supabase.rpc('join_family', { p_code: code, p_name: name() });
    if (error) return msg(/not found/i.test(error.message) ? "That code didn't match a family. Check it with the first parent." : /full/i.test(error.message) ? 'That family already has 4 people.' : error.message, true);
    start();
  };
}

let started = false, recovering = RECOVERY;
async function start() {
  const { data: { session } } = await supabase.auth.getSession();
  if (!session) return showSignIn();
  if (recovering) return showNewPassword();
  me = session.user.id;
  let mem = null, error = null;
  try { ({ data: mem, error } = await supabase.from('family_members').select('family_id, user_id, display_name, families(invite_code)').eq('user_id', me).limit(1)); } catch (e) { error = e; }
  if (error && isNetErr(error)) return startOffline();
  if (error) { screen(`<h1>Can't reach the database</h1><p class="muted">${esc(error.message)}</p><button class="primary" onclick="location.reload()">Try again</button>`); return; }
  if (!mem.length) return showFamilySetup();
  familyId = mem[0].family_id; family = mem[0].families;
  const { data: all } = await supabase.from('family_members').select('user_id, display_name').eq('family_id', familyId);
  members = all || [];
  idb.set('fam', { me, familyId, family, members });
  queue = (await idb.get('queue:' + familyId)) || [];
  if (queue.length) await flushQueue();
  try { await loadAll(); } catch (e) { if (isNetErr(e)) return startOffline(); screen(`<h1>Couldn't load your data</h1><p class="muted">${esc(e.message)}</p><button class="primary" onclick="location.reload()">Try again</button>`); return; }
  gate.hidden = true; gate.innerHTML = '';
  offline = false; syncPill(); exposeFamily();
  if (!started) { started = true; subscribe(); addFamilyPanel(); addAccountChip(); addAutoBackupPanel(); readyResolve(); }
  else if (!subscribed) subscribe();
}
function exposeFamily() {
  window.__family = { members: members.length, code: family?.invite_code || '', signupsClosed: window.__family?.signupsClosed ?? null };
  fetch(URL.replace(/\/+$/, '') + '/auth/v1/settings', { headers: { apikey: KEY } }).then(r => r.ok ? r.json() : null)
    .then(j => { if (j && 'disable_signup' in j) { window.__family.signupsClosed = !!j.disable_signup; notify(); } }).catch(() => {});
}
window.__familyRefresh = async () => { try { const { data } = await supabase.from('family_members').select('user_id, display_name').eq('family_id', familyId); if (data) members = data; } catch {} exposeFamily(); notify(); };
async function startOffline() {
  const fam = await idb.get('fam');
  if (!fam || fam.me !== me) { screen(`<h1>No connection</h1><p class="muted">The app needs the internet the first time you sign in on this phone. After that it works without signal.</p><button class="primary" onclick="location.reload()">Try again</button>`); return; }
  ({ familyId, family, members } = fam); offline = true; window.__family = { members: members.length, code: family?.invite_code || '', signupsClosed: null };
  queue = (await idb.get('queue:' + familyId)) || [];
  const snap = (await idb.get('snap:' + familyId)) || [];
  const m = new Map(snap); applyQueue(m); cache.clear(); m.forEach((v, k) => cache.set(k, v));
  gate.hidden = true; gate.innerHTML = '';
  if (!started) { started = true; addFamilyPanel(); addAccountChip(); readyResolve(); }
  notify(); syncPill();
}
async function goOnline() {
  if (!navigator.onLine) return;
  try {
    const { error } = await supabase.from('family_members').select('family_id').eq('user_id', me).limit(1);
    if (error) return;
    offline = false; await flushQueue(); await loadAll(); if (!subscribed) subscribe();
    if (!document.getElementById('autoBk')) addAutoBackupPanel();
    syncPill();
  } catch {}
}
function addAccountChip() {
  const top = document.querySelector('.hero-top'); if (!top || document.getElementById('acct')) return;
  const meName = members.find(m => m.user_id === me)?.display_name || 'You';
  const w = document.createElement('div'); w.id = 'acct'; w.className = 'acct';
  w.innerHTML = `<button type="button" class="acct-btn" aria-haspopup="true" aria-expanded="false"><span class="acct-av">${esc(meName.slice(0,1).toUpperCase())}</span><span class="acct-nm">${esc(meName)}</span></button>
    <div class="acct-menu" hidden><div class="acct-who">Signed in as <strong>${esc(meName)}</strong></div>
    <a href="#plan" class="acct-item">Family code &amp; members</a><button type="button" class="acct-item acct-out">Sign out</button></div>`;
  top.appendChild(w);
  const btn = w.querySelector('.acct-btn'), menu = w.querySelector('.acct-menu');
  const close = () => { menu.hidden = true; btn.setAttribute('aria-expanded', 'false'); };
  btn.onclick = (e) => { e.stopPropagation(); menu.hidden = !menu.hidden; btn.setAttribute('aria-expanded', String(!menu.hidden)); };
  document.addEventListener('click', (e) => { if (!w.contains(e.target)) close(); });
  w.querySelector('a.acct-item').onclick = () => { close(); setTimeout(() => document.querySelector('#v-plan .panel:last-child')?.scrollIntoView({ behavior: 'smooth' }), 150); };
  w.querySelector('.acct-out').onclick = async () => { w.querySelector('.acct-out').textContent = 'Signing out…'; await idb.clear(); await supabase.auth.signOut(); location.reload(); };
}
function addFamilyPanel() {
  const host = document.querySelector('#v-plan'); if (!host) return;
  const p = document.createElement('div'); p.className = 'panel';
  p.innerHTML = `<div><h2>Family &amp; account</h2><div class="muted">Give this code to his mum. She creates her own account, then taps <em>Join family</em> and enters it.</div></div>
    <div class="strip" style="grid-template-columns:repeat(auto-fit,minmax(150px,1fr))">
      <div class="stat"><span class="k">Family code</span><span class="v" style="user-select:all">${esc(family?.invite_code || '—')}</span><span class="n">Keep it private</span></div>
      <div class="stat"><span class="k">Members</span><span class="v">${members.length}</span><span class="n">${esc(members.map(m => m.user_id === me ? (m.display_name || 'You') + ' (you)' : m.display_name || 'Parent').join(', '))}</span></div>
    </div>
    <div class="btns"><button type="button" id="copyFam">Copy code</button>${navigator.share ? '<button type="button" id="shareFam">Share code</button>' : ''}<button type="button" id="signOut">Sign out</button></div>`;
  host.appendChild(p);
  p.querySelector('#copyFam').onclick = async () => { try { await navigator.clipboard.writeText(family?.invite_code || ''); } catch {} };
  const sf = p.querySelector('#shareFam'); if (sf) sf.onclick = () => navigator.share({ text: `Join our T1D Care Hub. Open ${location.origin}, create an account, tap Join family and enter this code: ${family?.invite_code}` }).catch(() => {});
  p.querySelector('#signOut').onclick = async () => { await idb.clear(); await supabase.auth.signOut(); location.reload(); };
}

/* ---------------- automatic daily backups (see supabase/backup-setup.sql) ---------------- */
const ukDay = (d = new Date()) => d.toLocaleDateString('en-CA', { timeZone: 'Europe/London' });   // YYYY-MM-DD
const niceDay = (day) => { const t = ukDay(), y = ukDay(new Date(Date.now() - 864e5)); return day === t ? 'Today' : day === y ? 'Yesterday' : new Date(day + 'T12:00:00').toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' }); };
const hhmm = (iso) => new Date(iso).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', timeZone: 'Europe/London' });
let abState = { list: null, missing: false, busy: '', msg: '' };
async function listBackups() {
  const { data, error } = await supabase.from('backups').select('day,created_at,items').eq('family_id', familyId).order('day', { ascending: false }).limit(31);
  if (error) { if (/backups|does not exist|schema cache|42P01|PGRST205/i.test(error.message + error.code)) { abState.missing = true; abState.list = []; return; } throw error; }
  abState.missing = false; abState.list = data || [];
}
async function backupNow() {
  const { data, error } = await supabase.rpc('backup_now');
  if (error) throw error;
  return data;
}
async function dailyBackupCheck() {
  try {
    await listBackups();
    if (abState.missing) return renderAutoBackup();
    const hasData = [...cache.keys()].some(k => !k.startsWith('meta/'));
    if (hasData && abState.list[0]?.day !== ukDay()) { await backupNow(); await listBackups(); }
  } catch (e) { console.warn('daily backup', e); }
  renderAutoBackup();
}
function renderAutoBackup() {
  const p = document.getElementById('autoBk'); if (!p) return;
  const L = abState.list;
  let body;
  if (L === null) body = '<p class="muted">Checking backups…</p>';
  else if (abState.missing) body = `<div class="band high"><strong>Automatic backups aren't switched on yet</strong><span>In Supabase, open <b>SQL Editor → New query</b>, paste the file <code>backup-setup.sql</code> from your GitHub folder and press <b>Run</b>. Then reopen this page.</span></div>`;
  else if (!L.length) body = '<p class="muted">No backups yet. The first one is made automatically once there is something to save.</p>';
  else {
    const L0 = L[0], fresh = L0.day === ukDay() || L0.day === ukDay(new Date(Date.now() - 864e5));
    body = `<div class="bk-status ${fresh ? '' : 'warn'}">${fresh ? '✓' : '!'} Last backup ${niceDay(L0.day).toLowerCase()} at ${hhmm(L0.created_at)}</div>
      <details class="bk-older"><summary>Go back to an earlier day (${L.length} kept)</summary><ul class="ab-list">${L.map(b => `<li><span class="ab-day"><b>${niceDay(b.day)}</b><small>${hhmm(b.created_at)} · ${Number(b.items).toLocaleString('en-GB')} item${b.items == 1 ? '' : 's'}</small></span><button type="button" class="small" data-abr="${b.day}">Restore</button></li>`).join('')}</ul></details>`;
  }
  p.querySelector('.ab-body').innerHTML = body;
  const st = p.querySelector('.ab-status'); st.hidden = !abState.msg; st.innerHTML = abState.msg;
  p.querySelector('#abNow').disabled = !!abState.busy || abState.missing;
  p.querySelector('#abNow').textContent = abState.busy === 'now' ? 'Backing up…' : 'Back up now';
  p.querySelectorAll('[data-abr]').forEach(b => { b.disabled = !!abState.busy; b.onclick = () => askRestore(b.dataset.abr); });
}
function askRestore(day) {
  abState.msg = `<strong>Restore the backup from ${niceDay(day)}?</strong> Anything deleted since then comes back, and items changed since then go back to how they were. Entries added after it are kept.
    <div class="btns" style="margin-top:8px"><button type="button" class="primary" id="abGo">Restore ${niceDay(day).toLowerCase()}'s backup</button><button type="button" id="abNo">Cancel</button></div>`;
  renderAutoBackup();
  document.getElementById('abNo').onclick = () => { abState.msg = ''; renderAutoBackup(); };
  document.getElementById('abGo').onclick = () => restoreBackup(day);
}
async function restoreBackup(day) {
  abState.busy = 'restore'; abState.msg = 'Getting the backup…'; renderAutoBackup();
  try {
    try { await backupNow(); } catch {}   // keep today's state too, so a restore can be undone
    const { data, error } = await supabase.from('backups').select('data').eq('family_id', familyId).eq('day', day).limit(1);
    if (error) throw error;
    const rows = (data?.[0]?.data || []).filter(r => r && r.c && r.i != null);
    const now = new Date().toISOString();
    for (let i = 0; i < rows.length; i += 400) {
      const chunk = rows.slice(i, i + 400).map(r => ({ family_id: familyId, collection: r.c, id: String(r.i), data: r.d, updated_by: me, updated_at: now }));
      const { error: e2 } = await supabase.from('records').upsert(chunk, { onConflict: 'family_id,collection,id' });
      if (e2) throw e2;
      abState.msg = `Restoring… ${Math.min(i + 400, rows.length)} of ${rows.length}`; renderAutoBackup();
    }
    await loadAll();
    abState.msg = `Restored ${rows.length.toLocaleString('en-GB')} items from ${niceDay(day).toLowerCase()}'s backup. Today's state was saved first, so you can go back to it from the list.`;
    await listBackups();
  } catch (e) { abState.msg = `Restore didn't finish: ${esc(e.message || e)}. Nothing was deleted. Try again.`; }
  abState.busy = ''; renderAutoBackup();
}
function addAutoBackupPanel() {
  const host = document.querySelector('#v-backup'); if (!host || document.getElementById('autoBk')) return;
  const p = document.createElement('div'); p.className = 'panel'; p.id = 'autoBk';
  p.innerHTML = `<div class="bk-h"><span class="bk-ic ok" aria-hidden="true"><svg class="ic" viewBox="0 0 24 24"><path d="M12 3l8 4v5c0 4.5-3.4 8.3-8 9-4.6-.7-8-4.5-8-9V7z"/><path d="M8.5 12l2.5 2.5 4.5-5"/></svg></span><div><h2>Automatic backups</h2><p class="muted">Every night, online, last 30 days kept. Nothing to do.</p></div></div>
    <div class="ab-body"></div><div class="ab-status notice" hidden></div>
    <div class="btns"><button type="button" class="small" id="abNow">Back up now</button></div>`;
  const head = host.querySelector('.view-head'); host.insertBefore(p, head ? head.nextSibling : host.firstChild);
  p.querySelector('#abNow').onclick = async () => {
    abState.busy = 'now'; abState.msg = ''; renderAutoBackup();
    try { const n = await backupNow(); await listBackups(); abState.msg = `Backed up ${Number(n).toLocaleString('en-GB')} item${n == 1 ? '' : 's'}.`; }
    catch (e) { abState.msg = `Couldn't back up: ${esc(e.message || e)}`; }
    abState.busy = ''; renderAutoBackup();
  };
  renderAutoBackup();
  dailyBackupCheck();
  document.addEventListener('visibilitychange', () => { if (!document.hidden && abState.list && !abState.missing && abState.list[0]?.day !== ukDay()) dailyBackupCheck(); });
}

if (!supabase) {
  screen(`<h1>Almost there</h1><p class="muted">This copy of the app isn't connected to its database yet. In Vercel, check that <code>VITE_SUPABASE_URL</code> and <code>VITE_SUPABASE_ANON_KEY</code> are set under Settings → Environment Variables, then redeploy.</p>`);
} else {
  supabase.auth.onAuthStateChange((event) => { if (event === 'PASSWORD_RECOVERY') { recovering = true; return showNewPassword(); } if (event === 'SIGNED_IN' && !started && !recovering) start(); if (event === 'SIGNED_OUT') { location.reload(); } });
  start();
}
