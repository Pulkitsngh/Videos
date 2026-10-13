/*
 * Navjyoti Tracker – LNB Group Microsoft 365 connector.
 *
 * The tracker page talks to its data through `window.claude.use('db' | 'user' | 'downloads')`.
 * This file provides those same calls on top of the company's own Microsoft 365:
 *   - people sign in with their work account (Microsoft Entra ID, via MSAL),
 *   - records are read from and written to the SharePoint lists on the Navjyoti site (Microsoft Graph),
 *   - every PC polls for changes, so all users see the same data.
 * Nothing here calls Claude or any service outside the company's Microsoft 365.
 *
 * Settings come from config.js (window.NAVJYOTI_CONFIG); the list and column names from fieldmap.js
 * (window.NAVJYOTI_FIELDS), which build_desktop_app.py generates from the Excel tracker's registers.
 */
(() => {
  const CFG = Object.assign({ pollSeconds: 30, graphBase: 'https://graph.microsoft.com/v1.0' }, window.NAVJYOTI_CONFIG || {});
  const MAP = window.NAVJYOTI_FIELDS || {};
  const VERSION = window.NAVJYOTI_VERSION || 'dev';
  const SCOPES = CFG.scopes || ['https://graph.microsoft.com/Sites.ReadWrite.All', 'User.Read'];

  /* ---------- small UI pieces (status banner, sidebar footer) ---------- */
  const ready = new Promise(r => document.readyState === 'loading' ? document.addEventListener('DOMContentLoaded', r) : r());
  function banner(html, kind = 'warn', actions = '', tag = '') {
    ready.then(() => {
      let b = document.getElementById('cloudMsg');
      if (!b) return;
      b.className = 'cloud-msg ' + kind; b.dataset.kind = tag;
      b.innerHTML = `<div>${html}</div>${actions}`;
      b.hidden = !html;
    });
  }
  const escH = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ---------- configuration check ---------- */
  const placeholder = v => !v || /PASTE|YOUR-|<|>/.test(v);
  if (placeholder(CFG.tenantId) || placeholder(CFG.clientId) || placeholder(CFG.siteUrl)) {
    banner('<b>Setup not finished.</b> Open <code>config.js</code> and fill in the tenant ID, the app (client) ID and the SharePoint site address, as described in the IT setup guide.', 'bad');
    window.claude = { use: async () => null };
    return;
  }

  /* ---------- sign-in (MSAL) ---------- */
  const msalApp = new msal.PublicClientApplication({
    auth: { clientId: CFG.clientId, authority: `https://login.microsoftonline.com/${CFG.tenantId}`, redirectUri: CFG.redirectUri || (location.origin + location.pathname.replace(/index\.html$/, '')) },
    cache: { cacheLocation: 'localStorage' },
  });
  let account = null;
  const signedIn = (async () => {
    await msalApp.initialize();
    const res = await msalApp.handleRedirectPromise();
    account = res?.account || msalApp.getActiveAccount() || msalApp.getAllAccounts()[0] || null;
    if (!account) { await msalApp.loginRedirect({ scopes: SCOPES, prompt: 'select_account' }); return new Promise(() => {}); }
    msalApp.setActiveAccount(account);
    ready.then(() => {
      const f = document.getElementById('sideFoot');
      if (f) f.innerHTML = `<span>Signed in as <b>${escH(account.name || account.username)}</b></span><button type="button" id="signOut">Sign out</button><span class="ver">Version ${escH(VERSION)}</span>`;
      document.getElementById('signOut')?.addEventListener('click', () => msalApp.logoutRedirect({ account }));
    });
    return account;
  })();

  async function token() {
    await signedIn;
    try { return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || e?.errorCode === 'interaction_required' || e?.errorCode === 'consent_required') {
        await msalApp.acquireTokenRedirect({ scopes: SCOPES, account }); return new Promise(() => {});
      }
      throw e;
    }
  }

  /* ---------- Microsoft Graph ---------- */
  async function graph(path, opts = {}, attempt = 0) {
    const url = path.startsWith('http') ? path : CFG.graphBase + path;
    let res;
    try {
      res = await fetch(url, { ...opts, headers: { Authorization: 'Bearer ' + await token(), 'Content-Type': 'application/json', ...(opts.headers || {}) } });
    } catch (e) { throw { code: 'network', message: 'cannot reach Microsoft 365 (check the internet connection)' }; }
    if ((res.status === 429 || res.status === 503) && attempt < 4) {
      await new Promise(r => setTimeout(r, (Number(res.headers.get('Retry-After')) || 2 ** attempt) * 1000));
      return graph(path, opts, attempt + 1);
    }
    if (!res.ok) {
      let detail = ''; try { detail = (await res.json())?.error?.message || ''; } catch {}
      // a 403 means SharePoint gave this person read access only (or none): the page then shows its view-only message
      throw { code: res.status === 403 ? 'invalid_argument' : res.status === 404 ? 'not_found' : 'graph_' + res.status, status: res.status, message: detail || res.statusText };
    }
    return res.status === 204 ? null : res.json();
  }

  let siteId = null, listIds = null;
  async function site() {
    if (listIds) return;
    const u = new URL(CFG.siteUrl);
    siteId = (await graph(`/sites/${u.hostname}:${u.pathname.replace(/\/$/, '')}`)).id;
    const lists = await graph(`/sites/${siteId}/lists?$select=id,displayName,name&$top=200`);
    listIds = Object.fromEntries(lists.value.map(l => [l.displayName, l.id]));
    const missing = Object.values(MAP).map(m => m.list).filter(n => !listIds[n]);
    if (missing.length) banner(`<b>Missing SharePoint lists:</b> ${missing.map(escH).join(', ')}. Run <code>Create-NavjyotiLists.ps1</code> on the site, then reload.`, 'bad');
  }

  /* ---------- field mapping between the page and SharePoint ---------- */
  const pad = n => String(n).padStart(2, '0');
  const localISO = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const isISO = s => /^\d{4}-\d{2}-\d{2}$/.test(s || '');
  // SharePoint keeps date-only columns as local midnight in UTC; reading them back in local time gives the right day
  const spToDate = v => v ? localISO(new Date(v)) : '';
  const dateToSp = v => isISO(v) ? new Date(v + 'T00:00:00').toISOString() : null;
  const plain = v => typeof v === 'string' && /<\/?(div|p|br)/i.test(v) ? v.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim() : v;

  function toPage(c, item) {
    const m = MAP[c], f = item.fields || {}, r = { _id: String(item.id) };
    for (const [key, col, kind] of m.fields) {
      const v = f[col];
      if (kind === 'date') r[key] = spToDate(v);
      else if (kind === 'num') r[key] = v == null || v === '' ? null : Number(v);
      else r[key] = v == null ? '' : plain(String(v));
    }
    if (m.turns) {
      r.turns = [];
      for (let i = 1; i <= m.turns; i++) {
        const t = { date: spToDate(f[`Turn${i}Date`]), temp: f[`Turn${i}TempC`] ?? null, moisture: f[`Turn${i}MoisturePct`] ?? null, watering: spToDate(f[`Turn${i}WateringDate`]) };
        if (t.date || t.temp != null || t.moisture != null || t.watering) r.turns.push(t);
      }
    }
    return r;
  }
  function toSp(c, body, partial) {
    const m = MAP[c], out = {};
    for (const [key, col, kind] of m.fields) {
      if (!(key in body)) continue;
      const v = body[key];
      if (kind === 'date') out[col] = dateToSp(v);
      else if (kind === 'num') out[col] = v == null || v === '' || !isFinite(v) ? null : Number(v);
      else out[col] = v == null || v === '' ? null : String(v);
    }
    if (m.turns && 'turns' in body) {
      const ts = body.turns || [];
      if (ts.length > m.turns) banner(`Only the first ${m.turns} turnings are kept in SharePoint; ${ts.length - m.turns} more were not saved.`, 'warn');
      for (let i = 1; i <= m.turns; i++) {
        const t = ts[i - 1] || {};
        out[`Turn${i}Date`] = dateToSp(t.date); out[`Turn${i}WateringDate`] = dateToSp(t.watering);
        out[`Turn${i}TempC`] = t.temp == null || t.temp === '' ? null : Number(t.temp);
        out[`Turn${i}MoisturePct`] = t.moisture == null || t.moisture === '' ? null : Number(t.moisture);
      }
    }
    if (!partial) for (const k of Object.keys(out)) if (out[k] === null) delete out[k];  // new items: leave empty columns out
    return out;
  }

  /* ---------- collections: the same calls the page already uses ---------- */
  const cache = {}, sig = {}, listeners = {};
  async function load(c) {
    await site();
    const lid = listIds[MAP[c].list];
    if (!lid) { cache[c] = []; return; }
    let url = `/sites/${siteId}/lists/${lid}/items?$expand=fields&$top=500`, items = [];
    while (url) { const page = await graph(url); items = items.concat(page.value); url = page['@odata.nextLink']; }
    const s = items.map(i => i.id + ':' + i.lastModifiedDateTime).sort().join('|');
    const changed = s !== sig[c];
    sig[c] = s;
    cache[c] = items.map(i => toPage(c, i));
    return changed;
  }
  const snapshot = c => ({ docs: cache[c].map(r => ({ id: r._id, data: () => { const o = { ...r }; delete o._id; return o; } })) });
  async function refresh(c, force) {
    try {
      const changed = await load(c);
      if (changed || force) (listeners[c] || []).forEach(l => l.cb(snapshot(c)));
      if (document.getElementById('cloudMsg')?.dataset.kind === 'net') banner('');
    } catch (e) {
      (listeners[c] || []).forEach(l => l.err && l.err(e));
      if (e.code === 'network') banner('<b>Offline.</b> Cannot reach LNB Group Microsoft 365. Records will refresh when the connection is back.', 'warn', '', 'net');
      else if (e.code === 'not_found') banner('<b>SharePoint site not found.</b> Check <code>siteUrl</code> in config.js.', 'bad');
      else banner(`<b>Could not load ${escH(MAP[c]?.list || c)}:</b> ${escH(e.message || e)}`, 'bad');
    }
  }
  const refreshAll = () => Object.keys(listeners).forEach(c => refresh(c));
  setInterval(() => { if (document.visibilityState === 'visible') refreshAll(); }, Math.max(10, CFG.pollSeconds) * 1000);
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') refreshAll(); });

  function collection(c) {
    if (!MAP[c]) throw new Error('Unknown register ' + c);
    return {
      onSnapshot(cb, err) {
        (listeners[c] = listeners[c] || []).push({ cb, err });
        if (cache[c]) cb(snapshot(c)); else refresh(c, true);
        return () => { listeners[c] = listeners[c].filter(l => l.cb !== cb); };
      },
      doc(id) {
        const known = () => id != null && (cache[c] || []).some(r => r._id === String(id));
        const items = async () => { await site(); return `/sites/${siteId}/lists/${listIds[MAP[c].list]}/items`; };
        return {
          async set(body) {
            if (known()) await graph(`${await items()}/${id}/fields`, { method: 'PATCH', body: JSON.stringify(toSp(c, body, true)) });
            else await graph(await items(), { method: 'POST', body: JSON.stringify({ fields: toSp(c, body, false) }) });
            await refresh(c, true);
          },
          async update(part) {
            await graph(`${await items()}/${id}/fields`, { method: 'PATCH', body: JSON.stringify(toSp(c, part, true)) });
            await refresh(c, true);
          },
          async delete() {
            await graph(`${await items()}/${id}`, { method: 'DELETE' });
            await refresh(c, true);
          },
        };
      },
    };
  }

  /* ---------- downloads (Export Excel, report downloads) ---------- */
  const downloads = {
    async save({ filename, data }) {
      const blob = data instanceof Blob ? data : new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    },
  };
  // SharePoint decides who may edit; a refused save turns the page view-only
  const user = { can: async () => null, isOwner: () => false };

  window.claude = {
    async use(name) {
      if (name === 'db') {
        try { await signedIn; await site(); return { collection }; }
        catch (e) { banner(`<b>Could not connect to LNB Group Microsoft 365:</b> ${escH(e.message || e.errorMessage || e)}`, 'bad'); return null; }
      }
      if (name === 'downloads') return downloads;
      if (name === 'user') return user;
      return null;
    },
  };

  /* ---------- new version check ---------- */
  async function checkVersion() {
    try {
      const v = (await (await fetch('version.json?ts=' + Date.now(), { cache: 'no-store' })).json()).version;
      if (v && v !== VERSION) banner(`<b>Version ${escH(v)} is available.</b> Finish what you are typing, then update. Your records are safe in SharePoint.`, 'info',
        '<button type="button" onclick="location.reload()">Update now</button>');
    } catch {}
  }
  setTimeout(checkVersion, 5000);
  setInterval(checkVersion, 10 * 60 * 1000);
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') checkVersion(); });
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
