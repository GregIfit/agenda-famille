// Synchronisation Outlook / Microsoft 365 (MSAL.js + Microsoft Graph)
import { loadScript, roundToDay, ymd } from '../utils.js';

const MSAL_SRC = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js';
const SCOPES = ['Calendars.ReadWrite'];
const GRAPH = 'https://graph.microsoft.com/v1.0';
export const CATEGORY = 'Agenda Famille';
let app = null;
let account = null;

async function getApp(cfg) {
  if (!cfg?.clientId) throw new Error('ID d’application Microsoft manquant dans js/config.js (voir README).');
  if (app) return app;
  await loadScript(MSAL_SRC);
  app = new msal.PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: cfg.authority || 'https://login.microsoftonline.com/common',
      redirectUri: location.origin + location.pathname,
    },
    cache: { cacheLocation: 'localStorage' },
  });
  if (app.initialize) await app.initialize();
  await app.handleRedirectPromise().catch(() => {});
  account = app.getAllAccounts()[0] || null;
  return app;
}

export async function connect(cfg) {
  const a = await getApp(cfg);
  const r = await a.loginPopup({ scopes: SCOPES, prompt: 'select_account' });
  account = r.account;
  a.setActiveAccount?.(account);
  return account.username;
}

export function disconnect() {
  account = null;
}

async function getToken(cfg, interactive) {
  const a = await getApp(cfg);
  if (!account) {
    if (!interactive) throw new Error('compte non connecté.');
    await connect(cfg);
  }
  try {
    return (await a.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
  } catch {
    if (!interactive) throw new Error('reconnexion nécessaire, cliquez sur « Synchroniser ».');
    return (await a.acquireTokenPopup({ scopes: SCOPES, account })).accessToken;
  }
}

async function graph(cfg, interactive, url, opts = {}) {
  const t = await getToken(cfg, interactive);
  const r = await fetch(url.startsWith('http') ? url : GRAPH + url, {
    ...opts,
    headers: { Authorization: 'Bearer ' + t, 'Content-Type': 'application/json', ...(opts.headers || {}) },
  });
  if (!r.ok) {
    const err = new Error(`Outlook ${r.status} : ${(await r.text()).slice(0, 160)}`);
    err.status = r.status;
    throw err;
  }
  return r.status === 204 ? null : r.json();
}

const utc = (s) => new Date(s.split('.')[0] + 'Z');

export async function fetchEvents(cfg, from, to, interactive = false) {
  const q = new URLSearchParams({
    startDateTime: from.toISOString(),
    endDateTime: to.toISOString(),
    $top: '200',
    $select: 'subject,start,end,isAllDay,showAs,categories,isCancelled',
  });
  let url = '/me/calendarView?' + q;
  const out = [];
  while (url) {
    const d = await graph(cfg, interactive, url, { headers: { Prefer: 'outlook.timezone="UTC"' } });
    for (const it of d.value || []) {
      if (it.isCancelled || (it.categories || []).includes(CATEGORY)) continue;
      let s = utc(it.start.dateTime), e = utc(it.end.dateTime);
      if (it.isAllDay) (s = roundToDay(s)), (e = roundToDay(e));
      out.push({
        start: s.toISOString(),
        end: e.toISOString(),
        allDay: !!it.isAllDay,
        title: it.subject || '(sans titre)',
        busy: it.showAs !== 'free',
        source: 'Outlook',
      });
    }
    url = d['@odata.nextLink'] || '';
  }
  return out;
}

export async function upsertEvent(cfg, ev, existingId) {
  const s = new Date(ev.start), e = new Date(ev.end);
  const tz = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  const body = {
    subject: ev.title,
    body: { contentType: 'text', content: ev.notes || '' },
    location: { displayName: ev.location || '' },
    categories: [CATEGORY],
    isAllDay: !!ev.allDay,
    start: ev.allDay ? { dateTime: ymd(s) + 'T00:00:00', timeZone: tz } : { dateTime: s.toISOString().slice(0, 19), timeZone: 'UTC' },
    end: ev.allDay ? { dateTime: ymd(e) + 'T00:00:00', timeZone: tz } : { dateTime: e.toISOString().slice(0, 19), timeZone: 'UTC' },
  };
  if (existingId) {
    try {
      return (await graph(cfg, true, '/me/events/' + encodeURIComponent(existingId), { method: 'PATCH', body: JSON.stringify(body) })).id;
    } catch (err) {
      if (err.status !== 404) throw err;
    }
  }
  return (await graph(cfg, true, '/me/events', { method: 'POST', body: JSON.stringify(body) })).id;
}
