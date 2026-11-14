// Synchronisation Google Agenda (OAuth « token client » de Google Identity Services)
import { loadScript, parseYmd, ymd } from '../utils.js';

export const GSI_SRC = 'https://accounts.google.com/gsi/client';
const SCOPE = 'https://www.googleapis.com/auth/calendar.events';
const API = 'https://www.googleapis.com/calendar/v3';
let token = null;
let expiresAt = 0;

export const isReady = () => !!token && Date.now() < expiresAt - 60_000;

export async function connect(clientId) {
  if (!clientId) throw new Error('ID client Google manquant dans js/config.js (voir README).');
  await loadScript(GSI_SRC);
  return new Promise((resolve, reject) => {
    const client = google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: SCOPE,
      callback: (r) => {
        if (r.error) return reject(new Error(r.error_description || r.error));
        token = r.access_token;
        expiresAt = Date.now() + (Number(r.expires_in) || 3600) * 1000;
        resolve();
      },
      error_callback: (e) => reject(new Error(e?.type === 'popup_closed' ? 'Connexion Google annulée' : e?.message || 'Connexion Google impossible')),
    });
    client.requestAccessToken();
  });
}

export function disconnect() {
  if (token && window.google?.accounts?.oauth2) google.accounts.oauth2.revoke(token, () => {});
  token = null;
  expiresAt = 0;
}

async function api(path, opts = {}) {
  if (!isReady()) throw new Error('session Google expirée, cliquez sur « Synchroniser ».');
  const r = await fetch(API + path, {
    ...opts,
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json', ...(opts.headers || {}) },
  });
  if (!r.ok) {
    const err = new Error(`Google ${r.status} : ${(await r.text()).slice(0, 160)}`);
    err.status = r.status;
    throw err;
  }
  return r.status === 204 ? null : r.json();
}

export async function fetchEvents(from, to) {
  const out = [];
  let pageToken = '';
  do {
    const q = new URLSearchParams({
      timeMin: from.toISOString(),
      timeMax: to.toISOString(),
      singleEvents: 'true',
      orderBy: 'startTime',
      maxResults: '2500',
    });
    if (pageToken) q.set('pageToken', pageToken);
    const d = await api('/calendars/primary/events?' + q);
    for (const it of d.items || []) {
      if (it.status === 'cancelled') continue;
      if (it.extendedProperties?.private?.agendaFamille) continue; // déjà dans l'agenda familial
      const allDay = !!it.start?.date;
      const s = allDay ? parseYmd(it.start.date) : new Date(it.start.dateTime);
      const e = allDay ? parseYmd(it.end.date) : new Date(it.end.dateTime);
      out.push({
        start: s.toISOString(),
        end: e.toISOString(),
        allDay,
        title: it.summary || '(sans titre)',
        busy: it.transparency !== 'transparent',
        source: 'Google',
      });
    }
    pageToken = d.nextPageToken;
  } while (pageToken);
  return out;
}

/** Crée ou met à jour l'événement familial dans l'agenda Google principal. Renvoie l'id Google. */
export async function upsertEvent(ev, existingId) {
  const s = new Date(ev.start), e = new Date(ev.end);
  const body = {
    summary: ev.title,
    location: ev.location || undefined,
    description: ev.notes || undefined,
    start: ev.allDay ? { date: ymd(s) } : { dateTime: s.toISOString() },
    end: ev.allDay ? { date: ymd(e) } : { dateTime: e.toISOString() },
    extendedProperties: { private: { agendaFamille: ev.id } },
  };
  if (existingId) {
    try {
      const r = await api('/calendars/primary/events/' + encodeURIComponent(existingId), { method: 'PUT', body: JSON.stringify(body) });
      if (r?.status !== 'cancelled') return r.id;
    } catch (err) {
      if (err.status !== 404 && err.status !== 410) throw err;
    }
  }
  const r = await api('/calendars/primary/events', { method: 'POST', body: JSON.stringify(body) });
  return r.id;
}
