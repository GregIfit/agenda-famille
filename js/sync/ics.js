// Lecture d'un agenda via un lien iCal (.ics / webcal://), récurrences comprises (ical.js)
import { loadScript } from '../utils.js';

const ICAL_SRC = 'https://cdn.jsdelivr.net/npm/ical.js@1.5.0/build/ical.min.js';

async function fetchText(url) {
  const r = await fetch(url, { cache: 'no-store' });
  if (!r.ok) throw new Error('HTTP ' + r.status);
  return r.text();
}

export async function fetchEvents(rawUrl, proxy, from, to) {
  await loadScript(ICAL_SRC);
  const url = rawUrl.trim().replace(/^webcal:\/\//i, 'https://');
  let text;
  try {
    text = await fetchText(url);
  } catch (e) {
    if (!proxy) throw new Error('lecture bloquée par le serveur (CORS). Configurez « icsProxy » dans js/config.js (voir README).');
    text = await fetchText(proxy + encodeURIComponent(url));
  }
  if (!/BEGIN:VCALENDAR/i.test(text)) throw new Error('ce lien ne renvoie pas un agenda iCal.');
  return parse(text, from, to);
}

export function parse(text, from, to) {
  const root = new ICAL.Component(ICAL.parse(text));
  for (const vtz of root.getAllSubcomponents('vtimezone')) {
    try {
      ICAL.TimezoneService.register(vtz);
    } catch {}
  }
  const masters = new Map();
  const exceptions = [];
  for (const v of root.getAllSubcomponents('vevent')) {
    const ev = new ICAL.Event(v);
    if (String(ev.uid || '').endsWith('@agenda-famille')) continue; // déjà dans l'agenda familial
    if (v.hasProperty('recurrence-id')) exceptions.push(ev);
    else masters.set(ev.uid + '|' + masters.size, ev);
  }
  for (const ex of exceptions) {
    const m = [...masters.values()].find((x) => x.uid === ex.uid);
    if (m) m.relateException(ex);
    else masters.set(ex.uid + '|ex' + masters.size, ex);
  }

  const out = [];
  const push = (ev, s, e) => {
    if (!(s < to && e > from)) return;
    const status = String(ev.component.getFirstPropertyValue('status') || '').toUpperCase();
    if (status === 'CANCELLED') return;
    const transp = String(ev.component.getFirstPropertyValue('transp') || '').toUpperCase();
    out.push({
      start: s.toISOString(),
      end: e.toISOString(),
      allDay: ev.startDate.isDate,
      title: ev.summary || '(sans titre)',
      busy: transp !== 'TRANSPARENT',
      source: 'iCal',
    });
  };

  for (const ev of masters.values()) {
    if (!ev.startDate) continue;
    const endOf = (start, end) => (end ? end.toJSDate() : new Date(start.toJSDate().getTime() + (start.isDate ? 864e5 : 0)));
    if (ev.isRecurring()) {
      const it = ev.iterator();
      let next, guard = 0;
      while ((next = it.next()) && guard++ < 3000) {
        const det = ev.getOccurrenceDetails(next);
        const s = det.startDate.toJSDate();
        if (s >= to) break;
        push(det.item, s, endOf(det.startDate, det.endDate));
      }
    } else {
      push(ev, ev.startDate.toJSDate(), endOf(ev.startDate, ev.endDate));
    }
  }
  return out;
}
