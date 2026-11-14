// Petits utilitaires : dates, HTML, iCal, chargement de scripts

export const uid = () =>
  Array.from(crypto.getRandomValues(new Uint8Array(8)), (b) => b.toString(16).padStart(2, '0')).join('');

export const pad = (n) => String(n).padStart(2, '0');
export const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
export const parseYmd = (s) => {
  const [y, m, d] = s.split('-').map(Number);
  return new Date(y, m - 1, d);
};
export const startOfDay = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
export const addDays = (d, n) => {
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
};
export const addMinutes = (d, n) => new Date(d.getTime() + n * 60000);
export const startOfWeek = (d) => {
  const x = startOfDay(d);
  return addDays(x, -((x.getDay() + 6) % 7)); // semaine commençant lundi
};
/** Arrondit à minuit le plus proche (utile pour les journées entières renvoyées en UTC). */
export const roundToDay = (d) => startOfDay(new Date(d.getTime() + 12 * 3600e3));
export const overlaps = (aS, aE, bS, bE) => +aS < +bE && +bS < +aE;
export const toLocalInput = (d) => `${ymd(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
export const isWeekend = (d) => d.getDay() === 0 || d.getDay() === 6;

const fmt = (o) => new Intl.DateTimeFormat('fr-FR', o);
const F = {
  time: fmt({ hour: '2-digit', minute: '2-digit' }),
  dayLong: fmt({ weekday: 'long', day: 'numeric', month: 'long' }),
  month: fmt({ month: 'long', year: 'numeric' }),
  short: fmt({ weekday: 'short', day: 'numeric', month: 'short' }),
  stamp: fmt({ day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }),
};
export const fmtTime = (d) => F.time.format(d);
export const fmtDayLong = (d) => F.dayLong.format(d);
export const fmtMonth = (d) => F.month.format(d);
export const fmtShort = (d) => F.short.format(d);
export const fmtStamp = (d) => F.stamp.format(d);
export const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);

export function fmtRange(s, e, allDay) {
  s = new Date(s);
  e = new Date(e);
  if (allDay) {
    const last = addDays(e, -1);
    return ymd(s) === ymd(last) ? `${cap(fmtDayLong(s))} · journée entière` : `Du ${fmtShort(s)} au ${fmtShort(last)}`;
  }
  if (ymd(s) === ymd(e)) return `${cap(fmtDayLong(s))} · ${fmtTime(s)}–${fmtTime(e)}`;
  return `${cap(fmtShort(s))} ${fmtTime(s)} → ${fmtShort(e)} ${fmtTime(e)}`;
}

export const esc = (s) =>
  String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

/** Fusionne des créneaux qui se chevauchent : [{start,end}] (ISO) → triés et fusionnés. */
export function mergeSlots(slots) {
  const s = slots
    .map((x) => ({ start: +new Date(x.start), end: +new Date(x.end) }))
    .filter((x) => x.end > x.start)
    .sort((a, b) => a.start - b.start);
  const out = [];
  for (const x of s) {
    const last = out[out.length - 1];
    if (last && x.start <= last.end) last.end = Math.max(last.end, x.end);
    else out.push({ ...x });
  }
  return out.map((x) => ({ start: new Date(x.start).toISOString(), end: new Date(x.end).toISOString() }));
}

// ── iCalendar (export) ──────────────────────────────────────
const icsEsc = (s) =>
  String(s ?? '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
const icsUtc = (d) => d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
const icsDate = (d) => ymd(d).replace(/-/g, '');
const fold = (line) => {
  const parts = [];
  while (line.length > 73) {
    parts.push(line.slice(0, 73));
    line = ' ' + line.slice(73);
  }
  parts.push(line);
  return parts.join('\r\n');
};

export function buildIcs(events, calName = 'Agenda Famille') {
  const L = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Agenda Famille//FR', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH', 'X-WR-CALNAME:' + icsEsc(calName)];
  const now = icsUtc(new Date());
  for (const ev of events) {
    const s = new Date(ev.start), e = new Date(ev.end);
    L.push('BEGIN:VEVENT', `UID:${ev.id}@agenda-famille`, 'DTSTAMP:' + now);
    if (ev.allDay) L.push('DTSTART;VALUE=DATE:' + icsDate(s), 'DTEND;VALUE=DATE:' + icsDate(e));
    else L.push('DTSTART:' + icsUtc(s), 'DTEND:' + icsUtc(e));
    L.push('SUMMARY:' + icsEsc(ev.title));
    if (ev.location) L.push('LOCATION:' + icsEsc(ev.location));
    if (ev.notes) L.push('DESCRIPTION:' + icsEsc(ev.notes));
    L.push('END:VEVENT');
  }
  L.push('END:VCALENDAR');
  return L.map(fold).join('\r\n') + '\r\n';
}

export function download(name, text, type = 'text/calendar;charset=utf-8') {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const a = Object.assign(document.createElement('a'), { href: url, download: name });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

const scripts = new Map();
export function loadScript(src) {
  if (!scripts.has(src)) {
    scripts.set(
      src,
      new Promise((resolve, reject) => {
        const s = Object.assign(document.createElement('script'), { src, async: true });
        s.onload = resolve;
        s.onerror = () => {
          scripts.delete(src);
          reject(new Error('Impossible de charger ' + src));
        };
        document.head.append(s);
      }),
    );
  }
  return scripts.get(src);
}
