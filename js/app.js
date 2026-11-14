import { CONFIG } from './config.js';
import * as U from './utils.js';
import { createLocalStore, createCloudStore } from './store.js';
import * as C from './cloud.js';
import * as G from './sync/google.js';
import * as O from './sync/outlook.js';
import * as ICS from './sync/ics.js';

const { esc } = U;
const COLORS = ['#e4572e', '#2e86ab', '#3bb273', '#9b5de5', '#f4a259', '#ef476f', '#0f8b8d', '#7a6c5d'];
const LS = { me: 'af:me', family: 'af:family', invite: 'af:invite', sources: 'af:sources', personal: 'af:personal', pushed: 'af:pushed', showBusy: 'af:showBusy' };
const ls = {
  get(k, d) {
    try {
      const v = localStorage.getItem(k);
      return v == null ? d : JSON.parse(v);
    } catch {
      return d;
    }
  },
  set(k, v) {
    try {
      localStorage.setItem(k, JSON.stringify(v));
    } catch {}
  },
};

const cloud = !!CONFIG.firebase?.apiKey;
const today = () => U.startOfDay(new Date());
const state = {
  view: 'agenda',
  pollId: null,
  cursor: today(),
  selected: today(),
  me: ls.get(LS.me, null),
  showBusy: ls.get(LS.showBusy, true),
  syncing: false,
};
let store;
let ctx = null; // session Firebase (mode cloud)
let familyId = null;
const app = document.getElementById('app');
const dlg = document.getElementById('modal');

// ── Helpers données ─────────────────────────────────────────
const members = () => store.list('members').sort((a, b) => (a.order ?? 0) - (b.order ?? 0) || a.name.localeCompare(b.name));
const member = (id) => store.get('members', id);
const fail = (e) => toast(e?.message || String(e), 'error');
const save = (k, o) => Promise.resolve(store.save(k, o)).catch(fail);
const remove = (k, id) => Promise.resolve(store.remove(k, id)).catch(fail);
const patch = (k, id, p, v) => Promise.resolve(store.patch(k, id, p, v)).catch(fail);
const update = (k, id, f) => Promise.resolve(store.update(k, id, f)).catch(fail);
const myAccess = () => (cloud && ctx?.user ? store.get('access', ctx.user.uid) : null);
/** En mode local, tout le monde gère tout ; en mode cloud, seulement les administrateurs. */
const isAdmin = () => !cloud || myAccess()?.role === 'admin';
const accessOfMember = (memberId) => store.list('access').find((a) => a.memberId === memberId);

function allSources() {
  return ls.get(LS.sources, {});
}
function sources(id = state.me) {
  return { google: false, outlook: false, outlookUser: '', ics: [], share: true, lastSync: null, errors: [], ...(allSources()[id] || {}) };
}
function setSources(s, id = state.me) {
  const all = allSources();
  all[id] = s;
  ls.set(LS.sources, all);
}
const hasSources = (s = sources()) => s.google || s.outlook || s.ics.length > 0;

let personalCache = null;
function personal(id = state.me) {
  if (!personalCache) personalCache = ls.get(LS.personal, {});
  return personalCache[id] || null;
}
function setPersonal(p, id = state.me) {
  personalCache = ls.get(LS.personal, {});
  if (p) personalCache[id] = p;
  else delete personalCache[id];
  ls.set(LS.personal, personalCache);
}
const personalItems = (id = state.me) => {
  const p = personal(id);
  return p ? Object.values(p.bySource || {}).flat() : null;
};

/** Créneaux occupés d'un membre (avec titres si c'est moi sur cet appareil). */
function busySlots(id) {
  if (id === state.me) {
    const p = personalItems(id);
    if (p) return p.filter((i) => i.busy);
  }
  return store.get('busy', id)?.slots || [];
}
const hasCalendar = (id) => (id === state.me && personal(id)) || !!store.get('busy', id);

/** Raisons pour lesquelles un membre n'est pas libre sur [s, e[. */
function conflictsOf(id, s, e, excludeEventId) {
  const out = [];
  for (const b of busySlots(id)) if (U.overlaps(s, e, new Date(b.start), new Date(b.end))) out.push(b.title || 'occupé (agenda perso)');
  for (const ev of store.list('events')) {
    if (ev.id === excludeEventId || !(ev.members || []).includes(id)) continue;
    if (U.overlaps(s, e, new Date(ev.start), new Date(ev.end))) out.push(ev.title + ' (famille)');
  }
  return out;
}

function avatar(m, extra = '') {
  if (!m) return '';
  return `<span class="avatar ${extra}" style="--c:${m.color}" title="${esc(m.name)}">${esc(m.name.trim().charAt(0).toUpperCase())}</span>`;
}

// ── Toasts & modale ─────────────────────────────────────────
function toast(msg, kind = 'info') {
  const el = Object.assign(document.createElement('div'), { className: 'toast ' + kind, textContent: msg });
  document.getElementById('toasts').append(el);
  setTimeout(() => el.classList.add('out'), kind === 'error' ? 6000 : 3000);
  setTimeout(() => el.remove(), kind === 'error' ? 6500 : 3500);
}
function openModal(html, setup) {
  dlg.innerHTML = `<div class="modal-card">${html}</div>`;
  if (!dlg.open) dlg.showModal();
  setup?.(dlg);
  dlg.querySelector('input:not([type=checkbox]):not([type=radio]), textarea')?.focus();
}
function closeModal() {
  dlg.close();
  dlg.innerHTML = '';
}
dlg.addEventListener('click', (e) => {
  if (e.target === dlg || e.target.closest('[data-close]')) closeModal();
});

// ── Routage ─────────────────────────────────────────────────
function parseHash() {
  const [path, query] = location.hash.replace(/^#\/?/, '').split('?');
  return { parts: path.split('/').filter(Boolean), params: new URLSearchParams(query || '') };
}
function route() {
  const { parts } = parseHash();
  state.view = ['agenda', 'sondages', 'famille'].includes(parts[0]) ? parts[0] : 'agenda';
  state.pollId = state.view === 'sondages' ? parts[1] || null : null;
  render();
  window.scrollTo(0, 0);
}
const go = (h) => (location.hash = h);

// ── Rendu ───────────────────────────────────────────────────
function render() {
  if (!store) return;
  if (cloud) state.me = myAccess()?.memberId || null; // en mode cloud, l'identité vient du compte connecté
  if (state.me && !member(state.me)) {
    if (cloud) state.me = null;
    else setMe(null, false);
  }
  app.innerHTML = `${header()}<main class="main">${state.view === 'sondages' ? viewPolls() : state.view === 'famille' ? viewFamily() : viewAgenda()}</main>`;
}

function header() {
  const ms = members();
  const tab = (v, label, icon) => `<a href="#/${v}" class="tab ${state.view === v ? 'active' : ''}"><span aria-hidden="true">${icon}</span>${label}</a>`;
  const openPolls = store.list('polls').filter((p) => p.status === 'open' && state.me && p.participants.includes(state.me) && !p.options.every((o) => p.votes?.[state.me]?.[o.id])).length;
  return `<header class="top">
    <div class="brand"><span class="logo" aria-hidden="true">◆</span>${esc(CONFIG.appName)}</div>
    <nav class="tabs">${tab('agenda', 'Agenda', '▦')}${tab('sondages', 'Sondages' + (openPolls ? ` <b class="badge">${openPolls}</b>` : ''), '✓')}${tab('famille', 'Famille', '☺')}</nav>
    <div class="who">
      ${state.me && hasSources() ? `<button class="btn ghost small" data-act="sync" ${state.syncing ? 'disabled' : ''} title="Synchroniser mes agendas">${state.syncing ? '<span class="spin"></span>' : '⟳'}<span class="hide-sm"> Synchro</span></button>` : ''}
      ${cloud
        ? `<button class="btn ghost small account" data-act="account" title="${esc(ctx.user.email || '')}">${state.me ? avatar(member(state.me)) : '<span class="avatar">?</span>'}<span class="hide-sm">${esc(member(state.me)?.name || 'Mon compte')}</span></button>`
        : ms.length ? `<label class="me-select">${state.me ? avatar(member(state.me)) : ''}<select data-change="me" aria-label="Qui êtes-vous ?">
        <option value="">Qui êtes-vous ?</option>${ms.map((m) => `<option value="${m.id}" ${m.id === state.me ? 'selected' : ''}>${esc(m.name)}</option>`).join('')}
      </select></label>` : ''}
    </div>
  </header>`;
}

// ── Vue Agenda ──────────────────────────────────────────────
function dayMap(from, to) {
  const map = new Map();
  const add = (d, type, item) => {
    const k = U.ymd(d);
    if (!map.has(k)) map.set(k, { events: [], busy: [] });
    map.get(k)[type].push(item);
  };
  const spread = (s, e, allDay, cb) => {
    if (!(s < to && e > from)) return;
    let d = U.startOfDay(s < from ? from : s);
    const last = allDay ? U.addDays(e, -1) : new Date(e.getTime() - 1);
    while (d <= last && d < to) cb(d), (d = U.addDays(d, 1));
  };
  for (const ev of store.list('events')) spread(new Date(ev.start), new Date(ev.end), ev.allDay, (d) => add(d, 'events', ev));
  if (state.showBusy)
    for (const m of members())
      for (const b of busySlots(m.id)) spread(new Date(b.start), new Date(b.end), b.allDay, (d) => add(d, 'busy', { ...b, member: m }));
  for (const v of map.values()) v.events.sort((a, b) => (b.allDay ? 1 : 0) - (a.allDay ? 1 : 0) || a.start.localeCompare(b.start));
  return map;
}

function viewAgenda() {
  const ms = members();
  if (!ms.length)
    return `<section class="empty card"><h1>Bienvenue 👋</h1><p>Commencez par ajouter les membres de la famille : chacun pourra ensuite relier son agenda personnel et répondre aux sondages.</p><button class="btn primary" data-act="add-member">Ajouter un membre</button></section>`;
  const first = new Date(state.cursor.getFullYear(), state.cursor.getMonth(), 1);
  const start = U.startOfWeek(first);
  const end = U.addDays(start, 42);
  const map = dayMap(start, end);
  const t = U.ymd(today()), sel = U.ymd(state.selected);
  const cells = [];
  for (let i = 0; i < 42; i++) {
    const d = U.addDays(start, i), k = U.ymd(d), info = map.get(k) || { events: [], busy: [] };
    const busyMembers = [...new Map(info.busy.map((b) => [b.member.id, b.member])).values()];
    cells.push(`<div class="day ${d.getMonth() !== first.getMonth() ? 'out' : ''} ${k === t ? 'today' : ''} ${k === sel ? 'sel' : ''}" data-act="select-day" data-day="${k}" role="button" tabindex="0" aria-label="${esc(U.fmtDayLong(d))}">
      <span class="num">${d.getDate()}</span>
      <div class="chips">${info.events.slice(0, 3).map((ev) => chip(ev)).join('')}${info.events.length > 3 ? `<span class="more">+${info.events.length - 3}</span>` : ''}</div>
      <div class="dots">${info.events.map((ev) => `<i style="--c:${member(ev.members?.[0])?.color || 'var(--accent)'}"></i>`).slice(0, 4).join('')}${busyMembers.map((m) => `<i class="ring" style="--c:${m.color}" title="${esc(m.name)} occupé·e"></i>`).join('')}</div>
    </div>`);
  }
  const wd = ['lun', 'mar', 'mer', 'jeu', 'ven', 'sam', 'dim'];
  return `<div class="agenda">
    <section class="card cal">
      <div class="cal-head">
        <h1>${U.cap(U.fmtMonth(first))}</h1>
        <div class="cal-nav">
          <button class="btn ghost small" data-act="prev" aria-label="Mois précédent">‹</button>
          <button class="btn ghost small" data-act="today">Aujourd’hui</button>
          <button class="btn ghost small" data-act="next" aria-label="Mois suivant">›</button>
        </div>
      </div>
      <div class="legend">
        ${ms.map((m) => `<span>${avatar(m, 'xs')} ${esc(m.name)}</span>`).join('')}
        <label class="switch"><input type="checkbox" data-change="show-busy" ${state.showBusy ? 'checked' : ''}> Afficher les occupations perso</label>
      </div>
      <div class="grid wd">${wd.map((w) => `<div>${w}</div>`).join('')}</div>
      <div class="grid month">${cells.join('')}</div>
    </section>
    ${dayPanel(map.get(sel) || dayMap(state.selected, U.addDays(state.selected, 1)).get(sel) || { events: [], busy: [] })}
  </div>`;
}

function chip(ev) {
  const m = member(ev.members?.[0]);
  const time = ev.allDay ? '' : `<b>${U.fmtTime(new Date(ev.start))}</b> `;
  return `<button class="chip ${ev.allDay ? 'allday' : ''}" style="--c:${m?.color || 'var(--accent)'}" data-act="open-event" data-id="${ev.id}">${time}${esc(ev.title)}</button>`;
}

function dayPanel(info) {
  const d = state.selected;
  const evs = info.events.map((ev) => {
    const ms = (ev.members || []).map(member).filter(Boolean);
    return `<li><button class="ev" data-act="open-event" data-id="${ev.id}" style="--c:${ms[0]?.color || 'var(--accent)'}">
      <span class="ev-time">${ev.allDay ? 'Journée' : `${U.fmtTime(new Date(ev.start))}<br><small>${U.fmtTime(new Date(ev.end))}</small>`}</span>
      <span class="ev-body"><b>${esc(ev.title)}</b>${ev.location ? `<small>📍 ${esc(ev.location)}</small>` : ''}<span class="avatars">${ms.map((m) => avatar(m, 'xs')).join('')}</span></span>
    </button></li>`;
  });
  const byMember = new Map();
  for (const b of info.busy) {
    if (!byMember.has(b.member.id)) byMember.set(b.member.id, { m: b.member, items: [] });
    byMember.get(b.member.id).items.push(b);
  }
  const busy = [...byMember.values()].map(
    ({ m, items }) =>
      `<li>${avatar(m, 'xs')} <span><b>${esc(m.name)}</b> · ${items
        .sort((a, b) => a.start.localeCompare(b.start))
        .map((b) => `${b.allDay ? 'journée' : `${U.fmtTime(new Date(b.start))}–${U.fmtTime(new Date(b.end))}`}${b.title ? ` <i>${esc(b.title)}</i>` : ''}`)
        .join(', ')}</span></li>`,
  );
  return `<aside class="card daypanel">
    <div class="panel-head"><h2>${U.cap(U.fmtDayLong(d))}</h2><button class="btn primary small" data-act="new-event" data-day="${U.ymd(d)}">+ Événement</button></div>
    ${evs.length ? `<ul class="evlist">${evs.join('')}</ul>` : `<p class="muted">Aucun événement familial ce jour-là.</p>`}
    ${state.showBusy ? `<h3>Occupations personnelles</h3>${busy.length ? `<ul class="busylist">${busy.join('')}</ul>` : `<p class="muted small">Personne n’a d’occupation connue.</p>`}` : ''}
  </aside>`;
}

// ── Modale événement ────────────────────────────────────────
function memberChecks(selected, name = 'members') {
  return `<div class="checks">${members()
    .map((m) => `<label class="pill" style="--c:${m.color}"><input type="checkbox" name="${name}" value="${m.id}" ${selected.includes(m.id) ? 'checked' : ''}><span>${esc(m.name)}</span></label>`)
    .join('')}</div>`;
}

function openEventModal(ev, day) {
  const isNew = !ev;
  if (!ev) {
    const base = day ? new Date(day.getFullYear(), day.getMonth(), day.getDate(), 18, 0) : U.addMinutes(today(), 18 * 60);
    ev = { id: U.uid(), title: '', start: base.toISOString(), end: U.addMinutes(base, 60).toISOString(), allDay: false, members: state.me ? [state.me] : [], location: '', notes: '' };
  }
  const s = new Date(ev.start), e = new Date(ev.end);
  const vS = ev.allDay ? U.ymd(s) : U.toLocalInput(s);
  const vE = ev.allDay ? U.ymd(U.addDays(e, -1)) : U.toLocalInput(e);
  const type = ev.allDay ? 'date' : 'datetime-local';
  openModal(
    `<form id="evf" novalidate>
      <h2>${isNew ? 'Nouvel événement' : 'Modifier l’événement'}</h2>
      <label>Titre<input name="title" required value="${esc(ev.title)}" placeholder="Anniversaire de Mamie" autocomplete="off"></label>
      <label class="check"><input type="checkbox" name="allDay" ${ev.allDay ? 'checked' : ''}> Journée entière</label>
      <div class="row2">
        <label>Début<input name="start" type="${type}" value="${vS}" required></label>
        <label>Fin<input name="end" type="${type}" value="${vE}" required></label>
      </div>
      <fieldset><legend>Participants</legend>${memberChecks(ev.members || [])}</fieldset>
      <label>Lieu<input name="location" value="${esc(ev.location)}" placeholder="Facultatif"></label>
      <label>Notes<textarea name="notes" rows="2" placeholder="Facultatif">${esc(ev.notes)}</textarea></label>
      <div id="conflicts"></div>
      <div class="modal-actions">
        ${!isNew ? `<button type="button" class="btn danger ghost" data-m="del">Supprimer</button>` : ''}
        <span class="grow"></span>
        <button type="button" class="btn ghost" data-close>Annuler</button>
        <button class="btn primary">Enregistrer</button>
      </div>
      ${!isNew ? `<div class="push"><span class="muted small">Copier dans un agenda perso :</span>
        <button type="button" class="btn small" data-m="push-google" ${CONFIG.google.clientId ? '' : 'disabled title="Non configuré"'}>Google</button>
        <button type="button" class="btn small" data-m="push-outlook" ${CONFIG.microsoft.clientId ? '' : 'disabled title="Non configuré"'}>Outlook</button>
        <button type="button" class="btn small" data-m="ics">Fichier .ics</button></div>` : ''}
    </form>`,
    (d) => {
      const f = d.querySelector('#evf');
      const el = (n) => f.elements.namedItem(n);
      const read = () => {
        const allDay = el('allDay').checked;
        let s, e;
        if (allDay) (s = U.parseYmd(el('start').value)), (e = U.addDays(U.parseYmd(el('end').value || el('start').value), 1));
        else (s = new Date(el('start').value)), (e = new Date(el('end').value));
        return {
          ...ev,
          title: el('title').value.trim(),
          allDay,
          start: isNaN(s) ? null : s.toISOString(),
          end: isNaN(e) ? null : e.toISOString(),
          members: [...f.querySelectorAll('input[name=members]:checked')].map((i) => i.value),
          location: el('location').value.trim(),
          notes: el('notes').value.trim(),
        };
      };
      const showConflicts = () => {
        const x = read();
        const box = d.querySelector('#conflicts');
        if (!x.start || !x.end) return (box.innerHTML = '');
        const lines = x.members
          .map((id) => ({ m: member(id), c: conflictsOf(id, new Date(x.start), new Date(x.end), x.id) }))
          .filter((r) => r.m && r.c.length)
          .map((r) => `<li>${avatar(r.m, 'xs')} <b>${esc(r.m.name)}</b> : ${esc(r.c.join(', '))}</li>`);
        box.innerHTML = lines.length ? `<div class="warn"><b>⚠ Conflits possibles</b><ul>${lines.join('')}</ul></div>` : '';
      };
      el('allDay').addEventListener('change', () => {
        const on = el('allDay').checked;
        for (const [n, h] of [['start', '18:00'], ['end', '19:00']]) {
          const i = el(n), v = i.value;
          i.type = on ? 'date' : 'datetime-local';
          i.value = on ? v.slice(0, 10) : `${v.slice(0, 10)}T${h}`;
        }
        showConflicts();
      });
      el('start').addEventListener('change', () => {
        const x = read();
        if (x.start && (!x.end || x.end <= x.start)) {
          el('end').value = x.allDay ? el('start').value : U.toLocalInput(U.addMinutes(new Date(x.start), 60));
        }
        showConflicts();
      });
      f.addEventListener('change', showConflicts);
      showConflicts();
      const valid = (x) => {
        if (!x.title) return toast('Donnez un titre à l’événement', 'error'), false;
        if (!x.start || !x.end || x.end <= x.start) return toast('La fin doit être après le début', 'error'), false;
        return true;
      };
      f.addEventListener('submit', (e) => {
        e.preventDefault();
        const x = read();
        if (!valid(x)) return;
        x.updatedAt = new Date().toISOString();
        x.createdBy ??= state.me;
        save('events', x);
        state.selected = U.startOfDay(new Date(x.start));
        state.cursor = state.selected;
        closeModal();
        toast(isNew ? 'Événement ajouté' : 'Événement mis à jour');
      });
      f.addEventListener('click', async (e) => {
        const b = e.target.closest('[data-m]');
        if (!b) return;
        const x = read();
        if (b.dataset.m === 'del') {
          if (!confirm(`Supprimer « ${ev.title} » pour toute la famille ?`)) return;
          remove('events', ev.id);
          closeModal();
          return toast('Événement supprimé');
        }
        if (!valid(x)) return;
        if (b.dataset.m === 'ics') return U.download(`${slug(x.title)}.ics`, U.buildIcs([x], CONFIG.appName));
        save('events', x);
        b.disabled = true;
        await pushEvent(b.dataset.m === 'push-google' ? 'google' : 'outlook', x);
        b.disabled = false;
      });
    },
  );
}
const slug = (s) => (s || 'evenement').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/gi, '-').toLowerCase().replace(/^-|-$/g, '');

async function pushEvent(provider, ev) {
  const map = ls.get(LS.pushed, {});
  const ids = map[ev.id] || {};
  try {
    if (provider === 'google') {
      if (!G.isReady()) await G.connect(CONFIG.google.clientId);
      ids.google = await G.upsertEvent(ev, ids.google);
    } else {
      ids.outlook = await O.upsertEvent(CONFIG.microsoft, ev, ids.outlook);
    }
    map[ev.id] = ids;
    ls.set(LS.pushed, map);
    toast(`« ${ev.title} » copié dans votre agenda ${provider === 'google' ? 'Google' : 'Outlook'}`);
  } catch (e) {
    fail(e);
  }
}

// ── Vue Famille & synchronisation ───────────────────────────
function viewFamily() {
  const ms = members();
  const me = member(state.me);
  const s = sources();
  const lastSync = s.lastSync ? U.fmtStamp(new Date(s.lastSync)) : 'jamais';
  const status = (on, label) => `<span class="status ${on ? 'on' : ''}">${on ? '● ' + label : '○ Non relié'}</span>`;
  const notConf = `<span class="muted small">Non configuré (voir README)</span>`;
  const admin = isAdmin();
  return `<div class="stack">
    <section class="card">
      <div class="panel-head"><h1>Membres</h1><button class="btn primary small" data-act="add-member">+ Ajouter</button></div>
      ${ms.length ? `<ul class="members">${ms
        .map((m) => {
          const b = store.get('busy', m.id);
          const acc = cloud ? accessOfMember(m.id) : null;
          const pending = cloud && admin ? store.list('invites').find((i) => i.memberId === m.id) : null;
          return `<li>${avatar(m)}<span class="grow"><b>${esc(m.name)}</b>${m.id === state.me ? ' <span class="tag">vous</span>' : ''}${acc?.role === 'admin' ? ' <span class="tag open">admin</span>' : ''}<br>
            <small class="muted">${cloud ? (acc ? `🔒 ${esc(acc.email || 'compte relié')} · ` : pending ? 'Invitation en attente · ' : 'Pas encore de compte · ') : ''}${b ? `agenda partagé, mis à jour ${U.fmtStamp(new Date(b.updatedAt))}` : 'aucun agenda relié'}</small></span>
            ${cloud && admin && !acc ? `<button class="btn small" data-act="invite" data-id="${m.id}">${pending ? 'Lien d’invitation' : 'Inviter'}</button>` : ''}
            <button class="btn ghost small" data-act="edit-member" data-id="${m.id}">Modifier</button></li>`;
        })
        .join('')}</ul>` : `<p class="muted">Aucun membre pour l’instant.</p>`}
    </section>

    <section class="card">
      <h1>Mes agendas personnels</h1>
      ${!me
        ? `<p class="muted">${cloud ? 'Votre compte n’est relié à aucun membre de la famille. Demandez à un administrateur de vous réinviter.' : 'Choisissez d’abord qui vous êtes (en haut à droite) pour relier vos agendas sur cet appareil.'}</p>`
        : `<p class="muted">Vos agendas restent sur cet appareil. ${s.share ? 'La famille ne voit que vos <b>créneaux occupés</b>, sans titre ni détail.' : 'Rien n’est partagé avec la famille.'}</p>
      <ul class="sources">
        <li><span class="src-name">Google Agenda</span>${CONFIG.google.clientId ? `${status(s.google, G.isReady() ? 'Connecté' : 'Relié')}
          ${s.google ? `<button class="btn ghost small" data-act="disconnect-google">Retirer</button>` : `<button class="btn small" data-act="connect-google">Connecter</button>`}` : notConf}</li>
        <li><span class="src-name">Outlook / Microsoft 365</span>${CONFIG.microsoft.clientId ? `${status(s.outlook, esc(s.outlookUser || 'Relié'))}
          ${s.outlook ? `<button class="btn ghost small" data-act="disconnect-outlook">Retirer</button>` : `<button class="btn small" data-act="connect-outlook">Connecter</button>`}` : notConf}</li>
        ${s.ics
          .map(
            (x, i) => `<li><span class="src-name">Lien iCal · ${esc(x.label || 'agenda')}</span><span class="status on" title="${esc(x.url)}">● ${esc(new URL(x.url.replace(/^webcal:/i, 'https:')).hostname)}</span>
            <button class="btn ghost small" data-act="remove-ics" data-i="${i}">Retirer</button></li>`,
          )
          .join('')}
        <li><span class="src-name">Autre agenda (iCloud, Proton, école…)</span><button class="btn small" data-act="add-ics">+ Lien iCal</button></li>
      </ul>
      <label class="switch"><input type="checkbox" data-change="share" ${s.share ? 'checked' : ''}> Partager mes créneaux occupés avec la famille (sans détails)</label>
      <div class="sync-row">
        <button class="btn primary" data-act="sync" ${!hasSources(s) || state.syncing ? 'disabled' : ''}>${state.syncing ? '<span class="spin"></span> Synchronisation…' : '⟳ Synchroniser maintenant'}</button>
        <span class="muted small">Dernière synchro : ${lastSync}${personalItems() ? ` · ${personalItems().length} événements lus` : ''}</span>
      </div>
      ${s.errors?.length ? `<div class="warn"><b>Problèmes lors de la dernière synchro</b><ul>${s.errors.map((e) => `<li>${esc(e)}</li>`).join('')}</ul></div>` : ''}`}
    </section>

    ${cloud ? accessSection() : ''}

    <section class="card">
      <h1>${cloud ? 'Sauvegarde' : 'Partage & sauvegarde'}</h1>
      ${cloud
        ? `<p class="muted">Les données sont partagées en temps réel, uniquement avec les comptes autorisés ci-dessus.</p>`
        : `<p class="warn">Mode <b>local</b> : les données sont enregistrées dans ce navigateur uniquement. Pour partager l’agenda entre plusieurs téléphones et ordinateurs, configurez Firebase (voir README).</p>`}
      <div class="btn-row">
        <button class="btn" data-act="export-ics">Exporter l’agenda familial (.ics)</button>
        <button class="btn ghost" data-act="export-json">Sauvegarde (.json)</button>
        <label class="btn ghost">Restaurer…<input type="file" accept="application/json" data-change="import-json" hidden></label>
      </div>
    </section>
  </div>`;
}

function accessSection() {
  const admin = isAdmin();
  const uid = ctx.user.uid;
  const accounts = store.list('access').sort((a, b) => (a.role === b.role ? (a.email || '').localeCompare(b.email || '') : a.role === 'admin' ? -1 : 1));
  const invites = admin ? store.list('invites').sort((a, b) => a.createdAt.localeCompare(b.createdAt)) : [];
  return `<section class="card">
    <h1>Accès & sécurité</h1>
    <p class="muted">Seuls les comptes Google ou Microsoft listés ici peuvent ouvrir l’agenda familial. ${admin ? 'Pour ajouter quelqu’un, créez son profil dans « Membres » puis cliquez sur « Inviter ».' : 'Les administrateurs gèrent les invitations.'}</p>
    <ul class="members">${accounts
      .map((a) => {
        const m = member(a.memberId);
        return `<li>${m ? avatar(m) : '<span class="avatar">?</span>'}<span class="grow"><b>${esc(a.email || a.name || 'Compte sans e-mail')}</b>${a.uid === uid ? ' <span class="tag">vous</span>' : ''}<br>
          <small class="muted">${m ? esc(m.name) : 'profil supprimé'} · ${a.role === 'admin' ? 'administrateur' : 'membre'} · depuis ${U.fmtStamp(new Date(a.joinedAt))}</small></span>
          ${admin && a.uid !== uid ? `<button class="btn ghost small" data-act="toggle-admin" data-id="${a.uid}">${a.role === 'admin' ? 'Retirer admin' : 'Rendre admin'}</button>
          <button class="btn ghost small danger" data-act="revoke" data-id="${a.uid}">Retirer l’accès</button>` : ''}</li>`;
      })
      .join('')}</ul>
    ${invites.length ? `<h3>Invitations en attente</h3><ul class="members">${invites
      .map((i) => {
        const expired = new Date(i.expiresAt) < new Date();
        return `<li>${avatar(member(i.memberId))}<span class="grow"><b>${esc(member(i.memberId)?.name || 'Profil supprimé')}</b><br>
          <small class="muted">${expired ? 'Expirée' : 'Valable jusqu’au ' + U.fmtStamp(new Date(i.expiresAt))} · usage unique</small></span>
          <button class="btn ghost small danger" data-act="cancel-invite" data-id="${i.id}">Annuler</button></li>`;
      })
      .join('')}</ul>` : ''}
    <div class="btn-row"><button class="btn ghost" data-act="sign-out">Se déconnecter (${esc(ctx.user.email || 'compte')})</button></div>
  </section>`;
}

function inviteLink(inviteId, memberId) {
  return `${baseUrl()}#/agenda?f=${encodeURIComponent(familyId)}&i=${encodeURIComponent(inviteId)}&m=${encodeURIComponent(memberId)}`;
}

async function openInviteModal(memberId) {
  const m = member(memberId);
  let inv = store.list('invites').find((i) => i.memberId === memberId && new Date(i.expiresAt) > new Date());
  try {
    if (!inv) {
      const id = await C.createInvite(ctx, familyId, memberId);
      inv = { id, expiresAt: U.addDays(new Date(), C.INVITE_DAYS).toISOString() };
    }
  } catch (e) {
    return fail(e);
  }
  const link = inviteLink(inv.id, memberId);
  openModal(
    `<h2>Inviter ${esc(m.name)}</h2>
    <p>Envoyez ce lien à <b>${esc(m.name)}</b> uniquement, par SMS ou messagerie privée. En l’ouvrant, ${esc(m.name)} se connecte avec son compte Google ou Microsoft, qui est alors autorisé.</p>
    <label>Lien d’invitation<input readonly value="${esc(link)}" onclick="this.select()"></label>
    <p class="muted small">Utilisable une seule fois, valable jusqu’au ${U.fmtStamp(new Date(inv.expiresAt))}. Une fois utilisé, le lien ne sert plus à rien : il ne donne aucun accès à une autre personne.</p>
    <div class="modal-actions"><span class="grow"></span><button class="btn ghost" data-close>Fermer</button><button class="btn primary" data-m="copy">Copier le lien</button></div>`,
    (d) => d.querySelector('[data-m=copy]').addEventListener('click', () => copy(link, 'Lien d’invitation copié')),
  );
}

function openAccountModal() {
  const a = myAccess();
  openModal(
    `<h2>Mon compte</h2>
    <p><b>${esc(ctx.user.displayName || '')}</b><br><span class="muted">${esc(ctx.user.email || '')}</span></p>
    <p class="small">Profil : ${esc(member(a?.memberId)?.name || '—')} · ${a?.role === 'admin' ? 'administrateur' : 'membre'}</p>
    <div class="modal-actions"><button class="btn ghost danger" data-m="forget">Quitter cet espace sur cet appareil</button><span class="grow"></span>
      <button class="btn ghost" data-close>Fermer</button><button class="btn primary" data-m="out">Se déconnecter</button></div>`,
    (d) => {
      d.querySelector('[data-m=out]').addEventListener('click', () => actions['sign-out']());
      d.querySelector('[data-m=forget]').addEventListener('click', () => {
        if (!confirm('Oublier cet espace familial sur cet appareil ? Votre accès reste valable : il suffira de rouvrir un lien de l’espace.')) return;
        localStorage.removeItem(LS.family);
        location.hash = '';
        location.reload();
      });
    },
  );
}

function openMemberModal(m) {
  const isNew = !m;
  const used = members().map((x) => x.color);
  m = m || { id: U.uid(), name: '', color: COLORS.find((c) => !used.includes(c)) || COLORS[0], order: members().length };
  openModal(
    `<form id="mf"><h2>${isNew ? 'Nouveau membre' : 'Modifier ' + esc(m.name)}</h2>
      <label>Prénom<input name="name" required value="${esc(m.name)}" autocomplete="off" placeholder="Ex. Léa"></label>
      <fieldset><legend>Couleur</legend><div class="swatches">${COLORS.map((c) => `<label style="--c:${c}"><input type="radio" name="color" value="${c}" ${c === m.color ? 'checked' : ''}><span></span></label>`).join('')}</div></fieldset>
      <div class="modal-actions">${!isNew && isAdmin() ? `<button type="button" class="btn danger ghost" data-m="del">Supprimer</button>` : ''}<span class="grow"></span>
        <button type="button" class="btn ghost" data-close>Annuler</button><button class="btn primary">Enregistrer</button></div>
    </form>`,
    (d) => {
      const f = d.querySelector('#mf');
      f.addEventListener('submit', (e) => {
        e.preventDefault();
        const name = f.elements.namedItem('name').value.trim();
        if (!name) return toast('Indiquez un prénom', 'error');
        save('members', { ...m, name, color: new FormData(f).get('color') || m.color });
        closeModal();
        if (isNew && !state.me && members().length <= 1) setMe(m.id);
      });
      f.querySelector('[data-m=del]')?.addEventListener('click', () => {
        const acc = cloud ? accessOfMember(m.id) : null;
        if (!confirm(`Retirer ${m.name} de la famille ?${acc ? ' Son compte perdra aussi l’accès.' : ''} Ses événements restent dans l’agenda.`)) return;
        if (acc && acc.uid !== ctx.user.uid) remove('access', acc.uid);
        remove('members', m.id);
        remove('busy', m.id);
        closeModal();
      });
    },
  );
}

function openIcsModal() {
  openModal(
    `<form id="icf"><h2>Ajouter un lien iCal</h2>
      <p class="muted small">Collez l’adresse « iCal » secrète de votre agenda (Google : Paramètres › Adresse secrète au format iCal ; Outlook.com : Paramètres › Agendas partagés › Publier ; iCloud : Partager › Agenda public).</p>
      <label>Adresse (https:// ou webcal://)<input name="url" type="url" required placeholder="https://…/basic.ics"></label>
      <label>Nom<input name="label" placeholder="Ex. Travail, Club de foot"></label>
      <div class="modal-actions"><span class="grow"></span><button type="button" class="btn ghost" data-close>Annuler</button><button class="btn primary">Ajouter</button></div>
    </form>`,
    (d) => {
      const f = d.querySelector('#icf');
      f.addEventListener('submit', (e) => {
        e.preventDefault();
        const url = f.elements.namedItem('url').value.trim();
        try {
          new URL(url.replace(/^webcal:/i, 'https:'));
        } catch {
          return toast('Adresse invalide', 'error');
        }
        const s = sources();
        s.ics.push({ url, label: f.elements.namedItem('label').value.trim() });
        setSources(s);
        closeModal();
        syncNow(true);
      });
    },
  );
}

// ── Synchronisation ─────────────────────────────────────────
async function syncNow(interactive) {
  if (!state.me || state.syncing) return;
  const s = sources();
  if (!hasSources(s)) return;
  state.syncing = true;
  render();
  const from = U.addDays(today(), -7);
  const to = U.addDays(today(), CONFIG.syncDays || 120);
  const prev = personal() || { bySource: {} };
  const bySource = {};
  const errors = [];
  const run = async (key, label, fn) => {
    try {
      bySource[key] = await fn();
    } catch (e) {
      errors.push(`${label} : ${e.message}`);
      if (prev.bySource[key]) bySource[key] = prev.bySource[key]; // on garde la dernière version connue
    }
  };
  if (s.google)
    await run('google', 'Google', async () => {
      if (!G.isReady()) {
        if (!interactive) throw new Error('reconnexion nécessaire, cliquez sur « Synchroniser ».');
        await G.connect(CONFIG.google.clientId);
      }
      return G.fetchEvents(from, to);
    });
  if (s.outlook) await run('outlook', 'Outlook', () => O.fetchEvents(CONFIG.microsoft, from, to, interactive));
  for (const x of s.ics) await run('ics:' + x.url, 'iCal ' + (x.label || ''), () => ICS.fetchEvents(x.url, CONFIG.icsProxy, from, to));

  setPersonal({ bySource, updatedAt: new Date().toISOString() });
  const items = Object.values(bySource).flat();
  if (s.share) {
    await save('busy', {
      id: state.me,
      slots: U.mergeSlots(items.filter((i) => i.busy)).map((x) => ({ ...x })),
      updatedAt: new Date().toISOString(),
    });
  }
  setSources({ ...sources(), lastSync: new Date().toISOString(), errors });
  state.syncing = false;
  render();
  if (errors.length) toast('Synchro partielle : ' + errors[0], 'error');
  else if (interactive) toast(`Synchronisé : ${items.length} événements lus`);
}

async function connectProvider(p) {
  try {
    const s = sources();
    if (p === 'google') {
      await G.connect(CONFIG.google.clientId);
      s.google = true;
    } else {
      s.outlookUser = await O.connect(CONFIG.microsoft);
      s.outlook = true;
    }
    setSources(s);
    await syncNow(true);
  } catch (e) {
    fail(e);
  }
}

async function dropSource(key, mutate) {
  const s = sources();
  mutate(s);
  setSources(s);
  const p = personal();
  if (p) {
    delete p.bySource[key];
    setPersonal(p);
  }
  if (hasSources(s)) await syncNow(false);
  else {
    setPersonal(null);
    remove('busy', state.me);
    render();
  }
}

// ── Sondages ────────────────────────────────────────────────
function pollStats(p) {
  const parts = p.participants.map(member).filter(Boolean);
  const rows = p.options.map((o) => {
    const c = { yes: 0, maybe: 0, no: 0 };
    for (const m of parts) {
      const v = p.votes?.[m.id]?.[o.id];
      if (v) c[v]++;
    }
    return { o, ...c, score: c.yes * 2 + c.maybe };
  });
  const max = Math.max(0, ...rows.map((r) => r.score));
  const pending = parts.filter((m) => !p.options.every((o) => p.votes?.[m.id]?.[o.id]));
  return { parts, rows, max, pending };
}

function viewPolls() {
  if (state.pollId) {
    const p = store.get('polls', state.pollId);
    return p ? viewPoll(p) : `<section class="card empty"><p>Ce sondage n’existe plus.</p><a class="btn" href="#/sondages">Retour aux sondages</a></section>`;
  }
  const polls = store.list('polls').sort((a, b) => (a.status === b.status ? b.createdAt.localeCompare(a.createdAt) : a.status === 'open' ? -1 : 1));
  return `<section class="card">
    <div class="panel-head"><h1>Sondages de dates</h1><button class="btn primary small" data-act="new-poll" ${members().length ? '' : 'disabled'}>+ Nouveau sondage</button></div>
    <p class="muted">Proposez plusieurs dates, chacun répond « oui / peut-être / non », et l’application signale qui est déjà occupé d’après son agenda.</p>
    ${polls.length ? `<ul class="polls">${polls
      .map((p) => {
        const st = pollStats(p);
        const chosen = p.options.find((o) => o.id === p.chosenOptionId);
        const mine = state.me && p.participants.includes(state.me) && st.pending.some((m) => m.id === state.me);
        return `<li><a href="#/sondages/${p.id}" class="poll-item">
          <span class="grow"><b>${esc(p.title)}</b><br><small class="muted">${
            p.status === 'open' ? `${p.options.length} dates · ${st.parts.length - st.pending.length}/${st.parts.length} réponses` : chosen ? '✓ ' + U.fmtRange(chosen.start, chosen.end, p.allDay) : 'Clos'
          }</small></span>
          ${mine ? '<span class="tag accent">À vous de répondre</span>' : ''}<span class="tag ${p.status === 'open' ? 'open' : ''}">${p.status === 'open' ? 'Ouvert' : 'Clos'}</span></a></li>`;
      })
      .join('')}</ul>` : `<p class="empty-inline">Aucun sondage pour l’instant.</p>`}
  </section>`;
}

const VOTE = { yes: ['Oui', '✓'], maybe: ['Peut-être', '~'], no: ['Non', '✕'] };

function viewPoll(p) {
  const { parts, rows, max, pending } = pollStats(p);
  const open = p.status === 'open';
  const amIn = state.me && p.participants.includes(state.me);
  const cards = rows.map(({ o, yes, maybe, no, score }) => {
    const s = new Date(o.start), e = new Date(o.end);
    const best = score > 0 && score === max;
    const chosen = p.chosenOptionId === o.id;
    const mine = p.votes?.[state.me]?.[o.id];
    const people = parts
      .map((m) => {
        const v = p.votes?.[m.id]?.[o.id];
        const c = conflictsOf(m.id, s, e, p.eventId);
        return `<span class="voter ${v || 'none'}" title="${esc(m.name)} : ${v ? VOTE[v][0] : 'pas encore répondu'}${c.length ? ' · occupé·e d’après son agenda' : ''}">${avatar(m, 'xs')}${esc(m.name)}<b>${v ? VOTE[v][1] : '·'}</b>${c.length ? '<em>⚠</em>' : ''}</span>`;
      })
      .join('');
    const myConf = amIn ? conflictsOf(state.me, s, e, p.eventId) : [];
    return `<li class="opt ${best ? 'best' : ''} ${chosen ? 'chosen' : ''}">
      <div class="opt-head"><b>${U.fmtRange(s, e, p.allDay)}</b>${chosen ? '<span class="tag open">Date retenue</span>' : best ? '<span class="tag accent">★ Meilleure option</span>' : ''}</div>
      <div class="bar" title="${yes} oui · ${maybe} peut-être · ${no} non"><i class="y" style="flex:${yes}"></i><i class="m" style="flex:${maybe}"></i><i class="n" style="flex:${no}"></i><i class="z" style="flex:${Math.max(0, parts.length - yes - maybe - no)}"></i></div>
      <div class="voters">${people}</div>
      ${amIn && open ? `<div class="myvote"><span class="muted small">Votre réponse${myConf.length ? ` <span class="warn-inline">⚠ ${esc(myConf.join(', '))}</span>` : ''}</span>
        <div class="seg">${Object.entries(VOTE).map(([k, [label]]) => `<button class="${k} ${mine === k ? 'on' : ''}" data-act="vote" data-opt="${o.id}" data-val="${k}">${label}</button>`).join('')}</div></div>` : ''}
      ${open ? `<button class="btn small ghost validate" data-act="choose" data-opt="${o.id}">Retenir cette date</button>` : ''}
    </li>`;
  });
  return `<section class="card">
    <a href="#/sondages" class="back">‹ Tous les sondages</a>
    <div class="panel-head"><h1>${esc(p.title)}</h1><span class="tag ${open ? 'open' : ''}">${open ? 'Ouvert' : 'Clos'}</span></div>
    ${p.description ? `<p>${esc(p.description)}</p>` : ''}
    <p class="muted small">${p.location ? '📍 ' + esc(p.location) + ' · ' : ''}Créé par ${esc(member(p.createdBy)?.name || '—')} · ${parts.length} participants</p>
    ${!state.me ? `<p class="warn">Choisissez qui vous êtes (en haut à droite) pour répondre.</p>` : !amIn ? `<p class="muted">Vous n’êtes pas invité·e à ce sondage.</p>` : ''}
    ${open && pending.length ? `<p class="small">En attente de : ${pending.map((m) => `${avatar(m, 'xs')} ${esc(m.name)}`).join(' ')}</p>` : ''}
    <div class="btn-row">
      ${amIn && open ? `<button class="btn small" data-act="autofill">Répondre d’après mon agenda</button>` : ''}
      <button class="btn small ghost" data-act="copy-poll">Copier le lien</button>
      ${open ? `<button class="btn small ghost" data-act="close-poll">Clore</button>` : `<button class="btn small ghost" data-act="reopen-poll">Rouvrir</button>`}
      <button class="btn small ghost danger" data-act="delete-poll">Supprimer</button>
    </div>
    <ul class="opts">${cards.join('')}</ul>
  </section>`;
}

const WINDOWS = { evening: [18 * 60, 22 * 60], lunch: [12 * 60, 14 * 60], day: [9 * 60, 18 * 60], all: [8 * 60, 22 * 60] };

function suggestSlots({ from, to, duration, allDay, win, weekendOnly, participants, count }) {
  const out = [];
  const now = new Date();
  for (let d = from; d <= to && out.length < count; d = U.addDays(d, 1)) {
    if (weekendOnly && !U.isWeekend(d)) continue;
    if (allDay) {
      const s = d, e = U.addDays(d, 1);
      if (s >= today() && participants.every((id) => !conflictsOf(id, s, e).length)) out.push({ s, e });
      continue;
    }
    const [a, b] = WINDOWS[win];
    for (let t = a; t + duration <= b; t += 30) {
      const s = U.addMinutes(d, t), e = U.addMinutes(s, duration);
      if (s < now) continue;
      if (participants.every((id) => !conflictsOf(id, s, e).length)) {
        out.push({ s, e });
        break;
      }
    }
  }
  return out;
}

function openPollModal() {
  let opts = [];
  const ms = members();
  openModal(
    `<form id="pf" novalidate><h2>Nouveau sondage</h2>
      <label>Objet<input name="title" required placeholder="Ex. Repas de Noël" autocomplete="off"></label>
      <label>Description<textarea name="description" rows="2" placeholder="Facultatif"></textarea></label>
      <div class="row2">
        <label>Durée<select name="duration">${[[60, '1 h'], [90, '1 h 30'], [120, '2 h'], [180, '3 h'], [240, '4 h'], [30, '30 min'], ['day', 'Journée entière']].map(([v, l]) => `<option value="${v}" ${v === 120 ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
        <label>Lieu<input name="location" placeholder="Facultatif"></label>
      </div>
      <fieldset><legend>Participants</legend>${memberChecks(ms.map((m) => m.id), 'participants')}</fieldset>
      <fieldset><legend>Dates proposées</legend>
        <ul id="optlist" class="optlist"></ul>
        <div class="row-add"><input id="optIn" type="datetime-local" aria-label="Date à proposer"><button type="button" class="btn small" data-m="add">Ajouter</button></div>
        <details class="suggest" open><summary>✨ Trouver des créneaux où tout le monde est libre</summary>
          <div class="row2"><label>Du<input type="date" id="sgFrom" value="${U.ymd(U.addDays(today(), 1))}"></label><label>Au<input type="date" id="sgTo" value="${U.ymd(U.addDays(today(), 21))}"></label></div>
          <div class="row2"><label>Moment<select id="sgWin"><option value="evening">Soirée (18 h–22 h)</option><option value="lunch">Midi (12 h–14 h)</option><option value="day">Journée (9 h–18 h)</option><option value="all">N’importe quand (8 h–22 h)</option></select></label>
          <label>Nombre<input type="number" id="sgCount" value="5" min="1" max="15"></label></div>
          <label class="check"><input type="checkbox" id="sgWe"> Week-ends uniquement</label>
          <button type="button" class="btn small" data-m="suggest">Proposer des créneaux</button>
          <p class="muted small" id="sgInfo"></p>
        </details>
      </fieldset>
      <div class="modal-actions"><span class="grow"></span><button type="button" class="btn ghost" data-close>Annuler</button><button class="btn primary">Créer le sondage</button></div>
    </form>`,
    (d) => {
      const f = d.querySelector('#pf');
      const el = (n) => f.elements.namedItem(n);
      const $ = (s) => d.querySelector(s);
      const dur = () => el('duration').value;
      const allDay = () => dur() === 'day';
      const participants = () => [...f.querySelectorAll('input[name=participants]:checked')].map((i) => i.value);
      const draw = () => {
        opts.sort((a, b) => a.start.localeCompare(b.start));
        $('#optlist').innerHTML = opts.length
          ? opts.map((o, i) => `<li>${U.fmtRange(o.start, o.end, allDay())}<button type="button" class="btn ghost small" data-m="rm" data-i="${i}" aria-label="Retirer">✕</button></li>`).join('')
          : '<li class="muted small">Ajoutez au moins deux dates, à la main ou avec la suggestion ci-dessous.</li>';
      };
      const addOpt = (s, e) => {
        if (opts.some((o) => o.start === s.toISOString())) return;
        opts.push({ id: U.uid(), start: s.toISOString(), end: e.toISOString() });
      };
      el('duration').addEventListener('change', () => {
        $('#optIn').type = allDay() ? 'date' : 'datetime-local';
        opts = opts.map((o) => {
          const s = allDay() ? U.startOfDay(new Date(o.start)) : new Date(o.start);
          return { ...o, start: s.toISOString(), end: (allDay() ? U.addDays(s, 1) : U.addMinutes(s, +dur())).toISOString() };
        });
        draw();
      });
      f.addEventListener('click', (e) => {
        const b = e.target.closest('[data-m]');
        if (!b) return;
        if (b.dataset.m === 'rm') opts.splice(+b.dataset.i, 1);
        if (b.dataset.m === 'add') {
          const v = $('#optIn').value;
          if (!v) return toast('Choisissez une date', 'error');
          const s = allDay() ? U.parseYmd(v) : new Date(v);
          addOpt(s, allDay() ? U.addDays(s, 1) : U.addMinutes(s, +dur()));
          $('#optIn').value = '';
        }
        if (b.dataset.m === 'suggest') {
          const ps = participants();
          const from = U.parseYmd($('#sgFrom').value || U.ymd(today()));
          const to = U.parseYmd($('#sgTo').value || U.ymd(U.addDays(today(), 21)));
          const found = suggestSlots({ from, to, duration: allDay() ? 0 : +dur(), allDay: allDay(), win: $('#sgWin').value, weekendOnly: $('#sgWe').checked, participants: ps, count: Math.min(15, +$('#sgCount').value || 5) });
          found.forEach(({ s, e }) => addOpt(s, e));
          const blind = ps.map(member).filter((m) => m && !hasCalendar(m.id));
          $('#sgInfo').innerHTML =
            (found.length ? `${found.length} créneau(x) libre(s) ajouté(s).` : 'Aucun créneau commun trouvé sur cette période.') +
            (blind.length ? ` Agenda inconnu pour : ${blind.map((m) => esc(m.name)).join(', ')} (considéré·es libres).` : '');
        }
        draw();
      });
      f.addEventListener('submit', (e) => {
        e.preventDefault();
        const title = el('title').value.trim();
        const ps = participants();
        if (!title) return toast('Donnez un objet au sondage', 'error');
        if (!ps.length) return toast('Choisissez au moins un participant', 'error');
        if (!opts.length) return toast('Proposez au moins une date', 'error');
        const p = {
          id: U.uid(),
          title,
          description: el('description').value.trim(),
          location: el('location').value.trim(),
          allDay: allDay(),
          participants: ps,
          options: opts,
          votes: {},
          status: 'open',
          createdBy: state.me,
          createdAt: new Date().toISOString(),
        };
        save('polls', p);
        closeModal();
        go('#/sondages/' + p.id);
      });
      draw();
    },
  );
}

// ── Actions ─────────────────────────────────────────────────
function setMe(id, rerender = true) {
  state.me = id || null;
  ls.set(LS.me, state.me);
  if (rerender) {
    render();
    if (state.me) syncNow(false);
  }
}

async function copy(text, msg) {
  try {
    await navigator.clipboard.writeText(text);
    toast(msg);
  } catch {
    prompt('Copiez ce lien :', text);
  }
}
const baseUrl = () => location.origin + location.pathname;
const familyParam = () => (cloud ? '?f=' + encodeURIComponent(familyId) : '');

const actions = {
  prev: () => ((state.cursor = new Date(state.cursor.getFullYear(), state.cursor.getMonth() - 1, 1)), render()),
  next: () => ((state.cursor = new Date(state.cursor.getFullYear(), state.cursor.getMonth() + 1, 1)), render()),
  today: () => ((state.cursor = state.selected = today()), render()),
  'select-day': (el) => {
    state.selected = U.parseYmd(el.dataset.day);
    if (state.selected.getMonth() !== state.cursor.getMonth()) state.cursor = state.selected;
    render();
    if (matchMedia('(max-width: 900px)').matches) document.querySelector('.daypanel')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  },
  'new-event': (el) => openEventModal(null, el.dataset.day ? U.parseYmd(el.dataset.day) : state.selected),
  'open-event': (el) => {
    const ev = store.get('events', el.dataset.id);
    if (ev) openEventModal(ev);
  },
  'add-member': () => openMemberModal(),
  'edit-member': (el) => openMemberModal(member(el.dataset.id)),
  sync: () => syncNow(true),
  'connect-google': () => connectProvider('google'),
  'connect-outlook': () => connectProvider('outlook'),
  'disconnect-google': () => (G.disconnect(), dropSource('google', (s) => (s.google = false))),
  'disconnect-outlook': () => (O.disconnect(), dropSource('outlook', (s) => ((s.outlook = false), (s.outlookUser = '')))),
  'add-ics': () => openIcsModal(),
  'remove-ics': (el) => {
    const url = sources().ics[+el.dataset.i]?.url;
    dropSource('ics:' + url, (s) => s.ics.splice(+el.dataset.i, 1));
  },
  'export-ics': () => U.download('agenda-famille.ics', U.buildIcs(store.list('events'), CONFIG.appName)),
  'export-json': () => {
    const data = { version: 1, exportedAt: new Date().toISOString(), members: store.list('members'), events: store.list('events'), polls: store.list('polls') };
    U.download(`agenda-famille-${U.ymd(new Date())}.json`, JSON.stringify(data, null, 2), 'application/json');
  },
  invite: (el) => openInviteModal(el.dataset.id),
  'cancel-invite': (el) => remove('invites', el.dataset.id),
  revoke: (el) => {
    const a = store.get('access', el.dataset.id);
    if (!confirm(`Retirer l’accès de ${a.email || 'ce compte'} ? Il ne pourra plus ouvrir l’agenda familial.`)) return;
    remove('access', a.uid);
    toast('Accès retiré');
  },
  'toggle-admin': (el) => {
    const a = store.get('access', el.dataset.id);
    patch('access', a.uid, ['role'], a.role === 'admin' ? 'member' : 'admin');
  },
  account: () => openAccountModal(),
  'sign-out': async () => {
    await C.signOut(ctx);
    location.hash = '';
    location.reload();
  },
  'new-poll': () => openPollModal(),
  vote: (el) => {
    const p = store.get('polls', state.pollId);
    const cur = p?.votes?.[state.me]?.[el.dataset.opt];
    patch('polls', state.pollId, ['votes', state.me, el.dataset.opt], cur === el.dataset.val ? undefined : el.dataset.val);
  },
  autofill: () => {
    const p = store.get('polls', state.pollId);
    if (!hasCalendar(state.me)) return toast('Reliez d’abord un agenda dans l’onglet Famille', 'error');
    const votes = { ...(p.votes?.[state.me] || {}) };
    for (const o of p.options) votes[o.id] = conflictsOf(state.me, new Date(o.start), new Date(o.end)).length ? 'no' : 'yes';
    patch('polls', p.id, ['votes', state.me], votes);
    toast('Réponses pré-remplies, vous pouvez les ajuster');
  },
  choose: (el) => {
    const p = store.get('polls', state.pollId);
    const o = p.options.find((x) => x.id === el.dataset.opt);
    if (!confirm(`Retenir le ${U.fmtRange(o.start, o.end, p.allDay)} et l’ajouter à l’agenda familial ?`)) return;
    const ev = { id: U.uid(), title: p.title, start: o.start, end: o.end, allDay: !!p.allDay, members: p.participants, location: p.location || '', notes: p.description || '', fromPoll: p.id, createdBy: state.me, updatedAt: new Date().toISOString() };
    save('events', ev);
    update('polls', p.id, { status: 'closed', chosenOptionId: o.id, eventId: ev.id });
    toast('Date retenue et ajoutée à l’agenda');
  },
  'close-poll': () => patch('polls', state.pollId, ['status'], 'closed'),
  'reopen-poll': () => patch('polls', state.pollId, ['status'], 'open'),
  'delete-poll': () => {
    if (!confirm('Supprimer ce sondage pour tout le monde ?')) return;
    remove('polls', state.pollId);
    go('#/sondages');
  },
  'copy-poll': () => copy(baseUrl() + '#/sondages/' + state.pollId + familyParam(), 'Lien du sondage copié'),
};

const changes = {
  me: (el) => setMe(el.value),
  'show-busy': (el) => ((state.showBusy = el.checked), ls.set(LS.showBusy, el.checked), render()),
  share: async (el) => {
    setSources({ ...sources(), share: el.checked });
    if (el.checked) await syncNow(false);
    else remove('busy', state.me);
    render();
  },
  'import-json': async (el) => {
    const file = el.files?.[0];
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (!confirm(`Restaurer ${data.members?.length || 0} membres, ${data.events?.length || 0} événements et ${data.polls?.length || 0} sondages ? Les éléments existants avec le même identifiant seront remplacés.`)) return;
      for (const k of ['members', 'events', 'polls']) for (const o of data[k] || []) if (o?.id) await save(k, o);
      toast('Sauvegarde restaurée');
    } catch (e) {
      fail(new Error('Fichier de sauvegarde illisible'));
    }
  },
};

app.addEventListener('click', (e) => {
  const el = e.target.closest('[data-act]');
  if (el && !el.disabled) actions[el.dataset.act]?.(el, e);
});
app.addEventListener('keydown', (e) => {
  if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('[data-act=select-day]')) {
    e.preventDefault();
    actions['select-day'](e.target);
  }
});
app.addEventListener('change', (e) => {
  const el = e.target.closest('[data-change]');
  if (el) changes[el.dataset.change]?.(el, e);
});

// ── Démarrage ───────────────────────────────────────────────
function gate(inner) {
  app.innerHTML = `<main class="main"><section class="card gate"><h1><span class="logo" aria-hidden="true">◆</span> ${esc(CONFIG.appName)}</h1>${inner}</section></main>`;
}
const pendingInvite = () => ls.get(LS.invite, null);

/** Écran de connexion : aucun accès sans compte Google ou Microsoft. */
function loginScreen() {
  const inv = pendingInvite();
  gate(`<p>${inv ? 'Vous avez reçu une invitation. Connectez-vous avec <b>votre</b> compte pour rejoindre l’agenda familial.' : 'Connectez-vous pour accéder à l’agenda familial.'}</p>
    <div class="btn-col">
      <button class="btn" data-gate="google"><b class="g">G</b> Continuer avec Google</button>
      <button class="btn" data-gate="microsoft"><b class="ms">▦</b> Continuer avec Microsoft</button>
    </div>
    <p class="muted small">Votre compte sert uniquement à vous identifier. L’accès à l’agenda de vos proches est réservé aux comptes invités.</p>`);
  for (const b of app.querySelectorAll('[data-gate]'))
    b.onclick = async () => {
      b.disabled = true;
      try {
        await C.signIn(ctx, b.dataset.gate);
        location.reload();
      } catch (e) {
        b.disabled = false;
        if (e.code !== 'auth/popup-closed-by-user' && e.code !== 'auth/cancelled-popup-request') fail(new Error(authError(e)));
      }
    };
}

function authError(e) {
  return (
    {
      'auth/unauthorized-domain': 'Ce site n’est pas autorisé dans Firebase (Authentication › Paramètres › Domaines autorisés).',
      'auth/operation-not-allowed': 'Ce mode de connexion n’est pas activé dans Firebase (Authentication › Sign-in method).',
      'auth/account-exists-with-different-credential': 'Cette adresse est déjà utilisée avec l’autre fournisseur : connectez-vous avec celui-ci.',
      'auth/popup-blocked': 'Le navigateur a bloqué la fenêtre de connexion : autorisez les pop-ups pour ce site.',
    }[e.code] || e.message
  );
}

/** Connecté, mais aucun espace familial connu sur cet appareil. */
function noFamilyScreen() {
  const first = (ctx.user.displayName || '').split(' ')[0];
  gate(`<p>Connecté·e en tant que <b>${esc(ctx.user.email || ctx.user.displayName || '')}</b>.</p>
    <div class="gate-grid">
      <form data-gate-form><h2>Créer l’espace familial</h2>
        <p class="muted small">Vous en serez l’administrateur, et vous inviterez ensuite chaque membre.</p>
        <label>Nom de l’espace<input name="family" required value="Famille" autocomplete="off"></label>
        <label>Votre prénom<input name="name" required value="${esc(first)}" autocomplete="off"></label>
        <button class="btn primary">Créer</button></form>
      <div><h2>Rejoindre</h2><p class="muted small">Pour rejoindre l’espace d’un proche, ouvrez le <b>lien d’invitation personnel</b> qu’il ou elle vous a envoyé.</p></div>
    </div>
    <div class="btn-row"><button class="btn ghost" data-gate="out">Changer de compte</button></div>`);
  app.querySelector('[data-gate=out]').onclick = () => C.signOut(ctx).then(() => location.reload());
  app.querySelector('[data-gate-form]').onsubmit = async (e) => {
    e.preventDefault();
    const f = e.target.elements;
    const btn = e.target.querySelector('button');
    btn.disabled = true;
    try {
      const id = await C.createFamily(ctx, { familyName: f.namedItem('family').value.trim() || 'Famille', memberName: f.namedItem('name').value.trim() || 'Moi', color: COLORS[0] });
      ls.set(LS.family, id);
      location.hash = '#/famille';
      location.reload();
    } catch (err) {
      btn.disabled = false;
      fail(err);
    }
  };
}

function deniedScreen(reason) {
  gate(`<p>${reason || `Le compte <b>${esc(ctx.user.email || '')}</b> n’a pas accès à cet agenda familial.`}</p>
    <p class="muted">Demandez à un administrateur de la famille de vous envoyer un lien d’invitation, puis ouvrez-le en étant connecté·e avec ce compte.</p>
    <div class="btn-row"><button class="btn" data-gate="out">Changer de compte</button><button class="btn ghost" data-gate="new">Créer mon propre espace</button></div>`);
  app.querySelector('[data-gate=out]').onclick = () => C.signOut(ctx).then(() => location.reload());
  app.querySelector('[data-gate=new]').onclick = () => {
    localStorage.removeItem(LS.family);
    noFamilyScreen();
  };
}

function errorScreen(e) {
  console.error(e);
  gate(`<p><b>Connexion impossible</b> : ${esc(e.message)}</p>
    <p class="muted small">Vérifiez la configuration Firebase (voir README) et votre connexion internet.</p>
    <button class="btn" onclick="location.reload()">Réessayer</button>`);
}

/** Démarrage en mode cloud : compte → espace familial → accès → données. */
async function startCloud() {
  ctx = await C.initCloud(CONFIG.firebase);
  // Les paramètres d'un lien (?f=espace&i=invitation&m=membre) sont mémorisés puis retirés de l'adresse
  const { parts, params } = parseHash();
  if (params.get('f')) ls.set(LS.family, params.get('f'));
  if (params.get('i')) ls.set(LS.invite, { f: params.get('f'), i: params.get('i'), m: params.get('m') });
  if ([...params.keys()].length) history.replaceState(null, '', location.pathname + '#/' + parts.join('/'));

  if (!ctx.user) return loginScreen();
  familyId = ls.get(LS.family, null);
  if (!familyId) return noFamilyScreen();

  let access = await C.getAccess(ctx, familyId);
  const inv = pendingInvite();
  if (inv && inv.f === familyId) {
    localStorage.removeItem(LS.invite);
    if (!access) {
      try {
        await C.joinWithInvite(ctx, familyId, inv.i, inv.m);
        access = await C.getAccess(ctx, familyId);
        toast('Bienvenue dans l’agenda familial !');
      } catch (e) {
        console.warn(e);
        return deniedScreen('Cette invitation n’est plus valable : elle a déjà été utilisée, a expiré ou a été annulée.');
      }
    }
  }
  if (!access) return deniedScreen();

  store = await createCloudStore(ctx, familyId, {
    admin: access.role === 'admin',
    onDenied: () => deniedScreen('Votre accès à cet agenda familial a été retiré.'),
  });
  return true;
}

async function init() {
  if (CONFIG.google.clientId) U.loadScript(G.GSI_SRC).catch(() => {});
  if (cloud) {
    try {
      if (!(await startCloud())) return;
    } catch (e) {
      return errorScreen(e);
    }
  } else {
    store = createLocalStore();
  }
  store.subscribe(render);
  addEventListener('hashchange', route);
  route();
  if (state.me && hasSources()) syncNow(false);
}

init();
