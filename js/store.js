// Stockage des données partagées de la famille.
//  - mode « local » : localStorage (un seul navigateur)
//  - mode « cloud » : Firebase Firestore (toute la famille, temps réel)
// API commune : list(kind), get(kind,id), save(kind,obj), remove(kind,id),
//               patch(kind,id,path[],value), subscribe(fn)

export const KINDS = ['members', 'events', 'polls', 'busy'];
const empty = () => Object.fromEntries(KINDS.map((k) => [k, {}]));
const clone = (o) => JSON.parse(JSON.stringify(o)); // retire aussi les undefined (refusés par Firestore)

function emitter() {
  const subs = new Set();
  let timer = null;
  return {
    subscribe(fn) {
      subs.add(fn);
      return () => subs.delete(fn);
    },
    emit() {
      if (timer) return;
      timer = setTimeout(() => {
        timer = null;
        subs.forEach((f) => f());
      }, 0);
    },
  };
}

function setPath(obj, path, value) {
  let t = obj;
  for (const key of path.slice(0, -1)) t = t[key] && typeof t[key] === 'object' ? t[key] : (t[key] = {});
  const last = path[path.length - 1];
  if (value === undefined) delete t[last];
  else t[last] = value;
}

export function createLocalStore() {
  return localStore();
}

function localStore() {
  const KEY = 'af:data:v1';
  const em = emitter();
  const load = () => {
    try {
      const d = JSON.parse(localStorage.getItem(KEY) || 'null');
      if (d) return { ...empty(), ...d };
    } catch {}
    return empty();
  };
  const data = load();
  const persist = () => {
    try {
      localStorage.setItem(KEY, JSON.stringify(data));
    } catch (e) {
      console.warn('Stockage local indisponible', e);
    }
    em.emit();
  };
  // Synchronise les onglets ouverts sur le même navigateur
  addEventListener('storage', (e) => {
    if (e.key !== KEY) return;
    const n = load();
    for (const k of KINDS) data[k] = n[k];
    em.emit();
  });
  return {
    mode: 'local',
    list: (k) => Object.values(data[k]),
    get: (k, id) => data[k][id] ?? null,
    subscribe: em.subscribe,
    async save(k, o) {
      data[k][o.id] = clone(o);
      persist();
    },
    async remove(k, id) {
      delete data[k][id];
      persist();
    },
    async patch(k, id, path, value) {
      const o = data[k][id];
      if (!o) return;
      setPath(o, path, value);
      persist();
    },
    async update(k, id, fields) {
      if (!data[k][id]) return;
      Object.assign(data[k][id], clone(fields));
      persist();
    },
  };
}

/**
 * Stockage partagé Firestore, temps réel.
 * @param ctx      contexte renvoyé par initCloud (utilisateur déjà connecté)
 * @param familyId identifiant de l'espace familial
 * @param opts     { admin: écoute aussi les invitations ; onDenied: accès retiré en cours de route }
 */
export async function createCloudStore(ctx, familyId, { admin = false, onDenied } = {}) {
  const { fs, db } = ctx;
  const kinds = [...KINDS, 'access', ...(admin ? ['invites'] : [])];
  const data = Object.fromEntries(kinds.map((k) => [k, {}]));
  const em = emitter();
  const col = (k) => fs.collection(db, 'families', familyId, k);
  const ref = (k, id) => fs.doc(db, 'families', familyId, k, id);
  const plain = (d) => {
    for (const [key, v] of Object.entries(d)) if (v && typeof v.toDate === 'function') d[key] = v.toDate().toISOString();
    return d;
  };

  await Promise.all(
    kinds.map(
      (k) =>
        new Promise((resolve, reject) => {
          let first = true;
          fs.onSnapshot(
            col(k),
            (snap) => {
              for (const ch of snap.docChanges()) {
                if (ch.type === 'removed') delete data[k][ch.doc.id];
                else data[k][ch.doc.id] = plain(ch.doc.data());
              }
              em.emit();
              if (first) (first = false), resolve();
            },
            (err) => {
              console.error(k, err);
              if (first) (first = false), reject(err);
              else if (err.code === 'permission-denied') onDenied?.();
            },
          );
        }),
    ),
  );

  return {
    mode: 'cloud',
    list: (k) => Object.values(data[k] || {}),
    get: (k, id) => data[k]?.[id] ?? null,
    subscribe: em.subscribe,
    save: (k, o) => fs.setDoc(ref(k, o.id), clone(o)),
    remove: (k, id) => fs.deleteDoc(ref(k, id)),
    patch: (k, id, path, value) =>
      fs.updateDoc(ref(k, id), path.join('.'), value === undefined ? fs.deleteField() : value),
    update: (k, id, fields) => fs.updateDoc(ref(k, id), clone(fields)),
  };
}
