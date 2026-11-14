// Mode « cloud » : Firebase (connexion par compte Google/Microsoft, espaces familiaux, invitations)
// La sécurité réelle est assurée par les règles Firestore (firestore.rules) ;
// ce module ne fait qu'appeler les opérations autorisées.
import { uid as rid, addDays } from './utils.js';

const BASE = 'https://www.gstatic.com/firebasejs/10.12.2/';
export const INVITE_DAYS = 7;

export async function initCloud(fbConfig) {
  const [appMod, fs, authMod] = await Promise.all([
    import(BASE + 'firebase-app.js'),
    import(BASE + 'firebase-firestore.js'),
    import(BASE + 'firebase-auth.js'),
  ]);
  const app = appMod.initializeApp(fbConfig);
  const auth = authMod.getAuth(app);
  let db;
  try {
    db = fs.initializeFirestore(app, { localCache: fs.persistentLocalCache() });
  } catch {
    db = fs.getFirestore(app);
  }
  const user = await new Promise((resolve) => {
    const stop = authMod.onAuthStateChanged(auth, (u) => {
      stop();
      resolve(u);
    });
  });
  return { app, auth, db, fs, authMod, user };
}

export async function signIn(ctx, provider) {
  const { authMod } = ctx;
  const p = provider === 'google' ? new authMod.GoogleAuthProvider() : new authMod.OAuthProvider('microsoft.com');
  p.setCustomParameters({ prompt: 'select_account' });
  const r = await authMod.signInWithPopup(ctx.auth, p);
  return r.user;
}

export const signOut = (ctx) => ctx.authMod.signOut(ctx.auth);

/** Adresse e-mail telle que les règles la voient (jeton d'identité). */
async function tokenEmail(ctx) {
  const t = await ctx.user.getIdTokenResult();
  return t.claims.email || '';
}

const ref = (ctx, ...path) => ctx.fs.doc(ctx.db, 'families', ...path);

export async function getAccess(ctx, familyId) {
  try {
    const s = await ctx.fs.getDocFromServer(ref(ctx, familyId, 'access', ctx.user.uid));
    return s.exists() ? s.data() : null;
  } catch (e) {
    if (e.code === 'permission-denied' || e.code === 'not-found') return null;
    throw e;
  }
}

/** Crée un espace familial dont l'utilisateur connecté est propriétaire et administrateur. */
export async function createFamily(ctx, { familyName, memberName, color }) {
  const { fs, user } = ctx;
  const familyId = rid() + rid(); // 32 caractères hexadécimaux
  const memberId = rid();
  const now = new Date().toISOString();
  await fs.setDoc(ref(ctx, familyId), { owner: user.uid, name: familyName, createdAt: now });
  await fs.setDoc(ref(ctx, familyId, 'access', user.uid), {
    uid: user.uid,
    email: await tokenEmail(ctx),
    name: user.displayName || '',
    role: 'admin',
    memberId,
    joinedAt: now,
  });
  await fs.setDoc(ref(ctx, familyId, 'members', memberId), { id: memberId, name: memberName, color, order: 0 });
  return familyId;
}

/** Utilise une invitation (usage unique : elle est supprimée dans la même transaction). */
export async function joinWithInvite(ctx, familyId, inviteId, memberId) {
  const { fs, db, user } = ctx;
  const b = fs.writeBatch(db);
  b.set(ref(ctx, familyId, 'access', user.uid), {
    uid: user.uid,
    email: await tokenEmail(ctx),
    name: user.displayName || '',
    role: 'member',
    memberId,
    invite: inviteId,
    joinedAt: new Date().toISOString(),
  });
  b.delete(ref(ctx, familyId, 'invites', inviteId));
  await b.commit();
}

export async function createInvite(ctx, familyId, memberId) {
  const id = rid() + rid();
  await ctx.fs.setDoc(ref(ctx, familyId, 'invites', id), {
    id,
    memberId,
    createdBy: ctx.user.uid,
    createdAt: new Date().toISOString(),
    expiresAt: ctx.fs.Timestamp.fromDate(addDays(new Date(), INVITE_DAYS)),
  });
  return id;
}
