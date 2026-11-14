// ─────────────────────────────────────────────────────────────
//  Configuration de l'Agenda Famille
//  Ces identifiants sont PUBLICS par nature (ils finissent dans le
//  navigateur). Ne mettez jamais ici de mot de passe ni de secret.
//  Voir README.md pour savoir comment les obtenir.
// ─────────────────────────────────────────────────────────────
export const CONFIG = {
  appName: 'Agenda Famille',

  // Google Agenda : ID client OAuth de type « Application Web »
  // ex. '1234567890-abc.apps.googleusercontent.com'
  google: {
    clientId: '',
  },

  // Outlook / Microsoft 365 : ID d'application (client) Entra ID, plateforme « SPA »
  microsoft: {
    clientId: '',
    authority: 'https://login.microsoftonline.com/common',
  },

  // Partage entre les membres de la famille (Firebase Firestore).
  // Laissez null pour un mode « local » (données dans ce navigateur uniquement).
  // ex. { apiKey: '...', authDomain: '...', projectId: '...', appId: '...' }
  firebase: null,

  // Proxy CORS optionnel pour lire les liens iCal (.ics) qui bloquent la lecture
  // depuis un navigateur. Le lien est ajouté encodé à la fin.
  // ex. 'https://mon-proxy.mon-compte.workers.dev/?url='
  icsProxy: '',

  // Nombre de jours synchronisés à partir d'aujourd'hui (plus 7 jours passés)
  syncDays: 120,
};
