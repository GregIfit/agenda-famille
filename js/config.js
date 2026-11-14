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
    clientId: '622799626716-7jj86ak62rvu9ofqv62l71b7gjep5ijc.apps.googleusercontent.com',
  },

  // Outlook / Microsoft 365 : ID d'application (client) Entra ID, plateforme « SPA »
  microsoft: {
    clientId: '',
    authority: 'https://login.microsoftonline.com/common',
  },

  // Partage sécurisé entre les membres de la famille (Firebase : connexion + Firestore).
  // Laissez null pour un mode « local » (données dans ce navigateur uniquement, sans connexion).
  // ex. { apiKey: '...', authDomain: '...', projectId: '...', appId: '...' }
  // N'oubliez pas de publier les règles de firestore.rules (voir README).
  firebase: {
  apiKey: "AIzaSyDpOUN0gLQs9iN9u2gJMW411HVkQm_BWcE",
  authDomain: "agenda-famille-5101d.firebaseapp.com",
  projectId: "agenda-famille-5101d",
  storageBucket: "agenda-famille-5101d.firebasestorage.app",
  messagingSenderId: "622799626716",
  appId: "1:622799626716:web:1211618af4aba19b8d6f19"
},

  // Proxy CORS optionnel pour lire les liens iCal (.ics) qui bloquent la lecture
  // depuis un navigateur. Le lien est ajouté encodé à la fin.
  // ex. 'https://mon-proxy.mon-compte.workers.dev/?url='
  icsProxy: '',

  // Nombre de jours synchronisés à partir d'aujourd'hui (plus 7 jours passés)
  syncDays: 120,
};
