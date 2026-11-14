# Agenda Famille

Une web app d'agenda familial partagé :

- **Agenda familial** (vue mensuelle + détail du jour) : événements, participants, lieu, notes.
- **Synchronisation avec les agendas personnels** de chaque membre :
  - **Google Agenda** et **Outlook / Microsoft 365** : lecture des événements, et copie d'un événement familial dans son agenda perso en un clic.
  - **Liens iCal (.ics / webcal://)** : pour iCloud, Proton, l'agenda de l'école ou du club, etc. Ces liens sont en lecture seule, récurrences comprises.
  - **Vie privée** : les agendas perso restent sur l'appareil de chacun. La famille ne voit que des **créneaux occupés**, sans titre ni détail, et le partage peut être désactivé.
- **Sondages de dates** :
  - proposition de dates à la main, ou **suggestion automatique de créneaux où tout le monde est libre** (soirée, midi, journée, week-ends) ;
  - réponses Oui / Peut-être / Non, avec un ⚠ quand quelqu'un est occupé d'après son agenda ;
  - « Répondre d'après mon agenda » remplit vos réponses automatiquement ;
  - « Retenir cette date » crée l'événement dans l'agenda familial et clôt le sondage.
- Export de l'agenda familial en `.ics`, plus une sauvegarde et une restauration en JSON.
- Interface en français, adaptée au mobile, avec thème clair et sombre.

C'est une application **100 % statique** (HTML/CSS/JS, sans étape de build). Elle s'héberge telle quelle sur **GitHub Pages**.

---

## 1. Essayer en local

```bash
python -m http.server 8080
```

Ouvrez ensuite <http://localhost:8080>. Sans configuration, l'app tourne en **mode local** : les données restent dans le navigateur. C'est pratique pour tester.

> Il faut passer par un serveur, pas ouvrir `index.html` en double-cliquant : les modules JavaScript et les connexions OAuth en ont besoin.

## 2. Configuration (`js/config.js`)

Tout se règle dans `js/config.js`. Les identifiants qu'on y met sont **publics par nature** : ils sont visibles dans le navigateur. On n'y met donc jamais de mot de passe.

| Clé | Rôle | Obligatoire ? |
|---|---|---|
| `firebase` | Partage des données entre tous les appareils de la famille | Oui, pour un usage familial réel |
| `google.clientId` | Connexion Google Agenda | Si quelqu'un utilise Google |
| `microsoft.clientId` | Connexion Outlook / Microsoft 365 | Si quelqu'un utilise Outlook |
| `icsProxy` | Lire les liens iCal qui bloquent les navigateurs (CORS) | Souvent oui pour iCal |

### 2.1 Partage familial sécurisé : Firebase (gratuit)

GitHub Pages ne sert que des fichiers : il faut une base de données pour partager l'agenda, et une connexion pour savoir qui y accède. Firebase (offre gratuite « Spark ») fournit les deux et suffit largement pour une famille.

1. Allez sur <https://console.firebase.google.com> › **Ajouter un projet** (Google Analytics n'est pas nécessaire).
2. **Build › Firestore Database** › Créer une base, en **mode production**, région `eur3` (Europe).
3. **Build › Authentication** › Commencer › onglet *Sign-in method* :
   - activez **Google** (il suffit de choisir une adresse d'assistance) ;
   - activez **Microsoft** si certains membres n'ont pas de compte Google : voir la section 2.3, partie « Connexion à l'app ».
   - **N'activez pas** « Anonyme ».
4. **Authentication › Settings › Domaines autorisés** : ajoutez `VOTRE-PSEUDO.github.io`.
5. **Firestore › Règles** : remplacez tout le contenu par celui du fichier **`firestore.rules`** fourni, puis cliquez sur *Publier*. **Cette étape est indispensable** : ce sont ces règles qui protègent vos données.
6. **Paramètres du projet** (roue dentée) › *Vos applications* › icône **Web `</>`** › copiez l'objet de configuration dans `js/config.js` :

   ```js
   firebase: {
     apiKey: 'AIza…',
     authDomain: 'mon-agenda.firebaseapp.com',
     projectId: 'mon-agenda',
     appId: '1:…:web:…',
   },
   ```

**Premier lancement**
1. Ouvrez l'app et connectez-vous avec votre compte Google ou Microsoft.
2. Cliquez sur **Créer l'espace familial** : vous en devenez l'administrateur.
3. Dans **Famille**, ajoutez le profil de chaque membre (prénom, couleur), puis cliquez sur **Inviter** à côté de chacun.
4. Envoyez à chaque personne **son** lien, en privé. Il est utilisable une seule fois et expire au bout de 7 jours.
5. La personne ouvre le lien et se connecte avec son compte : celui-ci est alors autorisé et relié à son profil.

### Comment l'accès est protégé

| Qui | Ce qu'il peut faire |
|---|---|
| Quelqu'un sans compte | Voir l'écran de connexion. Rien d'autre. |
| Un compte non invité | Voir « Ce compte n'a pas accès ». Aucune donnée. |
| Un membre | Voir et modifier l'agenda et les sondages, répondre **pour lui seul**, publier **ses** créneaux occupés. |
| Un administrateur | Tout ce que fait un membre, plus : inviter, retirer un accès, nommer d'autres administrateurs, supprimer un profil. |

- La protection est assurée **par le serveur** (les règles Firestore), pas par l'application : même en contournant l'interface, une personne non autorisée ne peut rien lire ni écrire. Ces règles sont couvertes par 36 tests automatisés, lancés dans l'émulateur Firebase.
- **Lien d'invitation intercepté ?** Il ne sert qu'une fois. Si quelqu'un d'autre l'utilise avant la bonne personne, son adresse apparaît dans **Famille › Accès & sécurité** : retirez-lui l'accès en un clic et réinvitez.
- **Retrait d'accès** : l'effet est immédiat sur tous ses appareils.
- Les agendas perso ne quittent jamais l'appareil de chacun ; seuls les créneaux « occupé » sont partagés.
- Le propriétaire de l'espace ne peut pas perdre son accès : il peut toujours se réinscrire.

### 2.2 Google Agenda

1. <https://console.cloud.google.com> › créez un projet.
2. **API et services › Bibliothèque** › activez **Google Calendar API**.
3. **Écran de consentement OAuth** : type *Externe*, avec le nom de l'app et votre e-mail. Ajoutez le champ d'application `.../auth/calendar.events`.
   Laissez l'app en mode **Test** et ajoutez l'adresse Gmail de **chaque membre de la famille** dans *Utilisateurs test* (100 au maximum). Pas besoin de vérification Google.
4. **Identifiants › Créer › ID client OAuth › Application Web**.
   *Origines JavaScript autorisées* : `https://VOTRE-PSEUDO.github.io` et `http://localhost:8080`.
5. Copiez l'ID client dans `google.clientId`.

> En mode Test, Google affiche un avertissement « application non vérifiée » : cliquez sur *Continuer*. La connexion Google dure environ 1 heure. Ensuite, un clic sur « ⟳ Synchro » la renouvelle.

### 2.3 Outlook / Microsoft 365 (comptes familiaux inclus)

1. <https://entra.microsoft.com> (ou portail Azure) › **Inscriptions d'applications › Nouvelle inscription**.
2. Types de comptes : **« Comptes dans un annuaire organisationnel et comptes Microsoft personnels »**. Cela couvre les comptes Microsoft 365 Famille (outlook.com, hotmail…) comme les comptes pro.
3. URI de redirection : plateforme **Application monopage (SPA)**, avec `https://VOTRE-PSEUDO.github.io/NOM-DU-DEPOT/` et `http://localhost:8080/`. Le `/` final doit correspondre exactement à l'adresse de l'app.
4. **Autorisations d'API** › Microsoft Graph › Déléguées › `Calendars.ReadWrite`.
5. Copiez l'**ID d'application (client)** dans `microsoft.clientId`.

**Connexion à l'app avec un compte Microsoft (Firebase).** Cette étape est facultative si tout le monde a un compte Google. On réutilise la même inscription d'application :

1. Dans Firebase › Authentication › Sign-in method › **Microsoft**, copiez l'*URL de redirection* affichée (`https://VOTRE-PROJET.firebaseapp.com/__/auth/handler`).
2. Dans Entra › votre application › **Authentification** › *Ajouter une plateforme* › **Web**, collez cette URL.
3. Allez dans **Certificats & secrets › Nouveau secret client** et copiez sa *valeur*.
4. Dans Firebase › Microsoft, collez l'**ID d'application** et ce **secret**, puis activez.

> Ce secret ne va **que** dans la console Firebase, jamais dans `config.js` ni dans le dépôt GitHub. Il expire au bout de 6 à 24 mois : notez la date pour le renouveler.

### 2.4 Liens iCal et proxy CORS

Beaucoup de serveurs d'agenda (Google, iCloud, Outlook.com…) refusent qu'une page web lise leurs `.ics` directement. Il faut alors un petit relais. Évitez les proxys publics : l'adresse iCal secrète y transiterait. Le plus simple est un **Cloudflare Worker** gratuit à vous :

```js
// Cloudflare › Workers › Créer › collez ce code › Déployer
export default {
  async fetch(req) {
    const url = new URL(req.url).searchParams.get('url');
    if (!url || !/^https:\/\//.test(url)) return new Response('url manquante', { status: 400 });
    const r = await fetch(url, { headers: { 'User-Agent': 'AgendaFamille' } });
    return new Response(r.body, {
      status: r.status,
      headers: { 'Content-Type': 'text/calendar; charset=utf-8', 'Access-Control-Allow-Origin': 'https://VOTRE-PSEUDO.github.io' },
    });
  },
};
```

Puis, dans `js/config.js` : `icsProxy: 'https://mon-proxy.VOTRE-COMPTE.workers.dev/?url='`.

## 3. Mise en ligne sur GitHub Pages

1. Créez un dépôt sur GitHub (ex. `agenda-famille`). Avec un compte gratuit, il doit être **public**. Ce n'est pas un problème : le code ne contient aucun secret, et les données sont dans Firebase.
2. Envoyez tous les fichiers de ce dossier (*Add file › Upload files*, ou `git push`).
3. **Settings › Pages** › *Deploy from a branch* › `main` / `(root)` › Save.
4. Quelques minutes plus tard, l'app est disponible sur `https://VOTRE-PSEUDO.github.io/agenda-famille/`.
5. Ajoutez cette adresse dans Firebase (domaines autorisés), dans Google (origines) et dans Microsoft (URI de redirection).

Sur téléphone : ouvrez l'adresse puis faites *Ajouter à l'écran d'accueil*.

### Et un hébergement sur Microsoft 365 / OneDrive ?

OneDrive et SharePoint ne servent pas de site web : un fichier `.html` y est téléchargé, pas exécuté. GitHub Pages est donc la bonne option pour l'app. Microsoft 365 reste pleinement utilisé, via la connexion Outlook de chaque membre.

## 4. Structure du code

```
index.html          page unique
css/styles.css      styles (clair/sombre, responsive)
firestore.rules     règles de sécurité à publier dans Firebase
js/config.js        votre configuration
js/app.js           interface : agenda, famille, sondages, synchro, accès
js/cloud.js         connexion Google/Microsoft, espaces familiaux, invitations
js/store.js         stockage : local (localStorage) ou Firebase Firestore temps réel
js/utils.js         dates, export .ics, utilitaires
js/sync/google.js   Google Calendar API
js/sync/outlook.js  Microsoft Graph (MSAL.js)
js/sync/ics.js      lecture iCal (ical.js, récurrences comprises)
```

## 5. Limites actuelles et pistes

- En mode **local** (sans Firebase), il n'y a pas de connexion : chacun choisit « qui il est ». Ce mode est réservé aux essais sur un seul appareil.
- Les membres peuvent modifier tous les événements familiaux. C'est voulu, pour qu'un parent puisse corriger l'événement d'un enfant.
- Les événements familiaux sont **copiés** dans l'agenda perso à la demande. Une modification ultérieure se recopie avec le même bouton, mais une suppression ne se propage pas.
- Seul l'agenda **principal** Google/Outlook est lu. Les autres agendas peuvent être ajoutés via leur lien iCal.
- La synchro se fait à l'ouverture de l'app et à la demande, pas en arrière-plan.
- Pistes : notifications et rappels, événements récurrents familiaux, PWA hors-ligne complète, choix des sous-agendas.
