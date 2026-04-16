# LoL Suivi Comportemental

Application personnelle de suivi pour League of Legends. Elle récupère automatiquement les parties ranked via l'API Riot, les stocke dans un fichier JSON local et génère un tableau de bord Excel avec plusieurs feuilles d'analyse. Le projet est utilisable en interface web ou en ligne de commande.

## Prérequis

- Node.js 16 ou supérieur
- Une clé API Riot Games valide
- Deux comptes à suivre sur le serveur configuré dans le projet

## Installation

```bash
npm install
```

## Configuration

Créer un fichier `.env` à la racine du projet :

```env
RIOT_API_KEY=RGAPI-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

La clé Riot de développement expire toutes les 24h. Il faut donc la renouveler régulièrement.

Les pseudos suivis se règlent dans [src/config.js](src/config.js) : `MY_NAME`, `MY_TAG`, `DUO_NAME` et `DUO_TAG`.

Les valeurs importantes par défaut sont aussi définies dans ce fichier :

- `REGION = europe`
- `REGION_PLATFORM = euw1`
- `QUEUE_FILTER = 420` pour les parties ranked solo/duo
- `MATCHES_TO_FETCH = 200`
- `API_DELAY_MS = 1500`

## Lancement

### Interface web

```bash
npm start
```

Le serveur démarre sur `http://localhost:3000` et expose :

- une page principale de contrôle avec les modes `all`, `fetch`, `migrate` et `excel`
- un terminal live via WebSocket pour suivre les logs en temps réel
- des endpoints API pour l'état, les données et l'arrêt du traitement
- le dashboard analytics sur `/dashboard.html`

### Ligne de commande

```bash
node lol.js all
node lol.js fetch
node lol.js migrate
node lol.js excel
```

Scripts npm disponibles :

```bash
npm run all
npm run fetch
npm run migrate
npm run excel
npm run legacy
```

Quand `all` ou `fetch` est lancé dans un terminal interactif, le script demande si tu veux modifier le nombre de parties à récupérer avant de démarrer.

## Interface web

La page principale sert de panneau de contrôle. Elle affiche l'état du traitement, les statistiques locales et permet de lancer chaque mode sans quitter le navigateur.

Le dashboard sur [public/dashboard.html](public/dashboard.html) propose une vue analytique complète avec 13 graphiques et un tableau détaillé des parties récentes. On y retrouve notamment :

- le win rate glissant
- le KDA des dernières parties
- la progression du rang
- les champions les plus joués
- la répartition des rôles
- la fatigue au fil de la session
- les signaux de tilt et de performance après victoire ou défaite
- la performance selon l'heure
- le CS/min et le GPM
- la répartition solo vs duo

Le tableau de détail permet d'explorer les dernières parties avec les informations clés comme le champion, le rôle, le KDA, le CS/min, le DPM, la vision et le rang.

## Fichiers générés

- [historique_matches.json](historique_matches.json) : base locale des parties
- [Suivi_Comportemental_Challenger.xlsx](Suivi_Comportemental_Challenger.xlsx) : export Excel généré par le pipeline

## Structure

```text
lol.js               CLI principale
server.js            Serveur HTTP + WebSocket
public/
  index.html         Page de contrôle
  dashboard.html     Dashboard analytics
src/
  config.js          Configuration et constantes
  utils.js           Utilitaires partagés
  analytics.js       Calculs d'analyse
  data-service.js    Lecture, écriture et enrichissement des données
  excel-service.js   Génération Excel
  pipeline.js        Orchestration des modes
```

## Limites connues

- La clé Riot doit être valide au moment de l'exécution.
- Le projet est configuré pour EUW par défaut. Pour changer de serveur, modifier `REGION` et `REGION_PLATFORM` dans [src/config.js](src/config.js).
- Si la timeline Riot est indisponible, certaines valeurs comme `csDiff` ou `goldDiff` peuvent rester nulles.
- Un seul mode peut tourner à la fois : le serveur applique un verrou pour éviter les exécutions concurrentes.



Prévision de rang avec wr et mmr
mettre les LP en plus du rang
préciser les mutlikills dans le dashboard


