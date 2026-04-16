# 🎮 LoL Suivi Comportemental — Tracker Ranked

Outil de tracking personnel pour League of Legends qui récupère automatiquement tes parties ranked via l'API Riot, les stocke en JSON et génère un fichier Excel d'analyse comportementale multi-feuilles.

Utilisable en **interface web** (`server.js`) ou en **ligne de commande** (`lol.js`).

---

## 📋 Prérequis

- **Node.js** v16 ou supérieur
- Un compte Riot Games avec un **API Key** (voir ci-dessous)
- Les deux comptes à tracker doivent exister sur le serveur EUW

---

## ⚙️ Installation

```bash
npm install
```

---

## 🔑 Configuration — fichier `.env`

Crée un fichier `.env` à la racine du projet :

```env
RIOT_API_KEY=RGAPI-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

> **Obtenir une clé API Riot :**
> [developer.riotgames.com](https://developer.riotgames.com) → connecte-toi → génère une clé de développement.
> ⚠️ Les clés de développement expirent toutes les **24h**.

---

## 👤 Paramètres joueurs (`src/config.js`)

```js
const MY_NAME  = 'xxxxx';
const MY_TAG   = 'EUW';
const DUO_NAME = 'xxxx';
const DUO_TAG  = 'EUW';
```

---

## 🖥️ Interface Web (recommandé)

```bash
node server.js
# → http://localhost:3000
```

### Page principale `/`
- 4 boutons pour déclencher chaque mode
- Terminal live (logs en temps réel via WebSocket)
- Stats : total parties, date de la dernière
- Indicateur d'état et verrou anti-concurrence

### Dashboard `/dashboard.html`
Visualisation complète des données avec **10 graphiques** :

| Graphique | Description |
|---|---|
| Win Rate glissant | Courbe du WR sur fenêtre de 10 parties + WR global |
| KDA par partie | Barres colorées victoire/défaite sur les 50 dernières |
| Progression du rang | Évolution du rank score dans le temps |
| Top Champions | Win rate + nombre de parties par champion (horizontal) |
| Distribution des rôles | Donut par rôle joué |
| Fatigue cognitive | Win rate par position dans la session (P1 → P4+) |
| Indicateur de Tilt | Comparaison WR/KDA/CS/DPM après victoire vs défaite |
| Performance horaire | Win rate par tranche (Nuit / Matin / Après-midi / Soir) |
| CS/Min & GPM | Double axe, évolution sur 50 parties |
| Solo vs Duo | Donut + win rate par contexte Yuumi |

Et un tableau des **50 dernières parties** avec : champion, rôle, KDA, K/D/A, CS/min, DPM, vision, rang.

---

## 💻 Ligne de commande

`lol.js` reste entièrement fonctionnel sans navigateur.

```bash
node lol.js          # Flux complet (fetch + migrate + excel)
node lol.js fetch    # Récupère les nouvelles parties uniquement
node lol.js migrate  # Complète les entrées incomplètes
node lol.js excel    # Régénère l'Excel (sans appel API)
```

Quand tu lances `fetch` ou `all` dans un terminal interactif, le script te demande si tu veux modifier le nombre de games avant de lancer la récupération.

---

## 📂 Fichiers générés

| Fichier | Description |
|---|---|
| `historique_matches.json` | Base de données locale |
| `Suivi_Comportemental_Challenger.xlsx` | Tableau de bord Excel (4 feuilles) |

---

## 📊 Contenu du fichier Excel

**4 feuilles :** Données Brutes · Par Champion · Solo vs Duo · Tendances Comportementales

Les Données Brutes incluent 8 groupes colorés : Identité, Combat (KDA/DMG/KP), Farm, Vision, Pings, Ganks/Lane diff, Rang, Divers.

Les Tendances incluent : vue d'ensemble, fatigue cognitive, performance horaire, indicateur de tilt, analyse des pings, impact des ganks, records personnels.

---

## 🏗️ Architecture

```
lol.js               ← CLI (inchangé)
server.js            ← Serveur web Express + WebSocket + route /api/data
public/
  ├── index.html     ← Page de contrôle (4 modes + terminal live)
  └── dashboard.html ← Dashboard analytics (10 graphiques Chart.js)
src/
  ├── config.js
  ├── utils.js
  ├── analytics.js
  ├── data-service.js
  ├── excel-service.js
  └── pipeline.js
```

---

## 📦 Dépendances

```bash
npm install express ws
```

`package.json` doit contenir : `axios`, `dotenv`, `exceljs`, `express`, `ws`.

---

## ⚠️ Limitations

- **Rate limit Riot** : géré automatiquement via le header `Retry-After`.
- **Clé dev Riot** : expire toutes les 24h.
- **Timeline API** : si indisponible, `csDiff` / `goldDiff` / `fatalGanksReceived` seront `null`.
- **Serveur** : configuré pour EUW. Modifie `REGION` et `REGION_PLATFORM` dans `src/config.js` pour un autre serveur.
- **Concurrence** : un seul mode peut tourner à la fois (verrou côté serveur + désactivation UI).




AJOUTER PROGRESSION
STOP EN COURS
