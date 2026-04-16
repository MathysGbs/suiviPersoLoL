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

> La commande installe automatiquement toutes les dépendances déclarées dans `package.json`, y compris `express` et `ws` pour l'interface web.

---

## 🔑 Configuration — fichier `.env`

Crée un fichier `.env` à la racine du projet (même niveau que `lol.js`) :

```env
RIOT_API_KEY=RGAPI-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

> **Obtenir une clé API Riot :**
> Rendez-toi sur [developer.riotgames.com](https://developer.riotgames.com), connecte-toi et génère une clé de développement.
> ⚠️ Les clés de développement expirent toutes les **24h** — tu devras la renouveler avant chaque session.
> Pour un usage permanent, fais une demande de clé de production sur le même portail.

---

## 👤 Paramètres joueurs

Les noms et tags des deux comptes sont définis dans `src/config.js` :

```js
const MY_NAME  = 'xxxxx';       // Ton Riot ID
const MY_TAG   = 'EUW';         // Ton tag

const DUO_NAME = 'xxx'; // Le Riot ID de ton duo
const DUO_TAG  = 'EUW';         // Le tag de ton duo
```

Modifie ces valeurs si tu changes de compte.

---

## 🖥️ Interface Web (recommandé)

Lance le serveur web :

```bash
node server.js
```

Puis ouvre [http://localhost:3000](http://localhost:3000) dans ton navigateur.

L'interface propose :
- **4 boutons** pour déclencher chaque mode en un clic
- **Terminal en temps réel** — les logs s'affichent au fur et à mesure via WebSocket
- **Stats** — nombre de parties enregistrées et date de la dernière
- **Indicateur d'état** — idle / en cours / succès / erreur
- Les boutons se **désactivent automatiquement** pendant qu'un mode tourne pour éviter les conflits

---

## 💻 Ligne de commande

`lol.js` reste entièrement fonctionnel en CLI si tu préfères sans navigateur.

### `all` — Flux complet *(par défaut)*

```bash
node lol.js
# ou explicitement :
node lol.js all
```

Enchaîne dans l'ordre : fetch → migrate → excel.
**Durée estimée :** 10–20 min pour 200 parties (délais API Riot inclus).

### `fetch` — Récupération des parties uniquement

```bash
node lol.js fetch
```

Contacte l'API Riot, met à jour `historique_matches.json`. Ne touche pas à l'Excel.

### `migrate` — Migration des entrées incomplètes

```bash
node lol.js migrate
```

Complète les anciennes entrées avec les champs manquants (rangs, pings, lane diff, ganks…).

### `excel` — Reconstruction Excel uniquement

```bash
node lol.js excel
```

Relit le JSON et régénère l'Excel. Aucun appel API — instantané.

---

## 📂 Fichiers générés

| Fichier | Description |
|---|---|
| `historique_matches.json` | Base de données locale de toutes tes parties |
| `Suivi_Comportemental_Challenger.xlsx` | Tableau de bord Excel final |

---

## 📊 Contenu du fichier Excel

Le fichier contient **4 feuilles** :

### 1. Données Brutes
Toutes les parties avec mise en forme conditionnelle par groupes colorés :

| Groupe | Colonnes |
|---|---|
| 🔵 Identité | Date, Heure, Session, Position/Session, Durée, Champion, Rôle, Type, Résultat |
| 🟣 Combat | KDA, K/D/A, % Mort, KP, % DMG, DMG/Gold, Ratio DMG |
| 🟢 Farm | CS/Min, GPM, DPM, Obj/Min |
| 🩵 Vision | Vision Score, Wards posées/détruites, Wards de contrôle |
| 🌑 Pings | Total, Négatifs, et tous les types détaillés (Danger, MIA, Retreat…) |
| 🟤 Ganks | Ganks mortels subis (early), CSD/GD à 10 et 15 min, Premier objet core |
| 🔴 Rang | Ton rang, rang moyen alliés, rang moyen ennemis, différence |
| ⚫ Divers | Multi-kill, First Blood, Yuumi duo/alliée, ADC adverse |

### 2. Par Champion
Statistiques agrégées par champion, triées par nombre de parties jouées.

### 3. Solo vs Duo
Comparaison de toutes les métriques entre tes parties solo, duo, avec/sans Yuumi.

### 4. Tendances Comportementales
- Vue d'ensemble globale, série actuelle, tendance récente
- **Fatigue cognitive** : performance par partie dans la session (P1 → P4+)
- Performance par tranche horaire (Nuit / Matin / Après-midi / Soir)
- **Indicateur de tilt** : win rate après victoire vs après défaite
- Analyse des pings selon le résultat, impact des ganks, records personnels

---

## 🏗️ Architecture du projet

```
lol.js               ← Point d'entrée CLI (inchangé)
server.js            ← Point d'entrée interface web (Express + WebSocket)
public/
  └── index.html     ← Interface web (HTML/CSS/JS vanilla)
src/
  ├── config.js      ← Constantes, couleurs, fonctions de rang
  ├── utils.js       ← Fonctions utilitaires pures
  ├── analytics.js   ← Calcul des analyses comportementales
  ├── data-service.js  ← Appels API Riot, lecture/écriture JSON
  ├── excel-service.js ← Construction du fichier Excel
  └── pipeline.js    ← Orchestration des modes
```

---

## 📦 Dépendances à ajouter

Si `express` et `ws` ne sont pas encore dans ton `package.json` :

```bash
npm install express ws
```

Vérifie que `package.json` contient bien :

```json
{
  "dependencies": {
    "axios":   "...",
    "dotenv":  "...",
    "exceljs": "...",
    "express": "^4.18.0",
    "ws":      "^8.0.0"
  }
}
```

---

## ⚠️ Limitations connues

- **Rate limit Riot** : géré automatiquement — le programme attend le délai indiqué par l'API (header `Retry-After`) avant de réessayer.
- **Clé dev Riot** : expire toutes les 24h, renouvelle-la avant chaque session.
- **Timeline API** : si indisponible pour un match, les champs `csDiff`, `goldDiff` et `fatalGanksReceived` seront `null` — non bloquant.
- **Serveur** : configuré pour **EUW**. Pour un autre serveur, modifie `REGION` et `REGION_PLATFORM` dans `src/config.js`.
- **Concurrence** : l'interface web empêche le lancement simultané de deux modes.
