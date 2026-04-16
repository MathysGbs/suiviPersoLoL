# 🎮 LoL Suivi Comportemental — Tracker Ranked

Outil de tracking personnel pour League of Legends qui récupère automatiquement tes parties ranked via l'API Riot, les stocke en JSON et génère un fichier Excel d'analyse comportementale multi-feuilles.

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

Crée un fichier `.env` à la racine du projet (même niveau que `lol.js`) :

```env
RIOT_API_KEY=RGAPI-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

> **Obtenir une clé API Riot :**
> Rendez-toi sur [developer.riotgames.com](https://developer.riotgames.com), connecte-toi et génère une clé de développement.
> ⚠️ Les clés de développement expirent toutes les **24h** — tu devras la renouveler régulièrement.
> Pour un usage permanent, fais une demande de clé de production sur le même portail.

---

## 👤 Paramètres joueurs

Les noms et tags des deux comptes sont définis dans `src/config.js` :

```js
const MY_NAME  = 'xxxx';       // Ton Riot ID
const MY_TAG   = 'EUW';         // Ton tag

const DUO_NAME = 'xxxx'; // Le Riot ID de ton duo
const DUO_TAG  = 'EUW';         // Le tag de ton duo
```

Modifie ces valeurs si tu changes de compte.

---

## 🚀 Modes d'exécution

### `all` — Flux complet *(par défaut)*

```bash
node lol.js
# ou explicitement :
node lol.js all
```

Enchaîne dans l'ordre :
1. Récupération des nouvelles parties via l'API Riot
2. Migration des anciennes entrées incomplètes
3. Reconstruction du fichier Excel

**Durée estimée :** 10–20 min pour 200 parties (délais API Riot inclus).

---

### `fetch` — Récupération des parties uniquement

```bash
node lol.js fetch
```

Contacte l'API Riot pour télécharger les nouvelles parties depuis la dernière exécution, met à jour le fichier `historique_matches.json`. Ne touche pas à l'Excel.

**Utile quand :** tu veux juste mettre à jour la base de données sans régénérer l'Excel.

---

### `migrate` — Migration des entrées incomplètes

```bash
node lol.js migrate
```

Parcourt le JSON existant et complète les parties qui manquent de champs (rangs, pings, lane diff, ganks subis…). Se produit automatiquement quand le schéma de données a évolué.

**Utile quand :** tu as des parties récupérées avec une ancienne version du programme et tu veux les enrichir avec les nouveaux champs.

---

### `excel` — Reconstruction Excel uniquement

```bash
node lol.js excel
```

Relit `historique_matches.json` et régénère intégralement `Suivi_Comportemental_Challenger.xlsx`. Aucun appel API.

**Utile quand :** tu veux modifier la mise en forme ou les analyses sans retoucher à l'API, ou que tu as déjà un JSON à jour.

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
Toutes les parties, une par ligne, avec mise en forme conditionnelle par groupes de colonnes colorés :

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
Analyses avancées :
- Vue d'ensemble globale et KDA/CS/DPM moyens
- Série actuelle et tendance récente (10 dernières parties)
- **Fatigue cognitive** : performance par partie dans la session (P1 → P4+)
- Performance par tranche horaire (Nuit / Matin / Après-midi / Soir)
- **Indicateur de tilt** : win rate après victoire vs après défaite
- Analyse des pings selon le résultat
- Impact des ganks subis sur le win rate
- Records personnels (meilleur KDA, DPM, CS/Min…)

---

## 🏗️ Architecture du projet

```
lol.js                  ← Point d'entrée (dispatch des modes)
src/
  ├── config.js         ← Constantes, couleurs, fonctions de rang
  ├── utils.js          ← Fonctions utilitaires pures (sleep, avgOf, secToMmSs…)
  ├── analytics.js      ← Calcul des analyses comportementales
  ├── data-service.js   ← Appels API Riot, lecture/écriture JSON
  ├── excel-service.js  ← Construction du fichier Excel
  └── pipeline.js       ← Orchestration des modes (runAll, runFetch…)
```

---

## ⚠️ Limitations connues

- **Rate limit Riot** : le programme gère automatiquement les erreurs 429 et attend le délai indiqué par l'API avant de réessayer.
- **Clé dev Riot** : expire toutes les 24h, pense à la renouveler avant chaque session.
- **Timeline API** : si l'API timeline est indisponible pour un match, les champs `csDiff10/15`, `goldDiff10/15` et `fatalGanksReceived` seront `null` pour cette partie — ce n'est pas bloquant.
- **Serveur** : le programme est configuré pour **EUW** (`euw1` / `europe`). Pour un autre serveur, modifie `REGION` et `REGION_PLATFORM` dans `src/config.js`.