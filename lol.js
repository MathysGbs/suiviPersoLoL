const axios  = require('axios');
const ExcelJS = require('exceljs');
const fs      = require('fs');
require('dotenv').config();

// ============================================================
// CONFIGURATION
// ============================================================
const RIOT_API_KEY     = process.env.RIOT_API_KEY;
const REGION           = 'europe';
const REGION_PLATFORM  = 'euw1'; // Platform pour les appels ranked

const MY_NAME          = 'enyαV';
const MY_TAG           = 'EUW';
const DUO_NAME         = 'Ahristocats';
const DUO_TAG          = 'EUW';

const EXCEL_FILENAME   = 'Suivi_Comportemental_Challenger.xlsx';
const JSON_FILENAME    = 'historique_matches.json';
const MATCHES_TO_FETCH = 200;
const QUEUE_FILTER     = 420;   // Ranked Solo/Duo uniquement. null = toutes files.
const API_DELAY_MS     = 1500;

// ============================================================
// PALETTE COULEURS (ARGB)
// ============================================================
const C = {
    // Headers par feuille
    hRaw:    'FF0F172A',
    hChamp:  'FF2E1065',
    hDuo:    'FF172554',
    hTend:   'FF052E16',
    subH:    'FF1E3A5F',
    white:   'FFFFFFFF',
    accent:  'FF4F46E5',   // indigo

    // Résultat
    winFg:   'FF166534', winBg:   'FFDCFCE7',
    lossFg:  'FF991B1B', lossBg:  'FFFEE2E2',

    // Alternance de lignes
    rowA:    'FFF9FAFB',
    rowB:    'FFFFFFFF',

    // Métriques par seuil
    good:    'FFD9F2DC',
    warn:    'FFFFF9C4',
    bad:     'FFFDE8E8',
    goodFg:  'FF166534',
    badFg:   'FF991B1B',
    purple:  'FFA855F7',
    amber:   'FFB45309',

    // Groupes colonnes Données Brutes
    grpIdentite: 'FF1E3A5F',   // bleu marine  — Date/Heure/Session/Rôle
    grpCombat:   'FF3B0764',   // violet foncé  — KDA/K/D/A
    grpFarm:     'FF052E16',   // vert foncé    — CS/GPM/DPM/ObjDPM
    grpVision:   'FF1C3035',   // teal foncé    — Vision/Wards
    grpDivers:   'FF1C1917',   // gris très foncé — MultiKill/FirstBlood/Yuumi
    grpRank:     'FF431407',   // brun foncé    — Rangs
    grpPings:    'FF1A1A2E',   // bleu nuit     — Pings
    grpGanks:    'FF2D1B4E',   // violet moyen  — Ganks

    // Rangs (couleurs Riot-like)
    rankIron:        'FF8B8B8B',
    rankBronze:      'FFB87333',
    rankSilver:      'FFA8A9AD',
    rankGold:        'FFFFD700',
    rankPlatinum:    'FF0BC4B5',
    rankEmerald:     'FF50C878',
    rankDiamond:     'FFB9F2FF',
    rankMaster:      'FF9B59B6',
    rankGrandmaster: 'FFE74C3C',
    rankChallenger:  'FFF0E68C',
};

// ============================================================
// CONSTANTES RANG
// ============================================================
const TIER_ORDER = {
    'IRON': 0, 'BRONZE': 4, 'SILVER': 8, 'GOLD': 12,
    'PLATINUM': 16, 'EMERALD': 20, 'DIAMOND': 24,
    'MASTER': 28, 'GRANDMASTER': 29, 'CHALLENGER': 30,
};
const DIV_ORDER = { 'IV': 0, 'III': 1, 'II': 2, 'I': 3 };

function rankToScore(tier, division) {
    if (!tier) return -1;
    const t = TIER_ORDER[tier.toUpperCase()] ?? -1;
    if (t === -1) return -1;
    // Master+ n'a pas de division
    if (t >= 28) return t;
    return t + (DIV_ORDER[division] ?? 0);
}

function scoreToLabel(score) {
    if (score < 0)  return 'Non classé';
    if (score >= 30) return 'Challenger';
    if (score >= 29) return 'GrandMaster';
    if (score >= 28) return 'Master';
    const tiers = ['Iron','Bronze','Silver','Gold','Platinum','Emerald','Diamond'];
    const divs  = ['IV','III','II','I'];
    const ti    = Math.floor(score / 4);
    const di    = score % 4;
    return tiers[ti] + ' ' + divs[di];
}

function rankColor(score) {
    if (score < 0)   return C.rankIron;
    if (score < 4)   return C.rankIron;
    if (score < 8)   return C.rankBronze;
    if (score < 12)  return C.rankSilver;
    if (score < 16)  return C.rankGold;
    if (score < 20)  return C.rankPlatinum;
    if (score < 24)  return C.rankEmerald;
    if (score < 28)  return C.rankDiamond;
    if (score < 29)  return C.rankMaster;
    if (score < 30)  return C.rankGrandmaster;
    return C.rankChallenger;
}

// ============================================================
// UTILITAIRES
// ============================================================
const sleep  = ms  => new Promise(r => setTimeout(r, ms));
const safeN  = (v, d = 0)   => { const n = parseFloat(v); return isNaN(n) ? d : n; };
const safeV  = (v, d = '-') => (v !== undefined && v !== null) ? v : d;
const avgOf  = (arr, k)     => arr.length ? arr.reduce((s, x) => s + safeN(x[k]), 0) / arr.length : 0;
const winPct = arr => arr.length ? (arr.filter(g => g.win === 'Victoire').length / arr.length) : 0;
const rowBg  = idx => idx % 2 === 0 ? C.rowA : C.rowB;

// Convertit des secondes en mm:ss
function secToMmSs(totalSeconds) {
    if (totalSeconds == null || isNaN(totalSeconds)) return '-';
    const s   = Math.round(totalSeconds);
    const m   = Math.floor(s / 60);
    const sec = s % 60;
    return m + ':' + String(sec).padStart(2, '0');
}

// Convertit des minutes décimales en mm:ss (backward-compat)
function minToMmSs(minutes) {
    return secToMmSs(Math.round(minutes * 60));
}

function pushTo(map, key, value) {
    if (!map[key]) map[key] = [];
    map[key].push(value);
}

function bestMultiKill(me) {
    if ((me.pentaKills  || 0) > 0) return 'Penta';
    if ((me.quadraKills || 0) > 0) return 'Quadra';
    if ((me.tripleKills || 0) > 0) return 'Triple';
    if ((me.doubleKills || 0) > 0) return 'Double';
    return '-';
}

function recalculateTimeline(data) {
    data.sort((a, b) => new Date(a.rawDate) - new Date(b.rawDate));
    let currentSessionId = 1;
    let gameInSession    = 1;
    for (let i = 0; i < data.length; i++) {
        if (i > 0) {
            const diffH = (new Date(data[i].rawDate) - new Date(data[i-1].rawDate)) / 3_600_000;
            if (diffH < 2) { gameInSession++; }
            else           { currentSessionId++; gameInSession = 1; }
        }
        data[i].sessionId     = currentSessionId;
        data[i].gameInSession = gameInSession;
    }
    return data;
}

// ============================================================
// CACHE RANG (summonerId → {tier, division, score, label})
// ============================================================
const rankCache = {};

async function fetchRankForSummoner(summonerId) {
    if (rankCache[summonerId] !== undefined) return rankCache[summonerId];
    
    try {
        const r = await axios.get(
            `https://${REGION_PLATFORM}.api.riotgames.com/lol/league/v4/entries/by-summoner/${summonerId}`,
            { headers: { 'X-Riot-Token': RIOT_API_KEY } }
        );
        
        const soloQ = r.data.find(e => e.queueType === 'RANKED_SOLO_5x5');
        if (!soloQ) {
            rankCache[summonerId] = { tier: null, division: null, score: -1, label: 'Non classé' };
        } else {
            const score = rankToScore(soloQ.tier, soloQ.rank);
            rankCache[summonerId] = {
                tier:     soloQ.tier,
                division: soloQ.rank,
                score,
                label:    scoreToLabel(score),
            };
        }
    } catch (e) {
        // Si l'API nous bloque (Rate Limit), on met en pause et on réessaie
        if (e.response && e.response.status === 429) {
            // Riot fournit un header "Retry-After" indiquant le temps d'attente requis en secondes.
            const retryAfter = e.response.headers['retry-after'] 
                ? parseInt(e.response.headers['retry-after']) * 1000 
                : 10000; // 10 secondes par défaut
            
            console.log(`\n      [!] Limite API (Rangs) atteinte. Attente de ${retryAfter / 1000}s...`);
            await sleep(retryAfter + 500); // Marge de sécurité
            return fetchRankForSummoner(summonerId); // Appel récursif après l'attente
        }
        
        // Pour toute autre erreur (ex: 404 joueur introuvable), on le marque Non classé
        rankCache[summonerId] = { tier: null, division: null, score: -1, label: 'Non classé' };
    }
    return rankCache[summonerId];
}

// ============================================================
// ANALYSE COMPORTEMENTALE
// ============================================================
function computeAnalytics(data) {
    if (!data.length) return null;

    const champMap        = {};
    const byGameInSession = {};
    const byRole          = {};
    const timeSlots       = {
        'Nuit (0h-6h)':         [],
        'Matin (6h-12h)':       [],
        'Apres-midi (12h-18h)': [],
        'Soir (18h-24h)':       [],
    };

    data.forEach(m => {
        pushTo(champMap, m.champion, m);
        const g = Math.min(safeN(m.gameInSession, 1), 4);
        pushTo(byGameInSession, g, m);
        if (m.role) pushTo(byRole, m.role, m);
        const h = new Date(m.rawDate).getHours();
        if      (h < 6)  timeSlots['Nuit (0h-6h)'].push(m);
        else if (h < 12) timeSlots['Matin (6h-12h)'].push(m);
        else if (h < 18) timeSlots['Apres-midi (12h-18h)'].push(m);
        else             timeSlots['Soir (18h-24h)'].push(m);
    });

    const byChampion = Object.entries(champMap)
        .map(([name, games]) => ({ name, games }))
        .sort((a, b) => b.games.length - a.games.length);

    const afterWin = [], afterLoss = [];
    for (let i = 1; i < data.length; i++) {
        (data[i-1].win === 'Victoire' ? afterWin : afterLoss).push(data[i]);
    }

    let streakCount = 0, streakType = '';
    for (let i = data.length - 1; i >= 0; i--) {
        const isW = data[i].win === 'Victoire';
        if (i === data.length - 1) { streakType = isW ? 'Victoires' : 'Defaites'; streakCount = 1; }
        else if ((isW && streakType === 'Victoires') || (!isW && streakType === 'Defaites')) { streakCount++; }
        else break;
    }

    const valid  = data.filter(m => m.kda != null);
    const topOf  = (arr, key) => arr.filter(m => m[key]).sort((a, b) => safeN(b[key]) - safeN(a[key]))[0];
    const records = {
        bestKDA:    topOf(valid, 'kda'),
        bestDPM:    topOf(valid, 'dpm'),
        bestCS:     topOf(valid, 'csPerMin'),
        bestGPM:    topOf(valid, 'gpm'),
        bestVision: topOf(valid, 'vision'),
    };

    // Analyse ganks (jungler uniquement)
    const gankData = data.filter(m => m.role === 'JUNGLE' && m.ganksPerformed != null);
    const byGankCount = {};
    gankData.forEach(m => {
        const bucket = Math.min(m.ganksPerformed || 0, 10);
        pushTo(byGankCount, bucket, m);
    });

    return {
        byChampion, byRole,
        soloGames:   data.filter(m => m.type === 'Solo'),
        duoGames:    data.filter(m => m.type === 'Duo'),
        yuumiGames:  data.filter(m => m.withYuumi === 'Oui'),
        noYuumi:     data.filter(m => m.type === 'Duo' && m.withYuumi !== 'Oui'),
        yuumiAlliee: data.filter(m => m.yuumiAlliee === 'Oui'),
        byGameInSession, byGankCount,
        timeSlots,
        afterWin, afterLoss,
        streak:   { count: streakCount, type: streakType },
        recent10: data.slice(-10),
        records,
        overall: {
            total:       data.length,
            winRate:     winPct(data),
            avgKDA:      avgOf(data, 'kda'),
            avgKills:    avgOf(data, 'kills'),
            avgDeaths:   avgOf(data, 'deaths'),
            avgAssists:  avgOf(data, 'assists'),
            avgCS:       avgOf(data, 'csPerMin'),
            avgDPM:      avgOf(data, 'dpm'),
            avgGPM:      avgOf(data, 'gpm'),
            avgDmgShare: avgOf(data, 'dmgShare'),
            avgKP:       avgOf(data, 'kp'),
            avgVision:   avgOf(data, 'vision'),
        },
    };
}

// ============================================================
// API RIOT
// ============================================================
async function getPlayerData(name, tag) {
    const r = await axios.get(
        `https://${REGION}.api.riotgames.com/riot/account/v1/accounts/by-riot-id/${encodeURIComponent(name)}/${encodeURIComponent(tag)}`,
        { headers: { 'X-Riot-Token': RIOT_API_KEY } }
    );
    return r.data;
}

async function getMatchIds(puuid) {
    let allIds = [];
    let start  = 0;
    const limit = 100;
    while (allIds.length < MATCHES_TO_FETCH) {
        const count = Math.min(limit, MATCHES_TO_FETCH - allIds.length);
        let url = `https://${REGION}.api.riotgames.com/lol/match/v5/matches/by-puuid/${puuid}/ids?start=${start}&count=${count}`;
        if (QUEUE_FILTER) url += `&queue=${QUEUE_FILTER}`;
        const r   = await axios.get(url, { headers: { 'X-Riot-Token': RIOT_API_KEY } });
        const ids = r.data;
        if (ids.length === 0) break;
        allIds = allIds.concat(ids);
        start += count;
        if (allIds.length < MATCHES_TO_FETCH) await sleep(API_DELAY_MS);
    }
    return allIds;
}

// ============================================================
// API TIMELINE (Pour l'analyse de lane)
// ============================================================
async function getMatchTimeline(matchId) {
    const r = await axios.get(
        `https://${REGION}.api.riotgames.com/lol/match/v5/matches/${matchId}/timeline`,
        { headers: { 'X-Riot-Token': RIOT_API_KEY } }
    );
    return r.data;
}


// ============================================================
// EXTRACTION MÉTRIQUES MATCH (v2 — avec rangs, rôle, pings, ganks)
// ============================================================
async function extractMatchMetrics(matchId, myPuuid, duoPuuid) {
    const r = await axios.get(
        `https://${REGION}.api.riotgames.com/lol/match/v5/matches/${matchId}`,
        { headers: { 'X-Riot-Token': RIOT_API_KEY } }
    );
    const match = r.data;
    const me    = match.info.participants.find(p => p.puuid === myPuuid);
    if (!me) throw new Error(`PUUID introuvable dans ${matchId}`);

    const duo      = match.info.participants.find(p => p.puuid === duoPuuid && p.teamId === me.teamId);
    const isDuo    = !!duo;
    const myTeam   = match.info.participants.filter(p => p.teamId  === me.teamId);
    const enemies  = match.info.participants.filter(p => p.teamId !== me.teamId);

    const minutes    = match.info.gameDuration / 60;
    const teamKills  = myTeam.reduce((s, p) => s + p.kills, 0);
    const teamDamage = myTeam.reduce((s, p) => s + p.totalDamageDealtToChampions, 0);
    const kp         = teamKills  === 0 ? 0 : ((me.kills + me.assists) / teamKills)  * 100;
    const dmgShare   = teamDamage === 0 ? 0 : (me.totalDamageDealtToChampions / teamDamage) * 100;
    const totalCS    = (me.totalMinionsKilled || 0) + (me.neutralMinionsKilled || 0);
    const gameDate   = new Date(match.info.gameCreation);

    // ── Pings ─────────────────────────────────────────────────
    const allInPings        = me.allInPings            || 0;
    const basicPings        = me.basicPings            || 0;
    const commandPings      = me.commandPings          || 0;
    const dangerPings       = me.dangerPings           || 0;
    const enemyMissingPings = me.enemyMissingPings     || 0;
    const holdPings         = me.holdPings             || 0;
    const needVisionPings   = me.needVisionPings       || 0;
    const onMyWayPings      = me.onMyWayPings          || 0;
    const pushPings         = me.pushPings             || 0;
    const retreatPings      = me.retreatPings          || 0;
    const visionClearedPings = me.visionClearedPings   || 0;
    const totalPings        = allInPings + basicPings + commandPings + dangerPings
                            + enemyMissingPings + holdPings + needVisionPings
                            + onMyWayPings + pushPings + retreatPings + visionClearedPings;
    // Pings "négatifs" = danger + retreat + enemyMissing (tilt/stress)
    const negativePings     = dangerPings + retreatPings + enemyMissingPings;

// ── Ganks Subis (Laner - via Timeline) ────────────────────
    console.log('   -> Analyse de la Timeline (Ganks subis)...');
    let fatalGanksReceived = 0;
    try {
        const timeline = await getMatchTimeline(matchId);
        // Identification du jungler ennemi
        const enemyJungler = match.info.participants.find(p => p.teamId !== me.teamId && p.teamPosition === 'JUNGLE');
        const enemyJunglerId = enemyJungler ? enemyJungler.participantId : null;
        
        if (enemyJunglerId) {
            for (const frame of timeline.info.frames) {
                if (frame.timestamp > 14 * 60 * 1000) break; // Phase de lane (14 min)
                for (const event of frame.events) {
                    if (event.type === 'CHAMPION_KILL' && event.victimId === me.participantId) {
                        // Le jungler ennemi est-il impliqué (kill ou assist) ?
                        if (event.killerId === enemyJunglerId || (event.assistingParticipantIds && event.assistingParticipantIds.includes(enemyJunglerId))) {
                            fatalGanksReceived++;
                        }
                    }
                }
            }
        }
    } catch (e) {
        console.log('      (Avertissement : Timeline indisponible)');
        fatalGanksReceived = null;
    }

    // ── Rangs ─────────────────────────────────────────────────
    // On récupère les summonerId depuis les participants
    const allySummonerIds  = myTeam.map(p => p.summonerId).filter(Boolean);
    const enemySummonerIds = enemies.map(p => p.summonerId).filter(Boolean);

    console.log('   -> Récupération rangs alliés...');
    const allyRanks = [];
    for (const sid of allySummonerIds) {
        await sleep(350);
        const rk = await fetchRankForSummoner(sid);
        if (rk.score >= 0) allyRanks.push(rk.score);
    }
    const myRankEntry = await fetchRankForSummoner(me.summonerId);
    // Rang moyen alliés (moi inclus)
    const avgAllyScore = allyRanks.length
        ? parseFloat((allyRanks.reduce((s, v) => s + v, 0) / allyRanks.length).toFixed(2))
        : -1;

    console.log('   -> Récupération rangs ennemis...');
    const enemyRanks = [];
    for (const sid of enemySummonerIds) {
        await sleep(350);
        const rk = await fetchRankForSummoner(sid);
        if (rk.score >= 0) enemyRanks.push(rk.score);
    }
    const avgEnemyScore = enemyRanks.length
        ? parseFloat((enemyRanks.reduce((s, v) => s + v, 0) / enemyRanks.length).toFixed(2))
        : -1;

    const rankDiff = (avgAllyScore >= 0 && avgEnemyScore >= 0)
        ? parseFloat((avgAllyScore - avgEnemyScore).toFixed(2))
        : null;

    return {
        matchId,
        rawDate:              gameDate.toISOString(),
        date:                 gameDate.toLocaleDateString('fr-FR'),
        time:                 gameDate.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }),
        sessionId:            0,
        gameInSession:        0,
        gameDuration:         secToMmSs(match.info.gameDuration),   // mm:ss désormais
        gameDurationSec:      match.info.gameDuration,              // en secondes pour les calculs
        champion:             me.championName,
        role,
        type:                 isDuo ? 'Duo' : 'Solo',
        win:                  me.win ? 'Victoire' : 'Defaite',
        kda:                  parseFloat(((me.kills + me.assists) / Math.max(1, me.deaths)).toFixed(2)),
        kills:                me.kills,
        deaths:               me.deaths,
        assists:              me.assists,
        csPerMin:             parseFloat((totalCS / minutes).toFixed(1)),
        gpm:                  parseFloat((me.goldEarned / minutes).toFixed(0)),
        dpm:                  parseFloat((me.totalDamageDealtToChampions / minutes).toFixed(0)),
        objDpm:               parseFloat(((me.damageDealtToObjectives || 0) / minutes).toFixed(0)),
        dmgShare:             parseFloat(dmgShare.toFixed(1)),
        kp:                   parseFloat(kp.toFixed(1)),
        vision:               me.visionScore || 0,
        wardsPlaced:          me.wardsPlaced || 0,
        wardsKilled:          me.wardsKilled || 0,
        controlWards:         me.visionWardsBoughtInGame || 0,
        pctDeadTime:          parseFloat(((me.totalTimeSpentDead || 0) / match.info.gameDuration * 100).toFixed(1)),
        bestMultiKill:        bestMultiKill(me),
        pentaKills:           me.pentaKills    || 0,
        quadraKills:          me.quadraKills   || 0,
        tripleKills:          me.tripleKills   || 0,
        doubleKills:          me.doubleKills   || 0,
        largestKillingSpree:  me.largestKillingSpree || 0,
        firstBlood:           me.firstBloodKill ? 'Oui' : 'Non',
        withYuumi:            (isDuo && duo.championName === 'Yuumi') ? 'Oui' : 'Non',
        yuumiAlliee:          myTeam.some(p => p.puuid !== myPuuid && p.championName === 'Yuumi') ? 'Oui' : 'Non',
        // Pings
        totalPings,
        negativePings,
        allInPings,
        basicPings,
        commandPings,
        dangerPings,
        enemyMissingPings,
        holdPings,
        needVisionPings,
        onMyWayPings,
        pushPings,
        retreatPings,
        visionClearedPings,
        // Ganks
        ganksPerformed,
        // Rangs
        myRank:           myRankEntry.label,
        myRankScore:      myRankEntry.score,
        avgAllyRank:      scoreToLabel(Math.round(avgAllyScore)),
        avgAllyRankScore: avgAllyScore,
        avgEnemyRank:     scoreToLabel(Math.round(avgEnemyScore)),
        avgEnemyRankScore: avgEnemyScore,
        rankDiff,
    };
}

// ============================================================
// MIGRATION — complète les entrées JSON incomplètes
// ============================================================
function needsMigration(m) {
    return m.gameDurationSec == null
        || m.yuumiAlliee     == null
        || m.firstBlood      == null
        || m.pctDeadTime     == null
        || m.objDpm          == null
        || m.role            == null
        || m.totalPings      == null
        || m.myRankScore     == null
        || m.fatalGanksReceived === undefined; // Déclenche la mise à jour
}

async function migrateOldEntries(data, myPuuid, duoPuuid) {
    const toFix = data.filter(needsMigration);
    if (toFix.length === 0) return 0;

    console.log(`   -> ${toFix.length} entree(s) incomplete(s). Migration en cours...`);
    let ok = 0, ko = 0;

    for (let i = 0; i < toFix.length; i++) {
        const entry = toFix[i];
        process.stdout.write(`   [Migration ${i+1}/${toFix.length}] ${entry.matchId}... `);
        try {
            const r     = await axios.get(
                `https://${REGION}.api.riotgames.com/lol/match/v5/matches/${entry.matchId}`,
                { headers: { 'X-Riot-Token': RIOT_API_KEY } }
            );
            const match  = r.data;
            const me     = match.info.participants.find(p => p.puuid === myPuuid);
            if (!me) throw new Error('PUUID introuvable');

            const duo    = match.info.participants.find(p => p.puuid === duoPuuid && p.teamId === me.teamId);
            const myTeam = match.info.participants.filter(p => p.teamId === me.teamId);
            const minutes = match.info.gameDuration / 60;

            // Durée
            entry.gameDurationSec = match.info.gameDuration;
            entry.gameDuration    = secToMmSs(match.info.gameDuration);

            // Champs manquants originaux
            entry.objDpm              = parseFloat(((me.damageDealtToObjectives || 0) / minutes).toFixed(0));
            entry.wardsPlaced         = me.wardsPlaced || 0;
            entry.wardsKilled         = me.wardsKilled || 0;
            entry.controlWards        = me.visionWardsBoughtInGame || 0;
            entry.pctDeadTime         = parseFloat(((me.totalTimeSpentDead || 0) / match.info.gameDuration * 100).toFixed(1));
            entry.bestMultiKill       = bestMultiKill(me);
            entry.pentaKills          = me.pentaKills          || 0;
            entry.quadraKills         = me.quadraKills         || 0;
            entry.tripleKills         = me.tripleKills         || 0;
            entry.doubleKills         = me.doubleKills         || 0;
            entry.largestKillingSpree = me.largestKillingSpree || 0;
            entry.firstBlood          = me.firstBloodKill ? 'Oui' : 'Non';
            entry.yuumiAlliee         = myTeam.some(p => p.puuid !== myPuuid && p.championName === 'Yuumi') ? 'Oui' : 'Non';
            entry.withYuumi           = (!!duo && duo.championName === 'Yuumi') ? 'Oui' : 'Non';

            // Nouveaux champs
            entry.role = me.teamPosition || me.individualPosition || '-';

            // Pings
            entry.allInPings         = me.allInPings            || 0;
            entry.basicPings         = me.basicPings            || 0;
            entry.commandPings       = me.commandPings          || 0;
            entry.dangerPings        = me.dangerPings           || 0;
            entry.enemyMissingPings  = me.enemyMissingPings     || 0;
            entry.holdPings          = me.holdPings             || 0;
            entry.needVisionPings    = me.needVisionPings       || 0;
            entry.onMyWayPings       = me.onMyWayPings          || 0;
            entry.pushPings          = me.pushPings             || 0;
            entry.retreatPings       = me.retreatPings          || 0;
            entry.visionClearedPings = me.visionClearedPings    || 0;
            entry.totalPings         = entry.allInPings + entry.basicPings + entry.commandPings
                                     + entry.dangerPings + entry.enemyMissingPings + entry.holdPings
                                     + entry.needVisionPings + entry.onMyWayPings + entry.pushPings
                                     + entry.retreatPings + entry.visionClearedPings;
            entry.negativePings      = entry.dangerPings + entry.retreatPings + entry.enemyMissingPings;

            // Ganks
            entry.ganksPerformed = entry.role === 'JUNGLE'
                ? (me.challenges?.killsOnLanersEarlyJungleAsJungler ?? null)
                : null;

            // Rangs (migration coûteuse — on essaie)
            if (entry.myRankScore == null) {
                await sleep(350);
                const myRk = await fetchRankForSummoner(me.summonerId);
                entry.myRank      = myRk.label;
                entry.myRankScore = myRk.score;

                const enemies    = match.info.participants.filter(p => p.teamId !== me.teamId);
                const allyScores = [], enemyScores = [];

                for (const p of myTeam) {
                    if (!p.summonerId) continue;
                    await sleep(300);
                    const rk = await fetchRankForSummoner(p.summonerId);
                    if (rk.score >= 0) allyScores.push(rk.score);
                }
                for (const p of enemies) {
                    if (!p.summonerId) continue;
                    await sleep(300);
                    const rk = await fetchRankForSummoner(p.summonerId);
                    if (rk.score >= 0) enemyScores.push(rk.score);
                }

                const avgA = allyScores.length  ? allyScores.reduce((s, v)  => s + v, 0) / allyScores.length  : -1;
                const avgE = enemyScores.length ? enemyScores.reduce((s, v) => s + v, 0) / enemyScores.length : -1;
                entry.avgAllyRank      = scoreToLabel(Math.round(avgA));
                entry.avgAllyRankScore = avgA >= 0 ? parseFloat(avgA.toFixed(2)) : -1;
                entry.avgEnemyRank     = scoreToLabel(Math.round(avgE));
                entry.avgEnemyRankScore = avgE >= 0 ? parseFloat(avgE.toFixed(2)) : -1;
                entry.rankDiff         = (avgA >= 0 && avgE >= 0) ? parseFloat((avgA - avgE).toFixed(2)) : null;
            }

            ok++;
            console.log('OK');
        } catch (e) {
            ko++;
            console.log(`KO (${e.message})`);
        }
        await sleep(API_DELAY_MS);
    }
    console.log(`   -> Migration : ${ok} corrigee(s)${ko ? ', ' + ko + ' echec(s)' : ''}.`);
    return ok;
}

// ============================================================
// HELPERS EXCEL — styles de base
// ============================================================
function writeHeaderRow(sheet, rowNum, labels, bg) {
    const row = sheet.getRow(rowNum);
    labels.forEach((label, i) => {
        const cell     = row.getCell(i + 1);
        cell.value     = label;
        cell.font      = { bold: true, color: { argb: C.white }, size: 10 };
        cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border    = { bottom: { style: 'thin', color: { argb: C.accent } } };
    });
    row.height = 26;
}

// Header multicolore par groupes
function writeMultiColorHeaderRow(sheet, rowNum, groups) {
    // groups = [{ label, bg, count }]
    const row = sheet.getRow(rowNum);
    let col   = 1;
    for (const grp of groups) {
        for (let i = 0; i < grp.count; i++) {
            const cell = row.getCell(col);
            cell.value     = grp.labels[i];
            cell.font      = { bold: true, color: { argb: C.white }, size: 9 };
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: grp.bg } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border    = { bottom: { style: 'thin', color: { argb: C.accent } } };
            col++;
        }
    }
    row.height = 28;
}

function sectionTitle(sheet, rowNum, text, span, bg) {
    sheet.mergeCells(rowNum, 1, rowNum, span);
    const cell     = sheet.getRow(rowNum).getCell(1);
    cell.value     = text;
    cell.font      = { bold: true, color: { argb: C.white }, size: 12 };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    sheet.getRow(rowNum).height = 28;
}

function colorResult(cell, win) {
    if (win === 'Victoire') {
        cell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg  } };
    } else {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
    }
}

function colorWR(cell, wr) {
    cell.numFmt = '0.0%';
    if (wr >= 0.55) {
        cell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg  } };
    } else if (wr <= 0.45) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
    }
}

function colorKDA(cell, kda) {
    cell.numFmt = '0.00';
    if (kda >= 4.0) {
        cell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good   } };
    } else if (kda < 1.5) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad    } };
    } else if (kda >= 2.5) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.warn   } };
    }
}

function colorDeaths(cell, d) {
    if (d >= 6) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad  } };
    } else if (d <= 2) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
    }
}

function colorCS(cell, cs) {
    cell.numFmt = '0.0';
    if (cs >= 9.0) {
        cell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
    } else if (cs < 7.0) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad  } };
    }
}

function colorDeadTime(cell, pct) {
    cell.numFmt = '0.0';
    if (pct > 20) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
    }
}

function colorRankCell(cell, score) {
    if (score < 0) return;
    const fg = rankColor(score);
    // Texte foncé sur fond clair, texte clair sur fond foncé
    const lightBgs = [C.rankGold, C.rankPlatinum, C.rankEmerald, C.rankDiamond, C.rankChallenger];
    const isLight  = lightBgs.includes(fg);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fg } };
    cell.font = { bold: true, color: { argb: isLight ? 'FF000000' : C.white }, size: 10 };
}

function colorRankDiff(cell, diff) {
    if (diff === null || diff === undefined) return;
    if (diff > 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg  } };
        cell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
    } else if (diff < -1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
    }
}

function fillRow(row, idx, firstLeft = false) {
    const bg = rowBg(idx);
    row.eachCell((cell, ci) => {
        if (!cell.fill || cell.fill.fgColor === undefined) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        }
        cell.alignment = { horizontal: (firstLeft && ci === 1) ? 'left' : 'center', vertical: 'middle' };
        if (!cell.font || (!cell.font.bold && !cell.font.color)) {
            cell.font = { size: 10 };
        }
    });
    row.height = 20;
}

// ============================================================
// SHEET 1 — DONNEES BRUTES (groupes colorés)
// ============================================================
function buildRawDataSheet(sheet, data) {
    // Définition des groupes et de leurs colonnes
    const GROUPS = [
        {
            bg: C.grpIdentite,
            labels: ['Date','Heure','Session','P/S','Duree','Champion','Role','Type','Resultat'],
            keys:   ['date','time','sessionId','gameInSession','gameDuration','champion','role','type','win'],
            widths: [11, 7, 9, 6, 8, 14, 10, 7, 10],
        },
        {
            bg: C.grpCombat,
            labels: ['KDA','K','D','A','% Mort','KP %','% DMG'],
            keys:   ['kda','kills','deaths','assists','pctDeadTime','kp','dmgShare'],
            widths: [7, 5, 5, 5, 8, 8, 8],
        },
        {
            bg: C.grpFarm,
            labels: ['CS/Min','GPM','DPM','Obj/Min'],
            keys:   ['csPerMin','gpm','dpm','objDpm'],
            widths: [8, 7, 8, 9],
        },
        {
            bg: C.grpVision,
            labels: ['Vision','Wards+','Wards-','Ctrl W.'],
            keys:   ['vision','wardsPlaced','wardsKilled','controlWards'],
            widths: [8, 8, 8, 9],
        },
        {
            bg: C.grpPings,
            labels: ['Pings Tot.','Pings Nég.','AllIn','Basic','Cmd','Danger','MIA','Hold','Vision?','OTW','Push','Retreat','ClearVis'],
            keys:   ['totalPings','negativePings','allInPings','basicPings','commandPings','dangerPings','enemyMissingPings','holdPings','needVisionPings','onMyWayPings','pushPings','retreatPings','visionClearedPings'],
            widths: [10, 10, 7, 7, 7, 8, 7, 7, 9, 8, 8, 9, 10],
        },
        {
            bg: C.grpGanks,
            labels: ['Ganks Subis (Early)'],
            keys:   ['fatalGanksReceived'],
            widths: [16],
        },
        {
            bg: C.grpRank,
            labels: ['Mon Rang','Rang Alliés','Rang Ennemis','Diff Rang'],
            keys:   ['myRank','avgAllyRank','avgEnemyRank','rankDiff'],
            widths: [14, 14, 14, 10],
        },
        {
            bg: C.grpDivers,
            labels: ['Multi-Kill','First Blood','Duo Yuumi','Yuumi Alli.'],
            keys:   ['bestMultiKill','firstBlood','withYuumi','yuumiAlliee'],
            widths: [11, 11, 10, 12],
        },
    ];

    // Aplatir pour construire sheet.columns
    const flatKeys   = GROUPS.flatMap(g => g.keys);
    const flatWidths = GROUPS.flatMap(g => g.widths);
    const NCOLS      = flatKeys.length;

    sheet.columns = flatKeys.map((key, i) => ({ key, width: flatWidths[i] }));
    sheet.views   = [{ state: 'frozen', ySplit: 1 }];

    // Header multicolore
    writeMultiColorHeaderRow(sheet, 1, GROUPS.map(g => ({ ...g, count: g.keys.length })));

    const reversed = [...data].reverse();

    reversed.forEach((m, idx) => {
        const rowData = {};
        flatKeys.forEach(k => { rowData[k] = safeV(m[k]); });
        const row = sheet.addRow(rowData);

        // Fond de ligne
        const bg = rowBg(idx);
        row.eachCell(cell => {
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.font      = { size: 10 };
        });
        row.getCell('date').alignment = { horizontal: 'left', vertical: 'middle' };
        row.height = 20;

        // Colorings
        colorResult(row.getCell('win'), m.win);
        if (typeof m.kda        === 'number') colorKDA(row.getCell('kda'), m.kda);
        if (typeof m.deaths     === 'number') colorDeaths(row.getCell('deaths'), m.deaths);
        if (typeof m.csPerMin   === 'number') colorCS(row.getCell('csPerMin'), m.csPerMin);
        if (typeof m.pctDeadTime === 'number') colorDeadTime(row.getCell('pctDeadTime'), m.pctDeadTime);

        if (typeof m.dmgShare === 'number') row.getCell('dmgShare').numFmt = '0.0';
        if (typeof m.kp       === 'number') row.getCell('kp').numFmt       = '0.0';
        if (typeof m.gpm      === 'number') row.getCell('gpm').numFmt      = '0';
        if (typeof m.dpm      === 'number') row.getCell('dpm').numFmt      = '0';

        // Rang
        if (typeof m.myRankScore === 'number')        colorRankCell(row.getCell('myRank'),      m.myRankScore);
        if (typeof m.avgAllyRankScore === 'number')   colorRankCell(row.getCell('avgAllyRank'),  m.avgAllyRankScore);
        if (typeof m.avgEnemyRankScore === 'number')  colorRankCell(row.getCell('avgEnemyRank'), m.avgEnemyRankScore);
        if (m.rankDiff !== null && m.rankDiff !== undefined) {
            row.getCell('rankDiff').value  = m.rankDiff;
            row.getCell('rankDiff').numFmt = '+0.0;-0.0;0.0';
            colorRankDiff(row.getCell('rankDiff'), m.rankDiff);
        }

        // Pings négatifs colorés en amber si élevés
        const negP = safeN(m.negativePings, 0);
        if (negP > 15) {
            row.getCell('negativePings').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.warn } };
            row.getCell('negativePings').font = { bold: true, color: { argb: C.amber }, size: 10 };
        }

        // First Blood
        if (m.firstBlood === 'Oui') {
            row.getCell('firstBlood').font = { bold: true, color: { argb: C.winFg }, size: 10 };
        }
        // Multi-kill
        const mk = String(m.bestMultiKill || '');
        if (mk === 'Triple' || mk === 'Quadra' || mk === 'Penta') {
            row.getCell('bestMultiKill').font = { bold: true, color: { argb: C.purple }, size: 10 };
        }
    });

    sheet.autoFilter = {
        from: { row: 1, column: 1 },
        to:   { row: 1, column: NCOLS },
    };
}

// ============================================================
// SHEET 2 — PAR CHAMPION
// ============================================================
function buildChampionSheet(sheet, data) {
    const LABELS = ['Champion','Parties','Victoires','Defaites','Win Rate %','KDA Moy.',
                    'K Moy.','D Moy.','A Moy.',
                    'CS/Min Moy.','GPM Moy.','DPM Moy.','% DMG Moy.','KP % Moy.',
                    'Vision Moy.','Solo','Duo','Duree Moy.'];
    const KEYS   = ['champion','games','wins','losses','winRate','kda',
                    'avgK','avgD','avgA',
                    'cs','gpm','dpm','dmg','kp','vision','solo','duo','duration'];
    const WIDTHS = [16, 9, 10, 10, 12, 9, 7, 7, 7, 11, 9, 9, 11, 10, 11, 7, 7, 11];
    const NCOLS  = LABELS.length;

    sheet.columns = KEYS.map((k, i) => ({ key: k, width: WIDTHS[i] }));
    sheet.views   = [{ state: 'frozen', ySplit: 2 }];

    sectionTitle(sheet, 1, 'Stats par Champion  —  trie par parties jouees', NCOLS, C.hChamp);
    writeHeaderRow(sheet, 2, LABELS, C.hChamp);

    const champMap = {};
    data.forEach(m => pushTo(champMap, m.champion, m));

    const list = Object.entries(champMap).map(([name, games]) => {
        const wins = games.filter(g => g.win === 'Victoire').length;
        return {
            champion: name,
            games:    games.length,
            wins,
            losses:   games.length - wins,
            winRate:  parseFloat(winPct(games)),
            kda:      parseFloat(avgOf(games, 'kda').toFixed(2)),
            avgK:     parseFloat(avgOf(games, 'kills').toFixed(1)),
            avgD:     parseFloat(avgOf(games, 'deaths').toFixed(1)),
            avgA:     parseFloat(avgOf(games, 'assists').toFixed(1)),
            cs:       parseFloat(avgOf(games, 'csPerMin').toFixed(1)),
            gpm:      Math.round(avgOf(games, 'gpm')),
            dpm:      Math.round(avgOf(games, 'dpm')),
            dmg:      parseFloat(avgOf(games, 'dmgShare').toFixed(1)),
            kp:       parseFloat(avgOf(games, 'kp').toFixed(1)),
            vision:   parseFloat(avgOf(games, 'vision').toFixed(1)),
            solo:     games.filter(g => g.type === 'Solo').length,
            duo:      games.filter(g => g.type === 'Duo').length,
            duration: games.filter(g => g.gameDurationSec)
                          .reduce((s, g) => s + (g.gameDurationSec || 0), 0)
                / Math.max(1, games.filter(g => g.gameDurationSec).length) / 60,
        };
    }).sort((a, b) => b.games - a.games);

    list.forEach((c, idx) => {
        const row = sheet.addRow(c);
        const bg  = rowBg(idx);
        row.eachCell(cell => {
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.font      = { size: 10 };
        });
        row.getCell('champion').font      = { bold: true, size: 10 };
        row.getCell('champion').alignment = { horizontal: 'left', vertical: 'middle' };
        row.height = 20;

        colorWR(row.getCell('winRate'), c.winRate);
        colorKDA(row.getCell('kda'), c.kda);
        colorCS(row.getCell('cs'), c.cs);
        colorDeaths(row.getCell('avgD'), Math.round(c.avgD));

        row.getCell('dmg').numFmt     = '0.0"%"';
        row.getCell('kp').numFmt      = '0.0"%"';
        row.getCell('gpm').numFmt     = '0';
        row.getCell('dpm').numFmt     = '0';
        // Durée en mm:ss
        const durSec  = c.duration * 60;
        row.getCell('duration').value = secToMmSs(durSec);
    });

    sheet.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: NCOLS } };
}

// ============================================================
// SHEET 3 — SOLO vs DUO
// ============================================================
function buildDuoSoloSheet(sheet, data) {
    const cats = [
        { label: 'Solo Queue',      games: data.filter(m => m.type === 'Solo') },
        { label: 'Duo (total)',     games: data.filter(m => m.type === 'Duo') },
        { label: 'Duo + Yuumi',    games: data.filter(m => m.withYuumi === 'Oui') },
        { label: 'Duo sans Yuumi', games: data.filter(m => m.type === 'Duo' && m.withYuumi !== 'Oui') },
        { label: 'Yuumi alliee',   games: data.filter(m => m.yuumiAlliee === 'Oui') },
    ];
    const SPAN = 1 + cats.length;

    sheet.columns = [
        { key: 'metric', width: 24 },
        ...cats.map(c => ({ key: c.label, width: 18 })),
    ];
    sheet.views = [{ state: 'frozen', ySplit: 2 }];

    sectionTitle(sheet, 1, 'Comparaison Solo vs Duo', SPAN, C.hDuo);
    writeHeaderRow(sheet, 2, ['Metrique', ...cats.map(c => c.label)], C.hDuo);

    const tableRows = [
        ['Parties',            ...cats.map(c => c.games.length)],
        ['Victoires',          ...cats.map(c => c.games.filter(g => g.win === 'Victoire').length)],
        ['Defaites',           ...cats.map(c => c.games.filter(g => g.win !== 'Victoire').length)],
        ['Win Rate %',         ...cats.map(c => parseFloat(winPct(c.games).toFixed(1)))],
        ['KDA Moyen',          ...cats.map(c => parseFloat(avgOf(c.games, 'kda').toFixed(2)))],
        ['K Moyen',            ...cats.map(c => parseFloat(avgOf(c.games, 'kills').toFixed(1)))],
        ['D Moyen',            ...cats.map(c => parseFloat(avgOf(c.games, 'deaths').toFixed(1)))],
        ['A Moyen',            ...cats.map(c => parseFloat(avgOf(c.games, 'assists').toFixed(1)))],
        ['CS/Min Moyen',       ...cats.map(c => parseFloat(avgOf(c.games, 'csPerMin').toFixed(1)))],
        ['DPM Moyen',          ...cats.map(c => Math.round(avgOf(c.games, 'dpm')))],
        ['GPM Moyen',          ...cats.map(c => Math.round(avgOf(c.games, 'gpm')))],
        ['% DMG Moyen',        ...cats.map(c => parseFloat(avgOf(c.games, 'dmgShare').toFixed(1)))],
        ['KP % Moyen',         ...cats.map(c => parseFloat(avgOf(c.games, 'kp').toFixed(1)))],
        ['Vision Moyen',       ...cats.map(c => parseFloat(avgOf(c.games, 'vision').toFixed(1)))],
        ['Obj/Min Moyen',      ...cats.map(c => Math.round(avgOf(c.games, 'objDpm')))],
        ['% Temps Mort Moyen', ...cats.map(c => parseFloat(avgOf(c.games, 'pctDeadTime').toFixed(1)))],
        ['Pings Tot. Moy.',    ...cats.map(c => Math.round(avgOf(c.games, 'totalPings')))],
        ['Pings Nég. Moy.',    ...cats.map(c => Math.round(avgOf(c.games, 'negativePings')))],
    ];

    tableRows.forEach((rowData, idx) => {
        const shRow = sheet.getRow(idx + 3);
        rowData.forEach((val, ci) => {
            const cell     = shRow.getCell(ci + 1);
            cell.value     = val;
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            cell.font      = { bold: ci === 0, size: 10 };
            cell.alignment = { horizontal: ci === 0 ? 'left' : 'center', vertical: 'middle' };
            if (rowData[0] === 'Win Rate %' && typeof val === 'number') cell.numFmt = '0.0"%"';
        });
        shRow.height = 22;

        if (rowData[0] === 'Win Rate %') {
            for (let ci = 1; ci <= cats.length; ci++) {
                const cell = shRow.getCell(ci + 1);
                colorWR(cell, (parseFloat(cell.value) || 0) / 100);
            }
        }
        if (rowData[0] === 'KDA Moyen') {
            for (let ci = 1; ci <= cats.length; ci++) {
                colorKDA(shRow.getCell(ci + 1), parseFloat(shRow.getCell(ci + 1).value) || 0);
            }
        }
    });
}

// ============================================================
// SHEET 4 — TENDANCES COMPORTEMENTALES (étendue)
// ============================================================
function buildTendancesSheet(sheet, data, a) {
    if (!a) {
        sheet.getRow(1).getCell(1).value = 'Aucune donnee disponible.';
        return;
    }

    sheet.columns = [
        { key: 'c1', width: 28 },
        { key: 'c2', width: 16 },
        { key: 'c3', width: 16 },
        { key: 'c4', width: 14 },
        { key: 'c5', width: 14 },
        { key: 'c6', width: 14 },
        { key: 'c7', width: 14 },
        { key: 'c8', width: 32 },
    ];
    const SPAN = 8;
    let r = 1;

    function kv(rowNum, label, value, idx, isPercent = false) {
        const row = sheet.getRow(rowNum);
        const bg  = rowBg(idx);
        const c1  = row.getCell(1);
        const c2  = row.getCell(2);
        c1.value = label;
        c1.font  = { bold: true, size: 10 };
        c1.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        c2.value = value;
        c2.font  = { size: 10 };
        c2.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        c2.alignment = { horizontal: 'center', vertical: 'middle' };
        if (isPercent && typeof value === 'number') c2.numFmt = '0.0%';
        row.height = 20;
    }

    function tRow(rowNum, values, idx, percentCols = []) {
        const row = sheet.getRow(rowNum);
        values.forEach((v, i) => {
            const cell = row.getCell(i + 1);
            cell.value = v;
            cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            cell.font  = { bold: i === 0, size: 10 };
            cell.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
            if (percentCols.includes(i + 1) && typeof v === 'number') cell.numFmt = '0.0%';
        });
        row.height = 20;
    }

    function colorWRCell(rowNum, colNum, wr) {
        const cell = sheet.getRow(rowNum).getCell(colNum);
        if (wr >= 0.55) {
            cell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg  } };
        } else if (wr <= 0.45) {
            cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
        }
    }

    // ── VUE D'ENSEMBLE ────────────────────────────────────────
    sectionTitle(sheet, r, "Vue d'Ensemble Globale", SPAN, C.hTend); r++;
    [
        ["Parties totales",      a.overall.total, false],
        ["Win Rate global",      a.overall.winRate, true],
        ["KDA moyen",            a.overall.avgKDA.toFixed(2), false],
        ["K / D / A moyens",     `${a.overall.avgKills.toFixed(1)} / ${a.overall.avgDeaths.toFixed(1)} / ${a.overall.avgAssists.toFixed(1)}`, false],
        ["CS/Min moyen",         a.overall.avgCS.toFixed(1), false],
        ["DPM moyen",            Math.round(a.overall.avgDPM), false],
        ["GPM moyen",            Math.round(a.overall.avgGPM), false],
        ["% DMG Equipe moyen",   a.overall.avgDmgShare.toFixed(1) + ' %', false],
        ["KP % moyen",           a.overall.avgKP.toFixed(1) + ' %', false],
        ["Vision Score moyen",   a.overall.avgVision.toFixed(1), false],
    ].forEach(([label, value, isPct], i) => { kv(r, label, value, i, isPct); r++; });
    r++;

    // ── SÉRIE & TENDANCE ──────────────────────────────────────
    sectionTitle(sheet, r, "Serie Actuelle & Tendance", SPAN, C.hTend); r++;
    const wr10   = winPct(a.recent10);
    const wrGlob = a.overall.winRate;
    const trend  = wr10 > wrGlob + 0.02 ? 'En progression' : wr10 < wrGlob - 0.02 ? 'En regression' : 'Stable';
    [
        ["Serie actuelle",          a.streak.count + ' ' + a.streak.type + ' consecutives'],
        ["Win Rate (10 derniers)",  wr10],
        ["Win Rate global",         wrGlob],
        ["Tendance recente",        trend],
        ["K/D/A (10 derniers)",     `${avgOf(a.recent10,'kills').toFixed(1)} / ${avgOf(a.recent10,'deaths').toFixed(1)} / ${avgOf(a.recent10,'assists').toFixed(1)}`],
        ["KDA (10 derniers)",       avgOf(a.recent10, 'kda').toFixed(2)],
        ["DPM (10 derniers)",       Math.round(avgOf(a.recent10, 'dpm'))],
    ].forEach(([label, value], i) => { kv(r, label, value, i, label.includes('Win Rate')); r++; });
    r++;

    // ── FATIGUE COGNITIVE ÉTENDUE ─────────────────────────────
    sectionTitle(sheet, r, "Fatigue Cognitive — Performance par Partie dans la Session", SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Partie/Session', 'Nb Parties', 'Win Rate', 'KDA Moy.', 'K Moy.', 'D Moy.', 'A Moy.', 'CS/Min | DPM'], C.subH); r++;

    // Calculer les valeurs globales pour la comparaison delta
    const globalWR  = a.overall.winRate;
    const globalKDA = a.overall.avgKDA;
    const globalCS  = a.overall.avgCS;
    const globalDPM = a.overall.avgDPM;

    [1, 2, 3, 4].forEach((gNum, idx) => {
        const games = a.byGameInSession[gNum] || [];
        const wr    = winPct(games);
        const kAvg  = avgOf(games, 'kills');
        const dAvg  = avgOf(games, 'deaths');
        const aAvg  = avgOf(games, 'assists');
        const kdaAvg = avgOf(games, 'kda');
        const csAvg  = avgOf(games, 'csPerMin');
        const dpmAvg = Math.round(avgOf(games, 'dpm'));

        const deltaWR  = games.length ? (wr - globalWR) * 100 : null;
        const labelWR  = deltaWR !== null ? (deltaWR >= 0 ? `+${deltaWR.toFixed(1)}%` : `${deltaWR.toFixed(1)}%`) : '-';

        tRow(r, [
            gNum < 4 ? `Partie ${gNum}` : 'Partie 4+',
            games.length,
            games.length ? wr : '-',
            games.length ? kdaAvg.toFixed(2) : '-',
            games.length ? kAvg.toFixed(1) : '-',
            games.length ? dAvg.toFixed(1) : '-',
            games.length ? aAvg.toFixed(1) : '-',
            games.length ? `${csAvg.toFixed(1)} | ${dpmAvg}` : '-',
        ], idx, [3]);

        if (games.length) {
            colorWRCell(r, 3, wr);

            // Delta vs global dans une colonne annexe (commentaire de cellule)
            const kdaCell = sheet.getRow(r).getCell(4);
            if (kdaAvg >= globalKDA + 0.3) {
                kdaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
            } else if (kdaAvg <= globalKDA - 0.3) {
                kdaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
            }

            // Deaths : colore si on meurt plus
            const dCell = sheet.getRow(r).getCell(6);
            if (dAvg > a.overall.avgDeaths + 0.5) {
                dCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
                dCell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
            } else if (dAvg <= a.overall.avgDeaths - 0.5) {
                dCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
            }
        }
        r++;
    });

    // Ligne de synthèse (diagnostic automatique)
    const g1 = a.byGameInSession[1] || [];
    const g4 = a.byGameInSession[4] || [];
    const fatigueWRDrop = g1.length && g4.length ? (winPct(g1) - winPct(g4)) * 100 : null;
    const fatigueDiag   = fatigueWRDrop === null ? 'Données insuffisantes'
        : fatigueWRDrop > 15  ? `⚠ Fatigue forte (-${fatigueWRDrop.toFixed(1)}% WR P1→P4+)`
        : fatigueWRDrop > 5   ? `! Legere fatigue (-${fatigueWRDrop.toFixed(1)}% WR P1→P4+)`
        :                       `✓ Stable (${fatigueWRDrop >= 0 ? '+' : ''}${fatigueWRDrop.toFixed(1)}% WR P1→P4+)`;
    const synthRow = sheet.getRow(r);
    sheet.mergeCells(r, 1, r, SPAN);
    synthRow.getCell(1).value     = 'Diagnostic : ' + fatigueDiag;
    synthRow.getCell(1).font      = { italic: true, size: 10, bold: fatigueWRDrop !== null && fatigueWRDrop > 5 };
    synthRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    synthRow.getCell(1).fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.warn } };
    synthRow.height = 22;
    r += 2;

    // ── PAR TRANCHE HORAIRE ───────────────────────────────────
    sectionTitle(sheet, r, "Performance par Tranche Horaire", SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Tranche Horaire', 'Parties', 'Win Rate', 'KDA', 'K/D/A', 'DPM', 'Pings Nég.', 'Diagnostic'], C.subH); r++;

    Object.entries(a.timeSlots).forEach(([slot, games], idx) => {
        const wr   = winPct(games);
        const diag = games.length < 3 ? 'Echantillon faible'
                   : wr >= 0.55       ? 'Tranche favorable'
                   : wr <= 0.45       ? 'Tranche defavorable'
                   :                    'Neutre';
        const kda  = games.length ? avgOf(games, 'kda').toFixed(2) : '-';
        const kda_str = games.length
            ? `${avgOf(games,'kills').toFixed(1)}/${avgOf(games,'deaths').toFixed(1)}/${avgOf(games,'assists').toFixed(1)}`
            : '-';
        tRow(r, [
            slot, games.length,
            games.length ? wr : '-',
            kda, kda_str,
            games.length ? Math.round(avgOf(games, 'dpm')) : '-',
            games.length ? Math.round(avgOf(games, 'negativePings')) : '-',
            diag,
        ], idx, [3]);
        if (games.length) colorWRCell(r, 3, wr);
        r++;
    });
    r++;

    // ── INDICATEUR DE TILT ────────────────────────────────────
    sectionTitle(sheet, r, "Indicateur de Tilt — Perf. apres Victoire vs Defaite", SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Contexte', 'Parties', 'Win Rate', 'KDA', 'K/D/A', 'DPM', 'Pings Nég.', 'Diagnostic'], C.subH); r++;

    const wrW   = winPct(a.afterWin);
    const wrL   = winPct(a.afterLoss);
    const delta = wrW - wrL;
    const tiltDiag = delta > 0.10 ? 'Tilt detecte — Stopper apres defaite'
                   : delta > 0.05 ? 'Legere regression post-defaite'
                   :                'Mental stable';
    [
        ['Apres une Victoire', a.afterWin.length,  wrW,
            avgOf(a.afterWin,'kda').toFixed(2),
            `${avgOf(a.afterWin,'kills').toFixed(1)}/${avgOf(a.afterWin,'deaths').toFixed(1)}/${avgOf(a.afterWin,'assists').toFixed(1)}`,
            Math.round(avgOf(a.afterWin,'dpm')),
            Math.round(avgOf(a.afterWin,'negativePings')),
            ''],
        ['Apres une Defaite',  a.afterLoss.length, wrL,
            avgOf(a.afterLoss,'kda').toFixed(2),
            `${avgOf(a.afterLoss,'kills').toFixed(1)}/${avgOf(a.afterLoss,'deaths').toFixed(1)}/${avgOf(a.afterLoss,'assists').toFixed(1)}`,
            Math.round(avgOf(a.afterLoss,'dpm')),
            Math.round(avgOf(a.afterLoss,'negativePings')),
            tiltDiag],
    ].forEach((rowData, idx) => {
        tRow(r, rowData, idx, [3]);
        if (typeof rowData[2] === 'number') colorWRCell(r, 3, rowData[2]);
        r++;
    });
    r++;

    // ── ANALYSE PINGS ─────────────────────────────────────────
    sectionTitle(sheet, r, "Analyse des Pings par Resultat", SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Contexte', 'Pings Tot.', 'Pings Nég.', 'Danger', 'MIA', 'Retreat', 'AllIn', 'OTW'], C.subH); r++;

    const pingCtxs = [
        { label: 'Victoires',           games: data.filter(m => m.win === 'Victoire') },
        { label: 'Defaites',            games: data.filter(m => m.win !== 'Victoire') },
        { label: 'Apres victoire',      games: a.afterWin },
        { label: 'Apres defaite',       games: a.afterLoss },
    ];
    pingCtxs.forEach(({ label, games }, idx) => {
        tRow(r, [
            label,
            Math.round(avgOf(games, 'totalPings')),
            Math.round(avgOf(games, 'negativePings')),
            Math.round(avgOf(games, 'dangerPings')),
            Math.round(avgOf(games, 'enemyMissingPings')),
            Math.round(avgOf(games, 'retreatPings')),
            Math.round(avgOf(games, 'allInPings')),
            Math.round(avgOf(games, 'onMyWayPings')),
        ], idx);
        // Colore pings négatifs si élevés
        const negVal = Math.round(avgOf(games, 'negativePings'));
        if (negVal > 10) {
            const cell = sheet.getRow(r).getCell(3);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.warn } };
            cell.font = { bold: true, color: { argb: C.amber }, size: 10 };
        }
        r++;
    });
    r++;

// ── WINRATE SELON LA PRESSION SUBIE (Bottom/Laner) ────────
    const lanerGames = data.filter(m => m.fatalGanksReceived !== null && m.fatalGanksReceived !== undefined);
    if (lanerGames.length >= 5) {
        sectionTitle(sheet, r, "Winrate par Ganks Mortels Subis (Phase de Lane)", SPAN, C.hTend); r++;
        writeHeaderRow(sheet, r, ['Ganks Subis', 'Parties', 'Win Rate', 'KDA', 'K/D/A', 'DPM', '', 'Tendance'], C.subH); r++;

        const gankBuckets = [0, 1, 2, 3]; // On regroupe au-delà de 3
        gankBuckets.forEach((gNum, idx) => {
            const games = gNum < 3
                ? lanerGames.filter(m => m.fatalGanksReceived === gNum)
                : lanerGames.filter(m => m.fatalGanksReceived >= 3);
            
            if (!games.length) return;
            const wr  = winPct(games);
            tRow(r, [
                gNum < 3 ? `${gNum} gank(s) mortel(s)` : '3+ ganks mortels',
                games.length,
                wr,
                avgOf(games, 'kda').toFixed(2),
                `${avgOf(games,'kills').toFixed(1)}/${avgOf(games,'deaths').toFixed(1)}/${avgOf(games,'assists').toFixed(1)}`,
                Math.round(avgOf(games, 'dpm')),
                '',
                wr >= 0.55 ? 'Excellente resistance' : wr <= 0.45 ? 'Impact critique' : 'Stable',
            ], idx, [3]);
            colorWRCell(r, 3, wr);
            r++;
        });
        r++;
    }

    // ── RECORDS PERSONNELS ────────────────────────────────────
    sectionTitle(sheet, r, "Records Personnels", SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Categorie', 'Valeur', 'Champion', 'Date', 'Resultat', 'Role', '', ''], C.subH); r++;

    [
        ['Meilleur KDA',    a.records.bestKDA,    'kda',      2],
        ['Meilleur DPM',    a.records.bestDPM,    'dpm',      0],
        ['Meilleur CS/Min', a.records.bestCS,     'csPerMin', 1],
        ['Meilleur GPM',    a.records.bestGPM,    'gpm',      0],
        ['Meilleur Vision', a.records.bestVision, 'vision',   0],
    ].forEach(([label, rec, key, dec], idx) => {
        const rRow = sheet.getRow(r);
        rRow.height = 20;
        if (rec) {
            const rawVal = safeN(rec[key]);
            const val    = dec === 0 ? Math.round(rawVal) : parseFloat(rawVal.toFixed(dec));
            const c1     = rRow.getCell(1);
            c1.value = label; c1.font = { bold: true, size: 10 };
            c1.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            const c2 = rRow.getCell(2);
            c2.value = val;
            c2.font  = { bold: true, size: 10, color: { argb: C.accent } };
            c2.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            c2.numFmt = dec === 2 ? '0.00' : dec === 1 ? '0.0' : '0';
            [rec.champion, rec.date, rec.win, rec.role || '-', '', ''].forEach((v, i) => {
                const cell = rRow.getCell(i + 3);
                cell.value = v;
                cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
                cell.font  = { size: 10 };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            });
            colorResult(rRow.getCell(5), rec.win);
        }
        r++;
    });
}

// ============================================================
// ORCHESTRATEUR EXCEL
// ============================================================
async function rebuildExcel(data) {
    console.log('  -> Reconstruction Excel (4 feuilles)...');
    const analytics = computeAnalytics(data);
    const wb        = new ExcelJS.Workbook();
    wb.creator      = 'LoL Tracker';
    wb.created      = new Date();

    const jobs = [
        { name: 'Donnees Brutes',             fn: () => buildRawDataSheet(wb.addWorksheet('Donnees Brutes'),             data) },
        { name: 'Par Champion',               fn: () => buildChampionSheet(wb.addWorksheet('Par Champion'),              data) },
        { name: 'Solo vs Duo',                fn: () => buildDuoSoloSheet(wb.addWorksheet('Solo vs Duo'),                data) },
        { name: 'Tendances Comportementales', fn: () => buildTendancesSheet(wb.addWorksheet('Tendances Comportementales'), data, analytics) },
    ];

    for (const { name, fn } of jobs) {
        try {
            fn();
            console.log('     OK : ' + name);
        } catch (e) {
            console.error('     ERREUR dans "' + name + '" : ' + e.message);
            console.error(e.stack);
        }
    }

    await wb.xlsx.writeFile(EXCEL_FILENAME);
    console.log('  -> Fichier sauvegarde : ' + EXCEL_FILENAME);
}

// ============================================================
// POINT D'ENTRÉE
// ============================================================
(async () => {
    try {
        console.log('1. Lecture de la base locale...');
        let data = [];
        if (fs.existsSync(JSON_FILENAME)) {
            data = JSON.parse(fs.readFileSync(JSON_FILENAME, 'utf8'));
        }
        console.log('   -> ' + data.length + ' partie(s) en memoire.');

        console.log('2. Connexion Riot Games...');
        const [player, duo] = await Promise.all([
            getPlayerData(MY_NAME, MY_TAG),
            getPlayerData(DUO_NAME, DUO_TAG),
        ]);
        console.log('   -> Joueur : ' + player.gameName + '#' + player.tagLine);
        console.log('   -> Duo    : ' + duo.gameName + '#' + duo.tagLine);

        console.log('3. Recuperation des IDs (queue: ' + (QUEUE_FILTER || 'toutes') + ')...');
        const allIds   = await getMatchIds(player.puuid);
        const savedIds = new Set(data.map(m => m.matchId));
        const newIds   = allIds.filter(id => !savedIds.has(id));

        if (newIds.length === 0) {
            console.log('\n-> Aucune nouvelle partie detectee.');
            const migrated = await migrateOldEntries(data, player.puuid, duo.puuid);
            data = recalculateTimeline(data);
            if (migrated > 0) {
                console.log('   -> Sauvegarde JSON apres migration...');
                fs.writeFileSync(JSON_FILENAME, JSON.stringify(data, null, 4));
            }
            await rebuildExcel(data);
            console.log('\n[SUCCES] Tableau de bord a jour.');
            return;
        }

        console.log('-> ' + newIds.length + ' nouvelle(s) partie(s).');
        console.log('4. Analyse des nouvelles parties...');
        let ok = 0, ko = 0;
        const newMatches = [];

        for (let i = 0; i < newIds.length; i++) {
            process.stdout.write(`   [${i+1}/${newIds.length}] ${newIds[i]}... `);
            try {
                const m = await extractMatchMetrics(newIds[i], player.puuid, duo.puuid);
                newMatches.push(m);
                ok++;
                console.log('OK  ' + m.champion + ' (' + m.role + ') — ' + m.win);
            } catch (e) {
                ko++;
                console.log('KO  ' + e.message);
            }
            await sleep(API_DELAY_MS);
        }

        console.log('\n5. Consolidation et recalcul chronologique...');
        data = recalculateTimeline(data.concat(newMatches));

        console.log('   -> Sauvegarde JSON...');
        fs.writeFileSync(JSON_FILENAME, JSON.stringify(data, null, 4));
        console.log('   -> ' + data.length + ' parties au total enregistrees.');

        const migrated = await migrateOldEntries(data, player.puuid, duo.puuid);
        if (migrated > 0) {
            fs.writeFileSync(JSON_FILENAME, JSON.stringify(data, null, 4));
        }

        console.log('6. Reconstruction Excel...');
        await rebuildExcel(data);
        console.log('\n[SUCCES] ' + ok + ' partie(s) ajoutee(s)' + (ko ? ', ' + ko + ' ignoree(s)' : '') + '.');

    } catch (err) {
        console.error('\n[ERREUR FATALE]');
        console.error(err.response ? err.response.data : err.message);
        process.exit(1);
    }
})();
