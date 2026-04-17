const axios = require('axios');
const fs = require('fs');

const {
    RIOT_API_KEY,
    REGION,
    REGION_PLATFORM,
    MY_NAME,
    MY_TAG,
    DUO_NAME,
    DUO_TAG,
    JSON_FILENAME,
    MATCHES_TO_FETCH,
    MAX_MATCHES_TO_FETCH,
    QUEUE_FILTER,
    API_DELAY_MS,
    HTTP_TIMEOUT_MS,
    rankToScore,
    scoreToLabel,
} = require('./config');

const {
    sleep,
    secToMmSs,
    bestMultiKill,
    recalculateTimeline,
} = require('./utils');

const rankCache = {};
let itemCatalogCache = null;
axios.defaults.timeout = HTTP_TIMEOUT_MS;

async function getItemCatalog() {
    if (itemCatalogCache) return itemCatalogCache;

    const versionsResp = await axios.get('https://ddragon.leagueoflegends.com/api/versions.json');
    const latest = Array.isArray(versionsResp.data) && versionsResp.data.length ? versionsResp.data[0] : null;
    if (!latest) throw new Error('Version Data Dragon introuvable');

    const itemResp = await axios.get(`https://ddragon.leagueoflegends.com/cdn/${latest}/data/fr_FR/item.json`);
    itemCatalogCache = itemResp.data?.data || {};
    return itemCatalogCache;
}

function isCoreItemId(itemId, itemCatalog) {
    if (!itemId) return false;
    const item = itemCatalog?.[String(itemId)];
    if (!item) return false;

    const tags = item.tags || [];
    const totalCost = item.gold?.total || 0;
    const purchasable = item.gold?.purchasable !== false;
    const onSummonersRift = item.maps?.['11'] !== false;
    const isConsumable = !!item.consumed || !!item.consumeOnFull || tags.includes('Consumable') || tags.includes('Trinket');
    const isStarter = tags.includes('Lane') || tags.includes('Jungle') || tags.includes('GoldPer') || tags.includes('Vision');
    const buildsIntoOtherItems = Array.isArray(item.into) && item.into.length > 0;

    return purchasable
        && onSummonersRift
        && !isConsumable
        && !isStarter
        && !buildsIntoOtherItems
        && totalCost >= 2200;
}

function computeFirstCoreItemMin(timeline, myPId, itemCatalog) {
    const candidateCorePurchases = [];

    for (const frame of timeline.info.frames) {
        const ts = frame.timestamp;
        for (const event of frame.events || []) {
            if (event.participantId !== myPId) continue;

            if (event.type === 'ITEM_PURCHASED' && isCoreItemId(event.itemId, itemCatalog)) {
                candidateCorePurchases.push({ itemId: event.itemId, ts });
            }

            if (event.type === 'ITEM_UNDO' && event.beforeId) {
                for (let i = candidateCorePurchases.length - 1; i >= 0; i--) {
                    if (candidateCorePurchases[i].itemId === event.beforeId) {
                        candidateCorePurchases.splice(i, 1);
                        break;
                    }
                }
            }
        }
    }

    if (!candidateCorePurchases.length) return null;
    return parseFloat((candidateCorePurchases[0].ts / 60_000).toFixed(1));
}

function hasSuspiciousFirstCore(entry) {
    return typeof entry.firstCoreItemMin === 'number' && entry.firstCoreItemMin > 0 && entry.firstCoreItemMin < 7;
}

function needsTimelineMigration(entry) {
    return entry.fatalGanksReceived === undefined
        || entry.csDiff10 === undefined
        || entry.goldDiff10 === undefined
        || entry.csDiff15 === undefined
        || entry.goldDiff15 === undefined
        || entry.firstCoreItemMin === undefined
        || hasSuspiciousFirstCore(entry);
}

function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
}

function normalizeTo100(value, min, max) {
    if (!Number.isFinite(value)) return 50;
    if (max <= min) return 50;
    return clamp(((value - min) / (max - min)) * 100, 0, 100);
}

function computeResponsabiliteScore(metrics) {
    if (metrics.isRemake) return null;

    const kdaScore = normalizeTo100(metrics.kda, 0.8, 5.2);
    const kpScore = normalizeTo100(metrics.kp, 20, 75);
    const dmgShareScore = normalizeTo100(metrics.dmgShare, 10, 35);
    const deadTimeScore = 100 - normalizeTo100(metrics.pctDeadTime, 2, 20);

    const gpmRatio = metrics.teamAvgGpm > 0 ? (metrics.gpm / metrics.teamAvgGpm) : 1;
    const gpmScore = normalizeTo100(gpmRatio, 0.75, 1.30);

    const pingEngaged = metrics.helpfulPings + metrics.negativePings;
    const pingQualityScore = pingEngaged > 0 ? (metrics.helpfulPings / pingEngaged) * 100 : 50;
    const pingVolumeScore = normalizeTo100(metrics.totalPings, 3, 24);
    const pingScore = clamp((pingQualityScore * 0.7) + (pingVolumeScore * 0.3), 0, 100);

    let multiKillScore = 0;
    if (metrics.pentaKills > 0) multiKillScore = 100;
    else if (metrics.quadraKills > 0) multiKillScore = 90;
    else if (metrics.tripleKills > 0) multiKillScore = 70;
    else if (metrics.doubleKills > 0) multiKillScore = 50;

    const rawScore = (kdaScore * 0.20)
        + (kpScore * 0.15)
        + (dmgShareScore * 0.15)
        + (deadTimeScore * 0.15)
        + (gpmScore * 0.15)
        + (pingScore * 0.10)
        + (multiKillScore * 0.10);

    return Math.round(clamp(rawScore, 0, 100));
}

function loadData() {
    if (!fs.existsSync(JSON_FILENAME)) return [];
    try {
        const parsed = JSON.parse(fs.readFileSync(JSON_FILENAME, 'utf8'));
        return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
        console.error(`[WARN] JSON local invalide (${JSON_FILENAME}) : ${e.message}`);
        return [];
    }
}

function saveData(data) {
    fs.writeFileSync(JSON_FILENAME, JSON.stringify(data, null, 4));
}

async function getPlayerData(name, tag) {
    const r = await axios.get(
        `https://${REGION}.api.riotgames.com/riot/account/v1/accounts/by-riot-id/${encodeURIComponent(name)}/${encodeURIComponent(tag)}`,
        { headers: { 'X-Riot-Token': RIOT_API_KEY } },
    );
    return r.data;
}

async function resolvePlayers() {
    const [player, duo] = await Promise.all([
        getPlayerData(MY_NAME, MY_TAG),
        getPlayerData(DUO_NAME, DUO_TAG),
    ]);
    return { player, duo };
}

function normalizeMatchLimit(value) {
    const parsed = Number.parseInt(value, 10);
    if (!Number.isFinite(parsed) || parsed <= 0) return MATCHES_TO_FETCH;
    return Math.min(parsed, MAX_MATCHES_TO_FETCH);
}

async function getMatchIds(puuid, matchesToFetch = MATCHES_TO_FETCH) {
    const targetCount = normalizeMatchLimit(matchesToFetch);
    let allIds = [];
    let start = 0;
    const limit = 100;

    while (allIds.length < targetCount) {
        const count = Math.min(limit, targetCount - allIds.length);
        let url = `https://${REGION}.api.riotgames.com/lol/match/v5/matches/by-puuid/${puuid}/ids?start=${start}&count=${count}`;
        if (QUEUE_FILTER) url += `&queue=${QUEUE_FILTER}`;

        const r = await axios.get(url, { headers: { 'X-Riot-Token': RIOT_API_KEY } });
        const ids = r.data;
        if (ids.length === 0) break;

        allIds = allIds.concat(ids);
        start += count;

        if (allIds.length < targetCount) await sleep(API_DELAY_MS);
    }

    return allIds;
}

async function getMatchTimeline(matchId) {
    const r = await axios.get(
        `https://${REGION}.api.riotgames.com/lol/match/v5/matches/${matchId}/timeline`,
        { headers: { 'X-Riot-Token': RIOT_API_KEY } },
    );
    return r.data;
}

// Utilise le PUUID (endpoint moderne) — summonerId déprécié dans Match v5
async function fetchRankForPuuid(puuid) {
    if (!puuid) return { tier: null, division: null, score: -1, label: 'Non classé' };
    if (rankCache[puuid] !== undefined) return rankCache[puuid];

    try {
        const r = await axios.get(
            `https://${REGION_PLATFORM}.api.riotgames.com/lol/league/v4/entries/by-puuid/${puuid}`,
            { headers: { 'X-Riot-Token': RIOT_API_KEY } },
        );

        const soloQ = r.data.find((e) => e.queueType === 'RANKED_SOLO_5x5');
        if (!soloQ) {
            rankCache[puuid] = { tier: null, division: null, score: -1, label: 'Non classé' };
        } else {
            const score = rankToScore(soloQ.tier, soloQ.rank);
            rankCache[puuid] = {
                tier: soloQ.tier,
                division: soloQ.rank,
                score,
                label: scoreToLabel(score),
            };
        }
    } catch (e) {
        if (e.response && e.response.status === 429) {
            const retryAfter = e.response.headers['retry-after']
                ? parseInt(e.response.headers['retry-after'], 10) * 1000
                : 10000;
            console.log(`\n      [!] Limite API (Rangs) atteinte. Attente de ${retryAfter / 1000}s...`);
            await sleep(retryAfter + 500);
            return fetchRankForPuuid(puuid);
        }
        rankCache[puuid] = { tier: null, division: null, score: -1, label: 'Non classé' };
    }

    return rankCache[puuid];
}

async function extractMatchMetrics(matchId, myPuuid, duoPuuid) {
    const r = await axios.get(
        `https://${REGION}.api.riotgames.com/lol/match/v5/matches/${matchId}`,
        { headers: { 'X-Riot-Token': RIOT_API_KEY } },
    );

    const match = r.data;
    const me = match.info.participants.find((p) => p.puuid === myPuuid);
    if (!me) throw new Error(`PUUID introuvable dans ${matchId}`);

    const duo = match.info.participants.find((p) => p.puuid === duoPuuid && p.teamId === me.teamId);
    const isDuo = !!duo;
    const myTeam = match.info.participants.filter((p) => p.teamId === me.teamId);
    const enemies = match.info.participants.filter((p) => p.teamId !== me.teamId);

    const minutes = match.info.gameDuration / 60;
    const teamKills = myTeam.reduce((s, p) => s + p.kills, 0);
    const teamDamage = myTeam.reduce((s, p) => s + p.totalDamageDealtToChampions, 0);
    const kp = teamKills === 0 ? 0 : ((me.kills + me.assists) / teamKills) * 100;
    const dmgShare = teamDamage === 0 ? 0 : (me.totalDamageDealtToChampions / teamDamage) * 100;
    const totalCS = (me.totalMinionsKilled || 0) + (me.neutralMinionsKilled || 0);
    const gameDate = new Date(match.info.gameCreation);

    const role = me.teamPosition || me.individualPosition || '-';
    const ganksPerformed = role === 'JUNGLE'
        ? (me.challenges?.killsOnLanersEarlyJungleAsJungler ?? null)
        : null;
    const dpg = parseFloat((me.totalDamageDealtToChampions / Math.max(1, me.goldEarned)).toFixed(3));
    const dmgRatio = parseFloat((me.totalDamageDealtToChampions / Math.max(1, me.totalDamageTaken)).toFixed(2));

    const enemyADC = enemies.find((p) => p.teamPosition === 'BOTTOM')?.championName || null;

    const allInPings = me.allInPings || 0;
    const basicPings = me.basicPings || 0;
    const commandPings = me.commandPings || 0;
    const dangerPings = me.dangerPings || 0;
    const enemyMissingPings = me.enemyMissingPings || 0;
    const holdPings = me.holdPings || 0;
    const needVisionPings = me.needVisionPings || 0;
    const onMyWayPings = me.onMyWayPings || 0;
    const pushPings = me.pushPings || 0;
    const retreatPings = me.retreatPings || 0;
    const visionClearedPings = me.visionClearedPings || 0;

    const totalPings = allInPings + basicPings + commandPings + dangerPings
        + enemyMissingPings + holdPings + needVisionPings
        + onMyWayPings + pushPings + retreatPings + visionClearedPings;
    const negativePings = dangerPings + retreatPings + enemyMissingPings;
    const helpfulPings = allInPings + commandPings + holdPings + needVisionPings
        + onMyWayPings + pushPings + visionClearedPings;

    console.log('   -> Analyse de la Timeline (Ganks, Lane phase, Core item)...');
    let fatalGanksReceived = 0;
    let csDiff10 = null;
    let goldDiff10 = null;
    let csDiff15 = null;
    let goldDiff15 = null;
    let firstCoreItemMin = null;

    try {
        const timeline = await getMatchTimeline(matchId);
        const myPId = me.participantId;
        const itemCatalog = await getItemCatalog();

        const enemyJungler = enemies.find((p) => p.teamPosition === 'JUNGLE');
        const enemyJunglerId = enemyJungler ? enemyJungler.participantId : null;

        const enemyLaner = enemies.find((p) => p.teamPosition === me.teamPosition);
        const enemyLanerId = enemyLaner ? enemyLaner.participantId : null;

        const MS_10 = 10 * 60 * 1000;
        const MS_15 = 15 * 60 * 1000;
        let snap10 = null;
        let snap15 = null;

        for (const frame of timeline.info.frames) {
            const ts = frame.timestamp;
            if (ts <= MS_10) snap10 = frame;
            if (ts <= MS_15) snap15 = frame;

            if (ts <= 14 * 60 * 1000 && enemyJunglerId) {
                for (const event of frame.events) {
                    if (event.type === 'CHAMPION_KILL' && event.victimId === myPId) {
                        if (event.killerId === enemyJunglerId || (event.assistingParticipantIds || []).includes(enemyJunglerId)) {
                            fatalGanksReceived++;
                        }
                    }
                }
            }

        }

        firstCoreItemMin = computeFirstCoreItemMin(timeline, myPId, itemCatalog);

        function laneStats(frame, pid) {
            const pf = frame?.participantFrames?.[String(pid)];
            if (!pf) return null;
            return {
                cs: (pf.minionsKilled || 0) + (pf.jungleMinionsKilled || 0),
                gold: pf.totalGold || 0,
            };
        }

        if (snap10 && enemyLanerId) {
            const me10 = laneStats(snap10, myPId);
            const enemy10 = laneStats(snap10, enemyLanerId);
            if (me10 && enemy10) {
                csDiff10 = me10.cs - enemy10.cs;
                goldDiff10 = me10.gold - enemy10.gold;
            }
        }

        if (snap15 && enemyLanerId) {
            const me15 = laneStats(snap15, myPId);
            const enemy15 = laneStats(snap15, enemyLanerId);
            if (me15 && enemy15) {
                csDiff15 = me15.cs - enemy15.cs;
                goldDiff15 = me15.gold - enemy15.gold;
            }
        }
    } catch (e) {
        console.log(`      (Avertissement : Timeline indisponible — ${e.message})`);
        fatalGanksReceived = null;
    }

    const allyPuuids  = myTeam.filter((p) => p.puuid !== myPuuid).map((p) => p.puuid).filter(Boolean);
    const enemyPuuids = enemies.map((p) => p.puuid).filter(Boolean);

    console.log('   -> Récupération rangs alliés...');
    const allyRanks = [];
    for (const puuid of allyPuuids) {
        await sleep(350);
        const rk = await fetchRankForPuuid(puuid);
        if (rk.score >= 0) allyRanks.push(rk.score);
    }

    const myRankEntry = await fetchRankForPuuid(myPuuid);
    const avgAllyScore = allyRanks.length
        ? parseFloat((allyRanks.reduce((s, v) => s + v, 0) / allyRanks.length).toFixed(2))
        : -1;

    console.log('   -> Récupération rangs ennemis...');
    const enemyRanks = [];
    for (const puuid of enemyPuuids) {
        await sleep(350);
        const rk = await fetchRankForPuuid(puuid);
        if (rk.score >= 0) enemyRanks.push(rk.score);
    }

    const avgEnemyScore = enemyRanks.length
        ? parseFloat((enemyRanks.reduce((s, v) => s + v, 0) / enemyRanks.length).toFixed(2))
        : -1;

    const rankDiff = (avgAllyScore >= 0 && avgEnemyScore >= 0)
        ? parseFloat((avgAllyScore - avgEnemyScore).toFixed(2))
        : null;

    const pctDeadTime = parseFloat((((me.totalTimeSpentDead || 0) / match.info.gameDuration) * 100).toFixed(1));
    const kda = parseFloat(((me.kills + me.assists) / Math.max(1, me.deaths)).toFixed(2));
    const gpm = parseFloat((me.goldEarned / minutes).toFixed(0));
    const teamAvgGpm = myTeam.length
        ? myTeam.reduce((sum, p) => sum + ((p.goldEarned || 0) / minutes), 0) / myTeam.length
        : gpm;
    const isRemake = Boolean(me.gameEndedInEarlySurrender) || match.info.gameDuration <= 300;

    const responsabilite = computeResponsabiliteScore({
        isRemake,
        kda,
        kp,
        dmgShare,
        pctDeadTime,
        gpm,
        teamAvgGpm,
        helpfulPings,
        negativePings,
        totalPings,
        pentaKills: me.pentaKills || 0,
        quadraKills: me.quadraKills || 0,
        tripleKills: me.tripleKills || 0,
        doubleKills: me.doubleKills || 0,
    });

    return {
        matchId,
        rawDate: gameDate.toISOString(),
        date: gameDate.toLocaleDateString('fr-FR'),
        time: gameDate.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }),
        sessionId: 0,
        gameInSession: 0,
        gameDuration: secToMmSs(match.info.gameDuration),
        gameDurationSec: match.info.gameDuration,
        isRemake,
        champion: me.championName,
        role,
        type: isDuo ? 'Duo' : 'Solo',
        win: me.win ? 'Victoire' : 'Defaite',
        kda,
        kills: me.kills,
        deaths: me.deaths,
        assists: me.assists,
        csPerMin: parseFloat((totalCS / minutes).toFixed(1)),
        gpm,
        dpm: parseFloat((me.totalDamageDealtToChampions / minutes).toFixed(0)),
        objDpm: parseFloat(((me.damageDealtToObjectives || 0) / minutes).toFixed(0)),
        dmgShare: parseFloat(dmgShare.toFixed(1)),
        kp: parseFloat(kp.toFixed(1)),
        vision: me.visionScore || 0,
        wardsPlaced: me.wardsPlaced || 0,
        wardsKilled: me.wardsKilled || 0,
        controlWards: me.visionWardsBoughtInGame || 0,
        pctDeadTime,
        bestMultiKill: bestMultiKill(me),
        pentaKills: me.pentaKills || 0,
        quadraKills: me.quadraKills || 0,
        tripleKills: me.tripleKills || 0,
        doubleKills: me.doubleKills || 0,
        largestKillingSpree: me.largestKillingSpree || 0,
        firstBlood: me.firstBloodKill ? 'Oui' : 'Non',
        withYuumi: (isDuo && duo.championName === 'Yuumi') ? 'Oui' : 'Non',
        yuumiAlliee: myTeam.some((p) => p.puuid !== myPuuid && p.championName === 'Yuumi') ? 'Oui' : 'Non',

        totalPings,
        helpfulPings,
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

        ganksPerformed,
        fatalGanksReceived,

        csDiff10,
        goldDiff10,
        csDiff15,
        goldDiff15,

        enemyADC,
        dpg,
        dmgRatio,
        firstCoreItemMin,

        myRank: myRankEntry.label,
        myRankScore: myRankEntry.score,
        avgAllyRank: scoreToLabel(Math.round(avgAllyScore)),
        avgAllyRankScore: avgAllyScore,
        avgEnemyRank: scoreToLabel(Math.round(avgEnemyScore)),
        avgEnemyRankScore: avgEnemyScore,
        rankDiff,
        responsabilite,
    };
}

function needsMigration(m) {
    return m.gameDurationSec == null
        || m.isRemake == null
        || m.yuumiAlliee == null
        || m.firstBlood == null
        || m.pctDeadTime == null
        || m.objDpm == null
        || m.role == null
        || m.totalPings == null
        || m.myRankScore == null
        || m.fatalGanksReceived === undefined
        || m.csDiff10 === undefined
        || m.dpg === undefined
        || m.dmgRatio === undefined
        || m.firstCoreItemMin === undefined
        || hasSuspiciousFirstCore(m)
        || m.enemyADC === undefined;
}

async function migrateOldEntries(data, myPuuid, duoPuuid) {
    const toFix = data.filter(needsMigration);
    if (toFix.length === 0) return 0;

    console.log(`   -> ${toFix.length} entree(s) incomplete(s). Migration en cours...`);
    let ok = 0;
    let ko = 0;

    for (let i = 0; i < toFix.length; i++) {
        const entry = toFix[i];
        process.stdout.write(`   [Migration ${i + 1}/${toFix.length}] ${entry.matchId}... `);

        try {
            const r = await axios.get(
                `https://${REGION}.api.riotgames.com/lol/match/v5/matches/${entry.matchId}`,
                { headers: { 'X-Riot-Token': RIOT_API_KEY } },
            );

            const match = r.data;
            const me = match.info.participants.find((p) => p.puuid === myPuuid);
            if (!me) throw new Error('PUUID introuvable');

            const duo = match.info.participants.find((p) => p.puuid === duoPuuid && p.teamId === me.teamId);
            const myTeam = match.info.participants.filter((p) => p.teamId === me.teamId);
            const minutes = match.info.gameDuration / 60;

            entry.gameDurationSec = match.info.gameDuration;
            entry.gameDuration = secToMmSs(match.info.gameDuration);
            entry.isRemake = Boolean(me.gameEndedInEarlySurrender) || match.info.gameDuration <= 300;
            entry.objDpm = parseFloat(((me.damageDealtToObjectives || 0) / minutes).toFixed(0));
            entry.wardsPlaced = me.wardsPlaced || 0;
            entry.wardsKilled = me.wardsKilled || 0;
            entry.controlWards = me.visionWardsBoughtInGame || 0;
            entry.pctDeadTime = parseFloat((((me.totalTimeSpentDead || 0) / match.info.gameDuration) * 100).toFixed(1));
            entry.bestMultiKill = bestMultiKill(me);
            entry.pentaKills = me.pentaKills || 0;
            entry.quadraKills = me.quadraKills || 0;
            entry.tripleKills = me.tripleKills || 0;
            entry.doubleKills = me.doubleKills || 0;
            entry.largestKillingSpree = me.largestKillingSpree || 0;
            entry.firstBlood = me.firstBloodKill ? 'Oui' : 'Non';
            entry.yuumiAlliee = myTeam.some((p) => p.puuid !== myPuuid && p.championName === 'Yuumi') ? 'Oui' : 'Non';
            entry.withYuumi = (!!duo && duo.championName === 'Yuumi') ? 'Oui' : 'Non';

            entry.role = me.teamPosition || me.individualPosition || '-';

            entry.allInPings = me.allInPings || 0;
            entry.basicPings = me.basicPings || 0;
            entry.commandPings = me.commandPings || 0;
            entry.dangerPings = me.dangerPings || 0;
            entry.enemyMissingPings = me.enemyMissingPings || 0;
            entry.holdPings = me.holdPings || 0;
            entry.needVisionPings = me.needVisionPings || 0;
            entry.onMyWayPings = me.onMyWayPings || 0;
            entry.pushPings = me.pushPings || 0;
            entry.retreatPings = me.retreatPings || 0;
            entry.visionClearedPings = me.visionClearedPings || 0;
            entry.totalPings = entry.allInPings + entry.basicPings + entry.commandPings
                + entry.dangerPings + entry.enemyMissingPings + entry.holdPings
                + entry.needVisionPings + entry.onMyWayPings + entry.pushPings
                + entry.retreatPings + entry.visionClearedPings;
            entry.helpfulPings = entry.allInPings + entry.commandPings + entry.holdPings
                + entry.needVisionPings + entry.onMyWayPings + entry.pushPings + entry.visionClearedPings;
            entry.negativePings = entry.dangerPings + entry.retreatPings + entry.enemyMissingPings;

            entry.ganksPerformed = entry.role === 'JUNGLE'
                ? (me.challenges?.killsOnLanersEarlyJungleAsJungler ?? null)
                : null;

            const teamKills = myTeam.reduce((s, p) => s + (p.kills || 0), 0);
            const teamDamage = myTeam.reduce((s, p) => s + (p.totalDamageDealtToChampions || 0), 0);
            const kp = teamKills === 0 ? 0 : ((me.kills + me.assists) / teamKills) * 100;
            const dmgShare = teamDamage === 0 ? 0 : (me.totalDamageDealtToChampions / teamDamage) * 100;
            const kda = parseFloat(((me.kills + me.assists) / Math.max(1, me.deaths)).toFixed(2));
            const gpm = parseFloat((me.goldEarned / minutes).toFixed(0));
            const teamAvgGpm = myTeam.length
                ? myTeam.reduce((sum, p) => sum + ((p.goldEarned || 0) / minutes), 0) / myTeam.length
                : gpm;

            entry.kp = parseFloat(kp.toFixed(1));
            entry.dmgShare = parseFloat(dmgShare.toFixed(1));
            entry.kda = kda;
            entry.gpm = gpm;
            entry.pctDeadTime = parseFloat((((me.totalTimeSpentDead || 0) / match.info.gameDuration) * 100).toFixed(1));
            entry.responsabilite = computeResponsabiliteScore({
                isRemake: entry.isRemake,
                kda: entry.kda,
                kp,
                dmgShare,
                pctDeadTime: entry.pctDeadTime,
                gpm,
                teamAvgGpm,
                helpfulPings: entry.helpfulPings,
                negativePings: entry.negativePings,
                totalPings: entry.totalPings,
                pentaKills: entry.pentaKills || 0,
                quadraKills: entry.quadraKills || 0,
                tripleKills: entry.tripleKills || 0,
                doubleKills: entry.doubleKills || 0,
            });

            if (entry.dpg === undefined) {
                entry.dpg = parseFloat((me.totalDamageDealtToChampions / Math.max(1, me.goldEarned)).toFixed(3));
                entry.dmgRatio = parseFloat((me.totalDamageDealtToChampions / Math.max(1, me.totalDamageTaken)).toFixed(2));
            }

            if (entry.enemyADC === undefined) {
                const enemies = match.info.participants.filter((p) => p.teamId !== me.teamId);
                entry.enemyADC = enemies.find((p) => p.teamPosition === 'BOTTOM')?.championName || null;
            }

            if (entry.myRankScore == null) {
                await sleep(350);
                const myRk = await fetchRankForPuuid(myPuuid);
                entry.myRank = myRk.label;
                entry.myRankScore = myRk.score;

                const enemies = match.info.participants.filter((p) => p.teamId !== me.teamId);
                const allyPlayers = myTeam.filter((p) => p.puuid !== myPuuid);
                const allyScores = [];
                const enemyScores = [];

                for (const p of allyPlayers) {
                    if (!p.puuid) continue;
                    await sleep(300);
                    const rk = await fetchRankForPuuid(p.puuid);
                    if (rk.score >= 0) allyScores.push(rk.score);
                }

                for (const p of enemies) {
                    if (!p.puuid) continue;
                    await sleep(300);
                    const rk = await fetchRankForPuuid(p.puuid);
                    if (rk.score >= 0) enemyScores.push(rk.score);
                }

                const avgA = allyScores.length ? allyScores.reduce((s, v) => s + v, 0) / allyScores.length : -1;
                const avgE = enemyScores.length ? enemyScores.reduce((s, v) => s + v, 0) / enemyScores.length : -1;
                entry.avgAllyRank = scoreToLabel(Math.round(avgA));
                entry.avgAllyRankScore = avgA >= 0 ? parseFloat(avgA.toFixed(2)) : -1;
                entry.avgEnemyRank = scoreToLabel(Math.round(avgE));
                entry.avgEnemyRankScore = avgE >= 0 ? parseFloat(avgE.toFixed(2)) : -1;
                entry.rankDiff = (avgA >= 0 && avgE >= 0) ? parseFloat((avgA - avgE).toFixed(2)) : null;
            }

            if (needsTimelineMigration(entry)) {
                try {
                    const timeline = await getMatchTimeline(entry.matchId);
                    const myPId = me.participantId;
                    const itemCatalog = await getItemCatalog();
                    const enemies = match.info.participants.filter((p) => p.teamId !== me.teamId);
                    const enemyLaner = enemies.find((p) => p.teamPosition === entry.role);
                    const enemyLanerId = enemyLaner ? enemyLaner.participantId : null;
                    const MS_10 = 10 * 60 * 1000;
                    const MS_15 = 15 * 60 * 1000;
                    let snap10 = null;
                    let snap15 = null;

                    entry.fatalGanksReceived = entry.fatalGanksReceived ?? 0;
                    const enemyJungler = enemies.find((p) => p.teamPosition === 'JUNGLE');
                    const enemyJunglerId = enemyJungler ? enemyJungler.participantId : null;

                    for (const frame of timeline.info.frames) {
                        const ts = frame.timestamp;
                        if (ts <= MS_10) snap10 = frame;
                        if (ts <= MS_15) snap15 = frame;

                        if (ts <= 14 * 60 * 1000 && enemyJunglerId) {
                            for (const event of frame.events) {
                                if (
                                    event.type === 'CHAMPION_KILL'
                                    && event.victimId === myPId
                                    && (event.killerId === enemyJunglerId || (event.assistingParticipantIds || []).includes(enemyJunglerId))
                                ) {
                                    entry.fatalGanksReceived++;
                                }
                            }
                        }

                    }

                    entry.firstCoreItemMin = computeFirstCoreItemMin(timeline, myPId, itemCatalog);

                    function laneStats(frame, pid) {
                        const pf = frame?.participantFrames?.[String(pid)];
                        if (!pf) return null;
                        return { cs: (pf.minionsKilled || 0) + (pf.jungleMinionsKilled || 0), gold: pf.totalGold || 0 };
                    }

                    if (snap10 && enemyLanerId) {
                        const m10 = laneStats(snap10, myPId);
                        const e10 = laneStats(snap10, enemyLanerId);
                        if (m10 && e10) {
                            entry.csDiff10 = m10.cs - e10.cs;
                            entry.goldDiff10 = m10.gold - e10.gold;
                        }
                    }

                    if (snap15 && enemyLanerId) {
                        const m15 = laneStats(snap15, myPId);
                        const e15 = laneStats(snap15, enemyLanerId);
                        if (m15 && e15) {
                            entry.csDiff15 = m15.cs - e15.cs;
                            entry.goldDiff15 = m15.gold - e15.gold;
                        }
                    }

                    entry.csDiff10 = entry.csDiff10 ?? null;
                    entry.goldDiff10 = entry.goldDiff10 ?? null;
                    entry.csDiff15 = entry.csDiff15 ?? null;
                    entry.goldDiff15 = entry.goldDiff15 ?? null;
                    await sleep(API_DELAY_MS);
                } catch (e) {
                    entry.csDiff10 = entry.csDiff10 ?? null;
                    entry.goldDiff10 = entry.goldDiff10 ?? null;
                    entry.csDiff15 = entry.csDiff15 ?? null;
                    entry.goldDiff15 = entry.goldDiff15 ?? null;
                    entry.firstCoreItemMin = entry.firstCoreItemMin ?? null;
                }
            }

            ok++;
            console.log('OK');
        } catch (e) {
            ko++;
            console.log(`KO (${e.message})`);
        }

        await sleep(API_DELAY_MS);
    }

    console.log(`   -> Migration : ${ok} corrigee(s)${ko ? `, ${ko} echec(s)` : ''}.`);
    return ok;
}

let programmaticStopRequested = false;

function requestStop() {
    programmaticStopRequested = true;
}

async function fetchNewMatches(data, playerPuuid, duoPuuid, options = {}) {
    programmaticStopRequested = false; // Reset au lancement
    const allIds = await getMatchIds(playerPuuid, options.matchesToFetch);
    const savedIds = new Set(data.map((m) => m.matchId));
    const newIds = allIds.filter((id) => !savedIds.has(id));

    if (newIds.length === 0) {
        return {
            newIds,
            newMatches: [],
            ok: 0,
            ko: 0,
            stopped: false,
        };
    }

    let ok = 0;
    let ko = 0;
    let stopRequested = false;
    const newMatches = [];

    const onSigint = () => {
        console.log('\n[!] Interruption (Ctrl+C) détectée. Arrêt du processus après le téléchargement du match en cours...');
        stopRequested = true;
    };
    process.on('SIGINT', onSigint);

    try {
        for (let i = 0; i < newIds.length; i++) {
            if (stopRequested || programmaticStopRequested) {
                console.log('\n[!] Arret du fetch. Sauvegarde des donnees recuperees en cours...');
                break;
            }

            const progressPercent = Math.round(((i + 1) / newIds.length) * 100);
            process.stdout.write(`   [${i + 1}/${newIds.length} - ${progressPercent}%] ${newIds[i]}... `);
            try {
                const m = await extractMatchMetrics(newIds[i], playerPuuid, duoPuuid);
                newMatches.push(m);
                ok++;
                console.log(`OK  ${m.champion} (${m.role}) — ${m.win}`);
            } catch (e) {
                ko++;
                console.log(`KO  ${e.message}`);
            }
            await sleep(API_DELAY_MS);
        }
    } finally {
        process.removeListener('SIGINT', onSigint);
    }

    return { newIds, newMatches, ok, ko, stopped: stopRequested || programmaticStopRequested };
}

module.exports = {
    loadData,
    saveData,
    resolvePlayers,
    fetchNewMatches,
    migrateOldEntries,
    recalculateTimeline,
    requestStop,
};
