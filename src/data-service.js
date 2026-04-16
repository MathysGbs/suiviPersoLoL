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
    QUEUE_FILTER,
    API_DELAY_MS,
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

function loadData() {
    if (!fs.existsSync(JSON_FILENAME)) return [];
    return JSON.parse(fs.readFileSync(JSON_FILENAME, 'utf8'));
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
    return parsed;
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

            if (firstCoreItemMin === null) {
                for (const event of frame.events) {
                    if (event.type === 'ITEM_PURCHASED' && event.participantId === myPId && ts >= 90_000) {
                        firstCoreItemMin = parseFloat((ts / 60_000).toFixed(1));
                        break;
                    }
                }
            }
        }

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
    
    // Calcul de la Responsabilité
    let resp = 50;
    
    // Impact KDA
    if (kda < 1.0) resp -= 20;
    else if (kda < 2.0) resp -= 10;
    else if (kda >= 3.0 && kda < 4.0) resp += 10;
    else if (kda >= 4.0) resp += 20;

    // Impact KP
    if (kp < 30) resp -= 15;
    else if (kp < 40) resp -= 5;
    else if (kp >= 50 && kp < 60) resp += 10;
    else if (kp >= 60) resp += 15;

    // Impact DMG %
    if (dmgShare < 15) resp -= 15;
    else if (dmgShare < 20) resp -= 5;
    else if (dmgShare >= 20 && dmgShare < 25) resp += 5;
    else if (dmgShare >= 25 && dmgShare < 30) resp += 10;
    else if (dmgShare >= 30) resp += 15;

    // Impact Temps mort
    if (pctDeadTime > 15) resp -= 15;
    else if (pctDeadTime > 10) resp -= 10;
    else if (pctDeadTime < 5) resp += 10;
    
    const responsabilite = Math.max(0, Math.min(100, Math.round(resp)));

    return {
        matchId,
        rawDate: gameDate.toISOString(),
        date: gameDate.toLocaleDateString('fr-FR'),
        time: gameDate.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }),
        sessionId: 0,
        gameInSession: 0,
        gameDuration: secToMmSs(match.info.gameDuration),
        gameDurationSec: match.info.gameDuration,
        champion: me.championName,
        role,
        type: isDuo ? 'Duo' : 'Solo',
        win: me.win ? 'Victoire' : 'Defaite',
        kda,
        kills: me.kills,
        deaths: me.deaths,
        assists: me.assists,
        csPerMin: parseFloat((totalCS / minutes).toFixed(1)),
        gpm: parseFloat((me.goldEarned / minutes).toFixed(0)),
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
            entry.negativePings = entry.dangerPings + entry.retreatPings + entry.enemyMissingPings;

            entry.ganksPerformed = entry.role === 'JUNGLE'
                ? (me.challenges?.killsOnLanersEarlyJungleAsJungler ?? null)
                : null;

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

            if (entry.csDiff10 === undefined) {
                try {
                    const timeline = await getMatchTimeline(entry.matchId);
                    const myPId = me.participantId;
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

                        if (entry.firstCoreItemMin == null) {
                            for (const event of frame.events) {
                                if (event.type === 'ITEM_PURCHASED' && event.participantId === myPId && ts >= 90_000) {
                                    entry.firstCoreItemMin = parseFloat((ts / 60_000).toFixed(1));
                                    break;
                                }
                            }
                        }
                    }

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
                    entry.csDiff10 = null;
                    entry.goldDiff10 = null;
                    entry.csDiff15 = null;
                    entry.goldDiff15 = null;
                    entry.firstCoreItemMin = null;
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

    process.removeListener('SIGINT', onSigint);

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
