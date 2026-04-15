const axios  = require('axios');
const ExcelJS = require('exceljs');
const fs      = require('fs');
require('dotenv').config();

// ============================================================
// CONFIGURATION
// ============================================================
const RIOT_API_KEY     = process.env.RIOT_API_KEY;
const REGION           = 'europe';

const MY_NAME          = 'enyαV';
const MY_TAG           = 'EUW';
const DUO_NAME         = 'Ahristocats';
const DUO_TAG          = 'EUW';

const EXCEL_FILENAME   = 'Suivi_Comportemental_Challenger.xlsx';
const JSON_FILENAME    = 'historique_matches.json';
const MATCHES_TO_FETCH = 50;
const QUEUE_FILTER     = 420;   // Ranked Solo/Duo uniquement. null = toutes files.
const API_DELAY_MS     = 1500;

// ============================================================
// PALETTE COULEURS (ARGB)
// ============================================================
const C = {
    // Headers par feuille
    hRaw:    'FF0F172A',   // Donnees Brutes
    hChamp:  'FF2E1065',   // Par Champion
    hDuo:    'FF172554',   // Solo vs Duo
    hTend:   'FF052E16',   // Tendances
    subH:    'FF1E3A5F',   // Sous-header interne
    white:   'FFFFFFFF',
    accent:  'FF4F46E5',   // indigo

    // Resultat
    winFg:   'FF166534', winBg:   'FFDCFCE7',
    lossFg:  'FF991B1B', lossBg:  'FFFEE2E2',

    // Alternance de lignes
    rowA:    'FFF9FAFB',
    rowB:    'FFFFFFFF',

    // Couleurs metriques (coloring manuel par seuil)
    good:    'FFD9F2DC',  // vert clair
    warn:    'FFFFF9C4',  // jaune clair
    bad:     'FFFDE8E8',  // rouge clair
    goodFg:  'FF166534',
    badFg:   'FF991B1B',
    purple:  'FFA855F7',
    amber:   'FFB45309',
};

// ============================================================
// UTILITAIRES
// ============================================================
const sleep  = ms  => new Promise(r => setTimeout(r, ms));
const safeN  = (v, d = 0)   => { const n = parseFloat(v); return isNaN(n) ? d : n; };
const safeV  = (v, d = '-') => (v !== undefined && v !== null) ? v : d;
const avgOf  = (arr, k)     => arr.length ? arr.reduce((s, x) => s + safeN(x[k]), 0) / arr.length : 0;
const winPct = arr           => arr.length ? (arr.filter(g => g.win === 'Victoire').length / arr.length) * 100 : 0;
const rowBg  = idx           => idx % 2 === 0 ? C.rowA : C.rowB;

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

function sessionInfo(currentDate, prev) {
    if (!prev) return { sessionId: 1, gameInSession: 1 };
    const diffH = (new Date(currentDate) - new Date(prev.rawDate)) / 3_600_000;
    return diffH < 2
        ? { sessionId: prev.sessionId,     gameInSession: prev.gameInSession + 1 }
        : { sessionId: prev.sessionId + 1, gameInSession: 1 };
}

// ============================================================
// ANALYSE COMPORTEMENTALE
// ============================================================
function computeAnalytics(data) {
    if (!data.length) return null;

    const champMap        = {};
    const byGameInSession = {};
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
        (data[i - 1].win === 'Victoire' ? afterWin : afterLoss).push(data[i]);
    }

    let streakCount = 0, streakType = '';
    for (let i = data.length - 1; i >= 0; i--) {
        const isW = data[i].win === 'Victoire';
        if (i === data.length - 1) {
            streakType = isW ? 'Victoires' : 'Defaites'; streakCount = 1;
        } else if ((isW && streakType === 'Victoires') || (!isW && streakType === 'Defaites')) {
            streakCount++;
        } else break;
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

    return {
        byChampion,
        soloGames:       data.filter(m => m.type === 'Solo'),
        duoGames:        data.filter(m => m.type === 'Duo'),
        yuumiGames:      data.filter(m => m.withYuumi === 'Oui'),
        noYuumi:         data.filter(m => m.type === 'Duo' && m.withYuumi !== 'Oui'),
        yuumiAlliee:     data.filter(m => m.yuumiAlliee === 'Oui'),
        byGameInSession,
        timeSlots,
        afterWin, afterLoss,
        streak:   { count: streakCount, type: streakType },
        recent10: data.slice(-10),
        records,
        overall: {
            total:       data.length,
            winRate:     winPct(data),
            avgKDA:      avgOf(data, 'kda'),
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
    let url = `https://${REGION}.api.riotgames.com/lol/match/v5/matches/by-puuid/${puuid}/ids?start=0&count=${MATCHES_TO_FETCH}`;
    if (QUEUE_FILTER) url += `&queue=${QUEUE_FILTER}`;
    const r = await axios.get(url, { headers: { 'X-Riot-Token': RIOT_API_KEY } });
    return r.data;
}

async function extractMatchMetrics(matchId, myPuuid, duoPuuid, prev) {
    const r = await axios.get(
        `https://${REGION}.api.riotgames.com/lol/match/v5/matches/${matchId}`,
        { headers: { 'X-Riot-Token': RIOT_API_KEY } }
    );
    const match = r.data;
    const me = match.info.participants.find(p => p.puuid === myPuuid);
    if (!me) throw new Error(`PUUID introuvable dans ${matchId}`);

    const duo   = match.info.participants.find(p => p.puuid === duoPuuid && p.teamId === me.teamId);
    const isDuo = !!duo;

    const minutes    = match.info.gameDuration / 60;
    const myTeam     = match.info.participants.filter(p => p.teamId === me.teamId);
    const teamKills  = myTeam.reduce((s, p) => s + p.kills, 0);
    const teamDamage = myTeam.reduce((s, p) => s + p.totalDamageDealtToChampions, 0);
    const kp         = teamKills  === 0 ? 0 : ((me.kills + me.assists) / teamKills)  * 100;
    const dmgShare   = teamDamage === 0 ? 0 : (me.totalDamageDealtToChampions / teamDamage) * 100;
    const totalCS    = (me.totalMinionsKilled || 0) + (me.neutralMinionsKilled || 0);
    const gameDate   = new Date(match.info.gameCreation);
    const sess       = sessionInfo(gameDate, prev);

    return {
        matchId,
        rawDate:             gameDate.toISOString(),
        date:                gameDate.toLocaleDateString('fr-FR'),
        time:                gameDate.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }),
        sessionId:           sess.sessionId,
        gameInSession:       sess.gameInSession,
        gameDuration:        parseFloat(minutes.toFixed(1)),
        champion:            me.championName,
        type:                isDuo ? 'Duo' : 'Solo',
        win:                 me.win ? 'Victoire' : 'Defaite',
        kda:                 parseFloat(((me.kills + me.assists) / Math.max(1, me.deaths)).toFixed(2)),
        kills:               me.kills,
        deaths:              me.deaths,
        assists:             me.assists,
        csPerMin:            parseFloat((totalCS / minutes).toFixed(1)),
        gpm:                 parseFloat((me.goldEarned / minutes).toFixed(0)),
        dpm:                 parseFloat((me.totalDamageDealtToChampions / minutes).toFixed(0)),
        objDpm:              parseFloat(((me.damageDealtToObjectives || 0) / minutes).toFixed(0)),
        dmgShare:            parseFloat(dmgShare.toFixed(1)),
        kp:                  parseFloat(kp.toFixed(1)),
        vision:              me.visionScore || 0,
        wardsPlaced:         me.wardsPlaced || 0,
        wardsKilled:         me.wardsKilled || 0,
        controlWards:        me.visionWardsBoughtInGame || 0,
        pctDeadTime:         parseFloat(((me.totalTimeSpentDead || 0) / match.info.gameDuration * 100).toFixed(1)),
        bestMultiKill:       bestMultiKill(me),
        pentaKills:          me.pentaKills    || 0,
        quadraKills:         me.quadraKills   || 0,
        tripleKills:         me.tripleKills   || 0,
        doubleKills:         me.doubleKills   || 0,
        largestKillingSpree: me.largestKillingSpree || 0,
        firstBlood:          me.firstBloodKill ? 'Oui' : 'Non',
        withYuumi:           (isDuo && duo.championName === 'Yuumi') ? 'Oui' : 'Non',
        // Yuumi dans l'équipe alliée (quel que soit le joueur)
        yuumiAlliee:         myTeam.some(p => p.puuid !== myPuuid && p.championName === 'Yuumi') ? 'Oui' : 'Non',
    };
}

// ============================================================
// MIGRATION — complète les entrées JSON incomplètes
// ============================================================

// Détecte si une entrée manque des champs ajoutés après la v1
function needsMigration(m) {
    return m.gameDuration  == null
        || m.yuumiAlliee   == null
        || m.firstBlood    == null
        || m.pctDeadTime   == null
        || m.objDpm        == null;
}

async function migrateOldEntries(data, myPuuid, duoPuuid) {
    const toFix = data.filter(needsMigration);
    if (toFix.length === 0) return 0;

    console.log(`   -> ${toFix.length} entree(s) incomplete(s) detectee(s). Migration en cours...`);
    let ok = 0, ko = 0;

    for (let i = 0; i < toFix.length; i++) {
        const entry = toFix[i];
        process.stdout.write(`   [Migration ${i + 1}/${toFix.length}] ${entry.matchId}... `);
        try {
            const r = await axios.get(
                `https://${REGION}.api.riotgames.com/lol/match/v5/matches/${entry.matchId}`,
                { headers: { 'X-Riot-Token': RIOT_API_KEY } }
            );
            const match   = r.data;
            const me      = match.info.participants.find(p => p.puuid === myPuuid);
            if (!me) throw new Error('PUUID introuvable');

            const duo     = match.info.participants.find(p => p.puuid === duoPuuid && p.teamId === me.teamId);
            const myTeam  = match.info.participants.filter(p => p.teamId === me.teamId);
            const minutes = match.info.gameDuration / 60;

            // On ne touche QU'aux champs manquants — sessionId, kda, etc. sont préservés
            entry.gameDuration        = parseFloat(minutes.toFixed(1));
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
            // Recalcule aussi withYuumi au cas où l'ancienne version était incorrecte
            entry.withYuumi           = (!!duo && duo.championName === 'Yuumi') ? 'Oui' : 'Non';

            ok++;
            console.log(`OK`);
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

// Ecrit + stylise une ligne de headers (crée les cellules avant de les styler)
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

// Titre de section fusionné
function sectionTitle(sheet, rowNum, text, span, bg) {
    sheet.mergeCells(rowNum, 1, rowNum, span);
    const cell     = sheet.getRow(rowNum).getCell(1);
    cell.value     = text;
    cell.font      = { bold: true, color: { argb: C.white }, size: 12 };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    sheet.getRow(rowNum).height = 28;
}

// Coloring résultat (Victoire/Defaite)
function colorResult(cell, win) {
    if (win === 'Victoire') {
        cell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg  } };
    } else {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
    }
}

// Coloring Win Rate par seuil
function colorWR(cell, wr) {
    cell.numFmt = '0.0"%"';
    if (wr >= 55) {
        cell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg  } };
    } else if (wr <= 45) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
    }
}

// Coloring KDA par seuil (1.5 mauvais / 2.5 bien / 4 excellent)
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

// Coloring Deaths par seuil (> 5 mauvais / <= 2 bon)
function colorDeaths(cell, d) {
    if (d >= 6) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad  } };
    } else if (d <= 2) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
    }
}

// Coloring CS/Min par seuil (< 7 mauvais / 9+ excellent)
function colorCS(cell, cs) {
    cell.numFmt = '0.0';
    if (cs >= 9.0) {
        cell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
    } else if (cs < 7.0) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad  } };
    }
}

// Coloring % Temps mort (> 20% mauvais)
function colorDeadTime(cell, pct) {
    cell.numFmt = '0.0';
    if (pct > 20) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
    }
}

// Fond de ligne standard
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
// SHEET 1 — DONNEES BRUTES
// ============================================================
function buildRawDataSheet(sheet, data) {
    // Colonnes — sans header (géré par writeHeaderRow)
    sheet.columns = [
        { key: 'date',          width: 11 },
        { key: 'time',          width: 7  },
        { key: 'sessionId',     width: 9  },
        { key: 'gameInSession', width: 9  },
        { key: 'gameDuration',  width: 11 },
        { key: 'champion',      width: 14 },
        { key: 'type',          width: 7  },
        { key: 'win',           width: 10 },
        { key: 'kda',           width: 7  },
        { key: 'kills',         width: 5  },
        { key: 'deaths',        width: 5  },
        { key: 'assists',       width: 5  },
        { key: 'csPerMin',      width: 8  },
        { key: 'gpm',           width: 7  },
        { key: 'dpm',           width: 8  },
        { key: 'objDpm',        width: 9  },
        { key: 'dmgShare',      width: 8  },
        { key: 'kp',            width: 7  },
        { key: 'vision',        width: 8  },
        { key: 'wardsPlaced',   width: 8  },
        { key: 'wardsKilled',   width: 8  },
        { key: 'controlWards',  width: 9  },
        { key: 'pctDeadTime',   width: 8  },
        { key: 'bestMultiKill', width: 11 },
        { key: 'firstBlood',    width: 11 },
        { key: 'withYuumi',     width: 10 },
        { key: 'yuumiAlliee',   width: 12 },
    ];
    const NCOLS = sheet.columns.length;
    sheet.views = [{ state: 'frozen', ySplit: 1 }];

    writeHeaderRow(sheet, 1, [
        'Date','Heure','Session','Partie/S','Duree','Champion','Type','Resultat',
        'KDA','K','D','A','CS/Min','GPM','DPM','Obj/Min',
        '% DMG','KP %','Vision','Wards+','Wards-','Ctrl W.','% Mort',
        'Multi-Kill','First Blood','Duo Yuumi','Yuumi Alliee',
    ], C.hRaw);

    const reversed = [...data].reverse();

    reversed.forEach((m, idx) => {
        const row = sheet.addRow({
            date:          safeV(m.date),
            time:          safeV(m.time),
            sessionId:     safeV(m.sessionId),
            gameInSession: safeV(m.gameInSession),
            gameDuration:  m.gameDuration  != null ? m.gameDuration  : '-',
            champion:      safeV(m.champion),
            type:          safeV(m.type),
            win:           safeV(m.win),
            kda:           m.kda      != null ? m.kda      : '-',
            kills:         safeV(m.kills,   0),
            deaths:        safeV(m.deaths,  0),
            assists:       safeV(m.assists, 0),
            csPerMin:      m.csPerMin  != null ? m.csPerMin  : '-',
            gpm:           m.gpm       != null ? m.gpm       : '-',
            dpm:           m.dpm       != null ? m.dpm       : '-',
            objDpm:        m.objDpm    != null ? m.objDpm    : '-',
            dmgShare:      m.dmgShare  != null ? m.dmgShare  : '-',
            kp:            m.kp        != null ? m.kp        : '-',
            vision:        safeV(m.vision, 0),
            wardsPlaced:   safeV(m.wardsPlaced,  0),
            wardsKilled:   safeV(m.wardsKilled,  0),
            controlWards:  safeV(m.controlWards, 0),
            pctDeadTime:   m.pctDeadTime != null ? m.pctDeadTime : '-',
            bestMultiKill: safeV(m.bestMultiKill, '-'),
            firstBlood:    safeV(m.firstBlood,    '-'),
            withYuumi:     safeV(m.withYuumi),
            yuumiAlliee:   safeV(m.yuumiAlliee,  '-'),
        });

        // Fond de ligne
        const bg = rowBg(idx);
        row.eachCell(cell => {
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.font      = { size: 10 };
        });
        row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
        row.height = 20;

        // Colorings par seuil (remplacent les color scales)
        colorResult(row.getCell('win'), m.win);
        if (typeof m.kda      === 'number') colorKDA(row.getCell('kda'), m.kda);
        if (typeof m.deaths   === 'number') colorDeaths(row.getCell('deaths'), m.deaths);
        if (typeof m.csPerMin === 'number') colorCS(row.getCell('csPerMin'), m.csPerMin);
        if (typeof m.pctDeadTime === 'number') colorDeadTime(row.getCell('pctDeadTime'), m.pctDeadTime);

        if (typeof m.dmgShare === 'number') row.getCell('dmgShare').numFmt = '0.0';
        if (typeof m.kp       === 'number') row.getCell('kp').numFmt       = '0.0';
        if (typeof m.gpm      === 'number') row.getCell('gpm').numFmt      = '0';
        if (typeof m.dpm      === 'number') row.getCell('dpm').numFmt      = '0';

        // First Blood
        if (m.firstBlood === 'Oui') {
            row.getCell('firstBlood').font = { bold: true, color: { argb: C.winFg }, size: 10 };
        }
        // Multi-kill > Double
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
                    'CS/Min Moy.','GPM Moy.','DPM Moy.','% DMG Moy.','KP % Moy.',
                    'Vision Moy.','Solo','Duo','Duree Moy.'];
    const KEYS   = ['champion','games','wins','losses','winRate','kda',
                    'cs','gpm','dpm','dmg','kp','vision','solo','duo','duration'];
    const WIDTHS = [16, 9, 10, 10, 12, 9, 11, 9, 9, 11, 10, 11, 7, 7, 11];
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
            winRate:  parseFloat(winPct(games).toFixed(1)),
            kda:      parseFloat(avgOf(games, 'kda').toFixed(2)),
            cs:       parseFloat(avgOf(games, 'csPerMin').toFixed(1)),
            gpm:      Math.round(avgOf(games, 'gpm')),
            dpm:      Math.round(avgOf(games, 'dpm')),
            dmg:      parseFloat(avgOf(games, 'dmgShare').toFixed(1)),
            kp:       parseFloat(avgOf(games, 'kp').toFixed(1)),
            vision:   parseFloat(avgOf(games, 'vision').toFixed(1)),
            solo:     games.filter(g => g.type === 'Solo').length,
            duo:      games.filter(g => g.type === 'Duo').length,
            duration: parseFloat(avgOf(games, 'gameDuration').toFixed(1)),
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

        // Colorings
        colorWR(row.getCell('winRate'), c.winRate);
        colorKDA(row.getCell('kda'),    c.kda);
        colorCS(row.getCell('cs'),      c.cs);

        row.getCell('dmg').numFmt  = '0.0"%"';
        row.getCell('kp').numFmt   = '0.0"%"';
        row.getCell('gpm').numFmt  = '0';
        row.getCell('dpm').numFmt  = '0';
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
        ['CS/Min Moyen',       ...cats.map(c => parseFloat(avgOf(c.games, 'csPerMin').toFixed(1)))],
        ['DPM Moyen',          ...cats.map(c => Math.round(avgOf(c.games, 'dpm')))],
        ['GPM Moyen',          ...cats.map(c => Math.round(avgOf(c.games, 'gpm')))],
        ['% DMG Moyen',        ...cats.map(c => parseFloat(avgOf(c.games, 'dmgShare').toFixed(1)))],
        ['KP % Moyen',         ...cats.map(c => parseFloat(avgOf(c.games, 'kp').toFixed(1)))],
        ['Vision Moyen',       ...cats.map(c => parseFloat(avgOf(c.games, 'vision').toFixed(1)))],
        ['Obj/Min Moyen',      ...cats.map(c => Math.round(avgOf(c.games, 'objDpm')))],
        ['% Temps Mort Moyen', ...cats.map(c => parseFloat(avgOf(c.games, 'pctDeadTime').toFixed(1)))],
    ];

    tableRows.forEach((rowData, idx) => {
        const shRow = sheet.getRow(idx + 3);
        rowData.forEach((val, ci) => {
            const cell     = shRow.getCell(ci + 1);
            cell.value     = val;
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            cell.font      = { bold: ci === 0, size: 10 };
            cell.alignment = { horizontal: ci === 0 ? 'left' : 'center', vertical: 'middle' };
        });
        shRow.height = 22;

        // Win Rate coloré
        if (rowData[0] === 'Win Rate %') {
            for (let ci = 1; ci < cats.length + 1; ci++) {
                colorWR(shRow.getCell(ci + 1), parseFloat(shRow.getCell(ci + 1).value) || 0);
            }
        }
        // KDA coloré
        if (rowData[0] === 'KDA Moyen') {
            for (let ci = 1; ci < cats.length + 1; ci++) {
                colorKDA(shRow.getCell(ci + 1), parseFloat(shRow.getCell(ci + 1).value) || 0);
            }
        }
    });
}

// ============================================================
// SHEET 4 — TENDANCES COMPORTEMENTALES
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
        { key: 'c4', width: 16 },
        { key: 'c5', width: 16 },
        { key: 'c6', width: 32 },
    ];
    const SPAN = 6;
    let r = 1;

    // Ligne clé/valeur (2 colonnes)
    function kv(rowNum, label, value, idx) {
        const row = sheet.getRow(rowNum);
        const bg  = rowBg(idx);
        const c1  = row.getCell(1);
        const c2  = row.getCell(2);
        c1.value     = label;
        c1.font      = { bold: true, size: 10 };
        c1.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        c1.alignment = { horizontal: 'left', vertical: 'middle' };
        c2.value     = value;
        c2.font      = { size: 10 };
        c2.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        c2.alignment = { horizontal: 'center', vertical: 'middle' };
        row.height   = 20;
    }

    // Ligne tableau multi-colonnes
    function tRow(rowNum, values, idx) {
        const row = sheet.getRow(rowNum);
        values.forEach((v, i) => {
            const cell     = row.getCell(i + 1);
            cell.value     = v;
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            cell.font      = { bold: i === 0, size: 10 };
            cell.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
        });
        row.height = 20;
    }

    // ── VUE D'ENSEMBLE ────────────────────────────────────────
    sectionTitle(sheet, r, "Vue d'Ensemble Globale", SPAN, C.hTend); r++;
    [
        ["Parties totales",      a.overall.total],
        ["Win Rate global",      a.overall.winRate.toFixed(1) + ' %'],
        ["KDA moyen",            a.overall.avgKDA.toFixed(2)],
        ["CS/Min moyen",         a.overall.avgCS.toFixed(1)],
        ["DPM moyen",            Math.round(a.overall.avgDPM)],
        ["GPM moyen",            Math.round(a.overall.avgGPM)],
        ["% DMG Equipe moyen",   a.overall.avgDmgShare.toFixed(1) + ' %'],
        ["KP % moyen",           a.overall.avgKP.toFixed(1) + ' %'],
        ["Vision Score moyen",   a.overall.avgVision.toFixed(1)],
    ].forEach(([label, value], i) => { kv(r, label, value, i); r++; });
    r++;

    // ── SERIE & TENDANCE ──────────────────────────────────────
    sectionTitle(sheet, r, "Serie Actuelle & Tendance", SPAN, C.hTend); r++;
    const wr10   = winPct(a.recent10);
    const wrGlob = a.overall.winRate;
    const trend  = wr10 > wrGlob + 2 ? 'En progression' : wr10 < wrGlob - 2 ? 'En regression' : 'Stable';
    [
        ["Serie actuelle",             a.streak.count + ' ' + a.streak.type + ' consecutives'],
        ["Win Rate (10 dernieres)",    wr10.toFixed(1) + ' %'],
        ["Win Rate global",            wrGlob.toFixed(1) + ' %'],
        ["Tendance recente",           trend],
        ["KDA (10 dernieres)",         avgOf(a.recent10, 'kda').toFixed(2)],
        ["DPM (10 dernieres)",         Math.round(avgOf(a.recent10, 'dpm'))],
        ["CS/Min (10 dernieres)",      avgOf(a.recent10, 'csPerMin').toFixed(1)],
    ].forEach(([label, value], i) => { kv(r, label, value, i); r++; });
    r++;

    // ── FATIGUE COGNITIVE ─────────────────────────────────────
    sectionTitle(sheet, r, "Fatigue Cognitive — Performance par Partie en Session", SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Partie dans Session', 'Nb Parties', 'Win Rate %', 'KDA Moy.', 'CS/Min', 'DPM'], C.subH); r++;
    [1, 2, 3, 4].forEach((gNum, idx) => {
        const games = a.byGameInSession[gNum] || [];
        const wr    = winPct(games);
        tRow(r, [
            gNum < 4 ? ('Partie ' + gNum) : 'Partie 4+',
            games.length,
            games.length ? wr.toFixed(1) + ' %' : '-',
            games.length ? avgOf(games, 'kda').toFixed(2) : '-',
            games.length ? avgOf(games, 'csPerMin').toFixed(1) : '-',
            games.length ? Math.round(avgOf(games, 'dpm')) : '-',
        ], idx);
        // Coloring Win Rate fatigue
        if (games.length) {
            const wrCell = sheet.getRow(r).getCell(3);
            if (wr >= 55) {
                wrCell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
                wrCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg  } };
            } else if (wr <= 45) {
                wrCell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
                wrCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
            }
        }
        r++;
    });
    r++;

    // ── PAR TRANCHE HORAIRE ───────────────────────────────────
    sectionTitle(sheet, r, "Performance par Tranche Horaire", SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Tranche Horaire', 'Parties', 'Win Rate %', 'KDA Moy.', 'DPM Moy.', 'Diagnostic'], C.subH); r++;
    Object.entries(a.timeSlots).forEach(([slot, games], idx) => {
        const wr   = winPct(games);
        const diag = games.length < 3 ? 'Echantillon faible'
                   : wr >= 55          ? 'Tranche favorable'
                   : wr <= 45          ? 'Tranche defavorable'
                   :                    'Neutre';
        tRow(r, [
            slot, games.length,
            games.length ? wr.toFixed(1) + ' %' : '-',
            games.length ? avgOf(games, 'kda').toFixed(2) : '-',
            games.length ? Math.round(avgOf(games, 'dpm')) : '-',
            diag,
        ], idx);
        // Coloring Win Rate heure
        if (games.length) {
            const wrCell = sheet.getRow(r).getCell(3);
            if (wr >= 55) {
                wrCell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
                wrCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg  } };
            } else if (wr <= 45) {
                wrCell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
                wrCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
            }
        }
        r++;
    });
    r++;

    // ── INDICATEUR DE TILT ────────────────────────────────────
    sectionTitle(sheet, r, "Indicateur de Tilt — Perf. apres Victoire vs Defaite", SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Contexte', 'Parties', 'Win Rate %', 'KDA Moy.', 'DPM Moy.', 'Diagnostic'], C.subH); r++;
    const wrW   = winPct(a.afterWin);
    const wrL   = winPct(a.afterLoss);
    const delta = wrW - wrL;
    const tiltDiag = delta > 10 ? 'Tilt detecte — Stopper apres defaite'
                   : delta > 5  ? 'Legere regression post-defaite'
                   :              'Mental stable';
    [
        ['Apres une Victoire', a.afterWin.length,  wrW.toFixed(1) + ' %', avgOf(a.afterWin,  'kda').toFixed(2), Math.round(avgOf(a.afterWin,  'dpm')), ''],
        ['Apres une Defaite',  a.afterLoss.length, wrL.toFixed(1) + ' %', avgOf(a.afterLoss, 'kda').toFixed(2), Math.round(avgOf(a.afterLoss, 'dpm')), tiltDiag],
    ].forEach((rowData, idx) => {
        tRow(r, rowData, idx);
        const wrCell = sheet.getRow(r).getCell(3);
        const wrVal  = parseFloat(wrCell.value) || 0;
        wrCell.numFmt = '0.0"%"';
        if (wrVal >= 55) {
            wrCell.font = { bold: true, color: { argb: C.winFg  }, size: 10 };
            wrCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg  } };
        } else if (wrVal <= 45) {
            wrCell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
            wrCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
        }
        r++;
    });
    r++;

    // ── RECORDS PERSONNELS ────────────────────────────────────
    sectionTitle(sheet, r, "Records Personnels", SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Categorie', 'Valeur', 'Champion', 'Date', 'Resultat', ''], C.subH); r++;

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

            // Cellule 1 : label
            const c1 = rRow.getCell(1);
            c1.value     = label;
            c1.font      = { bold: true, size: 10 };
            c1.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            c1.alignment = { horizontal: 'left', vertical: 'middle' };

            // Cellule 2 : valeur (accent indigo, jamais color: undefined)
            const c2 = rRow.getCell(2);
            c2.value     = val;
            c2.font      = { bold: true, size: 10, color: { argb: C.accent } };
            c2.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            c2.alignment = { horizontal: 'center', vertical: 'middle' };
            c2.numFmt    = dec === 2 ? '0.00' : dec === 1 ? '0.0' : '0';

            // Cellules 3-6 : champion, date, résultat
            [rec.champion, rec.date, rec.win, ''].forEach((v, i) => {
                const cell     = rRow.getCell(i + 3);
                cell.value     = v;
                cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
                cell.font      = { size: 10 };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            });
            colorResult(rRow.getCell(5), rec.win);
        } else {
            ['Categorie', '-', '-', '-', '-', ''].forEach((v, i) => {
                const cell = rRow.getCell(i + 1);
                cell.value = i === 0 ? label : '-';
                cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
                cell.font  = { bold: i === 0, size: 10 };
                cell.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
            });
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
// POINT D'ENTREE
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

            // Migration des entrées incomplètes (anciennes versions du script)
            const migrated = await migrateOldEntries(data, player.puuid, duo.puuid);
            if (migrated > 0) {
                console.log('   -> Sauvegarde JSON apres migration...');
                fs.writeFileSync(JSON_FILENAME, JSON.stringify(data, null, 4));
            }

            await rebuildExcel(data);
            console.log('\n[SUCCES] Tableau de bord a jour.');
            return;
        }

        newIds.reverse();
        console.log('-> ' + newIds.length + ' nouvelle(s) partie(s).');

        console.log('4. Analyse des nouvelles parties...');
        let prev = data.length > 0 ? data[data.length - 1] : null;
        let ok = 0, ko = 0;

        for (let i = 0; i < newIds.length; i++) {
            process.stdout.write('   [' + (i + 1) + '/' + newIds.length + '] ' + newIds[i] + '... ');
            try {
                const m = await extractMatchMetrics(newIds[i], player.puuid, duo.puuid, prev);
                data.push(m);
                prev = m;
                ok++;
                console.log('OK  ' + m.champion + ' — ' + m.win + ' (S' + m.sessionId + ' G' + m.gameInSession + ')');
            } catch (e) {
                ko++;
                console.log('KO  ' + e.message);
            }
            await sleep(API_DELAY_MS);
        }

        console.log('\n5. Sauvegarde JSON...');
        fs.writeFileSync(JSON_FILENAME, JSON.stringify(data, null, 4));
        console.log('   -> ' + data.length + ' parties enregistrees.');

        // Migration des entrées encore incomplètes (si besoin)
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