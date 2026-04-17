const ExcelJS = require('exceljs');

const { EXCEL_FILENAME, C, rankColor } = require('./config');
const { computeAnalytics } = require('./analytics');
const { safeN, safeV, avgOf, winPct, rowBg, secToMmSs, pushTo } = require('./utils');

function writeHeaderRow(sheet, rowNum, labels, bg) {
    const row = sheet.getRow(rowNum);
    labels.forEach((label, i) => {
        const cell = row.getCell(i + 1);
        cell.value = label;
        cell.font = { bold: true, color: { argb: C.white }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = { bottom: { style: 'thin', color: { argb: C.accent } } };
    });
    row.height = 26;
}

function writeMultiColorHeaderRow(sheet, rowNum, groups) {
    const row = sheet.getRow(rowNum);
    let col = 1;
    for (const grp of groups) {
        for (let i = 0; i < grp.count; i++) {
            const cell = row.getCell(col);
            cell.value = grp.labels[i];
            cell.font = { bold: true, color: { argb: C.white }, size: 9 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: grp.bg } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = { bottom: { style: 'thin', color: { argb: C.accent } } };
            col++;
        }
    }
    row.height = 28;
}

function sectionTitle(sheet, rowNum, text, span, bg) {
    sheet.mergeCells(rowNum, 1, rowNum, span);
    const cell = sheet.getRow(rowNum).getCell(1);
    cell.value = text;
    cell.font = { bold: true, color: { argb: C.white }, size: 12 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    sheet.getRow(rowNum).height = 28;
}

function colorResult(cell, win) {
    if (win === 'Victoire') {
        cell.font = { bold: true, color: { argb: C.winFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg } };
    } else {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
    }
}

function colorWR(cell, wr) {
    cell.numFmt = '0.0%';
    if (wr >= 0.55) {
        cell.font = { bold: true, color: { argb: C.winFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg } };
    } else if (wr <= 0.45) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
    }
}

function colorKDA(cell, kda) {
    cell.numFmt = '0.00';
    if (kda >= 4.0) {
        cell.font = { bold: true, color: { argb: C.winFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
    } else if (kda < 1.5) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
    } else if (kda >= 2.5) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.warn } };
    }
}

function colorResponsabilite(cell, resp) {
    if (!Number.isFinite(resp)) return;

    if (resp >= 100) {
        cell.font = { bold: true, color: { argb: C.white }, size: 10 };
        cell.fill = {
            type: 'gradient',
            gradient: 'angle',
            degree: 0,
            stops: [
                { position: 0, color: { argb: 'FFEF4444' } },
                { position: 0.2, color: { argb: 'FFF59E0B' } },
                { position: 0.4, color: { argb: 'FFEAB308' } },
                { position: 0.6, color: { argb: 'FF22C55E' } },
                { position: 0.8, color: { argb: 'FF3B82F6' } },
                { position: 1, color: { argb: 'FF8B5CF6' } },
            ],
        };
    } else if (resp > 80) {
        cell.font = { bold: true, color: { argb: 'FF713F12' }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFACC15' } };
    } else if (resp > 60) {
        cell.font = { bold: true, color: { argb: C.white }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16A34A' } };
    } else if (resp > 40) {
        cell.font = { bold: true, color: { argb: C.white }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF97316' } };
    } else if (resp > 20) {
        cell.font = { bold: true, color: { argb: C.white }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDC2626' } };
    } else {
        cell.font = { bold: true, color: { argb: C.white }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6B7280' } };
    }
}

function colorDeaths(cell, d) {
    if (d >= 6) {
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
    } else if (d <= 2) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
    }
}

function colorCS(cell, cs) {
    cell.numFmt = '0.0';
    if (cs >= 9.0) {
        cell.font = { bold: true, color: { argb: C.winFg }, size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
    } else if (cs < 7.0) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
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
    const lightBgs = [C.rankGold, C.rankPlatinum, C.rankEmerald, C.rankDiamond, C.rankChallenger];
    const isLight = lightBgs.includes(fg);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fg } };
    cell.font = { bold: true, color: { argb: isLight ? 'FF000000' : C.white }, size: 10 };
}

function colorRankDiff(cell, diff) {
    if (diff === null || diff === undefined) return;
    if (diff > 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg } };
        cell.font = { bold: true, color: { argb: C.winFg }, size: 10 };
    } else if (diff < -1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
        cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
    }
}

function buildRawDataSheet(sheet, data) {
    const GROUPS = [
        {
            bg: C.grpIdentite,
            labels: ['Date', 'Heure', 'Session', 'P/S', 'Duree', 'Champion', 'Role', 'Type', 'Resultat'],
            keys: ['date', 'time', 'sessionId', 'gameInSession', 'gameDuration', 'champion', 'role', 'type', 'win'],
            widths: [11, 7, 9, 6, 8, 14, 10, 7, 10],
        },
        {
            bg: C.grpCombat,
            labels: ['KDA', 'K', 'D', 'A', '% Mort', 'KP %', '% DMG', 'DMG/Gold', 'Ratio DMG', 'Resp.'],
            keys: ['kda', 'kills', 'deaths', 'assists', 'pctDeadTime', 'kp', 'dmgShare', 'dpg', 'dmgRatio', 'responsabilite'],
            widths: [7, 5, 5, 5, 8, 8, 8, 10, 10, 8],
        },
        {
            bg: C.grpFarm,
            labels: ['CS/Min', 'GPM', 'DPM', 'Obj/Min'],
            keys: ['csPerMin', 'gpm', 'dpm', 'objDpm'],
            widths: [8, 7, 8, 9],
        },
        {
            bg: C.grpVision,
            labels: ['Vision', 'Wards+', 'Wards-', 'Ctrl W.'],
            keys: ['vision', 'wardsPlaced', 'wardsKilled', 'controlWards'],
            widths: [8, 8, 8, 9],
        },
        {
            bg: C.grpPings,
            labels: ['Pings Tot.', 'Pings Nég.', 'AllIn', 'Basic', 'Cmd', 'Danger', 'MIA', 'Hold', 'Vision?', 'OTW', 'Push', 'Retreat', 'ClearVis'],
            keys: ['totalPings', 'negativePings', 'allInPings', 'basicPings', 'commandPings', 'dangerPings', 'enemyMissingPings', 'holdPings', 'needVisionPings', 'onMyWayPings', 'pushPings', 'retreatPings', 'visionClearedPings'],
            widths: [10, 10, 7, 7, 7, 8, 7, 7, 9, 8, 8, 9, 10],
        },
        {
            bg: C.grpGanks,
            labels: ['Ganks Subis (Early)', 'CSD@10', 'GD@10', 'CSD@15', 'GD@15', 'Core Item (min)'],
            keys: ['fatalGanksReceived', 'csDiff10', 'goldDiff10', 'csDiff15', 'goldDiff15', 'firstCoreItemMin'],
            widths: [16, 9, 10, 9, 10, 15],
        },
        {
            bg: C.grpRank,
            labels: ['Mon Rang', 'Rang Alliés', 'Rang Ennemis', 'Diff Rang'],
            keys: ['myRank', 'avgAllyRank', 'avgEnemyRank', 'rankDiff'],
            widths: [14, 14, 14, 10],
        },
        {
            bg: C.grpDivers,
            labels: ['Multi-Kill', 'First Blood', 'Duo Yuumi', 'Yuumi Alli.', 'ADC Adverse'],
            keys: ['bestMultiKill', 'firstBlood', 'withYuumi', 'yuumiAlliee', 'enemyADC'],
            widths: [11, 11, 10, 12, 14],
        },
    ];

    const flatKeys = GROUPS.flatMap((g) => g.keys);
    const flatWidths = GROUPS.flatMap((g) => g.widths);
    const NCOLS = flatKeys.length;

    sheet.columns = flatKeys.map((key, i) => ({ key, width: flatWidths[i] }));
    sheet.views = [{ state: 'frozen', ySplit: 1 }];

    writeMultiColorHeaderRow(sheet, 1, GROUPS.map((g) => ({ ...g, count: g.keys.length })));

    const reversed = [...data].reverse();
    reversed.forEach((m, idx) => {
        const rowData = {};
        flatKeys.forEach((k) => { rowData[k] = safeV(m[k]); });
        const row = sheet.addRow(rowData);

        const bg = rowBg(idx);
        row.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.font = { size: 10 };
        });
        row.getCell('date').alignment = { horizontal: 'left', vertical: 'middle' };
        row.height = 20;

        colorResult(row.getCell('win'), m.win);
        if (typeof m.kda === 'number') colorKDA(row.getCell('kda'), m.kda);
        if (typeof m.deaths === 'number') colorDeaths(row.getCell('deaths'), m.deaths);
        if (typeof m.csPerMin === 'number') colorCS(row.getCell('csPerMin'), m.csPerMin);
        if (typeof m.pctDeadTime === 'number') colorDeadTime(row.getCell('pctDeadTime'), m.pctDeadTime);
        if (typeof m.responsabilite === 'number') colorResponsabilite(row.getCell('responsabilite'), m.responsabilite);

        if (typeof m.dmgShare === 'number') row.getCell('dmgShare').numFmt = '0.0';
        if (typeof m.kp === 'number') row.getCell('kp').numFmt = '0.0';
        if (typeof m.gpm === 'number') row.getCell('gpm').numFmt = '0';
        if (typeof m.dpm === 'number') row.getCell('dpm').numFmt = '0';

        if (typeof m.myRankScore === 'number') colorRankCell(row.getCell('myRank'), m.myRankScore);
        if (typeof m.avgAllyRankScore === 'number') colorRankCell(row.getCell('avgAllyRank'), m.avgAllyRankScore);
        if (typeof m.avgEnemyRankScore === 'number') colorRankCell(row.getCell('avgEnemyRank'), m.avgEnemyRankScore);
        if (m.rankDiff !== null && m.rankDiff !== undefined) {
            row.getCell('rankDiff').value = m.rankDiff;
            row.getCell('rankDiff').numFmt = '+0.0;-0.0;0.0';
            colorRankDiff(row.getCell('rankDiff'), m.rankDiff);
        }

        if (typeof m.dpg === 'number') {
            row.getCell('dpg').numFmt = '0.000';
            if (m.dpg >= 1.5) {
                row.getCell('dpg').font = { bold: true, color: { argb: C.winFg }, size: 10 };
                row.getCell('dpg').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
            } else if (m.dpg < 0.9) {
                row.getCell('dpg').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
            }
        }

        if (typeof m.dmgRatio === 'number') {
            row.getCell('dmgRatio').numFmt = '0.00';
            if (m.dmgRatio >= 1.5) {
                row.getCell('dmgRatio').font = { bold: true, color: { argb: C.winFg }, size: 10 };
                row.getCell('dmgRatio').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
            } else if (m.dmgRatio < 0.8) {
                row.getCell('dmgRatio').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
            }
        }

        ['csDiff10', 'csDiff15'].forEach((key) => {
            const v = m[key];
            if (typeof v === 'number') {
                row.getCell(key).numFmt = '+0;-0;0';
                if (v > 0) row.getCell(key).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
                else if (v < 0) row.getCell(key).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
            }
        });

        ['goldDiff10', 'goldDiff15'].forEach((key) => {
            const v = m[key];
            if (typeof v === 'number') {
                row.getCell(key).numFmt = '+0;-0;0';
                if (v > 200) row.getCell(key).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
                else if (v < -200) row.getCell(key).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };
            }
        });

        const negP = safeN(m.negativePings, 0);
        if (negP > 15) {
            row.getCell('negativePings').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.warn } };
            row.getCell('negativePings').font = { bold: true, color: { argb: C.amber }, size: 10 };
        }

        if (m.firstBlood === 'Oui') {
            row.getCell('firstBlood').font = { bold: true, color: { argb: C.winFg }, size: 10 };
        }

        const mk = String(m.bestMultiKill || '');
        if (mk === 'Triple' || mk === 'Quadra' || mk === 'Penta') {
            row.getCell('bestMultiKill').font = { bold: true, color: { argb: C.purple }, size: 10 };
        }
    });

    sheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: NCOLS },
    };
}

function buildChampionSheet(sheet, data) {
    const LABELS = ['Champion', 'Parties', 'Victoires', 'Defaites', 'Win Rate %', 'KDA Moy.', 'K Moy.', 'D Moy.', 'A Moy.', 'CS/Min Moy.', 'GPM Moy.', 'DPM Moy.', '% DMG Moy.', 'KP % Moy.', 'Vision Moy.', 'Solo', 'Duo', 'Duree Moy.'];
    const KEYS = ['champion', 'games', 'wins', 'losses', 'winRate', 'kda', 'avgK', 'avgD', 'avgA', 'cs', 'gpm', 'dpm', 'dmg', 'kp', 'vision', 'solo', 'duo', 'duration'];
    const WIDTHS = [16, 9, 10, 10, 12, 9, 7, 7, 7, 11, 9, 9, 11, 10, 11, 7, 7, 11];
    const NCOLS = LABELS.length;

    sheet.columns = KEYS.map((k, i) => ({ key: k, width: WIDTHS[i] }));
    sheet.views = [{ state: 'frozen', ySplit: 2 }];

    sectionTitle(sheet, 1, 'Stats par Champion  —  trie par parties jouees', NCOLS, C.hChamp);
    writeHeaderRow(sheet, 2, LABELS, C.hChamp);

    const champMap = {};
    data.forEach((m) => pushTo(champMap, m.champion, m));

    const list = Object.entries(champMap).map(([name, games]) => {
        const wins = games.filter((g) => g.win === 'Victoire').length;
        return {
            champion: name,
            games: games.length,
            wins,
            losses: games.length - wins,
            winRate: parseFloat(winPct(games)),
            kda: parseFloat(avgOf(games, 'kda').toFixed(2)),
            avgK: parseFloat(avgOf(games, 'kills').toFixed(1)),
            avgD: parseFloat(avgOf(games, 'deaths').toFixed(1)),
            avgA: parseFloat(avgOf(games, 'assists').toFixed(1)),
            cs: parseFloat(avgOf(games, 'csPerMin').toFixed(1)),
            gpm: Math.round(avgOf(games, 'gpm')),
            dpm: Math.round(avgOf(games, 'dpm')),
            dmg: parseFloat(avgOf(games, 'dmgShare').toFixed(1)),
            kp: parseFloat(avgOf(games, 'kp').toFixed(1)),
            vision: parseFloat(avgOf(games, 'vision').toFixed(1)),
            solo: games.filter((g) => g.type === 'Solo').length,
            duo: games.filter((g) => g.type === 'Duo').length,
            duration: games.filter((g) => g.gameDurationSec).reduce((s, g) => s + (g.gameDurationSec || 0), 0)
                / Math.max(1, games.filter((g) => g.gameDurationSec).length) / 60,
        };
    }).sort((a, b) => b.games - a.games);

    list.forEach((c, idx) => {
        const row = sheet.addRow(c);
        const bg = rowBg(idx);
        row.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.font = { size: 10 };
        });
        row.getCell('champion').font = { bold: true, size: 10 };
        row.getCell('champion').alignment = { horizontal: 'left', vertical: 'middle' };
        row.height = 20;

        colorWR(row.getCell('winRate'), c.winRate);
        colorKDA(row.getCell('kda'), c.kda);
        colorCS(row.getCell('cs'), c.cs);
        colorDeaths(row.getCell('avgD'), Math.round(c.avgD));

        row.getCell('dmg').numFmt = '0.0"%"';
        row.getCell('kp').numFmt = '0.0"%"';
        row.getCell('gpm').numFmt = '0';
        row.getCell('dpm').numFmt = '0';
        row.getCell('duration').value = secToMmSs(c.duration * 60);
    });

    sheet.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: NCOLS } };
}

function buildDuoSoloSheet(sheet, data) {
    const cats = [
        { label: 'Solo Queue', games: data.filter((m) => m.type === 'Solo') },
        { label: 'Duo (total)', games: data.filter((m) => m.type === 'Duo') },
        { label: 'Duo + Yuumi', games: data.filter((m) => m.withYuumi === 'Oui') },
        { label: 'Duo sans Yuumi', games: data.filter((m) => m.type === 'Duo' && m.withYuumi !== 'Oui') },
        { label: 'Yuumi alliee', games: data.filter((m) => m.yuumiAlliee === 'Oui') },
    ];
    const SPAN = 1 + cats.length;

    sheet.columns = [
        { key: 'metric', width: 24 },
        ...cats.map((c) => ({ key: c.label, width: 18 })),
    ];
    sheet.views = [{ state: 'frozen', ySplit: 2 }];

    sectionTitle(sheet, 1, 'Comparaison Solo vs Duo', SPAN, C.hDuo);
    writeHeaderRow(sheet, 2, ['Metrique', ...cats.map((c) => c.label)], C.hDuo);

    const tableRows = [
        ['Parties', ...cats.map((c) => c.games.length)],
        ['Victoires', ...cats.map((c) => c.games.filter((g) => g.win === 'Victoire').length)],
        ['Defaites', ...cats.map((c) => c.games.filter((g) => g.win !== 'Victoire').length)],
        ['Win Rate %', ...cats.map((c) => parseFloat(winPct(c.games).toFixed(3)))],
        ['KDA Moyen', ...cats.map((c) => parseFloat(avgOf(c.games, 'kda').toFixed(2)))],
        ['K Moyen', ...cats.map((c) => parseFloat(avgOf(c.games, 'kills').toFixed(1)))],
        ['D Moyen', ...cats.map((c) => parseFloat(avgOf(c.games, 'deaths').toFixed(1)))],
        ['A Moyen', ...cats.map((c) => parseFloat(avgOf(c.games, 'assists').toFixed(1)))],
        ['CS/Min Moyen', ...cats.map((c) => parseFloat(avgOf(c.games, 'csPerMin').toFixed(1)))],
        ['DPM Moyen', ...cats.map((c) => Math.round(avgOf(c.games, 'dpm')))],
        ['GPM Moyen', ...cats.map((c) => Math.round(avgOf(c.games, 'gpm')))],
        ['% DMG Moyen', ...cats.map((c) => parseFloat(avgOf(c.games, 'dmgShare').toFixed(1)))],
        ['KP % Moyen', ...cats.map((c) => parseFloat(avgOf(c.games, 'kp').toFixed(1)))],
        ['Vision Moyen', ...cats.map((c) => parseFloat(avgOf(c.games, 'vision').toFixed(1)))],
        ['Obj/Min Moyen', ...cats.map((c) => Math.round(avgOf(c.games, 'objDpm')))],
        ['% Temps Mort Moyen', ...cats.map((c) => parseFloat(avgOf(c.games, 'pctDeadTime').toFixed(1)))],
        ['Pings Tot. Moy.', ...cats.map((c) => Math.round(avgOf(c.games, 'totalPings')))],
        ['Pings Nég. Moy.', ...cats.map((c) => Math.round(avgOf(c.games, 'negativePings')))],
    ];

    tableRows.forEach((rowData, idx) => {
        const shRow = sheet.getRow(idx + 3);
        rowData.forEach((val, ci) => {
            const cell = shRow.getCell(ci + 1);
            cell.value = val;
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            cell.font = { bold: ci === 0, size: 10 };
            cell.alignment = { horizontal: ci === 0 ? 'left' : 'center', vertical: 'middle' };
            if (rowData[0] === 'Win Rate %' && typeof val === 'number') cell.numFmt = '0.0%';
        });
        shRow.height = 22;

        if (rowData[0] === 'Win Rate %') {
            for (let ci = 1; ci <= cats.length; ci++) {
                const cell = shRow.getCell(ci + 1);
                colorWR(cell, parseFloat(cell.value) || 0);
            }
        }
        if (rowData[0] === 'KDA Moyen') {
            for (let ci = 1; ci <= cats.length; ci++) {
                colorKDA(shRow.getCell(ci + 1), parseFloat(shRow.getCell(ci + 1).value) || 0);
            }
        }
    });
}

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
        const bg = rowBg(idx);
        const c1 = row.getCell(1);
        const c2 = row.getCell(2);
        c1.value = label;
        c1.font = { bold: true, size: 10 };
        c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        c2.value = value;
        c2.font = { size: 10 };
        c2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        c2.alignment = { horizontal: 'center', vertical: 'middle' };
        if (isPercent && typeof value === 'number') c2.numFmt = '0.0%';
        row.height = 20;
    }

    function tRow(rowNum, values, idx, percentCols = []) {
        const row = sheet.getRow(rowNum);
        values.forEach((v, i) => {
            const cell = row.getCell(i + 1);
            cell.value = v;
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            cell.font = { bold: i === 0, size: 10 };
            cell.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
            if (percentCols.includes(i + 1) && typeof v === 'number') cell.numFmt = '0.0%';
        });
        row.height = 20;
    }

    function colorWRCell(rowNum, colNum, wr) {
        const cell = sheet.getRow(rowNum).getCell(colNum);
        if (wr >= 0.55) {
            cell.font = { bold: true, color: { argb: C.winFg }, size: 10 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.winBg } };
        } else if (wr <= 0.45) {
            cell.font = { bold: true, color: { argb: C.lossFg }, size: 10 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.lossBg } };
        }
    }

    sectionTitle(sheet, r, "Vue d'Ensemble Globale", SPAN, C.hTend); r++;
    [
        ['Parties totales', a.overall.total, false],
        ['Win Rate global', a.overall.winRate, true],
        ['KDA moyen', a.overall.avgKDA.toFixed(2), false],
        ['K / D / A moyens', `${a.overall.avgKills.toFixed(1)} / ${a.overall.avgDeaths.toFixed(1)} / ${a.overall.avgAssists.toFixed(1)}`, false],
        ['CS/Min moyen', a.overall.avgCS.toFixed(1), false],
        ['DPM moyen', Math.round(a.overall.avgDPM), false],
        ['GPM moyen', Math.round(a.overall.avgGPM), false],
        ['% DMG Equipe moyen', `${a.overall.avgDmgShare.toFixed(1)} %`, false],
        ['KP % moyen', `${a.overall.avgKP.toFixed(1)} %`, false],
        ['Vision Score moyen', a.overall.avgVision.toFixed(1), false],
        ['Responsabilité moyenne', a.overall.avgResp != null ? a.overall.avgResp.toFixed(1) : 'N/A', false],
    ].forEach(([label, value, isPct], i) => { kv(r, label, value, i, isPct); r++; });
    r++;

    sectionTitle(sheet, r, 'Serie Actuelle & Tendance', SPAN, C.hTend); r++;
    const wr10 = winPct(a.recent10);
    const wrGlob = a.overall.winRate;
    const trend = wr10 > wrGlob + 0.02 ? 'En progression' : wr10 < wrGlob - 0.02 ? 'En regression' : 'Stable';
    [
        ['Serie actuelle', `${a.streak.count} ${a.streak.type} consecutives`],
        ['Win Rate (10 derniers)', wr10],
        ['Win Rate global', wrGlob],
        ['Tendance recente', trend],
        ['K/D/A (10 derniers)', `${avgOf(a.recent10, 'kills').toFixed(1)} / ${avgOf(a.recent10, 'deaths').toFixed(1)} / ${avgOf(a.recent10, 'assists').toFixed(1)}`],
        ['KDA (10 derniers)', avgOf(a.recent10, 'kda').toFixed(2)],
        ['DPM (10 derniers)', Math.round(avgOf(a.recent10, 'dpm'))],
    ].forEach(([label, value], i) => { kv(r, label, value, i, label.includes('Win Rate')); r++; });
    r++;

    sectionTitle(sheet, r, 'Fatigue Cognitive — Performance par Partie dans la Session', SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Partie/Session', 'Nb Parties', 'Win Rate', 'KDA Moy.', 'K Moy.', 'D Moy.', 'A Moy.', 'CS/Min | DPM'], C.subH); r++;

    const globalWR = a.overall.winRate;
    const globalKDA = a.overall.avgKDA;

    [1, 2, 3, 4].forEach((gNum, idx) => {
        const games = a.byGameInSession[gNum] || [];
        const wr = winPct(games);
        const kAvg = avgOf(games, 'kills');
        const dAvg = avgOf(games, 'deaths');
        const aAvg = avgOf(games, 'assists');
        const kdaAvg = avgOf(games, 'kda');
        const csAvg = avgOf(games, 'csPerMin');
        const dpmAvg = Math.round(avgOf(games, 'dpm'));

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
            const kdaCell = sheet.getRow(r).getCell(4);
            if (kdaAvg >= globalKDA + 0.3) kdaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.good } };
            else if (kdaAvg <= globalKDA - 0.3) kdaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bad } };

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

    const g1 = a.byGameInSession[1] || [];
    const g4 = a.byGameInSession[4] || [];
    const fatigueWRDrop = g1.length && g4.length ? (winPct(g1) - winPct(g4)) * 100 : null;
    const fatigueDiag = fatigueWRDrop === null ? 'Données insuffisantes'
        : fatigueWRDrop > 15 ? `⚠ Fatigue forte (-${fatigueWRDrop.toFixed(1)}% WR P1→P4+)`
            : fatigueWRDrop > 5 ? `! Legere fatigue (-${fatigueWRDrop.toFixed(1)}% WR P1→P4+)`
                : `✓ Stable (${fatigueWRDrop >= 0 ? '+' : ''}${fatigueWRDrop.toFixed(1)}% WR P1→P4+)`;
    const synthRow = sheet.getRow(r);
    sheet.mergeCells(r, 1, r, SPAN);
    synthRow.getCell(1).value = `Diagnostic : ${fatigueDiag}`;
    synthRow.getCell(1).font = { italic: true, size: 10, bold: fatigueWRDrop !== null && fatigueWRDrop > 5 };
    synthRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    synthRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.warn } };
    synthRow.height = 22;
    r += 2;

    sectionTitle(sheet, r, 'Performance par Tranche Horaire', SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Tranche Horaire', 'Parties', 'Win Rate', 'KDA', 'K/D/A', 'DPM', 'Pings Nég.', 'Diagnostic'], C.subH); r++;

    Object.entries(a.timeSlots).forEach(([slot, games], idx) => {
        const wr = winPct(games);
        const diag = games.length < 3 ? 'Echantillon faible'
            : wr >= 0.55 ? 'Tranche favorable'
                : wr <= 0.45 ? 'Tranche defavorable'
                    : 'Neutre';
        const kda = games.length ? avgOf(games, 'kda').toFixed(2) : '-';
        const kdaStr = games.length
            ? `${avgOf(games, 'kills').toFixed(1)}/${avgOf(games, 'deaths').toFixed(1)}/${avgOf(games, 'assists').toFixed(1)}`
            : '-';

        tRow(r, [
            slot,
            games.length,
            games.length ? wr : '-',
            kda,
            kdaStr,
            games.length ? Math.round(avgOf(games, 'dpm')) : '-',
            games.length ? Math.round(avgOf(games, 'negativePings')) : '-',
            diag,
        ], idx, [3]);

        if (games.length) colorWRCell(r, 3, wr);
        r++;
    });
    r++;

    sectionTitle(sheet, r, 'Indicateur de Tilt — Perf. apres Victoire vs Defaite', SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Contexte', 'Parties', 'Win Rate', 'KDA', 'K/D/A', 'DPM', 'Pings Nég.', 'Diagnostic'], C.subH); r++;

    const wrW = winPct(a.afterWin);
    const wrL = winPct(a.afterLoss);
    const delta = wrW - wrL;
    const tiltDiag = delta > 0.10 ? 'Tilt detecte — Stopper apres defaite'
        : delta > 0.05 ? 'Legere regression post-defaite'
            : 'Mental stable';

    [
        ['Apres une Victoire', a.afterWin.length, wrW, avgOf(a.afterWin, 'kda').toFixed(2), `${avgOf(a.afterWin, 'kills').toFixed(1)}/${avgOf(a.afterWin, 'deaths').toFixed(1)}/${avgOf(a.afterWin, 'assists').toFixed(1)}`, Math.round(avgOf(a.afterWin, 'dpm')), Math.round(avgOf(a.afterWin, 'negativePings')), ''],
        ['Apres une Defaite', a.afterLoss.length, wrL, avgOf(a.afterLoss, 'kda').toFixed(2), `${avgOf(a.afterLoss, 'kills').toFixed(1)}/${avgOf(a.afterLoss, 'deaths').toFixed(1)}/${avgOf(a.afterLoss, 'assists').toFixed(1)}`, Math.round(avgOf(a.afterLoss, 'dpm')), Math.round(avgOf(a.afterLoss, 'negativePings')), tiltDiag],
    ].forEach((rowData, idx) => {
        tRow(r, rowData, idx, [3]);
        if (typeof rowData[2] === 'number') colorWRCell(r, 3, rowData[2]);
        r++;
    });
    r++;

    sectionTitle(sheet, r, 'Analyse des Pings par Resultat', SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Contexte', 'Pings Tot.', 'Pings Nég.', 'Danger', 'MIA', 'Retreat', 'AllIn', 'OTW'], C.subH); r++;

    const pingCtxs = [
        { label: 'Victoires', games: data.filter((m) => m.win === 'Victoire') },
        { label: 'Defaites', games: data.filter((m) => m.win !== 'Victoire') },
        { label: 'Apres victoire', games: a.afterWin },
        { label: 'Apres defaite', games: a.afterLoss },
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

        const negVal = Math.round(avgOf(games, 'negativePings'));
        if (negVal > 10) {
            const cell = sheet.getRow(r).getCell(3);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.warn } };
            cell.font = { bold: true, color: { argb: C.amber }, size: 10 };
        }
        r++;
    });
    r++;

    const lanerGames = data.filter((m) => m.fatalGanksReceived !== null && m.fatalGanksReceived !== undefined);
    if (lanerGames.length >= 5) {
        sectionTitle(sheet, r, 'Winrate par Ganks Mortels Subis (Phase de Lane)', SPAN, C.hTend); r++;
        writeHeaderRow(sheet, r, ['Ganks Subis', 'Parties', 'Win Rate', 'KDA', 'K/D/A', 'DPM', '', 'Tendance'], C.subH); r++;

        [0, 1, 2, 3].forEach((gNum, idx) => {
            const games = gNum < 3
                ? lanerGames.filter((m) => m.fatalGanksReceived === gNum)
                : lanerGames.filter((m) => m.fatalGanksReceived >= 3);

            if (!games.length) return;
            const wr = winPct(games);
            tRow(r, [
                gNum < 3 ? `${gNum} gank(s) mortel(s)` : '3+ ganks mortels',
                games.length,
                wr,
                avgOf(games, 'kda').toFixed(2),
                `${avgOf(games, 'kills').toFixed(1)}/${avgOf(games, 'deaths').toFixed(1)}/${avgOf(games, 'assists').toFixed(1)}`,
                Math.round(avgOf(games, 'dpm')),
                '',
                wr >= 0.55 ? 'Excellente resistance' : wr <= 0.45 ? 'Impact critique' : 'Stable',
            ], idx, [3]);
            colorWRCell(r, 3, wr);
            r++;
        });
        r++;
    }

    sectionTitle(sheet, r, 'Records Personnels', SPAN, C.hTend); r++;
    writeHeaderRow(sheet, r, ['Categorie', 'Valeur', 'Champion', 'Date', 'Resultat', 'Role', '', ''], C.subH); r++;

    [
        ['Meilleur KDA', a.records.bestKDA, 'kda', 2],
        ['Meilleur DPM', a.records.bestDPM, 'dpm', 0],
        ['Meilleur CS/Min', a.records.bestCS, 'csPerMin', 1],
        ['Meilleur GPM', a.records.bestGPM, 'gpm', 0],
        ['Meilleur Vision', a.records.bestVision, 'vision', 0],
    ].forEach(([label, rec, key, dec], idx) => {
        const rRow = sheet.getRow(r);
        rRow.height = 20;
        if (rec) {
            const rawVal = safeN(rec[key]);
            const val = dec === 0 ? Math.round(rawVal) : parseFloat(rawVal.toFixed(dec));
            const c1 = rRow.getCell(1);
            c1.value = label;
            c1.font = { bold: true, size: 10 };
            c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            const c2 = rRow.getCell(2);
            c2.value = val;
            c2.font = { bold: true, size: 10, color: { argb: C.accent } };
            c2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
            c2.numFmt = dec === 2 ? '0.00' : dec === 1 ? '0.0' : '0';
            [rec.champion, rec.date, rec.win, rec.role || '-', '', ''].forEach((v, i) => {
                const cell = rRow.getCell(i + 3);
                cell.value = v;
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg(idx) } };
                cell.font = { size: 10 };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            });
            colorResult(rRow.getCell(5), rec.win);
        }
        r++;
    });
}

async function rebuildExcel(data) {
    console.log('  -> Reconstruction Excel (4 feuilles)...');

    const analytics = computeAnalytics(data);
    const wb = new ExcelJS.Workbook();
    wb.creator = 'LoL Tracker';
    wb.created = new Date();

    const jobs = [
        { name: 'Donnees Brutes', fn: () => buildRawDataSheet(wb.addWorksheet('Donnees Brutes'), data) },
        { name: 'Par Champion', fn: () => buildChampionSheet(wb.addWorksheet('Par Champion'), data) },
        { name: 'Solo vs Duo', fn: () => buildDuoSoloSheet(wb.addWorksheet('Solo vs Duo'), data) },
        { name: 'Tendances Comportementales', fn: () => buildTendancesSheet(wb.addWorksheet('Tendances Comportementales'), data, analytics) },
    ];

    for (const { name, fn } of jobs) {
        try {
            fn();
            console.log(`     OK : ${name}`);
        } catch (e) {
            console.error(`     ERREUR dans "${name}" : ${e.message}`);
            console.error(e.stack);
        }
    }

    try {
        await wb.xlsx.writeFile(EXCEL_FILENAME);
    } catch (e) {
        throw new Error(`Echec sauvegarde Excel (${EXCEL_FILENAME}) : ${e.message}`);
    }
    console.log(`  -> Fichier sauvegarde : ${EXCEL_FILENAME}`);
}

module.exports = {
    rebuildExcel,
};
