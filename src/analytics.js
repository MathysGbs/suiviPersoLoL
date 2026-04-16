const { safeN, avgOf, winPct, pushTo } = require('./utils');

function computeAnalytics(data) {
    if (!data.length) return null;

    const champMap = {};
    const byGameInSession = {};
    const byRole = {};
    const timeSlots = {
        'Nuit (0h-6h)': [],
        'Matin (6h-12h)': [],
        'Apres-midi (12h-18h)': [],
        'Soir (18h-24h)': [],
    };

    data.forEach((m) => {
        pushTo(champMap, m.champion, m);
        const g = Math.min(safeN(m.gameInSession, 1), 4);
        pushTo(byGameInSession, g, m);
        if (m.role) pushTo(byRole, m.role, m);

        const h = new Date(m.rawDate).getHours();
        if (h < 6) timeSlots['Nuit (0h-6h)'].push(m);
        else if (h < 12) timeSlots['Matin (6h-12h)'].push(m);
        else if (h < 18) timeSlots['Apres-midi (12h-18h)'].push(m);
        else timeSlots['Soir (18h-24h)'].push(m);
    });

    const byChampion = Object.entries(champMap)
        .map(([name, games]) => ({ name, games }))
        .sort((a, b) => b.games.length - a.games.length);

    const afterWin = [];
    const afterLoss = [];
    for (let i = 1; i < data.length; i++) {
        (data[i - 1].win === 'Victoire' ? afterWin : afterLoss).push(data[i]);
    }

    let streakCount = 0;
    let streakType = '';
    for (let i = data.length - 1; i >= 0; i--) {
        const isW = data[i].win === 'Victoire';
        if (i === data.length - 1) {
            streakType = isW ? 'Victoires' : 'Defaites';
            streakCount = 1;
        } else if ((isW && streakType === 'Victoires') || (!isW && streakType === 'Defaites')) streakCount++;
        else break;
    }

    const valid = data.filter((m) => m.kda != null);
    const topOf = (arr, key) => arr.filter((m) => m[key]).sort((a, b) => safeN(b[key]) - safeN(a[key]))[0];

    const records = {
        bestKDA: topOf(valid, 'kda'),
        bestDPM: topOf(valid, 'dpm'),
        bestCS: topOf(valid, 'csPerMin'),
        bestGPM: topOf(valid, 'gpm'),
        bestVision: topOf(valid, 'vision'),
    };

    const gankData = data.filter((m) => m.role === 'JUNGLE' && m.ganksPerformed != null);
    const byGankCount = {};
    gankData.forEach((m) => {
        const bucket = Math.min(m.ganksPerformed || 0, 10);
        pushTo(byGankCount, bucket, m);
    });

    return {
        byChampion,
        byRole,
        soloGames: data.filter((m) => m.type === 'Solo'),
        duoGames: data.filter((m) => m.type === 'Duo'),
        yuumiGames: data.filter((m) => m.withYuumi === 'Oui'),
        noYuumi: data.filter((m) => m.type === 'Duo' && m.withYuumi !== 'Oui'),
        yuumiAlliee: data.filter((m) => m.yuumiAlliee === 'Oui'),
        byGameInSession,
        byGankCount,
        timeSlots,
        afterWin,
        afterLoss,
        streak: { count: streakCount, type: streakType },
        recent10: data.slice(-10),
        records,
        overall: {
            total: data.length,
            winRate: winPct(data),
            avgKDA: avgOf(data, 'kda'),
            avgKills: avgOf(data, 'kills'),
            avgDeaths: avgOf(data, 'deaths'),
            avgAssists: avgOf(data, 'assists'),
            avgCS: avgOf(data, 'csPerMin'),
            avgDPM: avgOf(data, 'dpm'),
            avgGPM: avgOf(data, 'gpm'),
            avgDmgShare: avgOf(data, 'dmgShare'),
            avgKP: avgOf(data, 'kp'),
            avgVision: avgOf(data, 'vision'),
            avgResp: avgOf(data, 'responsabilite'),
        },
    };
}

module.exports = { computeAnalytics };
