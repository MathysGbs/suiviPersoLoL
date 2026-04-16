const { C } = require('./config');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const safeN = (v, d = 0) => {
    const n = parseFloat(v);
    return Number.isNaN(n) ? d : n;
};
const safeV = (v, d = '-') => (v !== undefined && v !== null ? v : d);
const avgOf = (arr, key) => arr.length ? arr.reduce((s, x) => s + safeN(x[key]), 0) / arr.length : 0;
const winPct = (arr) => arr.length ? (arr.filter((g) => g.win === 'Victoire').length / arr.length) : 0;
const rowBg = (idx) => (idx % 2 === 0 ? C.rowA : C.rowB);

function secToMmSs(totalSeconds) {
    if (totalSeconds == null || Number.isNaN(totalSeconds)) return '-';
    const s = Math.round(totalSeconds);
    const m = Math.floor(s / 60);
    const sec = s % 60;
    return m + ':' + String(sec).padStart(2, '0');
}

function pushTo(map, key, value) {
    if (!map[key]) map[key] = [];
    map[key].push(value);
}

function bestMultiKill(me) {
    if ((me.pentaKills || 0) > 0) return 'Penta';
    if ((me.quadraKills || 0) > 0) return 'Quadra';
    if ((me.tripleKills || 0) > 0) return 'Triple';
    if ((me.doubleKills || 0) > 0) return 'Double';
    return '-';
}

function recalculateTimeline(data) {
    data.sort((a, b) => new Date(a.rawDate) - new Date(b.rawDate));
    let currentSessionId = 1;
    let gameInSession = 1;

    for (let i = 0; i < data.length; i++) {
        if (i > 0) {
            const diffH = (new Date(data[i].rawDate) - new Date(data[i - 1].rawDate)) / 3_600_000;
            if (diffH < 2) gameInSession++;
            else {
                currentSessionId++;
                gameInSession = 1;
            }
        }
        data[i].sessionId = currentSessionId;
        data[i].gameInSession = gameInSession;
    }

    return data;
}

module.exports = {
    sleep,
    safeN,
    safeV,
    avgOf,
    winPct,
    rowBg,
    secToMmSs,
    pushTo,
    bestMultiKill,
    recalculateTimeline,
};
