require('dotenv').config();

const RIOT_API_KEY = process.env.RIOT_API_KEY;
const REGION = 'europe';
const REGION_PLATFORM = 'euw1';

const MY_NAME = 'enyαV';
const MY_TAG = 'EUW';
const DUO_NAME = 'Ahristocats';
const DUO_TAG = 'EUW';

const EXCEL_FILENAME = 'Suivi_Comportemental_Challenger.xlsx';
const JSON_FILENAME = 'historique_matches.json';
const MATCHES_TO_FETCH = 200;
const QUEUE_FILTER = 420;
const API_DELAY_MS = 1500;

const C = {
    hRaw: 'FF0F172A',
    hChamp: 'FF2E1065',
    hDuo: 'FF172554',
    hTend: 'FF052E16',
    subH: 'FF1E3A5F',
    white: 'FFFFFFFF',
    accent: 'FF4F46E5',

    winFg: 'FF166534', winBg: 'FFDCFCE7',
    lossFg: 'FF991B1B', lossBg: 'FFFEE2E2',

    rowA: 'FFF9FAFB',
    rowB: 'FFFFFFFF',

    good: 'FFD9F2DC',
    warn: 'FFFFF9C4',
    bad: 'FFFDE8E8',
    goodFg: 'FF166534',
    badFg: 'FF991B1B',
    purple: 'FFA855F7',
    amber: 'FFB45309',

    grpIdentite: 'FF1E3A5F',
    grpCombat: 'FF3B0764',
    grpFarm: 'FF052E16',
    grpVision: 'FF1C3035',
    grpDivers: 'FF1C1917',
    grpRank: 'FF431407',
    grpPings: 'FF1A1A2E',
    grpGanks: 'FF2D1B4E',

    rankIron: 'FF8B8B8B',
    rankBronze: 'FFB87333',
    rankSilver: 'FFA8A9AD',
    rankGold: 'FFFFD700',
    rankPlatinum: 'FF0BC4B5',
    rankEmerald: 'FF50C878',
    rankDiamond: 'FFB9F2FF',
    rankMaster: 'FF9B59B6',
    rankGrandmaster: 'FFE74C3C',
    rankChallenger: 'FFF0E68C',
};

const TIER_ORDER = {
    IRON: 0, BRONZE: 4, SILVER: 8, GOLD: 12,
    PLATINUM: 16, EMERALD: 20, DIAMOND: 24,
    MASTER: 28, GRANDMASTER: 29, CHALLENGER: 30,
};

const DIV_ORDER = { IV: 0, III: 1, II: 2, I: 3 };

function rankToScore(tier, division) {
    if (!tier) return -1;
    const t = TIER_ORDER[tier.toUpperCase()] ?? -1;
    if (t === -1) return -1;
    if (t >= 28) return t;
    return t + (DIV_ORDER[division] ?? 0);
}

function scoreToLabel(score) {
    if (score < 0) return 'Non classé';
    if (score >= 30) return 'Challenger';
    if (score >= 29) return 'GrandMaster';
    if (score >= 28) return 'Master';
    const tiers = ['Iron', 'Bronze', 'Silver', 'Gold', 'Platinum', 'Emerald', 'Diamond'];
    const divs = ['IV', 'III', 'II', 'I'];
    const ti = Math.floor(score / 4);
    const di = score % 4;
    return tiers[ti] + ' ' + divs[di];
}

function rankColor(score) {
    if (score < 4) return C.rankIron;
    if (score < 8) return C.rankBronze;
    if (score < 12) return C.rankSilver;
    if (score < 16) return C.rankGold;
    if (score < 20) return C.rankPlatinum;
    if (score < 24) return C.rankEmerald;
    if (score < 28) return C.rankDiamond;
    if (score < 29) return C.rankMaster;
    if (score < 30) return C.rankGrandmaster;
    return C.rankChallenger;
}

module.exports = {
    RIOT_API_KEY,
    REGION,
    REGION_PLATFORM,
    MY_NAME,
    MY_TAG,
    DUO_NAME,
    DUO_TAG,
    EXCEL_FILENAME,
    JSON_FILENAME,
    MATCHES_TO_FETCH,
    QUEUE_FILTER,
    API_DELAY_MS,
    C,
    rankToScore,
    scoreToLabel,
    rankColor,
};
