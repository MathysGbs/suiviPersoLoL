require('dotenv').config();

const readline = require('node:readline/promises');
const { stdin, stdout } = require('node:process');

const { runAll, runFetchOnly, runMigrateOnly, runExcelOnly } = require('./src/pipeline');
const { loadData } = require('./src/data-service');
const { MATCHES_TO_FETCH, MAX_MATCHES_TO_FETCH } = require('./src/config');

const mode = (process.argv[2] || 'all').toLowerCase();

const modes = {
    all:     runAll,
    fetch:   runFetchOnly,
    migrate: runMigrateOnly,
    excel:   runExcelOnly,
};

async function askMatchesToFetch(defaultValue) {
    if (!stdin.isTTY || !stdout.isTTY) {
        return defaultValue;
    }

    const rl = readline.createInterface({ input: stdin, output: stdout });

    try {
        const shouldModify = await rl.question('Modifier le nombre de games avant la récupération ? (o/N) ');
        if (!/^o(ui)?$/i.test(shouldModify.trim())) {
            return defaultValue;
        }

        const answer = await rl.question(`Nombre de games à récupérer [${defaultValue}] `);
        const parsed = Number.parseInt(answer, 10);
        if (!Number.isFinite(parsed) || parsed <= 0) {
            console.log(`Valeur invalide, utilisation de ${defaultValue}.`);
            return defaultValue;
        }

        if (parsed > MAX_MATCHES_TO_FETCH) {
            console.log(`Valeur trop élevée, utilisation de ${MAX_MATCHES_TO_FETCH}.`);
            return MAX_MATCHES_TO_FETCH;
        }

        return parsed;
    } finally {
        rl.close();
    }
}

function getDefaultMatchesToFetch() {
    try {
        const historyCount = loadData().length;
        return historyCount > 0 ? historyCount : MATCHES_TO_FETCH;
    } catch {
        return MATCHES_TO_FETCH;
    }
}

async function main() {
    const runner = modes[mode];
    if (!runner) {
        console.log('Mode invalide. Utilise: all | fetch | migrate | excel');
        process.exitCode = 1;
        return;
    }

    const options = {};
    if (mode === 'fetch' || mode === 'all') {
        options.matchesToFetch = await askMatchesToFetch(getDefaultMatchesToFetch());
    }

    await runner(options);
}

main().catch((err) => {
        console.error('\n[ERREUR FATALE]');
        console.error(err.response ? err.response.data : err.message);
        process.exit(1);
    });