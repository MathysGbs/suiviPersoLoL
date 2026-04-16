require('dotenv').config();

const { runAll, runFetchOnly, runMigrateOnly, runExcelOnly } = require('./src/pipeline');

const mode = (process.argv[2] || 'all').toLowerCase();

const modes = {
    all:     runAll,
    fetch:   runFetchOnly,
    migrate: runMigrateOnly,
    excel:   runExcelOnly,
};

if (!modes[mode]) {
    console.log('Mode invalide. Utilise: all | fetch | migrate | excel');
    process.exitCode = 1;
} else {
    modes[mode]().catch((err) => {
        console.error('\n[ERREUR FATALE]');
        console.error(err.response ? err.response.data : err.message);
        process.exit(1);
    });
}