const {
    loadData,
    saveData,
    resolvePlayers,
    fetchNewMatches,
    migrateOldEntries,
    recalculateTimeline,
} = require('./data-service');
const { QUEUE_FILTER } = require('./config');
const { rebuildExcel } = require('./excel-service');

async function runFetchOnly() {
    console.log('1. Lecture de la base locale...');
    let data = loadData();
    console.log(`   -> ${data.length} partie(s) en memoire.`);

    console.log('2. Connexion Riot Games...');
    const { player, duo } = await resolvePlayers();
    console.log(`   -> Joueur : ${player.gameName}#${player.tagLine}`);
    console.log(`   -> Duo    : ${duo.gameName}#${duo.tagLine}`);

    console.log(`3. Recuperation des IDs (queue: ${QUEUE_FILTER || 'toutes'})...`);
    const { newIds, newMatches, ok, ko } = await fetchNewMatches(data, player.puuid, duo.puuid);

    if (newIds.length === 0) {
        console.log('\n-> Aucune nouvelle partie detectee.');
        return { data, added: 0, ok: 0, ko: 0, player, duo };
    }

    console.log('\n4. Consolidation et recalcul chronologique...');
    data = recalculateTimeline(data.concat(newMatches));

    console.log('   -> Sauvegarde JSON...');
    saveData(data);
    console.log(`   -> ${data.length} parties au total enregistrees.`);

    return { data, added: newIds.length, ok, ko, player, duo };
}

async function runMigrateOnly() {
    console.log('1. Lecture de la base locale...');
    let data = loadData();
    console.log(`   -> ${data.length} partie(s) en memoire.`);

    console.log('2. Connexion Riot Games...');
    const { player, duo } = await resolvePlayers();

    console.log('3. Migration des entrées incomplètes...');
    const migrated = await migrateOldEntries(data, player.puuid, duo.puuid);

    if (migrated > 0) {
        data = recalculateTimeline(data);
        saveData(data);
        console.log(`   -> ${migrated} entree(s) corrigee(s) et sauvegardee(s).`);
    } else {
        console.log('   -> Aucune migration necessaire.');
    }

    return { data, migrated };
}

async function runExcelOnly() {
    console.log('1. Lecture de la base locale...');
    const data = loadData();
    console.log(`   -> ${data.length} partie(s) chargee(s).`);

    console.log('2. Reconstruction Excel...');
    await rebuildExcel(data);
    console.log('\n[SUCCES] Fichier Excel reconstruit.');
}

async function runAll() {
    const fetchResult = await runFetchOnly();

    console.log('\n5. Migration des anciennes entrées (si besoin)...');
    const migrated = await migrateOldEntries(fetchResult.data, fetchResult.player.puuid, fetchResult.duo.puuid);

    let data = fetchResult.data;
    if (migrated > 0) {
        data = recalculateTimeline(data);
        saveData(data);
    }

    console.log('6. Reconstruction Excel...');
    await rebuildExcel(data);

    if (fetchResult.added === 0 && migrated === 0) {
        console.log('\n[SUCCES] Tableau de bord deja a jour.');
    } else {
        console.log(`\n[SUCCES] ${fetchResult.ok} partie(s) ajoutee(s)${fetchResult.ko ? `, ${fetchResult.ko} ignoree(s)` : ''}${migrated ? `, ${migrated} migree(s)` : ''}.`);
    }
}

module.exports = {
    runAll,
    runFetchOnly,
    runMigrateOnly,
    runExcelOnly,
};
