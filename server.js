require('dotenv').config();

const express   = require('express');
const http      = require('http');
const WebSocket = require('ws');
const fs        = require('fs');
const path      = require('path');

const { runAll, runFetchOnly, runMigrateOnly, runExcelOnly } = require('./src/pipeline');
const { JSON_FILENAME } = require('./src/config');

// ── Serveur HTTP + WebSocket ───────────────────────────────
const app    = express();
const server = http.createServer(app);
const wss    = new WebSocket.Server({ server });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── État global ────────────────────────────────────────────
let isRunning = false;

// ── Broadcast vers tous les clients WS connectés ──────────
function broadcast(type, data = {}) {
    const msg = JSON.stringify({ type, ...data });
    wss.clients.forEach((client) => {
        if (client.readyState === WebSocket.OPEN) client.send(msg);
    });
}

// ── Patch console.log → terminal WS ───────────────────────
function patchConsole() {
    const origLog   = console.log;
    const origError = console.error;

    console.log = (...args) => {
        origLog(...args);
        broadcast('log', { level: 'info', message: args.join(' ') });
    };
    console.error = (...args) => {
        origError(...args);
        broadcast('log', { level: 'error', message: args.join(' ') });
    };

    return () => {
        console.log   = origLog;
        console.error = origError;
    };
}

// ── Lecture des stats depuis le JSON local ─────────────────
function getStats() {
    try {
        if (!fs.existsSync(JSON_FILENAME)) return { total: 0, lastUpdated: null };
        const data = JSON.parse(fs.readFileSync(JSON_FILENAME, 'utf8'));
        const sorted = [...data].sort((a, b) => new Date(b.rawDate) - new Date(a.rawDate));
        return {
            total:       data.length,
            lastUpdated: sorted[0]?.rawDate ?? null,
        };
    } catch {
        return { total: 0, lastUpdated: null };
    }
}

// ── Factory de route pour chaque mode ─────────────────────
function runMode(modeFn, modeName) {
    return async (req, res) => {
        if (isRunning) {
            return res.status(409).json({ error: 'Un mode est déjà en cours d\'exécution.' });
        }

        isRunning = true;
        const unpatch = patchConsole();

        broadcast('start', { mode: modeName });
        res.json({ ok: true });

        try {
            await modeFn();
            broadcast('done', { mode: modeName, stats: getStats() });
        } catch (err) {
            broadcast('error', {
                message: err.response ? JSON.stringify(err.response.data) : err.message,
            });
        } finally {
            unpatch();
            isRunning = false;
        }
    };
}

// ── Routes API ─────────────────────────────────────────────
app.get('/api/status', (_req, res) => {
    res.json({ isRunning, stats: getStats() });
});

app.get('/api/data', (_req, res) => {
    try {
        if (!fs.existsSync(JSON_FILENAME)) return res.json([]);
        const data = JSON.parse(fs.readFileSync(JSON_FILENAME, 'utf8'));
        res.json(data);
    } catch (err) {
        res.status(500).json({ error: 'Impossible de lire les données : ' + err.message });
    }
});

app.post('/api/run/all',     runMode(runAll,         'all'));
app.post('/api/run/fetch',   runMode(runFetchOnly,   'fetch'));
app.post('/api/run/migrate', runMode(runMigrateOnly, 'migrate'));
app.post('/api/run/excel',   runMode(runExcelOnly,   'excel'));

// ── Log WS connect/disconnect ──────────────────────────────
wss.on('connection', () => {
    const origLog = console.log;
    origLog('  [WS] Client connecté');
});

// ── Démarrage ──────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    // eslint-disable-next-line no-console
    process.stdout.write(`\n  LoL Tracker → http://localhost:${PORT}\n\n`);
});
