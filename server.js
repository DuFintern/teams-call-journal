const express = require('express');
const bodyParser = require('body-parser');
const sqlite3 = require('sqlite3').verbose();
const { Parser } = require('json2csv');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// SQLite-Datenbank initialisieren
const db = new sqlite3.Database('./call_journal.db', (err) => {
    if (err) console.error('Fehler beim Öffnen der DB:', err);
    else console.log('SQLite Datenbank verbunden.');
});

// Tabelle anlegen
db.run(`CREATE TABLE IF NOT EXISTS calls (
    id TEXT PRIMARY KEY,
    userEmail TEXT,
    phoneNumber TEXT,
    direction TEXT,
    startTime TEXT,
    durationSeconds INTEGER,
    note TEXT,
    status TEXT DEFAULT 'PENDING'
)`);

// 1. Webhook Endpoint für Microsoft Graph (CallRecords Notification)
app.post('/api/lifecycle/notifications', (req, res) => {
    // Validierungsschritt für Microsoft Graph Webhook-Aktivierung
    if (req.query && req.query.validationToken) {
        res.set('Content-Type', 'text/plain');
        return res.status(200).send(req.query.validationToken);
    }

    const notifications = req.body.value;
    if (notifications) {
        notifications.forEach(notification => {
            // Hier empfängt der Server das Event für beendete Anrufe
            const callData = notification.resourceData;
            
            // Beispielsätze aus der Graph API verarbeiten
            const callId = callData.id || `call_${Date.now()}`;
            const userEmail = callData.organizer?.userPrincipalName || 'user@company.com';
            const phoneNumber = callData.participants?.[0]?.identity?.phone?.id || 'Unbekannt';
            const direction = callData.type === 'groupCall' ? 'INBOUND' : 'OUTBOUND';
            const startTime = new Date().toISOString();
            const durationSeconds = callData.durationSeconds || 0;

            const stmt = db.prepare(`INSERT OR IGNORE INTO calls (id, userEmail, phoneNumber, direction, startTime, durationSeconds) VALUES (?, ?, ?, ?, ?, ?)`);
            stmt.run(callId, userEmail, phoneNumber, direction, startTime, durationSeconds);
            stmt.finalize();
        });
    }
    res.status(202).send();
});

// 2. Anrufe für den aktuellen Teams-User abrufen
app.get('/api/calls', (req, res) => {
    const userEmail = req.query.email;
    db.all(`SELECT * FROM calls WHERE userEmail = ? ORDER BY startTime DESC`, [userEmail], (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(rows);
    });
});

// 3. Notiz zu einem Anruf speichern
app.post('/api/calls/note', (req, res) => {
    const { callId, note } = req.body;
    db.run(`UPDATE calls SET note = ?, status = 'COMPLETED' WHERE id = ?`, [note, callId], function(err) {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ success: true });
    });
});

// 4. Zentraler CSV-Export für den Admin
app.get('/api/export/csv', (req, res) => {
    db.all(`SELECT id, userEmail, phoneNumber, direction, startTime, durationSeconds, note, status FROM calls`, [], (err, rows) => {
        if (err) return res.status(500).send('Fehler beim Export');
        try {
            const json2csvParser = new Parser();
            const csv = json2csvParser.parse(rows);
            res.header('Content-Type', 'text/csv');
            res.attachment('Anruf_Journal_Export.csv');
            return res.send(csv);
        } catch (csvErr) {
            res.status(500).send('CSV Parsing Fehler');
        }
    });
});

app.listen(PORT, () => console.log(`Journal App Server läuft auf Port ${PORT}`));