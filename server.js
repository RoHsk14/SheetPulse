const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const QRCode = require('qrcode');
const express = require('express');

const app = express();
app.use(express.json());

let qrCodeData = null;
let clientStatus = 'initializing';

const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: {
        headless: true,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-accelerated-2d-canvas',
            '--no-first-run',
            '--no-zygote',
            '--disable-gpu',
            '--single-process'
        ]
    }
});

client.on('qr', async (qr) => {
    qrCodeData = qr;
    clientStatus = 'awaiting_scan';
    console.log('--- SCANNEZ LE QR CODE CI-DESSOUS AVEC WHATSAPP ---');
    qrcode.generate(qr, { small: true });
});

client.on('ready', () => {
    qrCodeData = null;
    clientStatus = 'connected';
    console.log('Félicitations, le bot WhatsApp est connecté et prêt !');
});

client.on('disconnected', (reason) => {
    clientStatus = 'disconnected';
    console.log('WhatsApp déconnecté:', reason);
});

app.get('/', (req, res) => {
    if (clientStatus === 'connected') {
        res.send(`
            <html>
            <head><meta charset="utf-8"><title>WS Sheets Bot</title>
            <style>body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background:#f5f5f5}
            .card{background:white;padding:40px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1);text-align:center}
            .status{color:#22c55e;font-size:18px;font-weight:600}
            </style></head>
            <body><div class="card"><div class="status">✅ Connecté à WhatsApp</div><p>Le bot est prêt à recevoir des commandes.</p></div></body></html>
        `);
    } else if (qrCodeData) {
        QRCode.toDataURL(qrCodeData, (err, url) => {
            if (err) return res.status(500).send('Erreur QR');
            res.send(`
                <html>
                <head><meta charset="utf-8"><title>WS Sheets Bot</title>
                <style>body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background:#f5f5f5}
                .card{background:white;padding:40px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1);text-align:center}
                img{width:300px;height:300px}
                p{color:#666}
                </style></head>
                <body><div class="card"><h2>Scannez ce QR code avec WhatsApp</h2>
                <img src="${url}" alt="QR Code"/>
                <p>Ouvrez WhatsApp > Appareils connectés > Connecter un appareil</p></div></body></html>
            `);
        });
    } else {
        res.send(`
            <html>
            <head><meta charset="utf-8"><title>WS Sheets Bot</title>
            <style>body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background:#f5f5f5}
            .card{background:white;padding:40px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1);text-align:center}
            .status{color:#f59e0b;font-size:18px;font-weight:600}
            </style></head>
            <body><div class="card"><div class="status">⏳ Initialisation en cours...</div><p>Veuillez patienter, le QR code va apparaître.</p></div></body></html>
        `);
    }
});

app.post('/webhook', async (req, res) => {
    const { message } = req.body;

    const groupId = process.env.GROUP_ID || '120363430608871862@g.us';

    try {
        await client.sendMessage(groupId, message);
        console.log('Message envoyé au groupe avec succès !');
        return res.status(200).json({ success: true });
    } catch (error) {
        console.error('Erreur lors de l\'envoi du message:', error);
        return res.status(500).json({ error: 'Échec de l\'envoi' });
    }
});

app.get('/status', (req, res) => {
    res.json({ status: clientStatus, hasQr: !!qrCodeData });
});

client.initialize();

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Serveur démarré sur le port ${PORT}`);
});

process.on('uncaughtException', (err) => {
    console.error('Erreur non rattrapée:', err);
});
