require('dotenv').config({ path: __dirname + '/.env' });
const path = require('path');
const os = require('os');
const authDataPath = path.join(os.homedir(), '.wwebjs_auth');
const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const QRCode = require('qrcode');
const express = require('express');
const { google } = require('googleapis');
const { createClient } = require('@supabase/supabase-js');
const WebSocket = require('ws');

const PORT = process.env.PORT || 3001;
const POLL_INTERVAL = 30000;
const SUPABASE_URL = process.env.NEXT_PUBLIC_SUPABASE_URL || '';
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || '';
const SA_KEY_PATH = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH || './service-account-key.json';

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY, { realtime: { transport: WebSocket } });

let qrCodeData = null;
let pairingCodeData = null;
let clientStatus = 'initializing';
let pollTimer = null;

function getGoogleAuth() {
  try {
    const fs = require('fs');
    const path = require('path');

    let credentials;
    const envKey = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
    if (envKey) {
      credentials = JSON.parse(envKey);
    } else {
      const filePath = path.resolve(SA_KEY_PATH);
      credentials = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    }

    return new google.auth.JWT({
      email: credentials.client_email,
      key: credentials.private_key,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
  } catch (e) {
    console.error('Google Auth init error:', e.message);
    return null;
  }
}

function extractSheetId(url) {
  if (!url) return null;
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!m) return null;
  const id = m[1];
  if (!/^[a-zA-Z0-9_-]{30,}$/.test(id)) return null;
  return id;
}

async function checkSheets() {
  if (clientStatus !== 'connected') return;

  try {
    const { data: shops, error } = await supabase
      .from('settings')
      .select('shop_slug, shop_name, google_sheet_url, whatsapp_group_id, google_sheet_columns')
      .eq('whatsapp_enabled', true)
      .not('google_sheet_url', 'is', null)
      .not('whatsapp_group_id', 'is', null);

    if (error) { console.error('Supabase error:', error.message); return; }
    if (!shops || shops.length === 0) return;

    const auth = getGoogleAuth();
    if (!auth) { console.error('Google Auth non disponible'); return; }
    const sheets = google.sheets({ version: 'v4', auth });

    for (const shop of shops) {
      const sheetId = extractSheetId(shop.google_sheet_url);
      if (!sheetId) continue;

      try {
        const res = await sheets.spreadsheets.values.get({
          spreadsheetId: sheetId,
          range: 'A:K',
          valueRenderOption: 'UNFORMATTED_VALUE',
        });

        const rows = res.data.values;
        if (!rows || rows.length < 2) continue;

        const headerRow = rows[0] || [];
        const statutCol = headerRow.findIndex(h => (h || '').toLowerCase().includes('statut'));
        if (statutCol === -1) continue;

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || !row[1]) continue;

          const statut = (row[statutCol] || '').trim().toLowerCase();
          if (statut === 'envoye' || statut === 'envoyé' || statut === 'envoye ') continue;

          const date = row[0] || '';
          const clientName = row[1] || '';
          const tel = row[2] || '';
          const adresse = row[3] || '';
          const quartier = row[4] || '';
          const produit = row[5] || '';
          const qte = row[6] || '';
          const total = row[7] || '';
          const devise = row[8] || 'FCFA';

          const msg = '*NOUVELLE COMMANDE* \n\n'
            + '*Client :* ' + clientName + '\n'
            + '*Tel :* ' + tel + '\n'
            + '*Adresse :* ' + adresse + ' (' + quartier + ')\n'
            + '*Produit :* ' + produit + ' x' + qte + '\n'
            + '*Total :* ' + total + ' ' + devise + '\n\n'
            + '⚡ _Commande du ' + date + '_';

          try {
            await client.sendMessage(shop.whatsapp_group_id, msg);
            console.log('[' + shop.shop_slug + '] Commande ligne ' + (i + 1) + ' envoyee');
          } catch (sendErr) {
            console.error('[' + shop.shop_slug + '] Erreur envoi ligne ' + (i + 1) + ': ' + sendErr.message);
            continue;
          }

          try {
            await sheets.spreadsheets.values.update({
              spreadsheetId: sheetId,
              range: String.fromCharCode(65 + statutCol) + (i + 1),
              valueInputOption: 'RAW',
              resource: { values: [['Envoye']] },
            });
          } catch (_) { }
        }
      } catch (sheetErr) {
        if (sheetErr.code === 404) console.error('[' + shop.shop_slug + '] Sheet introuvable');
        else if (sheetErr.code === 403) console.error('[' + shop.shop_slug + '] Acces refuse au sheet');
        else console.error('[' + shop.shop_slug + '] Erreur lecture sheet:', sheetErr.message);
      }
    }
  } catch (err) {
    console.error('Erreur checkSheets:', err.message);
  }
}

function startPolling() {
  stopPolling();
  console.log('Surveillance des sheets demarree (toutes les ' + (POLL_INTERVAL / 1000) + 's)');
  checkSheets();
  pollTimer = setInterval(checkSheets, POLL_INTERVAL);
}

function stopPolling() {
  if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
}

const router = express.Router();
router.use(express.json());

router.use(function (req, res, next) {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  next();
});

router.get('/status', function (req, res) {
  res.json({
    status: clientStatus,
    hasQr: !!qrCodeData,
    hasPairing: !!pairingCodeData,
  });
});

router.get('/qr-image', function (req, res) {
  if (!qrCodeData) return res.status(404).send('Aucun QR');
  QRCode.toDataURL(qrCodeData, function (err, url) {
    if (err) return res.status(500).send('Erreur');
    const b = url.replace(/^data:image\/png;base64,/, '');
    res.writeHead(200, { 'Content-Type': 'image/png', 'Cache-Control': 'no-cache' });
    res.end(Buffer.from(b, 'base64'));
  });
});

router.post('/pairing', async function (req, res) {
  const phone = req.body.phone;
  if (!phone) return res.status(400).json({ error: 'Numero requis' });
  if (clientStatus === 'connected') return res.status(400).json({ error: 'Deja connecte' });
  try { await client.destroy(); } catch (_) { }
  qrCodeData = null; pairingCodeData = null; clientStatus = 'initializing';
  client = new Client({
    authStrategy: new LocalAuth({ dataPath: authDataPath }),
    webVersion: '2.2401.1',
    webVersionCache: {
      type: 'local',
      path: path.join(__dirname, '.wwebjs_cache'),
    },
    puppeteer: {
      headless: true,
      executablePath: process.env.CHROME_PATH || undefined,
      protocolTimeout: 300000,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--no-first-run',
        '--no-zygote',
        '--disable-gpu',
        '--single-process',
        '--disable-gl-drawing-for-tests',
      ]
    }, pairWithPhoneNumber: { phoneNumber: phone },
  });
  attachClientEvents();
  client.initialize();
  res.json({ success: true, message: 'Code de couplage envoye' });
});

router.post('/reset', async function (req, res) {
  try { await client.destroy(); } catch (_) { }
  qrCodeData = null; pairingCodeData = null; clientStatus = 'initializing';
  client = new Client({
    authStrategy: new LocalAuth({ dataPath: authDataPath }),
    webVersion: '2.2401.1',
    webVersionCache: {
      type: 'local',
      path: path.join(__dirname, '.wwebjs_cache'),
    },
    puppeteer: {
      headless: true,
      protocolTimeout: 300000,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--no-first-run',
        '--no-zygote',
        '--disable-gpu',
        '--single-process',
        '--disable-gl-drawing-for-tests',
      ],
    },
  });
  attachClientEvents();
  client.initialize();
  res.json({ success: true, message: 'Bot réinitialise' });
});

router.get('/groups', async function (req, res) {
  if (clientStatus !== 'connected') return res.status(503).json({ error: 'WhatsApp non connecte' });
  try {
    const chats = await client.getChats();
    const groups = chats.filter(function (c) { return c.isGroup; }).map(function (c) {
      return { id: c.id._serialized, name: c.name || c.id._serialized };
    });
    res.json(groups);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

let client = null;

function attachClientEvents() {
  client.on('qr', function (qr) {
    qrCodeData = qr; pairingCodeData = null; clientStatus = 'awaiting_scan';
    console.log('--- QR CODE GENERATED ---');
    qrcode.generate(qr, { small: true });
  });
  client.on('authenticated', function () { console.log('WhatsApp auth OK'); });
  client.on('auth_failure', function (m) { console.error('WhatsApp auth FAIL:', m); });
  client.on('ready', async function () {
    qrCodeData = null; pairingCodeData = null; clientStatus = 'connected';
    console.log('WhatsApp connecte !');
    try {
      const chats = await client.getChats();
      const groups = chats.filter(function (c) { return c.isGroup; });
      if (groups.length > 0) {
        console.log('--- GROUPES DISPONIBLES (' + groups.length + ') ---');
        groups.forEach(function (g) { console.log('  ' + g.name + ' | ' + g.id._serialized); });
      }
    } catch (_) { }
    startPolling();
  });
  client.on('disconnected', function (reason) {
    clientStatus = 'disconnected';
    console.log('WhatsApp deconnecte:', reason);
    stopPolling();
  });
}

client = new Client({
  authStrategy: new LocalAuth({ dataPath: authDataPath }),
  webVersion: '2.2401.1',
  webVersionCache: {
    type: 'local',
    path: path.join(__dirname, '.wwebjs_cache'),
  },
  puppeteer: {
    headless: true,
    protocolTimeout: 300000,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-accelerated-2d-canvas',
      '--no-first-run',
      '--no-zygote',
      '--disable-gpu',
      '--single-process',
      '--disable-gl-drawing-for-tests',
    ],
  },
});
attachClientEvents();
client.initialize();



process.on('uncaughtException', function (err) {
  console.error('Erreur:', err.message);
});

module.exports = router;
