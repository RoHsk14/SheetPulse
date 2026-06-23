const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const QRCode = require('qrcode');
const express = require('express');
const fs = require('fs');
const { google } = require('googleapis');

const CONFIG_PATH = './config.json';
const SA_PATH = './service-account.json';
const POLL_INTERVAL = 30000;

function loadConfig() {
    try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')); }
    catch { return { groupId: '', sheetId: '', publicUrl: '' }; }
}
function saveConfig(c) {
    fs.writeFileSync(CONFIG_PATH, JSON.stringify(c, null, 2));
}

const app = express();
app.use(express.json());
const cors = require('cors');
app.use(cors({
  origin: '*',
  methods: ['GET','POST','PUT','DELETE','OPTIONS'],
  allowedHeaders: ['Content-Type','Authorization']
}));

let config = loadConfig();
let qrCodeData = null;
let pairingCodeData = null;
let clientStatus = 'initializing';
let pollTimer = null;
let sheetsService = null;

function initSheets() {
    try {
        if (!fs.existsSync(SA_PATH)) { console.log('service-account.json manquant'); return null; }
        var auth = new google.auth.GoogleAuth({
            keyFile: SA_PATH,
            scopes: ['https://www.googleapis.com/auth/spreadsheets']
        });
        return google.sheets({ version: 'v4', auth: auth });
    } catch (e) { console.log('Erreur init Sheets:', e.message); return null; }
}

async function checkSheet() {
    if (clientStatus !== 'connected' || !config.sheetId || !sheetsService) return;
    try {
        var res = await sheetsService.spreadsheets.values.get({
            spreadsheetId: config.sheetId,
            range: 'A:K',
            valueRenderOption: 'UNFORMATTED_VALUE'
        });
        var rows = res.data.values;
        if (!rows || rows.length < 2) return;

        var lastRow = config.lastProcessedRow || 1;
        var newRows = [];
        for (var i = lastRow; i < rows.length; i++) {
            if (rows[i] && rows[i][1]) newRows.push({ index: i + 1, data: rows[i] });
        }
        if (newRows.length === 0) return;

        for (var n = 0; n < newRows.length; n++) {
            var r = newRows[n];
            var row = r.data;
            var date = row[0] || '';
            var clientName = row[1] || '';
            var tel = row[2] || '';
            var adresse = row[3] || '';
            var quartier = row[4] || '';
            var produit = row[5] || '';
            var qte = row[6] || '';
            var total = row[7] || '';
            var devise = row[8] || 'FCFA';
            var statut = row[9] || 'En attente';

            var msg = '📦 *NOUVELLE COMMANDE* 📦\n\n'
                + '*Client :* ' + clientName + '\n'
                + '*Tel :* ' + tel + '\n'
                + '*Adresse :* ' + adresse + ' (' + quartier + ')\n'
                + '*Produit :* ' + produit + ' x' + qte + '\n'
                + '*Total :* ' + total + ' ' + devise + '\n'
                + '*Statut :* ' + statut + '\n\n'
                + '⚡ _Commande du ' + date + '_';

            try {
                await client.sendMessage(config.groupId, msg);
                console.log('Commande ' + r.index + ' envoyee au groupe');
            } catch (e) {
                console.log('Erreur envoi commande ' + r.index + ': ' + e.message);
            }

            try {
                await sheetsService.spreadsheets.values.update({
                    spreadsheetId: config.sheetId,
                    range: 'J' + r.index,
                    valueInputOption: 'RAW',
                    resource: { values: [['Envoye']] }
                });
            } catch (_) {}
        }
        config.lastProcessedRow = rows.length;
        saveConfig(config);
    } catch (e) {
        if (e.code === 404) { console.log('Sheet introuvable, verifie l ID'); }
        else if (e.code === 403) { console.log('Acces refuse. Partage le sheet avec ' + getSaEmail()); }
        else { console.log('Erreur lecture sheet:', e.message); }
    }
}

function getSaEmail() {
    try { return JSON.parse(fs.readFileSync(SA_PATH, 'utf8')).client_email; }
    catch { return 'service account'; }
}

function startPolling() {
    stopPolling();
    sheetsService = initSheets();
    if (sheetsService && config.sheetId) {
        console.log('Surveillance du sheet demarree (toutes les 30s)');
        checkSheet();
        pollTimer = setInterval(checkSheet, POLL_INTERVAL);
    }
}

function stopPolling() {
    if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
}

function baseHtml(title, content) {
    return '<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
        + '<title>' + title + ' - SheetPulse</title>'
        + '<script src="https://cdn.tailwindcss.com"></script></head>'
        + '<body class="bg-gray-50 min-h-screen">'
        + '<nav class="bg-white shadow-sm border-b">'
        + '<div class="max-w-5xl mx-auto px-4 py-3 flex items-center justify-between">'
        + '<a href="/" class="text-xl font-bold text-gray-800">SheetPulse</a>'
        + '<div class="flex gap-4 text-sm">'
        + '<a href="/" class="text-gray-600 hover:text-gray-900' + (title === 'Dashboard' ? ' font-semibold text-blue-600' : '') + '">Dashboard</a>'
        + '<a href="/config" class="text-gray-600 hover:text-gray-900' + (title === 'Configuration' ? ' font-semibold text-blue-600' : '') + '">Configuration</a>'
        + '</div></div></nav>'
        + '<div class="max-w-5xl mx-auto px-4 py-8">' + content + '</div></body></html>';
}

app.get('/', function (req, res) {
    if (clientStatus === 'connected') {
        var saEmail = getSaEmail();
        var allGood = config.sheetId && config.groupId;
        res.send(baseHtml('Dashboard',
            (allGood
                ? '<div class="bg-white rounded-xl shadow-sm border-2 border-green-300 p-8 text-center">'
                + '<div class="w-20 h-20 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">'
                + '<svg class="w-10 h-10 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg></div>'
                + '<h1 class="text-2xl font-bold text-green-800 mb-2">✓ Configuration complete</h1>'
                + '<p class="text-green-700 mb-1">WhatsApp connecte et Google Sheet configure.</p>'
                + '<p class="text-green-600 text-sm mb-4">Le bot surveille les nouvelles commandes toutes les 30 secondes.</p>'
                + '<div class="bg-green-50 rounded-lg p-4 text-sm text-green-800 text-left max-w-md mx-auto space-y-1">'
                + '<p><span class="font-medium">Groupe :</span> <span class="font-mono">' + config.groupId + '</span></p>'
                + '<p><span class="font-medium">Sheet :</span> <span class="font-mono">' + config.sheetId + '</span></p></div>'
                + '<div class="mt-6"><a href="/config" class="px-6 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 font-medium">Modifier la configuration</a></div>'
                + '</div>'
                : '<div class="bg-white rounded-xl shadow-sm border p-8 text-center">'
                + '<div class="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">'
                + '<svg class="w-8 h-8 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"/></svg></div>'
                + '<h1 class="text-2xl font-bold text-gray-800 mb-2">WhatsApp connecte</h1>'
                + '<p class="text-gray-500 mb-4">WhatsApp est pret, configure le Google Sheet pour commencer.</p>'
                + '<div class="mt-2 flex justify-center gap-4">'
                + '<a href="/config" class="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">Configuration</a></div>'
                + '<div class="mt-6 p-4 bg-yellow-50 border border-yellow-200 rounded-lg text-sm text-yellow-800 text-left">'
                + '<p class="font-medium mb-1">Pour commencer :</p>'
                + '<ol class="list-decimal list-inside space-y-1">'
                + '<li>Va dans <strong>Configuration</strong></li>'
                + '<li>Entre ton ID de Google Sheet</li>'
                + '<li>Partage le sheet avec <strong>' + saEmail + '</strong></li>'
                + '<li>Le bot detectera automatiquement les nouvelles commandes</li></ol></div>'
                + '</div>')
        ));
    } else if (clientStatus === 'awaiting_scan' && pairingCodeData) {
        res.send(baseHtml('Dashboard',
            '<div class="bg-white rounded-xl shadow-sm border p-8 text-center">'
            + '<div class="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mx-auto mb-4">'
            + '<svg class="w-8 h-8 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 6v6m0 0v6m0-6h6m-6 0H6"/></svg></div>'
            + '<h1 class="text-2xl font-bold text-gray-800 mb-2">Code de couplage</h1>'
            + '<p class="text-gray-500 mb-4">WhatsApp > Appareils connectes > Connecter un appareil</p>'
            + '<div class="bg-gray-100 rounded-lg p-6 mb-4 inline-block">'
            + '<span class="text-3xl font-bold font-mono tracking-widest text-gray-800" id="pairingCodeDisplay">' + pairingCodeData + '</span></div>'
            + '<p class="text-sm text-gray-400">Saisis ce code dans WhatsApp pour connecter</p>'
            + '<p class="mt-4 text-sm text-gray-400">La page se met a jour automatiquement.</p></div>'
            + '<script>setInterval(function(){ fetch("/status").then(r=>r.json()).then(d=>{if(d.status==="connected")location.reload()}); }, 3000);</script>'
        ));
    } else if (qrCodeData) {
        res.send(baseHtml('Dashboard',
            '<div class="bg-white rounded-xl shadow-sm border p-8 text-center">'
            + '<h1 class="text-2xl font-bold text-gray-800 mb-2">Scannez le QR code</h1>'
            + '<p class="text-gray-500 mb-6">Ouvrez WhatsApp > Appareils connectes > Connecter un appareil</p>'
            + '<img id="qrImg" src="/qr-image" alt="QR Code" class="w-64 h-64 mx-auto rounded-lg border p-2"/>'
            + '<p class="mt-4 text-sm text-gray-400">Mise a jour automatique...</p>'
            + '</div>'
            + '<div class="bg-white rounded-xl shadow-sm border p-8 mt-6 text-center">'
            + '<h2 class="text-lg font-semibold text-gray-800 mb-2">Alternative : code de couplage</h2>'
            + '<p class="text-sm text-gray-500 mb-4">Si le QR ne marche pas, entre ton numero pour recevoir un code</p>'
            + '<form id="pairingForm" class="flex gap-2 max-w-md mx-auto">'
            + '<input type="tel" name="phone" placeholder="229XXXXXXXX" required'
            + ' class="flex-1 px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 font-mono text-sm">'
            + '<button type="submit" class="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm font-medium">Envoyer</button></form>'
            + '<p id="pairingError" class="mt-2 text-sm text-red-500 hidden"></p>'
            + '</div>'
            + '<script>'
            + 'setInterval(function(){ document.getElementById("qrImg").src = "/qr-image?" + Date.now(); }, 2000);'
            + 'setInterval(function(){ fetch("/status").then(r=>r.json()).then(function(d){if(d.status==="connected")location.reload()}); }, 3000);'
            + 'document.getElementById("pairingForm").addEventListener("submit", async function(e){'
            + '  e.preventDefault();'
            + '  var phone = e.target.phone.value, btn = e.target.querySelector("button");'
            + '  btn.disabled = true; btn.textContent = "Envoi...";'
            + '  try {'
            + '    var r = await fetch("/api/pairing", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({phone:phone}) });'
            + '    if (r.ok) { location.reload(); }'
            + '    else { var err = await r.json(); document.getElementById("pairingError").textContent = err.error; document.getElementById("pairingError").classList.remove("hidden"); }'
            + '  } catch(e) { document.getElementById("pairingError").textContent = "Erreur reseau"; document.getElementById("pairingError").classList.remove("hidden"); }'
            + '  btn.disabled = false; btn.textContent = "Envoyer";'
            + '});'
            + '</script>'
        ));
    } else {
        res.send(baseHtml('Dashboard',
            '<div class="bg-white rounded-xl shadow-sm border p-8 text-center">'
            + '<div class="w-16 h-16 bg-yellow-100 rounded-full flex items-center justify-center mx-auto mb-4">'
            + '<svg class="w-8 h-8 text-yellow-600 animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg></div>'
            + '<h1 class="text-2xl font-bold text-gray-800 mb-2">Initialisation en cours...</h1>'
            + '<p class="text-gray-500">Preparation de WhatsApp, veuillez patienter.</p>'
            + '<p class="mt-4 text-sm text-gray-400">La page se recharge automatiquement.</p></div>'
            + '<script>setTimeout(function(){ location.reload(); }, 3000);</script>'
        ));
    }
});

app.get('/config', function (req, res) {
    var selectedGroup = config.groupId || '';
    var saEmail = getSaEmail();

    res.send(baseHtml('Configuration',
        '<h1 class="text-2xl font-bold text-gray-800 mb-6">Configuration</h1>'
        + '<form id="configForm" class="space-y-6">'

        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">1. Groupe WhatsApp</h2>'
        + '<input type="text" id="groupSearch" placeholder="Rechercher un groupe..."'
        + ' class="w-full px-4 py-2 border rounded-lg mb-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"'
        + (clientStatus !== 'connected' ? ' disabled' : '') + '>'
        + '<select name="groupId" id="groupSelect" class="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white" size="6"'
        + (clientStatus !== 'connected' ? ' disabled' : '') + '>'
        + (clientStatus === 'connected'
            ? '<option value="">Chargement des groupes...</option>'
            : '<option value="">Connecte WhatsApp d abord</option>')
        + '</select>'
        + '<p id="groupStatus" class="mt-1 text-sm text-gray-400">'
        + (clientStatus === 'connected' ? 'Chargement...' : 'WhatsApp non connecte')
        + '</p></div>'

                    + '<div class="bg-white rounded-xl shadow-sm border p-6">'
                    + '<h2 class="text-lg font-semibold text-gray-800 mb-2">2. Google Sheet</h2>'
                    + '<label class="block text-sm font-medium text-gray-700 mb-1">Lien ou ID du Google Sheet</label>'
                    + '<input type="text" name="sheetId" value="' + (config.sheetId || '').replace(/"/g, '&quot;') + '"'
                    + ' placeholder="https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit"'
                    + ' class="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 font-mono text-sm">'
                    + '<p class="mt-1 text-sm text-gray-400">Colle le lien complet ou seulement l ID</p>'
                    + '</div>'

        + '<div class="bg-white rounded-xl shadow-sm border p-6 bg-blue-50 border-blue-200">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">3. Partager l acces</h2>'
        + '<p class="text-sm text-gray-600 mb-3">Ajoute cet email comme <strong>editeur</strong> dans ton Google Sheet :</p>'
        + '<div class="flex gap-2">'
        + '<input type="text" value="' + saEmail + '" readonly class="flex-1 px-4 py-2 bg-white border rounded-lg font-mono text-sm">'
        + '<button onclick="navigator.clipboard.writeText(\'' + saEmail + '\')" class="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm whitespace-nowrap">Copier</button></div>'
        + '<p class="mt-2 text-sm text-gray-500">Va dans ton Sheet > Partager > ajoute cet email en editeur.</p>'
        + '</div>'

        + '<button type="submit" class="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-medium">Enregistrer</button></form>'
        + '<div id="toast" class="hidden fixed bottom-4 right-4 px-6 py-3 rounded-lg shadow-lg z-50"></div>'

                    + '<script>'
                    + (clientStatus === 'connected' ? (
                        'fetch("/api/groups").then(function(r){'
                        + '  if (!r.ok) return r.json().then(function(e){ throw new Error(e.error); });'
                        + '  return r.json();'
                        + '}).then(function(groups){'
                        + '  var sel = document.getElementById("groupSelect"), search = document.getElementById("groupSearch");'
                        + '  var selected = "' + selectedGroup.replace(/"/g, '\\"') + '";'
                        + '  function render(f){'
                        + '    sel.options.length = 0;'
                        + '    var q = f ? f.toLowerCase() : "", c = 0;'
                        + '    groups.forEach(function(g){'
                        + '      if (q && g.name.indexOf(q) === -1 && g.id.indexOf(q) === -1) return;'
                        + '      var opt = new Option(g.name + " (" + g.id + ")", g.id);'
                        + '      if (g.id === selected) opt.selected = true;'
                        + '      sel.options[sel.options.length] = opt;'
                        + '      c++;'
                        + '    });'
                        + '    document.getElementById("groupStatus").textContent = c + " / " + groups.length + " groupe(s)";'
                        + '  }'
                        + '  search.addEventListener("input", function(){ render(this.value); });'
                        + '  render("");'
                        + '  sel.disabled = false; search.disabled = false;'
                        + '}).catch(function(e){ document.getElementById("groupStatus").textContent = "Erreur: " + e.message; });'
                    ) : '')
                    + 'function extractSheetId(v){'
                    + '  var m = v.match(/\\/spreadsheets\\/d\\/([a-zA-Z0-9_-]+)/);'
                    + '  return m ? m[1] : v;'
                    + '}'
                    + 'document.querySelector("[name=sheetId]").addEventListener("change", function(){'
                    + '  this.value = extractSheetId(this.value);'
                    + '});'
                    + 'document.getElementById("configForm").addEventListener("submit", async function(e){'
                    + '  e.preventDefault();'
                    + '  var sheetVal = extractSheetId(document.querySelector("[name=sheetId]").value);'
                    + '  var data = { groupId: document.getElementById("groupSelect").value, sheetId: sheetVal };'
                    + '  var r = await fetch("/api/config", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(data) });'
                    + '  if (r.ok) {'
                    + '    var t = document.getElementById("toast");'
                    + '    t.className = "fixed bottom-4 right-4 px-6 py-3 bg-green-600 text-white rounded-lg shadow-lg z-50";'
                    + '    t.innerHTML = "✓ Configuration enregistree ! <a href=\'/\' class=\'underline font-medium\'>Voir le dashboard</a>";'
                    + '    setTimeout(function(){ t.classList.add("hidden"); }, 6000);'
                    + '  }'
                    + '});'
        + '</script>'
    ));
});

function extractSheetId(v) {
    if (!v) return v;
    var m = v.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    return m ? m[1] : v;
}

app.post('/api/config', function (req, res) {
    var d = req.body;
    if (d.groupId !== undefined) config.groupId = d.groupId;
    if (d.sheetId !== undefined) config.sheetId = extractSheetId(d.sheetId);
    if (d.publicUrl !== undefined) config.publicUrl = d.publicUrl;
    if (d.groupId !== undefined || d.sheetId !== undefined || d.publicUrl !== undefined) saveConfig(config);
    if (d.sheetId !== undefined) startPolling();
    res.json({ success: true });
});

app.get('/api/config', function (req, res) {
    res.json(config);
});

app.get('/api/groups', async function (req, res) {
    if (clientStatus !== 'connected') return res.status(503).json({ error: 'WhatsApp non connecte' });
    try {
        var groups = await client.pupPage.evaluate(function() {
            var chats = window.require('WAWebCollections').Chat.getModelsArray();
            return chats.filter(function(c) { return c.groupMetadata; }).map(function(c) {
                return { id: c.id._serialized, name: c.formattedTitle || c.name || c.id._serialized };
            });
        });
        res.json(groups);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/pairing', async function (req, res) {
    var phone = req.body.phone;
    if (!phone) return res.status(400).json({ error: 'Numero requis' });
    if (clientStatus === 'connected') return res.status(400).json({ error: 'Deja connecte' });
    try { await client.destroy(); } catch (_) {}
    qrCodeData = null; pairingCodeData = null; clientStatus = 'initializing';
    client = new Client({
        authStrategy: new LocalAuth(),
        puppeteer: { headless: true, protocolTimeout: 120000, args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-accelerated-2d-canvas', '--no-first-run', '--no-zygote', '--disable-gpu', '--single-process'] },
        pairWithPhoneNumber: { phoneNumber: phone }
    });
    attachClientEvents();
    client.initialize();
    res.json({ success: true, message: 'Code de couplage envoye' });
});

app.get('/qr-image', function (req, res) {
    if (!qrCodeData) return res.status(404).send('Aucun QR');
    QRCode.toDataURL(qrCodeData, function (err, url) {
        if (err) return res.status(500).send('Erreur');
        var b = url.replace(/^data:image\/png;base64,/, '');
        res.writeHead(200, { 'Content-Type': 'image/png', 'Cache-Control': 'no-cache' });
        res.end(Buffer.from(b, 'base64'));
    });
});

app.get('/debug', function (req, res) {
    res.json({ clientStatus, hasQr: !!qrCodeData, qrLen: qrCodeData ? qrCodeData.length : 0, hasPairing: !!pairingCodeData, sheetConfigured: !!config.sheetId, config });
});

app.get('/status', function (req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
    res.json({ status: clientStatus, hasQr: !!qrCodeData, hasPairing: !!pairingCodeData, sheetConfigured: !!config.sheetId });
});

var PORT = process.env.PORT || 3000;

function attachClientEvents() {
    client.on('qr', function (qr) {
        qrCodeData = qr; pairingCodeData = null; clientStatus = 'awaiting_scan';
        console.log('--- SCANNEZ LE QR CODE CI-DESSOUS AVEC WHATSAPP ---');
        qrcode.generate(qr, { small: true });
        console.log('Ou ouvrez http://localhost:' + PORT + ' pour le voir sur une page web');
    });
        client.on('authenticated', function() { console.log('Auth OK'); });
    client.on('auth_failure', function(m) { console.log('Auth FAIL:', m); });
    client.on('code', function (code) {
        pairingCodeData = code; qrCodeData = null; clientStatus = 'awaiting_scan';
        console.log('--- CODE DE COUPLAGE ---');
        console.log('Code:', code);
        console.log('Saisis ce code dans WhatsApp > Appareils connectes');
    });
    client.on('ready', async function () {
        qrCodeData = null; pairingCodeData = null; clientStatus = 'connected';
        console.log('WhatsApp connecte !');
        try {
            var groups = await client.pupPage.evaluate(function() {
                var chats = window.require('WAWebCollections').Chat.getModelsArray();
                return chats.filter(function(c) { return c.groupMetadata; }).map(function(c) {
                    return { id: c.id._serialized, name: c.formattedTitle || c.name || c.id._serialized };
                });
            });
            if (groups.length > 0) {
                console.log('--- LISTE DES GROUPES ---');
                groups.forEach(function(g) { console.log('Nom: ' + g.name + ' | ID: ' + g.id); });
            }
        } catch (_) {}
        startPolling();
    });
    client.on('disconnected', function (reason) {
        clientStatus = 'disconnected';
        console.log('WhatsApp deconnecte:', reason);
        stopPolling();
    });
}

var client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: { headless: true, protocolTimeout: 120000, args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-accelerated-2d-canvas', '--no-first-run', '--no-zygote', '--disable-gpu', '--single-process'] }
});

attachClientEvents();
client.initialize();

app.listen(PORT, function () {
    console.log('SheetPulse demarre sur http://localhost:' + PORT);
});

process.on('uncaughtException', function (err) {
    console.error('Erreur:', err.message);
});
