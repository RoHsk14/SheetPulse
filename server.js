const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const QRCode = require('qrcode');
const express = require('express');
const fs = require('fs');

const CONFIG_PATH = './config.json';

function loadConfig() {
    try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')); }
    catch { return { groupId: '', columns: {} }; }
}
function saveConfig(config) {
    fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2));
}

const app = express();
app.use(express.json());

let config = loadConfig();
let qrCodeData = null;
let pairingCodeData = null;
let clientStatus = 'initializing';

function generateAppsScript(baseUrl, cols) {
    return [
        'function onEdit(e) {',
        '  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();',
        '  var lastRow = sheet.getLastRow();',
        '  var id = sheet.getRange(lastRow, ' + (cols.id || 1) + ').getValue();',
        '  var client = sheet.getRange(lastRow, ' + (cols.client || 2) + ').getValue();',
        '  var produits = sheet.getRange(lastRow, ' + (cols.produits || 3) + ').getValue();',
        '  var total = sheet.getRange(lastRow, ' + (cols.total || 4) + ').getValue();',
        '  if (!id) return;',
        '  var msg = "📦 *NOUVELLE COMMANDE* 📦\\n\\n"',
        '    + "• ID : " + id + "\\n"',
        '    + "• Client : " + client + "\\n"',
        '    + "• Commande : " + produits + "\\n"',
        '    + "• Total : " + total + " FCFA\\n\\n"',
        '    + "⚡ À traiter rapidement !";',
        '  var opts = {',
        '    method: "post",',
        '    contentType: "application/json",',
        '    payload: JSON.stringify({ message: msg }),',
        '    muteHttpExceptions: true',
        '  };',
        '  UrlFetchApp.fetch("' + baseUrl + '/webhook", opts);',
        '}'
    ].join('\n');
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
        + '<a href="/integration" class="text-gray-600 hover:text-gray-900' + (title === 'Integration' ? ' font-semibold text-blue-600' : '') + '">Integration</a>'
        + '</div></div></nav>'
        + '<div class="max-w-5xl mx-auto px-4 py-8">' + content + '</div></body></html>';
}

app.get('/', function (req, res) {
    if (clientStatus === 'connected') {
        res.send(baseHtml('Dashboard',
            '<div class="bg-white rounded-xl shadow-sm border p-8 text-center">'
            + '<div class="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">'
            + '<svg class="w-8 h-8 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"/></svg></div>'
            + '<h1 class="text-2xl font-bold text-gray-800 mb-2">WhatsApp connecte</h1>'
            + '<p class="text-gray-500 mb-6">Le bot est pret a recevoir des commandes.</p>'
            + '<div class="flex justify-center gap-4">'
            + '<a href="/config" class="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">Configurer</a>'
            + '<a href="/integration" class="px-6 py-2 border border-gray-300 rounded-lg hover:bg-gray-50">Voir l integration</a></div>'
            + (config.groupId ? '<p class="mt-6 text-sm text-gray-400">Groupe cible : <span class="font-mono text-gray-600">' + config.groupId + '</span></p>' : '')
            + '</div>'
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
            + '</div>'
            + '<script>setInterval(function(){ location.reload(); }, 3000);</script>'
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
            + 'document.getElementById("pairingForm").addEventListener("submit", async function(e){'
            + '  e.preventDefault();'
            + '  var phone = e.target.phone.value;'
            + '  var btn = e.target.querySelector("button");'
            + '  btn.disabled = true; btn.textContent = "Envoi...";'
            + '  try {'
            + '    var r = await fetch("/api/pairing", { method: "POST", headers: {"Content-Type":"application/json"}, body: JSON.stringify({phone: phone}) });'
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
    var c = config.columns || {};
    var fields = [
        { key: 'id', label: 'ID commande', ph: '1' },
        { key: 'client', label: 'Nom client', ph: '2' },
        { key: 'produits', label: 'Produits commandes', ph: '3' },
        { key: 'total', label: 'Total (FCFA)', ph: '4' },
        { key: 'statut', label: 'Statut', ph: '5' }
    ];
    var inputs = fields.map(function (f) {
        return '<div><label class="block text-sm font-medium text-gray-700 mb-1">' + f.label + '</label>'
            + '<input type="number" name="col_' + f.key + '" value="' + (c[f.key] || '') + '" placeholder="' + f.ph + '"'
            + ' class="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"></div>';
    }).join('');

    res.send(baseHtml('Configuration',
        '<h1 class="text-2xl font-bold text-gray-800 mb-6">Configuration</h1>'
        + '<form id="configForm" class="space-y-6">'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-4">Groupe WhatsApp</h2>'
        + '<label class="block text-sm font-medium text-gray-700 mb-1">ID du groupe</label>'
        + '<input type="text" name="groupId" value="' + (config.groupId || '') + '" placeholder="120363430608871862@g.us"'
        + ' class="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 font-mono text-sm">'
        + '<p class="mt-1 text-sm text-gray-400">L ID qui termine par <span class="font-mono">@g.us</span></p></div>'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-4">Mapping des colonnes</h2>'
        + '<p class="text-sm text-gray-500 mb-4">Indique le numero de chaque colonne (A=1, B=2, C=3...).</p>'
        + '<div class="grid grid-cols-2 gap-4">' + inputs + '</div></div>'
        + '<button type="submit" class="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-medium">Enregistrer</button></form>'
        + '<div id="toast" class="hidden fixed bottom-4 right-4 px-6 py-3 bg-green-600 text-white rounded-lg shadow-lg"></div>'
        + '<script>'
        + 'document.getElementById("configForm").addEventListener("submit", async function(e){'
        + '  e.preventDefault();'
        + '  var fd = new FormData(e.target);'
        + '  var data = { groupId: fd.get("groupId"), columns: {} };'
        + '  for (var pair of fd.entries()) {'
        + '    if (pair[0].startsWith("col_")) data.columns[pair[0].replace("col_", "")] = parseInt(pair[1]);'
        + '  }'
        + '  var r = await fetch("/api/config", { method: "POST", headers: {"Content-Type":"application/json"}, body: JSON.stringify(data) });'
        + '  if (r.ok) {'
        + '    var t = document.getElementById("toast");'
        + '    t.textContent = "Configuration enregistree ✓";'
        + '    t.classList.remove("hidden");'
        + '    setTimeout(function(){ t.classList.add("hidden"); }, 3000);'
        + '  }'
        + '});'
        + '</script>'
    ));
});

app.get('/integration', function (req, res) {
    var baseUrl = 'http://localhost:' + (process.env.PORT || 3000);
    var scriptCode = generateAppsScript(baseUrl, config.columns);

    res.send(baseHtml('Integration',
        '<h1 class="text-2xl font-bold text-gray-800 mb-6">Integration Google Sheets</h1>'
        + '<div class="space-y-6">'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">URL du webhook</h2>'
        + '<div class="flex gap-2">'
        + '<input type="text" value="' + baseUrl + '/webhook" readonly class="flex-1 px-4 py-2 bg-gray-50 border rounded-lg font-mono text-sm">'
        + '<button onclick="navigator.clipboard.writeText(\'' + baseUrl + '/webhook\')" class="px-4 py-2 bg-gray-100 border rounded-lg hover:bg-gray-200 text-sm">Copier</button></div></div>'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">Apps Script a copier</h2>'
        + '<p class="text-sm text-gray-500 mb-4">Colle ce code dans <span class="font-mono bg-gray-100 px-2 py-0.5 rounded">Extensions > Apps Script</span>.</p>'
        + '<div class="relative">'
        + '<textarea id="scriptCode" rows="14" readonly class="w-full px-4 py-3 bg-gray-50 border rounded-lg font-mono text-sm">' + scriptCode + '</textarea>'
        + '<button onclick="navigator.clipboard.writeText(document.getElementById(\'scriptCode\').value)" class="absolute top-2 right-2 px-3 py-1.5 bg-blue-600 text-white rounded text-xs hover:bg-blue-700">Copier</button></div></div>'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">Instructions</h2>'
        + '<ol class="list-decimal list-inside space-y-2 text-gray-600">'
        + '<li>Ouvre ton Google Sheet de commandes</li>'
        + '<li>Va dans <strong>Extensions > Apps Script</strong></li>'
        + '<li>Supprime le code par defaut et colle le code ci-dessus</li>'
        + '<li>Clique sur <strong>Enregistrer</strong> (icone disquette)</li>'
        + '<li>Ajoute une ligne dans ton Sheet → le message sera envoye automatiquement</li></ol></div></div>'
    ));
});

app.post('/api/config', function (req, res) {
    var data = req.body;
    if (data.groupId) config.groupId = data.groupId;
    if (data.columns) config.columns = data.columns;
    saveConfig(config);
    res.json({ success: true });
});

app.get('/api/config', function (req, res) {
    res.json(config);
});

app.post('/api/pairing', async function (req, res) {
    var phone = req.body.phone;
    if (!phone) return res.status(400).json({ error: 'Numero requis' });
    if (clientStatus === 'connected') return res.status(400).json({ error: 'Deja connecte' });
    try {
        console.log('Demande de code de couplage pour:', phone);
        await client.destroy();
    } catch (_) {}
    qrCodeData = null;
    pairingCodeData = null;
    clientStatus = 'initializing';
    client = new Client({
        authStrategy: new LocalAuth(),
        puppeteer: {
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-accelerated-2d-canvas', '--no-first-run', '--no-zygote', '--disable-gpu', '--single-process']
        },
        pairWithPhoneNumber: { phoneNumber: phone }
    });
    attachClientEvents();
    client.initialize();
    res.json({ success: true, message: 'Code de couplage envoye a WhatsApp', code: 'En attente...' });
});

app.get('/qr-image', function (req, res) {
    if (!qrCodeData) return res.status(404).send('Aucun QR');
    QRCode.toDataURL(qrCodeData, function (err, url) {
        if (err) return res.status(500).send('Erreur');
        var base64 = url.replace(/^data:image\/png;base64,/, '');
        var img = Buffer.from(base64, 'base64');
        res.writeHead(200, { 'Content-Type': 'image/png', 'Cache-Control': 'no-cache' });
        res.end(img);
    });
});

app.get('/status', function (req, res) {
    res.json({ status: clientStatus, hasQr: !!qrCodeData, hasPairing: !!pairingCodeData });
});

var PORT = process.env.PORT || 3000;

function attachClientEvents() {
    client.on('qr', function (qr) {
        qrCodeData = qr;
        pairingCodeData = null;
        clientStatus = 'awaiting_scan';
        console.log('--- SCANNEZ LE QR CODE CI-DESSOUS AVEC WHATSAPP ---');
        qrcode.generate(qr, { small: true });
        console.log('Ou ouvrez http://localhost:' + PORT + ' pour le voir sur une page web');
    });

    client.on('code', function (code) {
        pairingCodeData = code;
        qrCodeData = null;
        clientStatus = 'awaiting_scan';
        console.log('--- CODE DE COUPLAGE ---');
        console.log('Code:', code);
        console.log('Saisis ce code dans WhatsApp > Appareils connectes');
    });

    client.on('ready', function () {
        qrCodeData = null;
        pairingCodeData = null;
        clientStatus = 'connected';
        console.log('WhatsApp connecte !');
    });

    client.on('disconnected', function (reason) {
        clientStatus = 'disconnected';
        console.log('WhatsApp deconnecte:', reason);
    });
}

var client = new Client({
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

attachClientEvents();

app.post('/webhook', async function (req, res) {
    var message = req.body.message;
    var groupId = config.groupId;
    if (!groupId) return res.status(400).json({ error: 'Groupe non configure' });
    if (clientStatus !== 'connected') return res.status(503).json({ error: 'WhatsApp non connecte' });
    try {
        await client.sendMessage(groupId, message);
        console.log('Message envoye au groupe');
        return res.json({ success: true });
    } catch (err) {
        console.error('Erreur envoi:', err);
        return res.status(500).json({ error: 'Echec de l envoi' });
    }
});

app.listen(PORT, function () {
    console.log('SheetPulse demarre sur http://localhost:' + PORT);
    console.log('Lancement de WhatsApp...');
    client.initialize();
});

process.on('uncaughtException', function (err) {
    console.error('Erreur:', err.message);
});
