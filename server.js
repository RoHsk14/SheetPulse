process.on('uncaughtException', function (err) {
    console.error('UNCAUGHT:', err.message, err.stack);
});
process.on('unhandledRejection', function (reason) {
    console.error('UNHANDLED:', reason);
});

require('dotenv').config();

const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const QRCode = require('qrcode');
const express = require('express');
const fs = require('fs');
const { google } = require('googleapis');
const { createClient } = require('@supabase/supabase-js');
const WebSocket = require('ws');
const path = require('path');
const os = require('os');

const CONFIG_PATH = './config.json';
const SA_PATH = './service-account.json';
const POLL_INTERVAL = 30000;

function loadConfig() {
    try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')); }
    catch { return { groupId: '', sheetId: '', publicUrl: '', lastProcessedRow: 0 }; }
}
function saveConfig(config) {
    fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2));
}

const app = express();
app.use(express.json());
app.use(function (req, res, next) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type,Authorization');
    if (req.method === 'OPTIONS') return res.sendStatus(200);
    next();
});

let config = loadConfig();
const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
const supabase = supabaseUrl && supabaseKey ? createClient(supabaseUrl, supabaseKey, { realtime: { transport: WebSocket } }) : null;
const groupId = process.env.GROUP_ID || null;

let pollTimer = null;
let qrCodeData = null;
let pairingCodeData = null;
let clientStatus = 'initializing';
let sheetsService = null;

function generateAppsScript(baseUrl) {
    return [
        'function sheetPulseOnEdit(e) {',
        '  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();',
        '  var lastRow = sheet.getLastRow();',
        '  var date = sheet.getRange(lastRow, 1).getValue();',
        '  var client = sheet.getRange(lastRow, 2).getValue();',
        '  var tel = sheet.getRange(lastRow, 3).getValue();',
        '  var adresse = sheet.getRange(lastRow, 4).getValue();',
        '  var quartier = sheet.getRange(lastRow, 5).getValue();',
        '  var produit = sheet.getRange(lastRow, 6).getValue();',
        '  var qte = sheet.getRange(lastRow, 7).getValue();',
        '  var total = sheet.getRange(lastRow, 8).getValue();',
        '  var devise = sheet.getRange(lastRow, 9).getValue() || "FCFA";',
        '  var statut = sheet.getRange(lastRow, 10).getValue() || "En attente";',
        '  if (!client) return;',
        '  var msg = "📦 *NOUVELLE COMMANDE* 📦\\n\\n"',
        '    + "*Client :* " + client + "\\n"',
        '    + "*Tel :* " + tel + "\\n"',
        '    + "*Adresse :* " + adresse + " (" + quartier + ")\\n"',
        '    + "*Produit :* " + produit + " x" + qte + "\\n"',
        '    + "*Total :* " + total + " " + devise + "\\n"',
        '    + "*Statut :* " + statut + "\\n\\n"',
        '    + "⚡ _Commande du " + date + "_";',
        '  var opts = { method: "post", contentType: "application/json",',
        '    payload: JSON.stringify({ message: msg }), muteHttpExceptions: true };',
        '  UrlFetchApp.fetch("' + baseUrl + '/webhook", opts);',
        '}',
        '',
        'function sheetPulseSetup() {',
        '  var sheet = SpreadsheetApp.getActiveSpreadsheet();',
        '  var triggers = ScriptApp.getProjectTriggers();',
        '  for (var i = 0; i < triggers.length; i++) {',
        '    if (triggers[i].getHandlerFunction() === "sheetPulseOnEdit") {',
        '      ScriptApp.deleteTrigger(triggers[i]);',
        '    }',
        '  }',
        '  ScriptApp.newTrigger("sheetPulseOnEdit")',
        '    .forSpreadsheet(sheet)',
        '    .onEdit()',
        '    .create();',
        '}'
    ].join('\n');
}

function extractSheetId(v) {
    if (!v) return v;
    var m = v.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    return m ? m[1] : v;
}

function initSheets() {
    try {
        const { google } = require('googleapis');
        var auth;
        var envKey = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
        if (envKey) {
            var creds = JSON.parse(envKey);
            auth = new google.auth.JWT(
                creds.client_email,
                null,
                creds.private_key,
                ['https://www.googleapis.com/auth/spreadsheets']
            );
        } else {
            if (!fs.existsSync(SA_PATH)) {
                console.log('service-account.json manquant');
                return null;
            }
            auth = new google.auth.GoogleAuth({
                keyFile: SA_PATH,
                scopes: ['https://www.googleapis.com/auth/spreadsheets']
            });
        }
        return google.sheets({ version: 'v4', auth: auth });
    } catch (e) {
        console.log('Erreur init Sheets:', e.message);
        return null;
    }
}

function getSaEmail() {
    try {
        var envKey = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
        if (envKey) return JSON.parse(envKey).client_email;
        return JSON.parse(fs.readFileSync(SA_PATH, 'utf8')).client_email;
    } catch { return 'service account'; }
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
    }
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
            + '<h1 class="text-2xl font-bold text-gray-800 mb-2">WhatsApp connecté</h1>'
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
            + '<p class="text-sm text-gray-500 mb-4">Si le QR ne marche pas, entrez votre numero(Sans le + ou  le 00) pour recevoir un code</p>'
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
            + '<p class="text-gray-500">Preparation de WhatsApp, veuillez patienter quelques instants.</p>'
            + '<p class="mt-4 text-sm text-gray-400">La page se recharge automatiquement.</p></div>'
            + '<script>setTimeout(function(){ location.reload(); }, 3000);</script>'
        ));
    }
});

app.get('/config', function (req, res) {
    var selected = config.groupId || '';
    var saEmail = getSaEmail();
    res.send(baseHtml('Configuration',
        '<h1 class="text-2xl font-bold text-gray-800 mb-6">Configuration</h1>'
        + '<form id="configForm" class="space-y-6">'

        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">Groupe WhatsApp</h2>'
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
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">Google Sheet</h2>'
        + '<label class="block text-sm font-medium text-gray-700 mb-1">ID du Google Sheet</label>'
        + '<input type="text" name="sheetId" id="sheetIdInput" value="' + (config.sheetId || '').replace(/"/g, '&quot;') + '"'
        + ' placeholder="https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit"'
        + ' class="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 font-mono text-sm">'
        + '<p class="mt-1 text-sm text-gray-400">Colle le lien complet ou seulement l ID</p>'
        + '</div>'

        + '<div class="bg-white rounded-xl shadow-sm border p-6 bg-blue-50 border-blue-200">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">Partager l acces</h2>'
        + '<p class="text-sm text-gray-600 mb-3">Ajoute cet email comme <strong>editeur</strong> dans ton Google Sheet :</p>'
        + '<div class="flex gap-2">'
        + '<input type="text" id="saEmail" value="' + saEmail + '" readonly class="flex-1 px-4 py-2 bg-white border rounded-lg font-mono text-sm">'
        + '<button onclick="navigator.clipboard.writeText(document.getElementById(\'saEmail\').value)" class="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm whitespace-nowrap">Copier</button></div>'
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
            + '  var selected = "' + selected.replace(/"/g, '\\"') + '";'
            + '  function render(f){'
            + '    sel.options.length = 0;'
            + '    var q = f ? f.toLowerCase() : "", c = 0;'
            + '    groups.forEach(function(g){'
            + '      if (q && g.name.toLowerCase().indexOf(q) === -1 && g.id.toLowerCase().indexOf(q) === -1) return;'
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
        + 'document.getElementById("sheetIdInput").addEventListener("change", function(){'
        + '  this.value = extractSheetId(this.value);'
        + '});'
        + 'document.getElementById("configForm").addEventListener("submit", async function(e){'
        + '  e.preventDefault();'
        + '  var sheetVal = extractSheetId(document.getElementById("sheetIdInput").value);'
        + '  var data = { groupId: document.getElementById("groupSelect").value, sheetId: sheetVal };'
        + '  var r = await fetch("/api/config", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(data) });'
        + '  if (r.ok) {'
        + '    var t = document.getElementById("toast");'
        + '    t.className = "fixed bottom-4 right-4 px-6 py-3 bg-green-600 text-white rounded-lg shadow-lg z-50";'
        + '    t.innerHTML = "Configuration enregistree ! <a href=\'/\' class=\'underline font-medium\'>Voir le dashboard</a>";'
        + '    setTimeout(function(){ t.classList.add("hidden"); }, 6000);'
        + '  }'
        + '});'
        + '</script>'
    ));
});

app.get('/integration', function (req, res) {
    var baseUrl = process.env.PUBLIC_URL || config.publicUrl || 'http://localhost:' + (process.env.PORT || 3000);
    var scriptCode = generateAppsScript(baseUrl);

    res.send(baseHtml('Integration',
        '<h1 class="text-2xl font-bold text-gray-800 mb-6">Integration Google Sheets</h1>'
        + '<div class="space-y-6">'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">URL publique du serveur</h2>'
        + '<p class="text-sm text-gray-500 mb-3">Utilise ton URL ngrok ou Render pour que Google Sheets puisse joindre le serveur.</p>'
        + '<form id="urlForm" class="flex gap-2">'
        + '<input type="url" name="publicUrl" value="' + (config.publicUrl || 'https://pacific-shy-primary.ngrok-free.dev') + '" placeholder="https://ton-url.ngrok-free.dev"'
        + ' class="flex-1 px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 font-mono text-sm">'
        + '<button type="submit" class="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm font-medium">OK</button></form>'
        + '<p id="urlSuccess" class="mt-1 text-sm text-green-600 hidden">URL enregistree ✓</p>'
        + '</div>'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">URL du webhook</h2>'
        + '<div class="flex gap-2">'
        + '<input type="text" id="webhookUrl" value="' + baseUrl + '/webhook" readonly class="flex-1 px-4 py-2 bg-gray-50 border rounded-lg font-mono text-sm">'
        + '<button onclick="navigator.clipboard.writeText(document.getElementById(\'webhookUrl\').value)" class="px-4 py-2 bg-gray-100 border rounded-lg hover:bg-gray-200 text-sm">Copier</button></div></div>'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">Structure du Google Sheet</h2>'
        + '<p class="text-sm text-gray-500 mb-4">Structure attendue de ton Google Sheet :</p>'
        + '<div class="overflow-x-auto"><table class="w-full text-sm border-collapse">'
        + '<thead><tr class="bg-gray-100">'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">Colonne</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">A</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">B</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">C</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">D</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">E</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">F</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">G</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">H</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">I</th>'
        + '<th class="px-4 py-2 border text-left font-medium text-gray-700">J</th>'
        + '</tr></thead><tbody>'
        + '<tr><td class="px-4 py-2 border font-medium text-gray-600">Contenu</td>'
        + '<td class="px-4 py-2 border">Date</td>'
        + '<td class="px-4 py-2 border">Nom client</td>'
        + '<td class="px-4 py-2 border">Telephone</td>'
        + '<td class="px-4 py-2 border">Adresse</td>'
        + '<td class="px-4 py-2 border">Quartier</td>'
        + '<td class="px-4 py-2 border">Produit</td>'
        + '<td class="px-4 py-2 border">Quantite</td>'
        + '<td class="px-4 py-2 border">Total</td>'
        + '<td class="px-4 py-2 border">Devise</td>'
        + '<td class="px-4 py-2 border">Statut</td></tr>'
        + '</tbody></table></div>'
        + '<p class="mt-2 text-sm text-gray-400">Ligne 1 = en-tetes. Les donnees commencent a la ligne 2.</p>'
        + '</div>'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">Apps Script a copier</h2>'
        + '<p class="text-sm text-gray-500 mb-4">Colle ce code dans <span class="font-mono bg-gray-100 px-2 py-0.5 rounded">Extensions > Apps Script</span> sur ton Sheet.</p>'
        + '<div class="relative">'
        + '<textarea id="scriptCode" rows="22" readonly class="w-full px-4 py-3 bg-gray-50 border rounded-lg font-mono text-sm">' + scriptCode + '</textarea>'
        + '<button onclick="navigator.clipboard.writeText(document.getElementById(\'scriptCode\').value)" class="absolute top-2 right-2 px-3 py-1.5 bg-blue-600 text-white rounded text-xs hover:bg-blue-700">Copier</button></div></div>'
        + '<div class="bg-white rounded-xl shadow-sm border p-6">'
        + '<h2 class="text-lg font-semibold text-gray-800 mb-2">Instructions</h2>'
        + '<ol class="list-decimal list-inside space-y-2 text-gray-600">'
        + '<li>Ouvre ton Google Sheet de commandes</li>'
        + '<li>Va dans <strong>Extensions > Apps Script</strong></li>'
        + '<li>Supprime le code par defaut et colle le code ci-dessus</li>'
        + '<li>Clique sur <strong>Enregistrer</strong> (icone disquette)</li>'
        + '<li>Dans la liste des fonctions, choisis <strong>sheetPulseSetup</strong> puis clique <strong>Run ▶️</strong></li>'
        + '<li>Autorise les permissions demandees</li>'
        + '<li>Une fois le setup fait, ajoute une ligne dans le Sheet → le message sera envoye automatiquement</li></ol></div></div>'
        + '<script>'
        + 'document.getElementById("urlForm").addEventListener("submit", async function(e){'
        + '  e.preventDefault();'
        + '  var url = e.target.publicUrl.value;'
        + '  var r = await fetch("/api/config", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({publicUrl: url}) });'
        + '  if (r.ok) {'
        + '    document.getElementById("urlSuccess").classList.remove("hidden");'
        + '    document.getElementById("webhookUrl").value = url.replace(/\\/$/,"") + "/webhook";'
        + '    setTimeout(function(){ document.getElementById("urlSuccess").classList.add("hidden"); }, 3000);'
        + '  }'
        + '});'
        + '</script>'
    ));
});

async function getGroupsFromPage(page) {
    return await page.evaluate(function() {
        var chats = window.require('WAWebCollections').Chat.getModelsArray();
        return chats.filter(function(c) { return c.groupMetadata; }).map(function(c) {
            return { id: c.id._serialized, name: c.formattedTitle || c.name || c.id._serialized };
        });
    });
}

app.get('/api/groups', async function (req, res) {
    if (clientStatus !== 'connected') return res.status(503).json({ error: 'WhatsApp non connecte' });
    try {
        console.log('Chargement des groupes...');
        var groups;
        try {
            groups = await getGroupsFromPage(client.pupPage);
        } catch (_) {
            await new Promise(function(r) { setTimeout(r, 1000); });
            groups = await getGroupsFromPage(client.pupPage);
        }
        console.log('Groupes trouves:', groups.length);
        if (groups.length > 0) {
            console.log('--- LISTE DES GROUPES ---');
            groups.forEach(function(g) { console.log('Nom: ' + g.name + ' | ID: ' + g.id); });
        }
        res.json(groups);
    } catch (err) {
        console.error('Erreur chargement groupes:', err.message);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/config', function (req, res) {
    var data = req.body;
    if (data.groupId !== undefined) config.groupId = data.groupId;
    if (data.sheetId !== undefined) config.sheetId = extractSheetId(data.sheetId);
    if (data.publicUrl !== undefined) config.publicUrl = data.publicUrl;
    if (data.groupId !== undefined || data.sheetId !== undefined || data.publicUrl !== undefined) saveConfig(config);
    if (data.sheetId !== undefined) startPolling();
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
        stopPolling();
        await client.destroy();
    } catch (_) { }
    qrCodeData = null;
    pairingCodeData = null;
    clientStatus = 'initializing';
    client = new Client({
        authStrategy: new LocalAuth(),
        puppeteer: {
            headless: true,
            protocolTimeout: 120000,
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
    res.json({ status: clientStatus, hasQr: !!qrCodeData, hasPairing: !!pairingCodeData, sheetConfigured: !!config.sheetId });
});

app.get('/debug', function (req, res) {
    res.json({ clientStatus, hasQr: !!qrCodeData, qrLen: qrCodeData ? qrCodeData.length : 0, hasPairing: !!pairingCodeData, sheetConfigured: !!config.sheetId, config, commit: process.env.RENDER_GIT_COMMIT || process.env.HF_SPACE || 'local' });
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

    client.on('ready', async function () {
        qrCodeData = null;
        pairingCodeData = null;
        clientStatus = 'connected';
        console.log('WhatsApp connecte !');
        try {
            var groups;
            try { groups = await getGroupsFromPage(client.pupPage); }
            catch (_) {
                await new Promise(function(r) { setTimeout(r, 1000); });
                groups = await getGroupsFromPage(client.pupPage);
            }
            if (groups.length > 0) {
                console.log('--- LISTE DES GROUPES ---');
                groups.forEach(function(g) { console.log('Nom: ' + g.name + ' | ID: ' + g.id); });
            }
        } catch (e) { console.log('Impossible de lister les groupes:', e.message); }
        startPolling();
    });

    client.on('disconnected', function (reason) {
        clientStatus = 'disconnected';
        stopPolling();
        console.log('WhatsApp deconnecte:', reason);
    });
}

var client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: {
        headless: true,
        protocolTimeout: 120000,
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

app.listen(PORT, '0.0.0.0', function () {
    console.log('STARTUP: listening on http://0.0.0.0:' + PORT);
    client.initialize().catch(function (err) {
        console.error('Erreur init WhatsApp:', err.message);
    });
});
