import Anthropic from '@anthropic-ai/sdk';
import { Client } from '@microsoft/microsoft-graph-client';
import { ConfidentialClientApplication } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── CONFIG LOCALE (PIN + sociétés) ──
const CONFIG_PATH = '/tmp/config.json';
function loadConfig() {
  try {
    if (fs.existsSync(CONFIG_PATH)) return JSON.parse(fs.readFileSync(CONFIG_PATH,'utf8'));
  } catch(e) {}
  return { sal_pin: process.env.SAL_PIN||'1234', adm_pin: process.env.ADM_PIN||'5678', societies: ['STE SINDIS MARKET','SLIDIS MARKET','LIVRIZY','INDRE SHOP'] };
}
function saveConfig(cfg) {
  try { fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg)); } catch(e) {}
}

// ── MSAL OneDrive ──
const msalApp = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
  }
});

async function getOneDriveClient() {
  const TOKEN_PATH = '/tmp/od_token.json';
  let tokenData = null;
  try { tokenData = JSON.parse(fs.readFileSync(TOKEN_PATH,'utf8')); } catch(e) {}
  if (!tokenData) throw new Error('OneDrive non autorisé — connectez-vous d\'abord via /api/auth');

  // Refresh si expiré
  if (Date.now() > tokenData.expiresAt - 60000) {
    const result = await msalApp.acquireTokenByRefreshToken({
      refreshToken: tokenData.refreshToken,
      scopes: ['https://graph.microsoft.com/Files.ReadWrite','offline_access'],
    });
    tokenData = { accessToken: result.accessToken, refreshToken: result.refreshToken||tokenData.refreshToken, expiresAt: result.expiresOn.getTime() };
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokenData));
  }

  return Client.initWithMiddleware({
    authProvider: { getAccessToken: async () => tokenData.accessToken }
  });
}

export default async function handler(req, res) {
  const url = req.url.split('?')[0];

  // ── GET /api/config ──
  if (url === '/api/config' && req.method === 'GET') {
    const cfg = loadConfig();
    return res.json({ societies: cfg.societies }); // Ne pas exposer les PIN
  }

  // ── POST /api/config ──
  if (url === '/api/config' && req.method === 'POST') {
    const cfg = loadConfig();
    if (req.body.sal_pin) cfg.sal_pin = req.body.sal_pin;
    if (req.body.adm_pin) cfg.adm_pin = req.body.adm_pin;
    if (req.body.societies) cfg.societies = req.body.societies;
    saveConfig(cfg);
    return res.json({ ok: true });
  }

  // ── GET /api/auth — Lance l'auth OneDrive ──
  if (url === '/api/auth' && req.method === 'GET') {
    const authUrl = await msalApp.getAuthCodeUrl({
      scopes: ['https://graph.microsoft.com/Files.ReadWrite','offline_access','User.Read'],
      redirectUri: process.env.REDIRECT_URI,
    });
    return res.redirect(authUrl);
  }

  // ── GET /api/auth/callback ──
  if (url === '/api/auth/callback' && req.method === 'GET') {
    const code = req.query.code;
    const result = await msalApp.acquireTokenByCode({
      code, scopes: ['https://graph.microsoft.com/Files.ReadWrite','offline_access','User.Read'],
      redirectUri: process.env.REDIRECT_URI,
    });
    const tokenData = { accessToken: result.accessToken, refreshToken: result.refreshToken, expiresAt: result.expiresOn.getTime() };
    fs.writeFileSync('/tmp/od_token.json', JSON.stringify(tokenData));
    return res.send('<h2 style="font-family:sans-serif;text-align:center;margin-top:40px">✅ OneDrive connecté avec succès !<br><br><a href="/">Retour à l\'app</a></h2>');
  }

  // ── POST /api/analyser — Analyse IA de la facture ──
  if (url === '/api/analyser' && req.method === 'POST') {
    try {
      const chunks = [];
      for await (const chunk of req) chunks.push(chunk);
      const buffer = Buffer.concat(chunks);

      // Extraire l'image du multipart (simple)
      const boundary = req.headers['content-type'].split('boundary=')[1];
      const parts = buffer.toString('binary').split('--' + boundary);
      let imageBase64 = null, imageType = 'image/jpeg';

      for (const part of parts) {
        if (part.includes('name="image"')) {
          const headerEnd = part.indexOf('\r\n\r\n');
          if (headerEnd > -1) {
            const headers = part.substring(0, headerEnd);
            const match = headers.match(/Content-Type: ([^\r\n]+)/);
            if (match) imageType = match[1].trim();
            const rawData = part.substring(headerEnd + 4, part.lastIndexOf('\r\n'));
            imageBase64 = Buffer.from(rawData, 'binary').toString('base64');
          }
        }
      }

      if (!imageBase64) return res.status(400).json({ error: 'Image manquante' });

      const message = await anthropic.messages.create({
        model: 'claude-opus-4-5',
        max_tokens: 500,
        messages: [{
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: imageType, data: imageBase64 } },
            { type: 'text', text: `Analyse cette facture ou ce ticket de caisse et extrait les informations suivantes en JSON uniquement, sans texte autour :
{
  "montant": "montant TTC avec € ex: 47,80 €",
  "date": "date au format JJ/MM/AAAA",
  "tva": "montant TVA avec € si visible, sinon null",
  "fournisseur": "nom du fournisseur/magasin si lisible",
  "description": "courte description en 5 mots max"
}
Si une information n'est pas lisible, mets null.` }
          ]
        }]
      });

      let result = {};
      try {
        const text = message.content[0].text.trim();
        const clean = text.replace(/```json|```/g, '').trim();
        result = JSON.parse(clean);
      } catch(e) {
        result = { montant: null, date: null, tva: null, fournisseur: null, description: null };
      }

      return res.json(result);
    } catch(err) {
      console.error('Erreur analyse:', err);
      return res.status(500).json({ error: err.message });
    }
  }

  // ── POST /api/onedrive — Upload vers OneDrive ──
  if (url === '/api/onedrive' && req.method === 'POST') {
    try {
      const { societe, mode_paiement, montant, date, description, filename } = req.body;
      const client = await getOneDriveClient();

      const now = new Date();
      const month = now.toISOString().slice(0, 7);
      const slug = societe.replace(/\s+/g, '_').toUpperCase().slice(0, 12);
      const modeTag = mode_paiement === 'especes' ? 'ESPECES' : 'CARTE';
      const dateTag = date ? date.replace(/\//g,'-') : now.toISOString().slice(0,10);
      const montantTag = montant ? montant.replace(/[€\s]/g,'').replace(',','.') : 'INCONNU';
      const folderPath = `/Factures/${slug}/${month}`;
      const safeName = `${dateTag}_${modeTag}_${montantTag}EUR_${description||'facture'}.jpg`.replace(/[^a-zA-Z0-9_\-\.]/g,'_');
      const fullPath = `${folderPath}/${safeName}`;

      // Créer les dossiers (ignore les erreurs si déjà existants)
      try {
        await client.api('/me/drive/root:/Factures:/children').post({ name: slug, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' });
      } catch(e) {}
      try {
        await client.api(`/me/drive/root:/Factures/${slug}:/children`).post({ name: month, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' });
      } catch(e) {}

      // On crée un fichier JSON de métadonnées (la vraie image est gérée séparément)
      const meta = { societe, mode_paiement, montant, date, description, uploaded_at: new Date().toISOString() };
      const metaName = safeName.replace('.jpg', '_meta.json');
      await client.api(`/me/drive/root:/${fullPath.replace('.jpg','_meta.json')}:/content`)
        .put(Buffer.from(JSON.stringify(meta, null, 2)));

      return res.json({ ok: true, path: fullPath });
    } catch(err) {
      console.error('Erreur OneDrive:', err);
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(404).json({ error: 'Route inconnue' });
}
