import Anthropic from '@anthropic-ai/sdk';
import { Client } from '@microsoft/microsoft-graph-client';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { Redis } from '@upstash/redis';

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── REDIS (stockage permanent des tokens) ──
const redis = new Redis({
  url: process.env.KV_REST_API_URL,
  token: process.env.KV_REST_API_TOKEN,
});

// ── MSAL OneDrive ──
const msalApp = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/consumers`,
  }
});

async function getOneDriveClient() {
  let tokenData = await redis.get('od_token');
  if (!tokenData) throw new Error('OneDrive non connecté — allez sur /api/auth');

  if (Date.now() > tokenData.expiresAt - 60000) {
    const result = await msalApp.acquireTokenByRefreshToken({
      refreshToken: tokenData.refreshToken,
      scopes: ['https://graph.microsoft.com/Files.ReadWrite', 'offline_access'],
    });
    tokenData = {
      accessToken: result.accessToken,
      refreshToken: result.refreshToken || tokenData.refreshToken,
      expiresAt: result.expiresOn.getTime()
    };
    await redis.set('od_token', tokenData);
  }

  return Client.initWithMiddleware({
    authProvider: { getAccessToken: async () => tokenData.accessToken }
  });
}

export default async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  const url = req.url.split('?')[0];

  // ── GET /api/config ──
  if (url === '/api/config' && req.method === 'GET') {
    const cfg = await redis.get('config') || {};
    return res.json({ societies: cfg.societies || ['STE SINDIS MARKET','SLIDIS MARKET','LIVRIZY','INDRE SHOP'] });
  }

  // ── POST /api/config ──
  if (url === '/api/config' && req.method === 'POST') {
    const cfg = await redis.get('config') || {};
    if (req.body.sal_pin) cfg.sal_pin = req.body.sal_pin;
    if (req.body.adm_pin) cfg.adm_pin = req.body.adm_pin;
    if (req.body.societies) cfg.societies = req.body.societies;
    await redis.set('config', cfg);
    return res.json({ ok: true });
  }

  // ── GET /api/auth ──
  if (url === '/api/auth' && req.method === 'GET') {
    const authUrl = await msalApp.getAuthCodeUrl({
      scopes: ['https://graph.microsoft.com/Files.ReadWrite', 'offline_access', 'User.Read'],
      redirectUri: process.env.REDIRECT_URI,
    });
    return res.redirect(authUrl);
  }

  // ── GET /api/auth/callback ──
  if (url === '/api/auth/callback' && req.method === 'GET') {
    try {
      const code = req.query.code;
      const result = await msalApp.acquireTokenByCode({
        code,
        scopes: ['https://graph.microsoft.com/Files.ReadWrite', 'offline_access', 'User.Read'],
        redirectUri: process.env.REDIRECT_URI,
      });
      const tokenData = {
        accessToken: result.accessToken,
        refreshToken: result.refreshToken,
        expiresAt: result.expiresOn.getTime()
      };
      await redis.set('od_token', tokenData);
      return res.send(`
        <html><body style="font-family:sans-serif;text-align:center;padding:60px;background:#0f0f11;color:#f0eff4">
          <div style="font-size:48px">✅</div>
          <h2>OneDrive connecté avec succès !</h2>
          <p style="color:#7a7a8a">Vous pouvez fermer cette page et retourner à l'application.</p>
          <a href="/" style="display:inline-block;margin-top:20px;padding:12px 24px;background:#6c63ff;color:white;border-radius:12px;text-decoration:none">Retour à l'app</a>
        </body></html>
      `);
    } catch(err) {
      return res.status(500).send(`<html><body style="font-family:sans-serif;text-align:center;padding:60px">
        <div style="font-size:48px">❌</div>
        <h2>Erreur de connexion</h2>
        <p>${err.message}</p>
        <a href="/api/auth">Réessayer</a>
      </body></html>`);
    }
  }

  // ── POST /api/analyser ──
  if (url === '/api/analyser' && req.method === 'POST') {
    try {
      const chunks = [];
      for await (const chunk of req) chunks.push(chunk);
      const buffer = Buffer.concat(chunks);
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
            { type: 'text', text: `Analyse cette facture ou ce ticket de caisse. Réponds UNIQUEMENT avec ce JSON, sans texte avant ou après :
{
  "montant": "montant TTC avec symbole € ex: 47,80 €",
  "date": "date au format JJ/MM/AAAA",
  "tva": "montant TVA avec € si visible sinon null",
  "fournisseur": "nom du fournisseur ou magasin",
  "description": "type achat en 3 mots max"
}` }
          ]
        }]
      });

      let result = {};
      try {
        const text = message.content[0].text.trim().replace(/```json|```/g, '').trim();
        result = JSON.parse(text);
      } catch(e) {
        result = { montant: 'Non lisible', date: 'Non lisible', tva: null, fournisseur: 'Inconnu', description: 'facture' };
      }

      return res.json(result);
    } catch(err) {
      return res.status(500).json({ error: err.message });
    }
  }

  // ── POST /api/onedrive ──
  if (url === '/api/onedrive' && req.method === 'POST') {
    try {
      const { societe, mode_paiement, montant, date, description, fournisseur } = req.body;
      const client = await getOneDriveClient();

      const now = new Date();
      const month = now.toISOString().slice(0, 7);
      const slug = societe.replace(/[^a-zA-Z0-9]/g, '_').toUpperCase().slice(0, 12);
      const modeTag = mode_paiement === 'especes' ? 'ESPECES' : 'CARTE';
      const dateTag = date ? date.replace(/\//g, '-') : now.toISOString().slice(0, 10);
      const montantTag = montant ? montant.replace(/[^0-9,\.]/g, '').replace(',', '.') : '0';
      const descTag = (description || 'facture').replace(/[^a-zA-Z0-9]/g, '_').slice(0, 20);
      const fileName = `${dateTag}_${modeTag}_${montantTag}EUR_${descTag}_meta.json`;
      const filePath = `/me/drive/root:/Factures/${slug}/${month}/${fileName}:/content`;

      const meta = {
        societe, mode_paiement, montant, date, tva: req.body.tva,
        fournisseur, description,
        uploaded_at: new Date().toISOString(),
        uploaded_by: 'Factures Pro App'
      };

      await client.api(filePath).put(Buffer.from(JSON.stringify(meta, null, 2)));

      return res.json({ ok: true, path: `/Factures/${slug}/${month}/${fileName}` });
    } catch(err) {
      console.error('OneDrive error:', err);
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(404).json({ error: 'Route inconnue' });
}
